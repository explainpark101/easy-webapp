import React, { useState, useRef, useEffect } from 'react';
import { Upload, FileText, Play, Download, Loader2, CheckCircle2, AlertCircle, Layout, Eraser, Settings, Move, RotateCcw } from 'lucide-react';

// 외부 라이브러리 로드 스크립트
const LIB_SCRIPTS = [
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
  'https://cdn.jsdelivr.net/gh/gitbrent/PptxGenJS@3.12.0/dist/pptxgen.bundle.js'
];

export default function App() {
  const [file, setFile] = useState(null);
  const [fileUrl, setFileUrl] = useState(null);
  const [status, setStatus] = useState('idle'); // idle, processing, completed, error
  const [progress, setProgress] = useState(0);
  const [logs, setLogs] = useState([]);
  const [error, setError] = useState(null);
  
  // 요청하신 정밀 디폴트 값 설정
  const DEFAULT_SETTINGS = {
    widthRatio: 9.2,   
    heightRatio: 3.1, 
    xRatio: 96.9,      
    yRatio: 98.4     
  };

  const [maskSettings, setMaskSettings] = useState({ ...DEFAULT_SETTINGS });

  useEffect(() => {
    const loadScripts = async () => {
      for (const src of LIB_SCRIPTS) {
        if (!document.querySelector(`script[src="${src}"]`)) {
          const script = document.createElement('script');
          script.src = src;
          script.async = true;
          document.head.appendChild(script);
        }
      }
    };
    loadScripts();
  }, []);

  const addLog = (message) => {
    setLogs(prev => [...prev, { time: new Date().toLocaleTimeString(), message }]);
  };

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile && selectedFile.type === 'application/pdf') {
      const url = URL.createObjectURL(selectedFile);
      setFile(selectedFile);
      setFileUrl(url);
      setStatus('idle');
      setError(null);
      setLogs([]);
      addLog(`파일 로드됨: ${selectedFile.name}`);
    } else {
      setError('PDF 파일만 업로드 가능합니다.');
    }
  };

  const resetToDefault = () => {
    setMaskSettings({ ...DEFAULT_SETTINGS });
    addLog('마스킹 설정이 기본값으로 초기화되었습니다.');
  };

  const handleSliderChange = (key, value) => {
    setMaskSettings(prev => ({ ...prev, [key]: parseFloat(value) }));
  };

  const processPdf = async () => {
    if (!file) return;

    try {
      setStatus('processing');
      setProgress(0);
      addLog('프로세스 시작...');

      const pdfjsLib = window['pdfjs-dist/build/pdf'];
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const numPages = pdf.numPages;
      addLog(`총 ${numPages}페이지 분석 및 마스킹 수행 중...`);

      const pptx = new window.PptxGenJS();
      pptx.layout = 'LAYOUT_16x9';

      for (let i = 1; i <= numPages; i++) {
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 2.5 });
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({ canvasContext: context, viewport }).promise;

        const m = maskSettings;
        const w = (canvas.width * m.widthRatio) / 100;
        const h = (canvas.height * m.heightRatio) / 100;
        const x = (canvas.width * m.xRatio) / 100 - (w / 2);
        const y = (canvas.height * m.yRatio) / 100 - (h / 2);

        // 색상 샘플링 및 마스킹
        const sampleX = Math.max(0, x - 2);
        const sampleY = Math.max(0, y - 2);
        const pixelData = context.getImageData(sampleX, sampleY, 1, 1).data;
        const bgColor = `rgb(${pixelData[0]}, ${pixelData[1]}, ${pixelData[2]})`;

        context.fillStyle = bgColor;
        context.fillRect(x, y, w, h);

        const imageData = canvas.toDataURL('image/png');
        const slide = pptx.addSlide();
        slide.addImage({ data: imageData, x: 0, y: 0, w: '100%', h: '100%' });

        const currentProgress = Math.round((i / numPages) * 100);
        setProgress(currentProgress);
        if (i % 5 === 0 || i === numPages) addLog(`${i}페이지 처리 중 (${currentProgress}%)...`);
      }

      addLog('PPTX 파일 저장 중...');
      await pptx.writeFile({ fileName: `Cleaned_${file.name.replace('.pdf', '')}.pptx` });
      
      setStatus('completed');
      addLog('변환이 성공적으로 완료되었습니다.');

    } catch (err) {
      console.error(err);
      setError('오류 발생: ' + err.message);
      setStatus('error');
    }
  };

  const handleResetApp = () => {
    setFile(null);
    setFileUrl(null);
    setStatus('idle');
    setProgress(0);
    setLogs([]);
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 md:p-8 font-sans text-slate-900">
      <div className="max-w-5xl mx-auto">
        <header className="mb-8 text-center">
          <div className="inline-flex items-center justify-center p-3 bg-blue-600 rounded-2xl mb-4 text-white shadow-lg">
            <Layout size={32} />
          </div>
          <h1 className="text-3xl font-bold text-slate-800 tracking-tight">PDF to Image PPT</h1>
          <p className="text-slate-500 mt-2">NotebookLM 로고를 완벽하게 제거하여 깔끔한 슬라이드로 변환합니다.</p>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
          {/* Settings Panel */}
          <div className="lg:col-span-4 space-y-6">
            <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
              <div className="flex items-center justify-between mb-6 border-b pb-3">
                <div className="flex items-center text-blue-600 font-bold">
                  <Settings size={20} className="mr-2" />
                  마스킹 정밀 설정
                </div>
                <button 
                  onClick={resetToDefault}
                  className="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-600 px-2 py-1 rounded flex items-center transition-colors"
                >
                  <RotateCcw size={12} className="mr-1" />
                  DEFAULT
                </button>
              </div>
              
              <div className="space-y-6">
                <div className="space-y-2">
                  <div className="flex justify-between text-xs font-bold text-slate-600 uppercase">
                    <span>가로 위치 (X)</span>
                    <span className="text-blue-600">{maskSettings.xRatio}%</span>
                  </div>
                  <input type="range" min="0" max="100" step="0.1" value={maskSettings.xRatio} 
                    onChange={(e) => handleSliderChange('xRatio', e.target.value)}
                    className="w-full h-2 bg-slate-100 rounded-lg appearance-none cursor-pointer accent-blue-600" />
                </div>

                <div className="space-y-2">
                  <div className="flex justify-between text-xs font-bold text-slate-600 uppercase">
                    <span>세로 위치 (Y)</span>
                    <span className="text-blue-600">{maskSettings.yRatio}%</span>
                  </div>
                  <input type="range" min="0" max="100" step="0.1" value={maskSettings.yRatio} 
                    onChange={(e) => handleSliderChange('yRatio', e.target.value)}
                    className="w-full h-2 bg-slate-100 rounded-lg appearance-none cursor-pointer accent-blue-600" />
                </div>

                <div className="pt-2 grid grid-cols-2 gap-4 border-t border-slate-50 mt-4">
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">가로 사이즈</label>
                    <input type="number" step="0.1" value={maskSettings.widthRatio} 
                      onChange={(e) => handleSliderChange('widthRatio', e.target.value)}
                      className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm font-mono focus:ring-2 focus:ring-blue-500 outline-none" />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">세로 사이즈</label>
                    <input type="number" step="0.1" value={maskSettings.heightRatio} 
                      onChange={(e) => handleSliderChange('heightRatio', e.target.value)}
                      className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm font-mono focus:ring-2 focus:ring-blue-500 outline-none" />
                  </div>
                </div>
              </div>
            </div>
            
            <div className="bg-slate-900 rounded-3xl p-6 text-slate-400 shadow-xl overflow-hidden">
              <div className="flex justify-between items-center mb-4 border-b border-slate-800 pb-3">
                <span className="font-mono text-[10px] uppercase tracking-widest font-bold">Process Logs</span>
              </div>
              <div className="overflow-y-auto space-y-2 font-mono text-[10px] h-48 scrollbar-hide">
                {logs.length === 0 && <p className="text-slate-600">준비 상태...</p>}
                {logs.map((log, idx) => (
                  <div key={idx} className="flex gap-2">
                    <span className="text-blue-500 font-bold shrink-0">{log.time}</span>
                    <span className="text-slate-300 break-all">{log.message}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          {/* Main Content Area */}
          <div className="lg:col-span-8 space-y-6">
            <div className="bg-white p-8 rounded-3xl shadow-sm border border-slate-200 min-h-[450px] flex flex-col justify-center">
              {!file ? (
                <div 
                  className="border-2 border-dashed border-slate-200 rounded-3xl p-16 text-center hover:border-blue-400 hover:bg-blue-50 transition-all cursor-pointer group"
                  onClick={() => document.getElementById('fileInput').click()}
                >
                  <div className="w-20 h-20 bg-blue-50 rounded-full flex items-center justify-center mx-auto mb-6 group-hover:scale-105 transition-transform">
                    <Upload className="text-blue-500" size={32} />
                  </div>
                  <p className="text-xl font-semibold text-slate-700">PDF 파일을 업로드하여 시작하세요</p>
                  <input id="fileInput" type="file" className="hidden" accept="application/pdf" onChange={handleFileChange} />
                </div>
              ) : (
                <div className="space-y-8 animate-in fade-in duration-500">
                  <div className="flex items-center p-5 bg-blue-50/50 rounded-2xl border border-blue-100 shadow-sm">
                    <div className="p-3 bg-white rounded-xl shadow-sm mr-4">
                      <FileText className="text-blue-600" size={24} />
                    </div>
                    <div className="flex-1 overflow-hidden text-left">
                      <p className="font-bold text-slate-800 truncate text-lg">{file.name}</p>
                      <p className="text-sm text-slate-500">{(file.size / (1024 * 1024)).toFixed(2)} MB • {status === 'completed' ? '변환 완료' : '처리 준비됨'}</p>
                    </div>
                    {status !== 'processing' && (
                      <button onClick={handleResetApp} className="text-slate-400 hover:text-red-500 p-2 transition-colors">취소</button>
                    )}
                  </div>

                  {status === 'processing' ? (
                    <div className="space-y-6 px-4">
                      <div className="flex justify-between items-end">
                        <div className="text-left">
                          <p className="text-sm font-bold text-slate-700">고해상도 캡쳐 및 마스킹 중</p>
                          <p className="text-xs text-slate-400">잠시만 기다려주세요...</p>
                        </div>
                        <span className="text-3xl font-black text-blue-600">{progress}%</span>
                      </div>
                      <div className="w-full bg-slate-100 rounded-full h-4 overflow-hidden shadow-inner border border-slate-200">
                        <div 
                          className="bg-gradient-to-r from-blue-500 to-indigo-600 h-full transition-all duration-300"
                          style={{ width: `${progress}%` }}
                        ></div>
                      </div>
                    </div>
                  ) : status === 'completed' ? (
                    <div className="p-8 bg-emerald-50 border border-emerald-100 rounded-3xl text-center space-y-6 animate-in zoom-in duration-300">
                      <div className="w-16 h-16 bg-white rounded-full flex items-center justify-center mx-auto shadow-sm">
                        <CheckCircle2 className="text-emerald-500" size={32} />
                      </div>
                      <div>
                        <h3 className="text-2xl font-black text-emerald-900">변환 작업 완료!</h3>
                        <p className="text-emerald-700 mt-1">파일이 다운로드되었습니다.</p>
                      </div>
                      
                      <div className="flex flex-col sm:flex-row gap-4 justify-center pt-2">
                        <button 
                          onClick={processPdf}
                          className="flex-1 bg-white text-emerald-700 border border-emerald-200 px-6 py-3 rounded-xl font-bold hover:bg-emerald-100 transition-all flex items-center justify-center"
                        >
                          <RotateCcw className="mr-2" size={18} />
                          현재파일 다시하기
                        </button>
                        <button 
                          onClick={handleResetApp}
                          className="flex-1 bg-emerald-600 text-white px-6 py-3 rounded-xl font-bold hover:bg-emerald-700 shadow-lg shadow-emerald-200 transition-all"
                        >
                          새파일로 계속하기
                        </button>
                      </div>
                    </div>
                  ) : status === 'error' ? (
                    <div className="p-6 bg-red-50 border border-red-100 rounded-2xl flex items-start text-red-700">
                      <AlertCircle className="mr-3 shrink-0" size={24} />
                      <div className="text-left">
                        <p className="font-bold">오류가 발생했습니다</p>
                        <p className="text-sm opacity-80">{error}</p>
                      </div>
                    </div>
                  ) : (
                    <button 
                      onClick={processPdf}
                      className="w-full bg-blue-600 text-white py-5 rounded-2xl font-black text-xl hover:bg-blue-700 shadow-xl shadow-blue-200 transition-all flex items-center justify-center group"
                    >
                      <Play className="mr-3 group-hover:scale-110 transition-transform" size={24} fill="currentColor" />
                      PPT 변환 시작
                    </button>
                  )}

                  {/* Uploaded PDF Download Link - Always visible when file exists */}
                  <div className="pt-4 border-t border-slate-100">
                    <a 
                      href={fileUrl} 
                      download={file.name}
                      className="inline-flex items-center text-xs font-bold text-slate-400 hover:text-blue-500 transition-colors"
                    >
                      <Download size={14} className="mr-1" />
                      원본 PDF 다운로드 ({file.name})
                    </a>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}