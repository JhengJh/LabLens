import React, { useState, useRef, useEffect, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import { GoogleGenAI, Type } from "@google/genai";
import { 
  Camera, Upload, RefreshCw, X, Loader2, ArrowLeft, 
  FlaskConical, TrendingUp, Calculator, MousePointer2, 
  TestTube2, Plus, Trash2, CheckCircle2, MoreHorizontal, 
  ChevronRight, Download, FileSpreadsheet, Sun, Moon,
  LayoutGrid, Settings2, Grip, Keyboard, Eraser, Languages,
  AlertTriangle, CheckCircle, XCircle,
  Key, Save, Eye, EyeOff, Lock, Settings, BarChart3, Activity, Beaker,
  Sparkles, Zap
} from 'lucide-react';
import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";

// --- Translations ---

const TRANSLATIONS = {
  en: {
    appTitle: "LabLens",
    subtitle: "Microplate Assistant",
    desc: "AI-powered analysis for your 96-well assays.",
    scanPlate: "Scan",
    useCamera: "Camera",
    uploadImage: "Upload",
    fromGallery: "Gallery",
    manualEntry: "Manual",
    noImage: "No Image",
    analyzing: "Analyzing Grid Structure...",
    discard: "Discard",
    analyze: "Analyze Plate",
    plateMap: "Plate Map Explorer",
    calibration: "Regression Analysis",
    samples: "Sample Management",
    selectRange: "Select Wells",
    selecting: "Selecting...",
    noDataSelected: "No data points selected",
    conc: "Conc",
    od: "OD",
    addRow: "Add",
    unknowns: "Unknown Samples",
    newGroup: "New Group",
    noGroups: "No sample groups created.",
    sampleName: "Sample Name",
    mean: "Mean",
    sd: "SD",
    cv: "CV%",
    setAllDil: "Batch Dilution",
    dil: "Dil",
    addManualWell: "+ Add Manual Well",
    export: "Report",
    slope: "Slope",
    intercept: "Intercept",
    r2: "R² Score",
    cameraError: "Camera access denied or unavailable.",
    return: "Return",
    concAxis: "Concentration",
    odAxis: "OD Value",
    reportTitle: "LabLens Analysis Report",
    rawMatrix: "RAW DATA MATRIX",
    stdCurve: "STANDARD CURVE",
    sampleTable: "SAMPLES",
    fitGood: "Excellent Fit",
    fitWarn: "Acceptable Fit",
    fitPoor: "Poor Fit",
    settings: "Settings",
    analysisFailed: "Analysis Failed",
    summaryR2: "R2",
    summarySamples: "Samples",
    summaryGroups: "Groups",
    activeGroup: "Active Group",
    well: "Well",
    presets: "Presets",
    selectPreset: "Select...",
    custom: "Custom",
    modelFlash: "Gemini Flash (Fast)",
    modelPro: "Gemini 3 Pro (Smart)",
    apiKeyLabel: "API Key",
    apiKeyPlaceholder: "Enter your Gemini API Key",
    save: "Save",
    saved: "Saved",
    missingKey: "API Key is missing. Please add it in Settings.",
    toggleView: "Toggle View"
  },
  zh: {
    appTitle: "LabLens",
    subtitle: "微孔板助手",
    desc: "AI 驱动的 96 孔板数据分析工具",
    scanPlate: "扫描",
    useCamera: "相机",
    uploadImage: "上传",
    fromGallery: "相册",
    manualEntry: "手动",
    noImage: "填报",
    analyzing: "正在分析网格结构...",
    discard: "放弃",
    analyze: "分析孔板",
    plateMap: "孔板浏览器",
    calibration: "回归分析",
    samples: "样本管理",
    selectRange: "选择孔位",
    selecting: "正在选择...",
    noDataSelected: "未选择数据点",
    conc: "浓度",
    od: "OD值",
    addRow: "添加",
    unknowns: "未知样本",
    newGroup: "新建组",
    noGroups: "暂无样本组",
    sampleName: "样本名称",
    mean: "平均值",
    sd: "标准差",
    cv: "变异系数",
    setAllDil: "统一稀释",
    dil: "稀释",
    addManualWell: "+ 手动添加孔位",
    export: "导出报告",
    slope: "斜率",
    intercept: "截距",
    r2: "R² 评分",
    cameraError: "无法访问相机",
    return: "返回",
    concAxis: "浓度",
    odAxis: "OD值",
    reportTitle: "LabLens 分析报告",
    rawMatrix: "原始数据矩阵",
    stdCurve: "标准曲线",
    sampleTable: "样本数据",
    fitGood: "拟合极佳",
    fitWarn: "拟合尚可",
    fitPoor: "拟合较差",
    settings: "设置",
    analysisFailed: "分析失败",
    summaryR2: "R2",
    summarySamples: "样本数",
    summaryGroups: "分组数",
    activeGroup: "当前组",
    well: "孔号",
    presets: "预设方案",
    selectPreset: "选择...",
    custom: "自定义",
    modelFlash: "Gemini Flash (快速)",
    modelPro: "Gemini 3 Pro (智能)",
    apiKeyLabel: "API Key",
    apiKeyPlaceholder: "输入您的 Gemini API Key",
    save: "保存",
    saved: "已保存",
    missingKey: "缺少 API Key，请在设置中填写。",
    toggleView: "切换视图"
  }
};

// --- Types ---

interface ExtractedItem {
  value: string;
  row: string;
  col: number;
}

interface StdCurvePoint {
  x: string;
  y: string;
  id: string;
  sourceWellId?: string;
}

interface FitResult {
  slope: number;
  intercept: number;
  r2: number;
}

interface SelectedWell {
  row: string;
  col: number;
  od: number;
  dilution: number;
  id?: string;
}

interface UnknownGroup {
  id: string;
  name: string;
  commonDilution: number;
  color: string;
  wells: SelectedWell[];
}

// --- Constants ---

const ROW_HEADERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
const COL_COUNT = 12;

const PRESETS: Record<string, string[]> = {
  'BCA': ['0', '0.025', '0.05', '0.1', '0.2', '0.3', '0.4', '0.5'],
  'Bradford': ['0', '0.0625', '0.125', '0.25', '0.5', '0.75', '1', '1.5']
};

// Modified: Default standard curve points now fixed to 10 entries
const DEFAULT_CONCS = [...PRESETS['BCA'], '', '']; 

const GROUP_COLORS = [
  { name: 'emerald', bg: 'bg-emerald-500', border: 'border-emerald-500', text: 'text-emerald-700 dark:text-emerald-400', ring: 'ring-emerald-500', softBg: 'bg-emerald-50 dark:bg-emerald-900/20' },
  { name: 'violet', bg: 'bg-violet-500', border: 'border-violet-500', text: 'text-violet-700 dark:text-violet-400', ring: 'ring-violet-500', softBg: 'bg-violet-50 dark:bg-violet-900/20' },
  { name: 'pink', bg: 'bg-pink-500', border: 'border-pink-500', text: 'text-pink-700 dark:text-pink-400', ring: 'ring-pink-500', softBg: 'bg-pink-50 dark:bg-pink-900/20' },
  { name: 'orange', bg: 'bg-orange-500', border: 'border-orange-500', text: 'text-orange-700 dark:text-orange-400', ring: 'ring-orange-500', softBg: 'bg-orange-50 dark:bg-orange-900/20' },
  { name: 'cyan', bg: 'bg-cyan-500', border: 'border-cyan-500', text: 'text-cyan-700 dark:text-cyan-400', ring: 'ring-cyan-500', softBg: 'bg-cyan-50 dark:bg-cyan-900/20' },
];

// --- Helper Functions ---

const wait = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const compressImage = (dataUrl: string): Promise<string> => {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let width = img.width;
      let height = img.height;
      const MAX_DIMENSION = 1024;
      if (width > MAX_DIMENSION || height > MAX_DIMENSION) {
        if (width > height) { height = Math.round((height / width) * MAX_DIMENSION); width = MAX_DIMENSION; }
        else { width = Math.round((width / height) * MAX_DIMENSION); height = MAX_DIMENSION; }
      }
      canvas.width = width; canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (!ctx) { resolve(dataUrl); return; }
      ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, width, height); ctx.drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL('image/jpeg', 0.8));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
};

function calculateLinearFit(points: StdCurvePoint[]): FitResult | null {
  const validPoints = points.map(p => ({ x: parseFloat(p.x), y: parseFloat(p.y) })).filter(p => !isNaN(p.x) && !isNaN(p.y));
  const n = validPoints.length;
  if (n < 2) return null;
  const sumX = validPoints.reduce((sum, p) => sum + p.x, 0);
  const sumY = validPoints.reduce((sum, p) => sum + p.y, 0);
  const sumXY = validPoints.reduce((sum, p) => sum + p.x * p.y, 0);
  const sumXX = validPoints.reduce((sum, p) => sum + p.x * p.x, 0);
  const denom = n * sumXX - sumX * sumX;
  if (denom === 0) return null;
  const slope = (n * sumXY - sumX * sumY) / denom;
  const intercept = (sumY - slope * sumX) / n;
  const yMean = sumY / n;
  let ssRes = 0, ssTot = 0;
  for (const p of validPoints) {
    const yPred = slope * p.x + intercept;
    ssRes += Math.pow(p.y - yPred, 2);
    ssTot += Math.pow(p.y - yMean, 2);
  }
  return { slope, intercept, r2: ssTot === 0 ? 1 : 1 - (ssRes / ssTot) };
}

function calculateConc(od: number, fit: FitResult | null, dilution: number) {
  if (!fit) return 0;
  return ((od - fit.intercept) / fit.slope) * dilution;
}

function calculateStats(values: number[]) {
  if (values.length === 0) return { mean: 0, sd: 0, cv: 0 };
  const mean = values.reduce((a, b) => a + b, 0) / values.length;
  if (values.length === 1) return { mean, sd: 0, cv: 0 };
  const variance = values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (values.length - 1);
  const sd = Math.sqrt(variance);
  return { mean, sd, cv: mean === 0 ? 0 : (sd / mean) * 100 };
}

function useMediaQuery(query: string) {
  const [matches, setMatches] = useState(() => window.matchMedia(query).matches);
  useEffect(() => {
    const media = window.matchMedia(query);
    const listener = (e: MediaQueryListEvent) => setMatches(e.matches);
    media.addEventListener('change', listener);
    return () => media.removeEventListener('change', listener);
  }, [query]);
  return matches;
}

function getFitQuality(r2: number) {
  if (r2 >= 0.98) return 'good';
  if (r2 >= 0.90) return 'warn';
  return 'poor';
}

// --- Components ---

function CameraView({ onCapture, onBack, t }: { onCapture: (img: string) => void, onBack: () => void, t: any }) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [error, setError] = useState<string | null>(null);
  useEffect(() => {
    let stream: MediaStream | null = null;
    const startCamera = async () => {
      try { stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } }); if (videoRef.current) { videoRef.current.srcObject = stream; } }
      catch (err) { console.error("Error accessing camera:", err); setError(t.cameraError); }
    };
    startCamera();
    return () => { if (stream) { stream.getTracks().forEach(track => track.stop()); } };
  }, [t]);
  const handleCapture = () => {
    if (videoRef.current && canvasRef.current) {
      const video = videoRef.current; const canvas = canvasRef.current; canvas.width = video.videoWidth; canvas.height = video.videoHeight;
      const ctx = canvas.getContext('2d'); if (ctx) { ctx.drawImage(video, 0, 0, canvas.width, canvas.height); onCapture(canvas.toDataURL('image/jpeg')); }
    }
  };
  return (
    <div className="relative w-full h-full bg-black flex flex-col items-center justify-center overflow-hidden">
      {error ? (
        <div className="text-white text-center p-4"><p className="text-red-400 mb-4">{error}</p><button onClick={onBack} className="px-6 py-2 bg-white text-black rounded-full font-medium">{t.return}</button></div>
      ) : (
        <>
          <video ref={videoRef} autoPlay playsInline className="w-full h-full object-cover"/>
          <canvas ref={canvasRef} className="hidden" />
          <div className="absolute top-0 left-0 p-6 z-10"><button onClick={onBack} className="text-white hover:opacity-70 transition-opacity"><X size={32} /></button></div>
          <div className="absolute bottom-0 left-0 right-0 p-10 flex justify-center items-center bg-gradient-to-t from-black/90 to-transparent"><button onClick={handleCapture} className="w-20 h-20 rounded-full border-[6px] border-white/30 bg-white flex items-center justify-center hover:scale-105 active:scale-95 transition-all shadow-2xl"/></div>
        </>
      )}
    </div>
  );
}

function StandardCurveChart({ points, slope, intercept, t }: { points: StdCurvePoint[], slope?: number, intercept?: number, t: any }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [dimensions, setDimensions] = useState({ width: 0, height: 0 });
  useEffect(() => {
    if (!containerRef.current) return;
    const resizeObserver = new ResizeObserver((entries) => { if (entries.length) setDimensions({ width: entries[0].contentRect.width, height: entries[0].contentRect.height }); });
    resizeObserver.observe(containerRef.current);
    return () => resizeObserver.disconnect();
  }, []);
  const data = useMemo(() => points.map(p => ({ x: parseFloat(p.x), y: parseFloat(p.y) })).filter(p => !isNaN(p.x) && !isNaN(p.y)), [points]);
  const calculateAxisRange = (min: number, max: number) => {
    let range = max - min; if (range <= 1e-9) { if (min === 0) { max = 1; } else { min *= 0.9; max *= 1.1; } range = max - min; }
    const paddingVal = range * 0.1; let niceMin = min - paddingVal, niceMax = max + paddingVal;
    const roughStep = (niceMax - niceMin) / 5; const exponent = Math.floor(Math.log10(roughStep)); const fraction = roughStep / Math.pow(10, exponent);
    let niceFraction = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
    const step = niceFraction * Math.pow(10, exponent); const startTick = Math.floor(niceMin / step) * step, endTick = Math.ceil(niceMax / step) * step;
    const ticks = []; for (let t = startTick; t <= endTick + (step * 0.001); t += step) ticks.push(t);
    return { min: startTick, max: endTick, ticks };
  };
  if (dimensions.width === 0) return <div ref={containerRef} className="w-full h-full" />;
  const xMinData = data.length ? Math.min(...data.map(p => p.x)) : 0, xMaxData = data.length ? Math.max(...data.map(p => p.x)) : 1;
  const xScaleInfo = calculateAxisRange(xMinData, xMaxData);
  let yMinData = data.length ? Math.min(...data.map(p => p.y)) : 0, yMaxData = data.length ? Math.max(...data.map(p => p.y)) : 1;
  if (typeof slope === 'number' && typeof intercept === 'number') { yMinData = Math.min(yMinData, slope * xScaleInfo.min + intercept, slope * xScaleInfo.max + intercept); yMaxData = Math.max(yMaxData, slope * xScaleInfo.min + intercept, slope * xScaleInfo.max + intercept); }
  const yScaleInfo = calculateAxisRange(yMinData, yMaxData);
  const padding = { top: 20, right: 30, bottom: 40, left: 50 };
  const scaleX = (val: number) => padding.left + ((val - xScaleInfo.min) / (xScaleInfo.max - xScaleInfo.min)) * (dimensions.width - padding.left - padding.right);
  const scaleY = (val: number) => dimensions.height - padding.bottom - ((val - yScaleInfo.min) / (yScaleInfo.max - yScaleInfo.min)) * (dimensions.height - padding.top - padding.bottom);
  return (
    <div ref={containerRef} className="w-full h-full flex items-center justify-center font-mono text-[10px] select-none">
      <svg width={dimensions.width} height={dimensions.height} className="overflow-visible">
        {yScaleInfo.ticks.map(tick => (<g key={`y-${tick}`}><line x1={padding.left} y1={scaleY(tick)} x2={dimensions.width - padding.right} y2={scaleY(tick)} stroke="currentColor" className="text-gray-100 dark:text-zinc-800" strokeDasharray="4"/><text x={padding.left - 8} y={scaleY(tick) + 3} textAnchor="end" className="fill-gray-400">{tick.toFixed(2).replace(/\.00$/, '')}</text></g>))}
        {xScaleInfo.ticks.map(tick => (<g key={`x-${tick}`}><line x1={scaleX(tick)} y1={padding.top} x2={scaleX(tick)} y2={dimensions.height - padding.bottom} stroke="currentColor" className="text-gray-100 dark:text-zinc-800" strokeDasharray="4"/><text x={scaleX(tick)} y={dimensions.height - padding.bottom + 15} textAnchor="middle" className="fill-gray-400">{tick.toFixed(2).replace(/\.00$/, '')}</text></g>))}
        <line x1={padding.left} y1={padding.top} x2={padding.left} y2={dimensions.height - padding.bottom} stroke="currentColor" className="text-gray-200 dark:text-zinc-700"/><line x1={padding.left} y1={dimensions.height - padding.bottom} x2={dimensions.width - padding.right} y2={dimensions.height - padding.bottom} stroke="currentColor" className="text-gray-200 dark:text-zinc-700"/>
        {typeof slope === 'number' && typeof intercept === 'number' && (<line x1={scaleX(xScaleInfo.min)} y1={scaleY(slope * xScaleInfo.min + intercept)} x2={scaleX(xScaleInfo.max)} y2={scaleY(slope * xScaleInfo.max + intercept)} stroke="currentColor" className="text-indigo-500" strokeWidth="2.5" strokeLinecap="round"/>)}
        {data.map((p, i) => (<circle key={i} cx={scaleX(p.x)} cy={scaleY(p.y)} r="5" className="fill-white dark:fill-zinc-900 stroke-indigo-600" strokeWidth="2"/>))}
      </svg>
    </div>
  );
}

function SettingsModal({ isOpen, onClose, t, apiKey, setApiKey }: { isOpen: boolean, onClose: () => void, t: any, apiKey: string, setApiKey: (k: string) => void }) {
  const [showKey, setShowKey] = useState(false);
  const [localKey, setLocalKey] = useState(apiKey);
  useEffect(() => { setLocalKey(apiKey); }, [apiKey]);
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 z-[60] bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white dark:bg-zinc-900 w-full max-w-md rounded-2xl shadow-2xl border border-zinc-200 dark:border-zinc-800 overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        <div className="p-4 border-b border-zinc-100 dark:border-zinc-800 flex justify-between items-center bg-zinc-50/50 dark:bg-zinc-900">
          <h3 className="font-bold flex items-center gap-2 text-zinc-800 dark:text-zinc-100"><Settings size={18}/> {t.settings}</h3>
          <button onClick={onClose} className="p-1 rounded-full hover:bg-zinc-200 dark:hover:bg-zinc-800 transition-colors"><X size={18} className="text-zinc-500"/></button>
        </div>
        <div className="p-6 space-y-4">
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase tracking-wider">{t.apiKeyLabel}</label>
            <div className="relative">
              <Key size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400"/>
              <input type={showKey ? "text" : "password"} value={localKey} onChange={(e) => setLocalKey(e.target.value)} placeholder={t.apiKeyPlaceholder} className="w-full pl-10 pr-10 py-2.5 bg-zinc-50 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-xl text-sm font-mono outline-none focus:ring-2 focus:ring-indigo-500 transition-all text-zinc-800 dark:text-zinc-200"/>
              <button onClick={() => setShowKey(!showKey)} className="absolute right-3 top-1/2 -translate-y-1/2 text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200">{showKey ? <EyeOff size={16}/> : <Eye size={16}/>}</button>
            </div>
          </div>
        </div>
        <div className="p-4 border-t border-zinc-100 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-900 flex justify-end">
          <button onClick={() => { setApiKey(localKey); onClose(); }} className="flex items-center gap-2 px-6 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold text-sm shadow-lg shadow-indigo-500/20 active:scale-95 transition-all"><Save size={16}/> {t.save}</button>
        </div>
      </div>
    </div>
  );
}

// --- Main App ---

function App() {
  const [theme, setTheme] = useState<'dark' | 'light'>('dark');
  const [lang, setLang] = useState<'en' | 'zh'>('zh');
  const [view, setView] = useState<'home' | 'camera' | 'preview' | 'results'>('home');
  const [image, setImage] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [results, setResults] = useState<ExtractedItem[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [stdCurvePoints, setStdCurvePoints] = useState<StdCurvePoint[]>(DEFAULT_CONCS.map((conc, i) => ({ x: conc, y: '', id: `pt-${i}` })));
  const [unknownGroups, setUnknownGroups] = useState<UnknownGroup[]>([]);
  const [editingGroupId, setEditingGroupId] = useState<string | null>(null);
  const [selectionTarget, setSelectionTarget] = useState<'std' | 'sample'>('std');
  const [isStdSelecting, setIsStdSelecting] = useState(false);
  const [analysisTab, setAnalysisTab] = useState<'calibration' | 'samples'>('calibration');
  const [selectedModel, setSelectedModel] = useState<string>('gemini-flash-latest');
  const [customApiKey, setCustomApiKey] = useState(() => localStorage.getItem('lablens_api_key') || '');
  const [showSettings, setShowSettings] = useState(false);
  
  const isDesktop = useMediaQuery('(min-width: 768px)');
  const isMobilePortrait = useMediaQuery('(max-width: 767px) and (orientation: portrait)');
  const useStandardLayout = !isMobilePortrait;

  const t = TRANSLATIONS[lang];

  useEffect(() => { theme === 'dark' ? document.documentElement.classList.add('dark') : document.documentElement.classList.remove('dark'); }, [theme]);

  const handleSetApiKey = (key: string) => { setCustomApiKey(key); localStorage.setItem('lablens_api_key', key); };

  const processImage = async () => {
    if (!image) return;
    const apiKeyToUse = customApiKey || process.env.API_KEY;
    if (!apiKeyToUse) { setError(t.missingKey); setShowSettings(true); return; }
    setLoading(true); setError(null);
    try {
      const compressedImage = await compressImage(image);
      const ai = new GoogleGenAI({ apiKey: apiKeyToUse });
      const match = compressedImage.match(/^data:(.+);base64,(.+)$/);
      const mimeType = match ? match[1] : 'image/jpeg', base64Data = match ? match[2] : compressedImage.split(',')[1];
      const response = await ai.models.generateContent({
        model: selectedModel,
        contents: { parts: [{ inlineData: { mimeType, data: base64Data } }, { text: "Extract measurement values from 96-well microplate. Map to A-H, 1-12. Output JSON [{value, row, col}]." }] },
        config: { temperature: 0, responseMimeType: "application/json", responseSchema: { type: Type.ARRAY, items: { type: Type.OBJECT, properties: { value: { type: Type.STRING }, row: { type: Type.STRING }, col: { type: Type.INTEGER } }, required: ["value", "row", "col"] } } }
      });
      const data = JSON.parse(response.text || "[]");
      setResults(data.map((item: any) => ({ value: item.value, row: item.row?.toUpperCase() || '?', col: parseInt(item.col) || 0 })));
      setView('results');
    } catch (err: any) { console.error(err); setError(err.message || "Extraction Failed."); }
    finally { setLoading(false); }
  };

  const handleImageSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.[0]) { const reader = new FileReader(); reader.onload = (ev) => { if (ev.target?.result) { setImage(ev.target.result as string); setView('preview'); } }; reader.readAsDataURL(e.target.files[0]); }
  };

  const reset = () => { setImage(null); setResults([]); setError(null); setView('home'); setStdCurvePoints(DEFAULT_CONCS.map((conc, i) => ({ x: conc, y: '', id: `pt-${i}` }))); setUnknownGroups([]); setEditingGroupId(null); };

  const handleExport = () => {
    setExporting(true); 
    setTimeout(() => { 
      try {
        const wsData: any[][] = [[t.reportTitle, new Date().toLocaleString()], [], [t.rawMatrix]];
        const colLabels = ["", ...Array.from({length: 12}, (_, i) => (i+1).toString())]; wsData.push(colLabels);
        ROW_HEADERS.forEach(r => { const rd: (string | number)[] = [r]; for(let i=1; i<=12; i++) { const c = results.find(it => it.row === r && it.col === i); rd.push(c ? parseFloat(c.value) : ""); } wsData.push(rd); });
        wsData.push([], [t.sampleTable], [t.newGroup, "Well", t.dil, t.od, t.conc, t.mean, t.sd, t.cv]);
        unknownGroups.forEach(g => { const stats = calculateStats(g.wells.map(w => calculateConc(w.od, fitResult, w.dilution))); g.wells.forEach((w, i) => wsData.push([g.name, `${w.row}${w.col}`, w.dilution, w.od, calculateConc(w.od, fitResult, w.dilution), i===0?stats.mean:"", i===0?stats.sd:"", i===0?stats.cv:""])); });
        const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), "Analysis"); XLSX.writeFile(wb, `LabLens_${Date.now()}.xlsx`);
      } finally { setExporting(false); } 
    }, 100); 
  };

  const fitResult = useMemo(() => calculateLinearFit(stdCurvePoints), [stdCurvePoints]);
  const fitQuality = fitResult ? getFitQuality(fitResult.r2) : 'neutral';

  const updateWellValue = (row: string, col: number, value: string) => {
    setResults(prev => { const existingIdx = prev.findIndex(p => p.row === row && p.col === col); if (existingIdx !== -1) { const next = [...prev]; next[existingIdx] = { ...next[existingIdx], value }; return next; } else { return [...prev, { row, col, value }]; } });
    setStdCurvePoints(prev => prev.map(p => { if (p.sourceWellId === `${row}-${col}`) return { ...p, y: value }; return p; }));
    setUnknownGroups(prev => prev.map(g => ({ ...g, wells: g.wells.map(w => { if (w.row === row && w.col === col) return { ...w, od: parseFloat(value) || 0 }; return w; }) })));
  };

  const handleCellClick = (row: string, col: number) => {
    const item = results.find(it => it.row === row && it.col === col);
    const val = item ? item.value : "";
    const cellId = `${row}-${col}`;
    if (selectionTarget === 'std' && isStdSelecting) {
      const idx = stdCurvePoints.findIndex(p => p.sourceWellId === cellId);
      if (idx !== -1) { setStdCurvePoints(prev => { const next = [...prev]; next[idx] = { ...next[idx], y: '', sourceWellId: undefined }; return next; }); } 
      else { const emptyIdx = stdCurvePoints.findIndex(p => !p.sourceWellId); if (emptyIdx !== -1) { setStdCurvePoints(prev => { const next = [...prev]; next[emptyIdx] = { ...next[emptyIdx], y: val, sourceWellId: cellId }; return next; }); } }
    } else if (selectionTarget === 'sample' && editingGroupId) {
      setUnknownGroups(prev => prev.map(g => { if (g.id !== editingGroupId) return { ...g, wells: g.wells.filter(w => w.row !== row || w.col !== col) }; const exists = g.wells.find(w => w.row === row && w.col === col); return { ...g, wells: exists ? g.wells.filter(w => w.row !== row || w.col !== col) : [...g.wells, { row, col, od: parseFloat(val) || 0, dilution: g.commonDilution }].sort((a, b) => a.row.localeCompare(b.row) || a.col - b.col) }; }));
    }
  };

  const renderPlateCell = (row: string, col: number) => {
    const cellId = `${row}-${col}`;
    const item = results.find(it => it.row === row && it.col === col);
    const stdIdx = stdCurvePoints.findIndex(p => p.sourceWellId === cellId);
    const group = unknownGroups.find(g => g.wells.some(w => w.row === row && w.col === col));
    const gColor = group ? GROUP_COLORS.find(c => c.name === group.color) : null;
    const isSelecting = (selectionTarget === 'std' && isStdSelecting) || (selectionTarget === 'sample' && editingGroupId !== null);
    let bgClass = 'bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 shadow-sm';
    let textClass = 'text-zinc-600 dark:text-zinc-300 placeholder-zinc-300';
    if (stdIdx !== -1) { bgClass = 'bg-indigo-600 border-indigo-600 shadow-md shadow-indigo-500/30'; textClass = 'text-white placeholder-white/50 font-bold'; }
    else if (gColor) { bgClass = `${gColor.bg} ${gColor.border} shadow-md shadow-${gColor.name}-500/30`; textClass = 'text-white placeholder-white/50 font-bold'; }
    else if (!item) { bgClass = 'bg-zinc-50 dark:bg-zinc-800/40 border-zinc-100 dark:border-zinc-800'; }
    return (
      <div key={cellId} onClick={() => handleCellClick(row, col)} className={`relative w-[95%] aspect-square rounded-full flex items-center justify-center transition-all cursor-pointer hover:scale-[1.05] active:scale-95 z-0 hover:z-10 mx-auto ${bgClass}`}> 
        <input type="text" className={`w-full h-full bg-transparent text-center text-[12px] lg:text-[12px] font-mono outline-none p-0 transform origin-center rounded-full ${textClass} ${isSelecting ? 'cursor-pointer' : ''}`} value={item?.value || ''} placeholder="-" onChange={(e) => updateWellValue(row, col, e.target.value)}/>
        {stdIdx !== -1 && <span className="absolute top-0.5 right-0.5 w-3 h-3 bg-white text-indigo-600 text-[8px] rounded-full flex items-center justify-center font-bold z-10 shadow-sm">{stdIdx+1}</span>}
      </div>
    );
  };

  const handlePresetSelect = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const key = e.target.value;
    if (key === 'custom') {
      setStdCurvePoints(Array.from({ length: 10 }, (_, i) => ({ x: '', y: '', id: `pt-${Date.now()}-${i}` })));
    } else {
      const vals = PRESETS[key];
      if (vals) {
        setStdCurvePoints(Array.from({ length: 10 }, (_, i) => ({
          x: vals[i] || '',
          y: '',
          id: `pt-${Date.now()}-${i}`
        })));
      }
    }
  };

  return (
    <div className="h-screen h-[100dvh] w-full flex flex-col bg-[#f8fafc] dark:bg-[#09090b] text-zinc-900 dark:text-zinc-100 transition-colors overflow-hidden">
      <SettingsModal isOpen={showSettings} onClose={() => setShowSettings(false)} t={t} apiKey={customApiKey} setApiKey={handleSetApiKey} />
      
      {/* --- DASHBOARD HEADER --- */}
      <header className="flex-none h-16 px-6 bg-white dark:bg-zinc-900 border-b border-zinc-200 dark:border-zinc-800 flex items-center justify-between shadow-sm z-50">
        <div className="flex items-center gap-4">
          <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center shadow-indigo-200 dark:shadow-none shadow-lg"><FlaskConical className="text-white" size={22}/></div>
          <div><h1 className="font-extrabold text-xl tracking-tight leading-none">{t.appTitle}</h1><p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest mt-1">Created by Jheng</p></div>
        </div>
        <div className="flex items-center gap-2">
          {view === 'results' && (<button onClick={handleExport} disabled={exporting} className="hidden md:flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg text-sm font-bold shadow-lg shadow-emerald-500/20 hover:bg-emerald-700 active:scale-95 transition-all mr-2">{exporting ? <Loader2 size={16} className="animate-spin"/> : <FileSpreadsheet size={16}/>} {t.export}</button>)}
          <div className="hidden sm:flex items-center mr-2 bg-zinc-100 dark:bg-zinc-800 rounded-lg p-0.5">
            <select value={selectedModel} onChange={(e) => setSelectedModel(e.target.value)} className="bg-transparent text-[10px] font-bold text-zinc-600 dark:text-zinc-300 rounded-md px-2 py-1.5 outline-none border-none cursor-pointer hover:bg-zinc-200 dark:hover:bg-zinc-700 transition-colors">
               <option value="gemini-flash-latest">{t.modelFlash}</option>
               <option value="gemini-3-pro-preview">{t.modelPro}</option>
            </select>
          </div>
          <button onClick={() => setShowSettings(true)} className="p-2 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg text-zinc-500 transition-colors"><Settings size={20}/></button>
          <button onClick={() => setLang(l => l === 'en' ? 'zh' : 'en')} className="p-2 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg text-zinc-500 transition-colors"><Languages size={20}/></button>
          <button onClick={() => setTheme(th => th === 'dark' ? 'light' : 'dark')} className="p-2 hover:bg-zinc-100 dark:hover:bg-zinc-800 rounded-lg text-zinc-500 transition-colors">{theme === 'dark' ? <Sun size={20}/> : <Moon size={20}/>}</button>
          {view === 'results' && (<button onClick={reset} className="ml-2 p-2 text-zinc-400 hover:text-red-500 transition-colors"><RefreshCw size={20}/></button>)}
        </div>
      </header>

      <main className="flex-1 overflow-hidden relative">
        {view === 'home' && (
          // 1. 外层增加 overflow-y-auto 允许在高度不够时滚动
          // 2. 使用 min-h-full 替代 h-full，确保内容少时居中，内容多时可撑开
          <div className="w-full h-full overflow-y-auto no-scrollbar">
            <div className="min-h-full flex flex-col items-center justify-center p-6 animate-in fade-in duration-500 py-10 sm:py-0">
              
              {/* 3. 调整标题间距：手机上用 mb-4，大屏用 mb-8 */}
              <div className="text-center mb-4 sm:mb-8">
                <h2 className="text-3xl lg:text-4xl font-black mb-2 text-zinc-800 dark:text-zinc-100">{t.subtitle}</h2>
                <p className="text-zinc-400 font-medium text-sm">{t.desc}</p>
              </div>

              {/* 4. 关键修改：将 md:grid-cols-3 改为 sm:grid-cols-3 */}
              {/* 这样手机横屏（通常 >640px）就会横向排列按钮，不再垂直堆叠 */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 w-full max-w-2xl">
                {[
                  { icon: <Camera size={24}/>, title: t.scanPlate, sub: t.useCamera, action: () => setView('camera'), primary: true }, 
                  { icon: <Upload size={24}/>, title: t.uploadImage, sub: t.fromGallery, action: () => document.getElementById('fileInput')?.click() }, 
                  { icon: <Keyboard size={24}/>, title: t.manualEntry, sub: t.noImage, action: () => setView('results') }
                ].map((card, i) => (
                  <button key={i} onClick={card.action} className={`group p-6 rounded-2xl flex flex-col items-center gap-3 transition-all hover:-translate-y-1 shadow-lg shadow-zinc-200/50 dark:shadow-none border ${card.primary ? 'bg-indigo-600 border-indigo-500 text-white' : 'bg-white dark:bg-zinc-900 border-zinc-100 dark:border-zinc-800'}`}>
                    <div className={`p-3 rounded-xl ${card.primary ? 'bg-white/20' : 'bg-zinc-50 dark:bg-zinc-800 text-indigo-600'}`}>{card.icon}</div>
                    <div className="text-center"><div className="text-base font-bold">{card.title}</div><div className={`text-[10px] uppercase tracking-wider mt-1 opacity-70 ${card.primary ? 'text-indigo-100' : ''}`}>{card.sub}</div></div>
                  </button>
                ))}
                <input id="fileInput" type="file" accept="image/*" onChange={handleImageSelect} className="hidden"/>
              </div>

            </div>
          </div>
        )}

        {view === 'preview' && image && (
          <div className="h-full flex flex-col p-6 animate-in fade-in duration-300">
            <div className="flex-1 relative rounded-3xl overflow-hidden bg-black border border-zinc-800 shadow-2xl group">
              <img src={image} className="w-full h-full object-contain opacity-50 group-hover:opacity-70 transition-opacity duration-500" />
              <div className="absolute inset-0 flex flex-col items-center justify-center p-8 bg-gradient-to-t from-black/80 via-transparent to-black/20">
                {error ? (
                  <div className="bg-red-500/10 backdrop-blur-md border border-red-500/20 p-6 rounded-2xl max-w-md text-center"><AlertTriangle className="mx-auto text-red-400 mb-2" size={32}/><p className="text-red-200 font-medium">{error}</p><button onClick={processImage} className="mt-4 px-4 py-2 bg-red-500/20 hover:bg-red-500/30 text-red-200 rounded-lg text-sm font-bold transition-all">Retry</button></div>
                ) : (
                  <div className="text-center">{loading ? (<div className="flex flex-col items-center gap-4"><Loader2 size={48} className="text-white animate-spin"/><p className="text-white font-medium text-lg tracking-wide animate-pulse">{t.analyzing}</p></div>) : (<button onClick={processImage} className="group relative px-8 py-4 bg-white text-black rounded-full font-bold text-lg shadow-[0_0_40px_-10px_rgba(255,255,255,0.3)] hover:scale-105 active:scale-95 transition-all flex items-center gap-3"><Sparkles size={20} className="text-indigo-600"/> {t.analyze}<div className="absolute inset-0 rounded-full ring-2 ring-white/20 group-hover:ring-white/40 transition-all"/></button>)}</div>
                )}
              </div>
              <button onClick={reset} className="absolute top-6 left-6 p-3 bg-black/50 backdrop-blur text-white rounded-full hover:bg-black/70 transition-all"><ArrowLeft size={24}/></button>
            </div>
          </div>
        )}

        {view === 'camera' && <CameraView onCapture={(img) => { setImage(img); setView('preview'); }} onBack={() => setView('home')} t={t}/>}

        {view === 'results' && (
  <div className="h-full flex flex-col p-2 lg:p-6 gap-4 lg:gap-6 overflow-y-auto pb-24 no-scrollbar">
    <div className="flex-1 grid grid-cols-1 lg:grid-cols-12 gap-4 lg:gap-6 min-h-0">
      
      {/* PANEL 1: PLATE MAP */}
      <div className="lg:col-span-7 xl:col-span-8 bg-white dark:bg-zinc-900 rounded-3xl border border-zinc-200 dark:border-zinc-800 flex flex-col transition-all relative h-[500px] lg:h-[calc(100vh-140px)] xl:h-full overflow-hidden shadow-sm">
        
        {/* 1. 增加 padding (p-6 lg:p-10) 制造留白 */}
        {/* 2. Flex + Center 确保孔板居中 */}
        <div className="flex-1 w-full h-full min-h-0 p-6 lg:p-10 flex flex-col items-center justify-center bg-gray-10 dark:bg-zinc-950/20 overflow-hidden relative">
           
           {/* 3. 核心修复: 去掉 w-full h-full，改用 max-w-full max-h-full */}
           {/* 这会让孔板尽可能大，但在触碰到宽或高的边界时停止，绝对不会被裁剪 */}
           <div className={`transition-all duration-300 flex items-center justify-center w-full h-full`}>
              
              {!useStandardLayout ? (
                /* === Mobile Portrait Layout (8列 x 12行) === */
                /* 竖屏模式：保持 aspect-[2/3]，同时限制最大宽高 */
                <div className="w-auto h-auto max-w-full max-h-full aspect-[2/3] flex flex-col justify-center m-auto shadow-sm">
                  <div className="grid grid-cols-[repeat(8,1fr)_auto] gap-px sm:gap-1 mb-1">
                    {[...ROW_HEADERS].reverse().map((char, i) => (<div key={i} className="text-center text-[10px] sm:text-xs font-bold text-zinc-400 select-none">{char}</div>))}
                    <div className="w-6 sm:w-8"></div>
                  </div>
                  <div className="flex-1 flex flex-col justify-start gap-px sm:gap-1 min-h-0">
                    {Array.from({length: 12}).map((_, rIdx) => {
                      const rowNum = rIdx + 1;
                      return (
                        <div key={rowNum} className="flex-1 grid grid-cols-[repeat(8,1fr)_auto] gap-px sm:gap-1 items-center">
                          {[...ROW_HEADERS].reverse().map((rowChar) => renderPlateCell(rowChar, rowNum))}
                          <div className="w-6 sm:w-8 flex items-center justify-center text-[10px] sm:text-xs font-bold text-zinc-400 select-none">{rowNum}</div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ) : (
                /* === Tablet/Desktop Layout (12列 x 8行) === */
                /* 横屏模式：保持 aspect-[13/9]，同时限制最大宽高 */
                <div className="w-auto h-auto max-w-full max-h-full aspect-[13/9] flex flex-col justify-center m-auto shadow-sm">
                  <div className="grid grid-cols-[auto_repeat(12,1fr)] gap-px sm:gap-1 mb-1 flex-none">
                     <div className="w-6 sm:w-8"></div>
                     {Array.from({length: 12}).map((_, i) => (<div key={i} className="text-center text-[10px] sm:text-xs font-bold text-zinc-400 select-none">{i+1}</div>))}
                  </div>
                  <div className="flex-1 flex flex-col justify-start gap-px sm:gap-1 min-h-0">
                    {ROW_HEADERS.map((rowChar, rIdx) => (
                      <div key={rowChar} className="flex-1 grid grid-cols-[auto_repeat(12,1fr)] gap-px sm:gap-1 items-center">
                        <div className="w-6 sm:w-8 flex items-center justify-center text-[10px] sm:text-xs font-bold text-zinc-400 select-none">{rowChar}</div>
                        {Array.from({ length: 12 }).map((_, cIdx) => renderPlateCell(rowChar, cIdx + 1))}
                      </div>
                    ))}
                  </div>
                </div>
              )}
           </div>
        </div>
      </div>

      {/* --- PANEL 2: ANALYSIS --- */}
      <div className="lg:col-span-5 xl:col-span-4 bg-white dark:bg-zinc-900 rounded-3xl border border-zinc-200 dark:border-zinc-800 flex flex-col overflow-hidden shadow-sm h-[500px] lg:h-[calc(100vh-140px)] xl:h-full transition-all relative">
        <div className="p-2.5 border-b border-zinc-100 dark:border-zinc-800 flex justify-between items-center h-[60px] flex-none bg-white dark:bg-zinc-900">
           <div className="flex bg-zinc-100 dark:bg-zinc-800 rounded-lg p-1 gap-1">
              <button onClick={() => setAnalysisTab('calibration')} className={`flex items-center gap-2 px-3 py-1.5 rounded-md text-xs font-bold transition-all ${analysisTab === 'calibration' ? 'bg-white dark:bg-zinc-700 shadow-sm text-indigo-600 dark:text-indigo-300' : 'text-zinc-500 hover:text-zinc-700 dark:text-zinc-400 dark:hover:text-zinc-200'}`}><TrendingUp size={14} className={analysisTab === 'calibration' ? 'text-indigo-500' : 'opacity-50'}/>{t.calibration}</button>
              <button onClick={() => setAnalysisTab('samples')} className={`flex items-center gap-2 px-3 py-1.5 rounded-md text-xs font-bold transition-all ${analysisTab === 'samples' ? 'bg-white dark:bg-zinc-700 shadow-sm text-emerald-600 dark:text-emerald-300' : 'text-zinc-500 hover:text-zinc-700 dark:text-zinc-400 dark:hover:text-zinc-200'}`}><TestTube2 size={14} className={analysisTab === 'samples' ? 'text-emerald-500' : 'opacity-50'}/>{t.samples}</button>
           </div>
           {analysisTab === 'calibration' && (<button onClick={() => { setSelectionTarget('std'); setIsStdSelecting(!isStdSelecting); }} className={`px-3 py-1.5 rounded-full text-[10px] font-bold transition-all border ${isStdSelecting ? 'bg-indigo-600 border-indigo-500 text-white' : 'bg-white dark:bg-zinc-800 border-zinc-200 dark:border-zinc-700 text-zinc-500 hover:border-indigo-400'}`}>{isStdSelecting ? t.selecting : t.selectRange}</button>)}
           {analysisTab === 'samples' && (<button onClick={() => { setSelectionTarget('sample'); const nId = Date.now().toString(); setUnknownGroups(prev => [...prev, { id: nId, name: `${t.sampleName} ${prev.length+1}`, commonDilution: 1, color: GROUP_COLORS[prev.length % GROUP_COLORS.length].name, wells: [] }]); setEditingGroupId(nId); }} className="px-3 py-1.5 bg-zinc-900 dark:bg-white text-white dark:text-black rounded-lg text-[10px] font-bold hover:scale-105 active:scale-95 transition-all flex items-center gap-1"><Plus size={12}/> {t.newGroup}</button>)}
        </div>

        <div className="flex-1 overflow-hidden relative flex flex-col min-h-0">
          {analysisTab === 'calibration' && (
            <div className="flex-1 flex flex-col min-h-0 animate-in fade-in slide-in-from-right-2 duration-300">
               <div className="flex-1 p-2 relative bg-zinc-50/50 dark:bg-zinc-950/20 w-full min-h-0">
                    {fitResult && (<div className="absolute top-4 right-4 z-10 bg-white/80 dark:bg-zinc-800/80 backdrop-blur-md border border-zinc-200 dark:border-zinc-700 p-2.5 rounded-xl shadow-sm flex flex-col gap-1 pointer-events-none"><div className="flex items-center justify-between gap-4"><span className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">{t.summaryR2}</span><span className={`text-xs font-mono font-bold ${fitQuality === 'good' ? 'text-emerald-500' : fitQuality === 'warn' ? 'text-amber-500' : 'text-red-500'}`}>{fitResult.r2.toFixed(4)}</span></div><div className="text-[10px] font-medium text-zinc-500 font-mono">y={fitResult.slope.toFixed(3)}x {fitResult.intercept >= 0 ? '+' : ''}{fitResult.intercept.toFixed(3)}</div></div>)}
                    {stdCurvePoints.some(p => p.y) ? (<StandardCurveChart points={stdCurvePoints} slope={fitResult?.slope} intercept={fitResult?.intercept} t={t}/>) : (<div className="h-full flex flex-col items-center justify-center text-zinc-300 gap-2"><TrendingUp size={32} className="opacity-20"/><span className="text-xs italic">{t.noDataSelected}</span></div>)}
               </div>
               <div className="flex-none border-t border-zinc-100 dark:border-zinc-800 bg-white dark:bg-zinc-900 z-10 flex flex-col">
                  <div className="flex items-center justify-between px-4 py-2 border-b border-zinc-50 dark:border-zinc-800">
                     <span className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">{t.stdCurve}</span> 
                     <div className="flex items-center gap-2">
                        <span className="text-[10px] font-medium text-zinc-500">{t.presets}</span>
                        <select className="text-[10px] bg-zinc-50 dark:bg-zinc-800 border-none rounded px-2 py-1 outline-none font-bold text-zinc-700 dark:text-zinc-300 cursor-pointer hover:bg-zinc-100 dark:hover:bg-zinc-700 transition-colors" onChange={handlePresetSelect} defaultValue="">
                          <option value="" disabled>{t.selectPreset}</option>
                          <option value="BCA">BCA (0 - 0.5)</option>
                          <option value="Bradford">Bradford (0 - 1.5)</option>
                          <option value="custom">{t.custom}</option>
                        </select>
                     </div>
                  </div>
                  <div className="p-3 grid grid-cols-5 gap-2 no-scrollbar overflow-y-auto max-h-40">
                     {stdCurvePoints.map((pt, i) => (
                       <div key={pt.id} className="flex-none w-full flex flex-col gap-1 group">
                          <div className="flex justify-between items-center px-0.5"><span className="text-[9px] font-mono text-zinc-400">#{i+1}</span></div>
                          <div className={`relative rounded-md border transition-all ${pt.sourceWellId ? 'border-indigo-500 bg-indigo-50/20' : 'border-zinc-200 dark:border-zinc-700 bg-zinc-50 dark:bg-zinc-800 focus-within:border-zinc-400'}`}>
                            <input 
                              value={pt.x} 
                              onChange={e => setStdCurvePoints(prev => prev.map(p => p.id === pt.id ? {...p, x: e.target.value} : p))} 
                              className="w-full text-center text-[10px] font-mono font-bold bg-transparent outline-none py-1.5 text-zinc-700 dark:text-zinc-200 placeholder-zinc-300" 
                              placeholder="-"
                            />
                            {pt.sourceWellId && <div className="absolute top-0 right-0 w-1.5 h-1.5 bg-indigo-500 rounded-full translate-x-1/3 -translate-y-1/3 shadow-sm ring-1 ring-white dark:ring-zinc-900"></div>}
                          </div>
                          <div className="h-3 text-center">{pt.y && <span className="text-[9px] font-mono text-zinc-400">{pt.y}</span>}</div>
                        </div>
                     ))}
                  </div>
               </div>
            </div>
          )}
          {analysisTab === 'samples' && (
            <div className="flex-1 overflow-y-auto p-4 animate-in fade-in slide-in-from-right-2 duration-300 bg-zinc-50/50 dark:bg-zinc-950/20 no-scrollbar">
               {unknownGroups.length === 0 ? (<div className="h-full flex flex-col items-center justify-center text-zinc-400 py-10"><TestTube2 size={48} className="mb-3 opacity-20"/><p className="text-xs font-medium">{t.noGroups}</p></div>) : (
                <div className="space-y-3">
                  {unknownGroups.map(group => {
                    const color = GROUP_COLORS.find(c => c.name === group.color)!;
                    const stats = calculateStats(group.wells.map(w => calculateConc(w.od, fitResult, w.dilution)));
                    const isEditing = editingGroupId === group.id;
                    return (
                      <div key={group.id} onClick={() => { setEditingGroupId(group.id); setSelectionTarget('sample'); }} className={`bg-white dark:bg-zinc-900 rounded-xl border transition-all cursor-pointer shadow-sm overflow-hidden ${isEditing ? 'border-indigo-500 ring-1 ring-indigo-500 shadow-md shadow-indigo-500/10' : 'border-zinc-200 dark:border-zinc-800 hover:border-zinc-300'}`}>
                        <div className={`px-3 py-2.5 flex items-center gap-3 ${isEditing ? 'bg-zinc-50 dark:bg-zinc-800/50' : ''}`}>
                          <div className={`w-2.5 h-2.5 rounded-full ${color.bg}`}/>
                          <input value={group.name} onChange={e => setUnknownGroups(prev => prev.map(g => g.id === group.id ? {...g, name: e.target.value} : g))} onClick={(e) => e.stopPropagation()} className="font-bold text-xs sm:text-sm bg-transparent outline-none flex-1 text-zinc-700 dark:text-zinc-200"/>
                          {!isEditing && (<div className="flex items-center gap-2 text-[10px] text-zinc-400 font-mono"><span className="font-bold text-zinc-600 dark:text-zinc-300">{stats.mean.toFixed(2)}</span><span className="text-zinc-300">|</span><span>CV {stats.cv.toFixed(0)}%</span></div>)}
                          <div className="flex items-center gap-2"><span className="text-[9px] bg-zinc-100 dark:bg-zinc-800 px-1.5 py-0.5 rounded text-zinc-500 font-mono font-bold">n={group.wells.length}</span>{isEditing && <button onClick={(e) => { e.stopPropagation(); setUnknownGroups(prev => prev.filter(g => g.id !== group.id)); }} className="text-zinc-300 hover:text-red-500 p-0.5"><X size={14}/></button>}</div>
                        </div>
                        {isEditing && (
                          <div className="p-3 border-t border-zinc-100 dark:border-zinc-800 animate-in slide-in-from-top-1 duration-200">
                            <div className="flex gap-2 mb-3">
                               {[{ l: t.mean, v: stats.mean.toFixed(2) }, { l: t.sd, v: stats.sd.toFixed(2) }, { l: t.cv, v: stats.cv.toFixed(1) + '%' }].map((s, i) => (<div key={i} className="flex-1 bg-zinc-50 dark:bg-zinc-800/50 rounded-lg p-1.5 text-center border border-zinc-100 dark:border-zinc-800"><div className="text-[8px] text-zinc-400 font-bold uppercase tracking-tight">{s.l}</div><div className="text-xs font-mono font-bold text-zinc-600 dark:text-zinc-300">{s.v}</div></div>))}
                            </div>
                            <div className="flex items-center justify-between mb-3 px-1"><label className="text-[10px] font-bold text-zinc-400 uppercase tracking-wide flex items-center gap-1"><Beaker size={10}/> {t.setAllDil}</label><input type="number" value={group.commonDilution} onClick={(e)=>e.stopPropagation()} onChange={e => { const val = parseFloat(e.target.value)||1; setUnknownGroups(prev => prev.map(g => g.id === group.id ? {...g, commonDilution: val, wells: g.wells.map(w => ({...w, dilution: val}))} : g)); }} className="w-16 text-right text-xs bg-zinc-100 dark:bg-zinc-800 px-2 py-1 rounded border-none outline-none focus:ring-1 focus:ring-indigo-500 font-mono font-bold"/></div>
                            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 max-h-40 overflow-y-auto pr-1 no-scrollbar">
                              {group.wells.map((w, i) => (<div key={i} className="flex flex-col bg-zinc-50 dark:bg-zinc-800/30 border border-zinc-100 dark:border-zinc-800 rounded-md p-1.5 group/row relative overflow-hidden"><div className="flex justify-between items-start"><span className="font-mono font-bold text-[10px] text-zinc-400">{w.row}{w.col}</span><button onClick={(e) => { e.stopPropagation(); setUnknownGroups(prev => prev.map(g => g.id === group.id ? {...g, wells: g.wells.filter((_, idx) => idx !== i)} : g)); }} className="text-zinc-300 hover:text-red-500 opacity-0 group-hover/row:opacity-100 absolute top-1 right-1"><X size={10}/></button></div><div className="flex justify-between items-end mt-1"><span className="text-[9px] text-zinc-400">x{w.dilution}</span><span className="font-mono font-bold text-xs text-emerald-600 dark:text-emerald-400">{calculateConc(w.od, fitResult, w.dilution).toFixed(2)}</span></div></div>))}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
               )}
            </div>
          )}
        </div>
      </div>
    </div>
  </div>
)}
      </main>
    </div>
  );
}

const root = createRoot(document.getElementById('root')!);
root.render(<App />);