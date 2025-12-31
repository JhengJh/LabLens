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
  Key, Save, Eye, EyeOff, Lock, Settings
} from 'lucide-react';
import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";

// --- Translations ---

const TRANSLATIONS = {
  en: {
    appTitle: "LabLens",
    subtitle: "Microplate Assistant",
    desc: "AI-powered analysis for your 96-well assays.",
    scanPlate: "Scan Plate",
    useCamera: "Use Camera",
    uploadImage: "Upload Image",
    fromGallery: "From Gallery",
    manualEntry: "Manual Entry",
    noImage: "No Image",
    analyzing: "Analyzing Grid Structure...",
    discard: "Discard",
    analyze: "Analyze Plate",
    plateMap: "Plate Map",
    calibration: "Calibration",
    samples: "Samples",
    selectRange: "Select Wells",
    selecting: "Done",
    noDataSelected: "No data points selected",
    conc: "Conc",
    od: "OD",
    addRow: "Add Row",
    unknowns: "Unknowns",
    newGroup: "New Group",
    noGroups: "No sample groups created.",
    sampleName: "Sample Name",
    mean: "Mean",
    sd: "SD",
    cv: "CV%",
    setAllDil: "Set All Dilutions",
    dil: "Dil",
    addManualWell: "+ Add Manual Well",
    export: "Export",
    slope: "Slope",
    intercept: "Intercept",
    r2: "R²",
    cameraError: "Camera access denied or unavailable.",
    return: "Return",
    concAxis: "Concentration",
    odAxis: "OD Value",
    reportTitle: "LabLens Report",
    rawMatrix: "RAW DATA MATRIX",
    stdCurve: "STANDARD CURVE",
    sampleTable: "SAMPLES",
    fitGood: "Excellent Fit",
    fitWarn: "Acceptable Fit",
    fitPoor: "Poor Fit",
    settings: "Settings",
    apiKeyTitle: "Gemini API Key",
    apiKeyDesc: "Enter your Google Gemini API Key to enable AI features. The key is stored locally in your browser.",
    apiKeyPlaceholder: "Paste your API key here (AIza...)",
    save: "Save Key",
    getKey: "Get a free API key",
    missingKey: "API Key Missing",
    missingKeyMsg: "Please set your Gemini API Key in the settings to analyze images."
  },
  zh: {
    appTitle: "LabLens",
    subtitle: "微孔板助手",
    desc: "AI 驱动的 96 孔板数据分析工具",
    scanPlate: "扫描孔板",
    useCamera: "使用相机",
    uploadImage: "上传图片",
    fromGallery: "从相册选择",
    manualEntry: "手动录入",
    noImage: "无图片模式",
    analyzing: "正在分析网格结构...",
    discard: "放弃",
    analyze: "分析孔板",
    plateMap: "孔板视图",
    calibration: "标准曲线",
    samples: "样本分析",
    selectRange: "选择孔位",
    selecting: "完成",
    noDataSelected: "未选择数据点",
    conc: "浓度",
    od: "OD值",
    addRow: "添加行",
    unknowns: "未知样本",
    newGroup: "新建组",
    noGroups: "暂无样本组",
    sampleName: "样本名称",
    mean: "平均值",
    sd: "标准差",
    cv: "变异系数",
    setAllDil: "统一稀释倍数",
    dil: "稀释",
    addManualWell: "+ 添加手动读数",
    export: "导出报告",
    slope: "斜率",
    intercept: "截距",
    r2: "R²",
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
    apiKeyTitle: "Gemini API 密钥",
    apiKeyDesc: "请输入您的 Google Gemini API 密钥以启用 AI 功能。密钥仅存储在您的本地浏览器中。",
    apiKeyPlaceholder: "在此粘贴您的 API 密钥 (AIza...)",
    save: "保存密钥",
    getKey: "获取免费 API 密钥",
    missingKey: "缺少 API 密钥",
    missingKeyMsg: "请在设置中配置您的 Gemini API 密钥以开始分析图片。"
  }
};

// --- Types ---

interface ExtractedItem {
  value: string;
  row: string;
  col: number;
}

interface StdCurvePoint {
  x: string; // Concentration
  y: string; // OD Value
  id: string; // Unique ID for keying
  sourceWellId?: string; // Links to a specific well (row-col)
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
  dilution: number; // Per-well dilution
  id?: string; // Optional unique ID for manual entries
}

interface UnknownGroup {
  id: string;
  name: string;
  commonDilution: number; // Used as default for new wells and batch updating
  color: string; // Tailwind color class key
  wells: SelectedWell[];
}

// --- Constants ---

const ROW_HEADERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
const COL_COUNT = 12;

// Default concentrations as requested
const DEFAULT_CONCS = ['1.5', '1', '0.75', '0.5', '0.25', '0.125', '0.0625', '0'];

// Updated Colors for Dark/Light Mode Compatibility
const GROUP_COLORS = [
  { 
    name: 'emerald', 
    bg: 'bg-emerald-500', 
    border: 'border-emerald-500',
    text: 'text-emerald-700 dark:text-emerald-400', 
    ring: 'ring-emerald-500',
  },
  { 
    name: 'violet', 
    bg: 'bg-violet-500', 
    border: 'border-violet-500',
    text: 'text-violet-700 dark:text-violet-400', 
    ring: 'ring-violet-500',
  },
  { 
    name: 'pink', 
    bg: 'bg-pink-500', 
    border: 'border-pink-500',
    text: 'text-pink-700 dark:text-pink-400', 
    ring: 'ring-pink-500',
  },
  { 
    name: 'orange', 
    bg: 'bg-orange-500', 
    border: 'border-orange-500',
    text: 'text-orange-700 dark:text-orange-400', 
    ring: 'ring-orange-500',
  },
  { 
    name: 'cyan', 
    bg: 'bg-cyan-500', 
    border: 'border-cyan-500',
    text: 'text-cyan-700 dark:text-cyan-400', 
    ring: 'ring-cyan-500',
  },
];

// --- Helper Functions ---

function calculateLinearFit(points: StdCurvePoint[]): FitResult | null {
  const validPoints = points
    .map(p => ({ x: parseFloat(p.x), y: parseFloat(p.y) }))
    .filter(p => !isNaN(p.x) && !isNaN(p.y));

  const n = validPoints.length;
  if (n < 2) return null;

  const sumX = validPoints.reduce((sum, p) => sum + p.x, 0);
  const sumY = validPoints.reduce((sum, p) => sum + p.y, 0);
  const sumXY = validPoints.reduce((sum, p) => sum + p.x * p.y, 0);
  const sumXX = validPoints.reduce((sum, p) => sum + p.x * p.x, 0);
  const sumYY = validPoints.reduce((sum, p) => sum + p.y * p.y, 0);

  const denom = n * sumXX - sumX * sumX;
  if (denom === 0) return null; // Vertical line

  const slope = (n * sumXY - sumX * sumY) / denom;
  const intercept = (sumY - slope * sumX) / n;

  // Calculate R-squared
  const yMean = sumY / n;
  let ssRes = 0;
  let ssTot = 0;
  for (const p of validPoints) {
    const yPred = slope * p.x + intercept;
    ssRes += Math.pow(p.y - yPred, 2);
    ssTot += Math.pow(p.y - yMean, 2);
  }
  
  const r2 = ssTot === 0 ? 1 : 1 - (ssRes / ssTot);

  return { slope, intercept, r2 };
}

function calculateConc(od: number, fit: FitResult | null, dilution: number) {
  if (!fit) return 0;
  let rawConc = (od - fit.intercept) / fit.slope;
  return rawConc * dilution;
}

function calculateStats(values: number[]) {
  if (values.length === 0) return { mean: 0, sd: 0, cv: 0 };
  const sum = values.reduce((a, b) => a + b, 0);
  const mean = sum / values.length;
  if (values.length === 1) return { mean, sd: 0, cv: 0 };
  
  const variance = values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (values.length - 1);
  const sd = Math.sqrt(variance);
  const cv = mean === 0 ? 0 : (sd / mean) * 100;
  return { mean, sd, cv };
}

// Parses "A1", "B12" etc into {row, col}
function parseWellLabel(label: string): { row: string, col: number } | null {
    const match = label.toUpperCase().match(/^([A-H])([0-9]{1,2})$/);
    if (!match) return null;
    const row = match[1];
    const col = parseInt(match[2]);
    if (col < 1 || col > 12) return null;
    return { row, col };
}

// Media Query Hook
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

// Helper for Fit Quality
function getFitQuality(r2: number) {
  if (r2 >= 0.98) return 'good';
  if (r2 >= 0.90) return 'warn';
  return 'poor';
}

// --- Components ---

function ApiKeyModal({ isOpen, onClose, t }: { isOpen: boolean; onClose: () => void; t: any }) {
  const [key, setKey] = useState('');
  const [show, setShow] = useState(false);

  useEffect(() => {
    if (isOpen) {
        const stored = localStorage.getItem('gemini_api_key');
        if (stored) setKey(stored);
    }
  }, [isOpen]);

  const handleSave = () => {
    localStorage.setItem('gemini_api_key', key.trim());
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="bg-white dark:bg-zinc-900 rounded-2xl w-full max-w-md p-6 shadow-2xl border border-zinc-200 dark:border-zinc-800">
        <div className="flex justify-between items-center mb-4">
          <h3 className="text-xl font-bold flex items-center gap-2 text-zinc-900 dark:text-white">
            <Key size={20} className="text-blue-600"/> {t.apiKeyTitle}
          </h3>
          <button onClick={onClose} className="text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 transition-colors">
            <X size={20}/>
          </button>
        </div>
        <p className="text-sm text-zinc-500 dark:text-zinc-400 mb-6 leading-relaxed">
          {t.apiKeyDesc}
        </p>
        
        <div className="relative mb-6">
          <input 
            type={show ? "text" : "password"} 
            value={key}
            onChange={(e) => setKey(e.target.value)}
            placeholder={t.apiKeyPlaceholder}
            className="w-full bg-zinc-100 dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-xl px-4 py-3 pr-12 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 font-mono text-zinc-900 dark:text-zinc-100 placeholder-zinc-400"
          />
          <button 
            onClick={() => setShow(!show)}
            className="absolute right-3 top-1/2 -translate-y-1/2 text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 p-1"
          >
            {show ? <EyeOff size={18}/> : <Eye size={18}/>}
          </button>
        </div>

        <div className="flex gap-3">
          <a 
            href="https://aistudio.google.com/app/apikey" 
            target="_blank" 
            rel="noopener noreferrer"
            className="flex-1 py-2.5 px-4 rounded-xl border border-zinc-200 dark:border-zinc-700 text-sm font-medium text-center text-zinc-700 dark:text-zinc-300 hover:bg-zinc-50 dark:hover:bg-zinc-800 transition-colors"
          >
            {t.getKey}
          </a>
          <button 
            onClick={handleSave}
            className="flex-1 py-2.5 px-4 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold flex items-center justify-center gap-2 transition-colors shadow-lg shadow-blue-500/20 active:scale-95"
          >
            <Save size={16}/> {t.save}
          </button>
        </div>
        
        <div className="mt-4 flex items-center gap-2 justify-center text-xs text-zinc-400">
           <Lock size={12}/> 
           <span>Stored locally via localStorage</span>
        </div>
      </div>
    </div>
  );
}

function CameraView({ onCapture, onBack, t }: { onCapture: (img: string) => void, onBack: () => void, t: any }) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let stream: MediaStream | null = null;
    const startCamera = async () => {
      try {
        stream = await navigator.mediaDevices.getUserMedia({ 
          video: { facingMode: 'environment' } 
        });
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
        }
      } catch (err) {
        console.error("Error accessing camera:", err);
        setError(t.cameraError);
      }
    };

    startCamera();

    return () => {
      if (stream) {
        stream.getTracks().forEach(track => track.stop());
      }
    };
  }, [t]);

  const handleCapture = () => {
    if (videoRef.current && canvasRef.current) {
      const video = videoRef.current;
      const canvas = canvasRef.current;
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        const dataUrl = canvas.toDataURL('image/jpeg');
        onCapture(dataUrl);
      }
    }
  };

  return (
    <div className="relative w-full h-full bg-black flex flex-col items-center justify-center">
      {error ? (
        <div className="text-white text-center p-4">
          <p className="text-red-400 mb-4">{error}</p>
          <button onClick={onBack} className="px-6 py-2 bg-white text-black rounded-full font-medium">
            {t.return}
          </button>
        </div>
      ) : (
        <>
          <video 
            ref={videoRef} 
            autoPlay 
            playsInline 
            className="w-full h-full object-cover"
          />
          <canvas ref={canvasRef} className="hidden" />
          
          <div className="absolute top-0 left-0 p-6 z-10">
             <button onClick={onBack} className="text-white hover:opacity-70 transition-opacity">
               <X size={32} />
             </button>
          </div>

          <div className="absolute bottom-0 left-0 right-0 p-10 flex justify-center items-center bg-gradient-to-t from-black/90 to-transparent">
            <button 
              onClick={handleCapture}
              className="w-20 h-20 rounded-full border-[6px] border-white/30 bg-white flex items-center justify-center hover:scale-105 active:scale-95 transition-all shadow-2xl"
            />
          </div>
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
    
    // Use ResizeObserver for more robust sizing than window 'resize'
    const resizeObserver = new ResizeObserver((entries) => {
      if (!entries || entries.length === 0) return;
      const { width, height } = entries[0].contentRect;
      setDimensions({ width, height });
    });

    resizeObserver.observe(containerRef.current);

    return () => {
      resizeObserver.disconnect();
    };
  }, []);

  const data = useMemo(() => {
    return points
      .map(p => ({ x: parseFloat(p.x), y: parseFloat(p.y) }))
      .filter(p => !isNaN(p.x) && !isNaN(p.y));
  }, [points]);

  // Calculate nice axis range based on min/max values
  const calculateAxisRange = (min: number, max: number) => {
    let range = max - min;
    // Handle cases with no variation or single point or no data
    if (range <= 1e-9) { 
        if (min === 0) { max = 1; }
        else { min = min * 0.9; max = max * 1.1; }
        range = max - min;
    }
    
    // Add padding (approx 5-10% on each side)
    const paddingVal = range * 0.1;
    let niceMin = min - paddingVal;
    let niceMax = max + paddingVal;

    // Calculate tick step
    const targetTicks = 5;
    const roughStep = (niceMax - niceMin) / targetTicks;
    const exponent = Math.floor(Math.log10(roughStep));
    const fraction = roughStep / Math.pow(10, exponent);
    
    let niceFraction = 1;
    if (fraction <= 1) niceFraction = 1;
    else if (fraction <= 2) niceFraction = 2;
    else if (fraction <= 5) niceFraction = 5;
    else niceFraction = 10;
    
    const step = niceFraction * Math.pow(10, exponent);
    
    // Round min/max to step boundary
    const startTick = Math.floor(niceMin / step) * step;
    const endTick = Math.ceil(niceMax / step) * step;
    
    const ticks: number[] = [];
    // Ensure we cover the range, dealing with floating point issues
    for (let t = startTick; t <= endTick + (step * 0.001); t += step) {
        ticks.push(t);
    }

    return { min: startTick, max: endTick, ticks };
  };

  if (dimensions.width === 0) return <div ref={containerRef} className="w-full h-full" />;

  // 1. Calculate Data Ranges
  const xValues = data.map(p => p.x);
  const yValues = data.map(p => p.y);

  // X Axis Range
  const xMinData = xValues.length ? Math.min(...xValues) : 0;
  const xMaxData = xValues.length ? Math.max(...xValues) : 1;
  const xScaleInfo = calculateAxisRange(xMinData, xMaxData);

  // Y Axis Range
  let yMinData = yValues.length ? Math.min(...yValues) : 0;
  let yMaxData = yValues.length ? Math.max(...yValues) : 1;

  // Predict Y at min/max X (based on X axis view) to ensure fit line is visible
  if (typeof slope === 'number' && typeof intercept === 'number') {
    const yAtMinX = slope * xScaleInfo.min + intercept;
    const yAtMaxX = slope * xScaleInfo.max + intercept;
    yMinData = Math.min(yMinData, yAtMinX, yAtMaxX);
    yMaxData = Math.max(yMaxData, yAtMinX, yAtMaxX);
  }
    
  const yScaleInfo = calculateAxisRange(yMinData, yMaxData);

  const { width, height } = dimensions;
  // Padding for labels
  const padding = { top: 20, right: 30, bottom: 40, left: 50 };

  const scaleX = (val: number) => padding.left + ((val - xScaleInfo.min) / (xScaleInfo.max - xScaleInfo.min)) * (width - padding.left - padding.right);
  const scaleY = (val: number) => height - padding.bottom - ((val - yScaleInfo.min) / (yScaleInfo.max - yScaleInfo.min)) * (height - padding.top - padding.bottom);

  return (
    <div ref={containerRef} className="w-full h-full flex items-center justify-center select-none font-mono text-xs">
      <svg width={width} height={height} className="overflow-visible">
        {/* Y Axis Grid & Labels */}
        {yScaleInfo.ticks.map(tick => {
            const y = scaleY(tick);
            if (y < padding.top - 10 || y > height - padding.bottom + 10) return null;
            return (
                <g key={`y-${tick}`}>
                    <line x1={padding.left} y1={y} x2={width - padding.right} y2={y} stroke="currentColor" className="text-gray-200 dark:text-zinc-800" strokeDasharray="4" />
                    <text x={padding.left - 8} y={y + 3} textAnchor="end" className="fill-gray-400 font-medium">{tick.toFixed(2).replace(/\.00$/, '')}</text>
                </g>
            );
        })}

        {/* X Axis Grid & Labels */}
        {xScaleInfo.ticks.map(tick => {
            const x = scaleX(tick);
            if (x < padding.left - 10 || x > width - padding.right + 10) return null;
            return (
                <g key={`x-${tick}`}>
                    <line x1={x} y1={padding.top} x2={x} y2={height - padding.bottom} stroke="currentColor" className="text-gray-200 dark:text-zinc-800" strokeDasharray="4" />
                    <text x={x} y={height - padding.bottom + 15} textAnchor="middle" className="fill-gray-400 font-medium">{tick.toFixed(2).replace(/\.00$/, '')}</text>
                </g>
            );
        })}

        {/* Axis Lines */}
        <line x1={padding.left} y1={padding.top} x2={padding.left} y2={height - padding.bottom} stroke="currentColor" className="text-gray-300 dark:text-zinc-600" />
        <line x1={padding.left} y1={height - padding.bottom} x2={width - padding.right} y2={height - padding.bottom} stroke="currentColor" className="text-gray-300 dark:text-zinc-600" />

        {/* Fit Line */}
        {typeof slope === 'number' && typeof intercept === 'number' && (
             <line 
             x1={scaleX(xScaleInfo.min)} 
             y1={scaleY(slope * xScaleInfo.min + intercept)} 
             x2={scaleX(xScaleInfo.max)} 
             y2={scaleY(slope * xScaleInfo.max + intercept)} 
             stroke="currentColor"
             className="text-blue-500 dark:text-blue-400"
             strokeWidth="3" 
             strokeLinecap="round"
           />
        )}

        {/* Data Points */}
        {data.map((p, i) => (
          <g key={i} className="group cursor-pointer">
            <circle 
              cx={scaleX(p.x)} 
              cy={scaleY(p.y)} 
              r="6" 
              className="fill-zinc-50 dark:fill-zinc-900 stroke-blue-600 dark:stroke-blue-500 group-hover:fill-blue-100 dark:group-hover:fill-blue-900/30 group-hover:r-8 transition-all duration-300"
              strokeWidth="2.5"
            />
            {/* Tooltip */}
             <g className="opacity-0 group-hover:opacity-100 transition-opacity pointer-events-none z-50">
                <rect 
                    x={Math.min(scaleX(p.x) - 40, width - 85)} 
                    y={scaleY(p.y) - 40} 
                    width="80" 
                    height="28" 
                    rx="6"
                    className="fill-zinc-800 dark:fill-white shadow-lg" 
                />
                <text 
                    x={Math.min(scaleX(p.x), width - 45)} 
                    y={scaleY(p.y) - 22} 
                    textAnchor="middle" 
                    className="fill-white dark:fill-zinc-900 font-bold text-xs"
                >
                    {p.x}, {p.y}
                </text>
             </g>
          </g>
        ))}
        
        {/* Axis Titles */}
        <text x={(width - padding.left - padding.right) / 2 + padding.left} y={height - 5} textAnchor="middle" className="fill-zinc-500 dark:fill-zinc-400 font-bold text-sm">{t.concAxis}</text>
        <text x={15} y={(height - padding.top - padding.bottom) / 2 + padding.top} textAnchor="middle" transform={`rotate(-90, 15, ${(height - padding.top - padding.bottom) / 2 + padding.top})`} className="fill-zinc-500 dark:fill-zinc-400 font-bold text-sm">{t.odAxis}</text>
      </svg>
    </div>
  );
}

function App() {
  const [theme, setTheme] = useState<'dark' | 'light'>('dark');
  const [lang, setLang] = useState<'en' | 'zh'>('en'); // Language State
  const [view, setView] = useState<'home' | 'camera' | 'preview' | 'results'>('home');
  // Removed activeTab, we now show all columns side-by-side
  
  const [image, setImage] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [results, setResults] = useState<ExtractedItem[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [showSettings, setShowSettings] = useState(false);

  // Standard Curve State
  const [stdCurvePoints, setStdCurvePoints] = useState<StdCurvePoint[]>(
    DEFAULT_CONCS.map((conc, i) => ({ x: conc, y: '', id: `pt-${i}` }))
  );
  
  // Unknown Groups State
  const [unknownGroups, setUnknownGroups] = useState<UnknownGroup[]>([]);
  const [editingGroupId, setEditingGroupId] = useState<string | null>(null);

  // Selection Logic
  // 'target' determines where clicks go: 'std' (standard curve) or 'sample' (unknowns)
  const [selectionTarget, setSelectionTarget] = useState<'std' | 'sample'>('std');
  const [isStdSelecting, setIsStdSelecting] = useState(false);

  // Layout State
  const isDesktop = useMediaQuery('(min-width: 1024px)');
  
  const t = TRANSLATIONS[lang]; // Current translation

  // Effect to handle dark mode class on html/body
  useEffect(() => {
    if (theme === 'dark') {
        document.documentElement.classList.add('dark');
    } else {
        document.documentElement.classList.remove('dark');
    }
  }, [theme]);

  const toggleTheme = () => {
    setTheme(prev => prev === 'dark' ? 'light' : 'dark');
  };
  
  const toggleLang = () => {
      setLang(prev => prev === 'en' ? 'zh' : 'en');
  }

  const processImage = async () => {
    if (!image) return;
    
    // Check for API Key (LocalStorage first, then Env var)
    const storedKey = localStorage.getItem('gemini_api_key');
    const apiKey = storedKey || process.env.API_KEY;

    if (!apiKey) {
      setError(t.missingKeyMsg);
      setShowSettings(true);
      return;
    }

    setLoading(true);
    setError(null);
    try {
      const ai = new GoogleGenAI({ apiKey: apiKey });
      const base64Data = image.split(',')[1];
      const response = await ai.models.generateContent({
        model: 'gemini-2.0-flash',
        contents: {
          parts: [
            { inlineData: { mimeType: 'image/jpeg', data: base64Data } },
            {
              text: `
                You are an expert laboratory assistant reading a 96-well microplate reader data sheet.
                Your Task: Extract measurement values and map them to their specific Grid Coordinates (Row A-H, Column 1-12).
                CRITICAL RULES:
                1. **Grid Structure**: The data represents an 8-row (A-H) by 12-column (1-12) matrix.
                2. **Coordinates**: For every measurement found, you MUST identify its 'row' (A, B, C...) and 'col' (1, 2, 3...) based on its position.
                3. **Ignore Metadata**: Do NOT treat axis labels or indices as data. Only extract measurements (decimals).
                4. **Empty Wells**: If a well is empty, DO NOT include it.
                5. **Precision**: Maintain exact decimal places.
                Output: JSON Array of objects { value, row, col }.
              `
            }
          ]
        },
        config: {
          temperature: 0,
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                value: { type: Type.STRING },
                row: { type: Type.STRING },
                col: { type: Type.INTEGER }
              },
              required: ["value", "row", "col"]
            }
          }
        }
      });
      const jsonText = response.text || "[]";
      const data = JSON.parse(jsonText);
      const normalizedData = data.map((item: any) => ({
        value: item.value,
        row: item.row?.toUpperCase() || '?',
        col: parseInt(item.col) || 0
      }));
      setResults(normalizedData);
      setView('results');
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Failed to extract data.");
    } finally {
      setLoading(false);
    }
  };

  const handleImageSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const reader = new FileReader();
      reader.onload = (ev) => {
        if (ev.target?.result) {
          setImage(ev.target.result as string);
          setView('preview');
        }
      };
      reader.readAsDataURL(e.target.files[0]);
    }
  };

  const reset = () => {
    setImage(null);
    setResults([]);
    setError(null);
    setView('home');
    setStdCurvePoints(DEFAULT_CONCS.map((conc, i) => ({ x: conc, y: '', id: `pt-${i}` })));
    setUnknownGroups([]);
    setSelectionTarget('std');
    setIsStdSelecting(false);
    setEditingGroupId(null);
  };

  const getCellData = (row: string, col: number) => {
    return results.find(item => item.row === row && item.col === col);
  };

  const clearStdCurve = () => {
    setStdCurvePoints(prev => prev.map(p => ({ ...p, y: '', sourceWellId: undefined })));
    setIsStdSelecting(false);
  };

  const handleAddStdPoint = () => {
      setStdCurvePoints(prev => [...prev, { x: '', y: '', id: `manual-${Date.now()}` }]);
  };
  
  const handleDeleteStdPoint = (id: string) => {
      setStdCurvePoints(prev => prev.filter(p => p.id !== id));
  };

  const handleCellClick = (row: string, col: number) => {
    const cellId = `${row}-${col}`;
    const cellData = getCellData(row, col);
    // Even if no image data, allow selection for visual marking or manual entry placeholder logic
    const val = cellData ? cellData.value : "0"; 

    if (selectionTarget === 'std') {
        if (isStdSelecting) {
             const existingIdx = stdCurvePoints.findIndex(p => p.sourceWellId === cellId);
             if (existingIdx !== -1) {
                 // Deselect
                 setStdCurvePoints(prev => {
                     const next = [...prev];
                     next[existingIdx] = { ...next[existingIdx], y: '', sourceWellId: undefined };
                     return next;
                 });
             } else {
                 // Select next empty
                 const emptyIdx = stdCurvePoints.findIndex(p => !p.y);
                 if (emptyIdx !== -1) {
                     setStdCurvePoints(prev => {
                         const next = [...prev];
                         next[emptyIdx] = { ...next[emptyIdx], y: val, sourceWellId: cellId };
                         return next;
                     });
                 } else {
                     // Append
                     setStdCurvePoints(prev => [...prev, { x: '', y: val, id: `manual-${Date.now()}`, sourceWellId: cellId }]);
                 }
             }
        }
    } else {
        // Sample Mode
        if (editingGroupId) {
            // Check if part of standard curve to avoid conflict
            const isStd = stdCurvePoints.some(p => p.sourceWellId === cellId);
            if (isStd) return;

            setUnknownGroups(prevGroups => {
              return prevGroups.map(group => {
                if (group.id !== editingGroupId) {
                   // Remove from other groups
                   if (group.wells.some(w => w.row === row && w.col === col)) {
                       return { ...group, wells: group.wells.filter(w => w.row !== row || w.col !== col) };
                   }
                   return group;
                }
                // Toggle in current group
                const exists = group.wells.find(w => w.row === row && w.col === col);
                let newWells;
                if (exists) {
                  newWells = group.wells.filter(w => w.row !== row || w.col !== col);
                } else {
                  newWells = [...group.wells, { row, col, od: parseFloat(val), dilution: group.commonDilution }];
                }
                // Sort wells
                newWells.sort((a, b) => {
                   if (a.row !== b.row) return a.row.localeCompare(b.row);
                   return a.col - b.col;
                });
                return { ...group, wells: newWells };
              });
            });
        }
    }
  };

  const fitResult = useMemo(() => calculateLinearFit(stdCurvePoints), [stdCurvePoints]);
  
  const hasValidPoints = useMemo(() => {
      return stdCurvePoints.some(p => !isNaN(parseFloat(p.x)) && !isNaN(parseFloat(p.y)));
  }, [stdCurvePoints]);

  const addUnknownGroup = () => {
    setSelectionTarget('sample');
    const nextColorIdx = unknownGroups.length % GROUP_COLORS.length;
    const newGroup: UnknownGroup = {
      id: Date.now().toString(),
      name: `${t.sampleName} ${unknownGroups.length + 1}`,
      commonDilution: 1,
      color: GROUP_COLORS[nextColorIdx].name,
      wells: []
    };
    setUnknownGroups([...unknownGroups, newGroup]);
    setEditingGroupId(newGroup.id);
  };

  const handleManualEntry = () => {
      setView('results');
  };

  const handleExport = () => {
    setExporting(true);
    setTimeout(() => {
      try {
        const dateStr = new Date().toLocaleString();
        const wsData: (string | number | null)[][] = [];
        wsData.push([t.reportTitle, dateStr]);
        wsData.push([]);
        
        wsData.push([t.rawMatrix]);
        const colHeaders = ["", ...Array.from({ length: 12 }, (_, i) => (i + 1).toString())];
        wsData.push(colHeaders);
        ROW_HEADERS.forEach(row => {
          const rowData: (string | number)[] = [row];
          for (let i = 1; i <= 12; i++) {
            const cell = getCellData(row, i);
            rowData.push(cell ? parseFloat(cell.value) : "");
          }
          wsData.push(rowData);
        });

        wsData.push([]);
        wsData.push([t.stdCurve]);
        if (fitResult) {
            wsData.push([t.slope, fitResult.slope, t.intercept, fitResult.intercept, t.r2, fitResult.r2]);
        }
        wsData.push([t.conc, t.od]);
        stdCurvePoints.forEach(p => wsData.push([p.x, p.y]));

        wsData.push([]);
        wsData.push([t.sampleTable]);
        wsData.push([t.newGroup, "Well", t.dil, t.od, t.conc, t.mean, t.sd, t.cv]);
        unknownGroups.forEach(group => {
            const wellConcs = group.wells.map(w => calculateConc(w.od, fitResult, w.dilution));
            const stats = calculateStats(wellConcs);
            group.wells.forEach((well, idx) => {
                wsData.push([
                    group.name, `${well.row}${well.col}`, well.dilution, well.od, wellConcs[idx],
                    idx === 0 ? stats.mean : "", idx === 0 ? stats.sd : "", idx === 0 ? stats.cv : ""
                ]);
            });
        });

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        XLSX.utils.book_append_sheet(wb, ws, "Data");
        XLSX.writeFile(wb, `LabLens_Results_${Date.now()}.xlsx`);
      } catch (e) { console.error(e); } finally { setExporting(false); }
    }, 100);
  };

  // Helper component to render individual well button
  const renderPlateCell = (row: string, col: number) => {
     const cellId = `${row}-${col}`;
     const item = getCellData(row, col);
     
     // Determine states from data structures
     const stdIndex = stdCurvePoints.findIndex(p => p.sourceWellId === cellId);
     const isStd = stdIndex !== -1;
     
     const group = unknownGroups.find(g => g.wells.some(w => w.row === row && w.col === col));
     const gColor = group ? GROUP_COLORS.find(c => c.name === group.color) : null;
     
     const hasData = !!item;
     
     return (
         <button
             key={`${row}-${col}`}
             onClick={() => handleCellClick(row, col)}
             disabled={!item && selectionTarget !== 'std' && !editingGroupId}
             className={`
                 relative w-full aspect-square rounded-full flex items-center justify-center text-[11px] lg:text-xs font-mono leading-none transition-transform active:scale-90
                 ${!hasData ? 'bg-zinc-200/50 dark:bg-zinc-800/30 text-transparent' : 'bg-white dark:bg-zinc-800 text-zinc-500 shadow-sm border border-zinc-200 dark:border-zinc-700 hover:border-zinc-400 hover:scale-105'}
                 ${isStd ? '!bg-blue-600 !text-white !border-blue-600 shadow-blue-500/20' : ''}
                 ${gColor ? `!${gColor.bg} !text-white !border-transparent` : ''}
             `}
             title={`${row}${col}: ${item?.value}`}
         >
             {item?.value}
             {isStd && (
                <span className="absolute -top-1 -right-1 flex h-4 w-4 items-center justify-center rounded-full bg-blue-700 text-[9px] font-bold text-white shadow ring-2 ring-white dark:ring-zinc-900 z-10">
                  {stdIndex + 1}
                </span>
             )}
         </button>
     );
  };

  return (
    <div className={`h-[100dvh] w-full flex flex-col font-sans bg-gray-50 dark:bg-[#09090b] text-zinc-900 dark:text-zinc-100 transition-colors duration-300 overflow-hidden`}>
      <ApiKeyModal isOpen={showSettings} onClose={() => setShowSettings(false)} t={t} />

      {/* --- HEADER --- */}
      <header className="flex-none sticky top-0 z-40 bg-white/80 dark:bg-[#09090b]/80 backdrop-blur-md border-b border-zinc-200 dark:border-zinc-800 px-6 h-14 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="bg-zinc-900 dark:bg-white text-white dark:text-black p-1.5 rounded-lg shadow-sm">
            <FlaskConical size={18} strokeWidth={3} />
          </div>
          <h1 className="font-bold text-xl tracking-tight">{t.appTitle}</h1>
        </div>
        
        <div className="flex items-center gap-3">
           <button onClick={() => setShowSettings(true)} className="p-2 text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 transition-colors rounded-full hover:bg-zinc-100 dark:hover:bg-zinc-800" title={t.settings}>
             <Settings size={20}/>
           </button>
           <button onClick={toggleLang} className="p-2 flex items-center gap-1 text-base font-bold text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 transition-colors rounded-full hover:bg-zinc-100 dark:hover:bg-zinc-800">
             <Languages size={18}/> {lang.toUpperCase()}
           </button>
           <button onClick={toggleTheme} className="p-2 text-zinc-500 hover:text-zinc-900 dark:hover:text-zinc-100 transition-colors rounded-full hover:bg-zinc-100 dark:hover:bg-zinc-800">
             {theme === 'dark' ? <Sun size={20}/> : <Moon size={20}/>}
           </button>
           {view === 'results' && (
              <button 
                onClick={handleExport} disabled={exporting}
                className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg text-base font-semibold transition-all shadow-sm active:scale-95"
              >
                {exporting ? <Loader2 size={16} className="animate-spin"/> : <FileSpreadsheet size={16}/>}
                <span className="hidden sm:inline">{t.export}</span>
              </button>
           )}
           {view !== 'home' && (
             <button onClick={reset} className="p-2 text-zinc-400 hover:text-red-500 transition-colors">
               <X size={22} />
             </button>
           )}
        </div>
      </header>

      {/* --- MAIN CONTENT --- */}
      <main className="flex-1 overflow-hidden relative">
        
        {/* VIEW: HOME */}
        {view === 'home' && (
          <div className="h-full flex flex-col items-center justify-center p-6 gap-6 max-w-lg mx-auto w-full animate-in fade-in zoom-in duration-300">
            <div className="text-center space-y-2 mb-8">
              <h2 className="text-5xl font-extrabold tracking-tighter mb-2">{t.subtitle}</h2>
              <p className="text-zinc-500 dark:text-zinc-400 text-xl">{t.desc}</p>
            </div>

            <div className="grid w-full gap-4">
               <button 
                  onClick={() => setView('camera')}
                  className="group relative overflow-hidden bg-zinc-900 dark:bg-white text-white dark:text-black p-8 rounded-2xl flex items-center justify-between shadow-xl hover:shadow-2xl transition-all active:scale-[0.99]"
                >
                  <div className="flex flex-col items-start gap-1">
                    <span className="font-bold text-2xl">{t.scanPlate}</span>
                    <span className="text-zinc-400 dark:text-zinc-500 text-base font-medium">{t.useCamera}</span>
                  </div>
                  <Camera size={32} className="opacity-80 group-hover:scale-110 transition-transform"/>
               </button>
               
               <div className="relative group">
                  <input type="file" accept="image/*" onChange={handleImageSelect} className="absolute inset-0 opacity-0 cursor-pointer z-10"/>
                  <button className="w-full bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 text-zinc-900 dark:text-zinc-100 p-8 rounded-2xl flex items-center justify-between shadow-sm group-hover:bg-zinc-50 dark:group-hover:bg-zinc-800/50 transition-all active:scale-[0.99]">
                      <div className="flex flex-col items-start gap-1">
                        <span className="font-bold text-2xl">{t.uploadImage}</span>
                        <span className="text-zinc-400 dark:text-zinc-500 text-base font-medium">{t.fromGallery}</span>
                      </div>
                      <Upload size={32} className="opacity-50 group-hover:scale-110 transition-transform"/>
                  </button>
               </div>
               
               <button 
                  onClick={handleManualEntry}
                  className="w-full bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 text-zinc-900 dark:text-zinc-100 p-8 rounded-2xl flex items-center justify-between shadow-sm group-hover:bg-zinc-50 dark:group-hover:bg-zinc-800/50 transition-all active:scale-[0.99]"
                >
                  <div className="flex flex-col items-start gap-1">
                    <span className="font-bold text-2xl">{t.manualEntry}</span>
                    <span className="text-zinc-400 dark:text-zinc-500 text-base font-medium">{t.noImage}</span>
                  </div>
                  <Keyboard size={32} className="opacity-50 group-hover:scale-110 transition-transform"/>
               </button>
            </div>
          </div>
        )}

        {/* VIEW: CAMERA */}
        {view === 'camera' && (
           <CameraView onCapture={(img) => { setImage(img); setView('preview'); }} onBack={() => setView('home')} t={t} />
        )}

        {/* VIEW: PREVIEW */}
        {view === 'preview' && image && (
          <div className="h-full flex flex-col bg-zinc-100 dark:bg-[#050505]">
             <div className="flex-1 p-8 flex items-center justify-center relative min-h-0">
                <img src={image} className="max-w-full max-h-full rounded-lg shadow-2xl object-contain"/>
                {loading && (
                  <div className="absolute inset-0 bg-black/60 backdrop-blur-sm flex flex-col items-center justify-center text-white z-20">
                     <Loader2 size={48} className="animate-spin mb-4"/>
                     <p className="font-mono uppercase tracking-widest text-sm">{t.analyzing}</p>
                  </div>
                )}
             </div>
             {!loading && (
               <div className="bg-white dark:bg-[#09090b] p-6 border-t border-zinc-200 dark:border-zinc-800 flex justify-center gap-4 flex-none">
                  <button onClick={() => setView('home')} className="px-8 py-3 rounded-xl font-medium text-zinc-600 dark:text-zinc-400 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors">{t.discard}</button>
                  <button onClick={processImage} className="px-8 py-3 rounded-xl font-bold bg-blue-600 hover:bg-blue-700 text-white shadow-lg shadow-blue-500/20 transition-all flex items-center gap-2">
                    {t.analyze} <ChevronRight size={18}/>
                  </button>
               </div>
             )}
          </div>
        )}

        {/* VIEW: RESULTS DASHBOARD - 3 COLUMN GRID */}
        {view === 'results' && (
          <div className="h-full w-full grid grid-cols-1 lg:grid-cols-3 bg-zinc-50 dark:bg-[#050505] divide-y lg:divide-y-0 lg:divide-x divide-zinc-200 dark:divide-zinc-800 overflow-y-auto lg:overflow-hidden">
            
            {/* === COL 1: PLATE GRID === */}
            <div className="flex flex-col min-h-[500px] lg:h-full lg:min-h-0 overflow-hidden bg-zinc-100 dark:bg-[#0a0a0a] min-w-0">
               {/* Column Header */}
               <div className="flex-none h-14 px-4 border-b border-zinc-200 dark:border-zinc-800 bg-white dark:bg-[#09090b] flex justify-between items-center z-10">
                  <h3 className="font-bold text-base flex items-center gap-2">
                    <LayoutGrid size={18} className="text-zinc-500"/> {t.plateMap}
                  </h3>
               </div>

               {/* Responsive Grid Container */}
               <div className="flex-1 overflow-y-auto overflow-x-hidden p-2 lg:p-4 flex items-center justify-center">
                  <div className="w-full h-full max-w-[600px] flex flex-col items-center">
                      
                      {isDesktop ? (
                        // === DESKTOP LAYOUT (Rotated 90 deg / Transposed) ===
                        // 8 Columns (A-H), 12 Rows (1-12)
                        <div className="w-full max-w-[420px] flex flex-col">
                             {/* Transposed Header (A-H) */}
                             <div className="grid grid-cols-[auto_repeat(8,1fr)] gap-1 mb-1 pr-0.5">
                                <div className="w-5"></div>
                                {ROW_HEADERS.map((r, i) => (
                                    <div key={i} className="text-center text-xs font-bold text-zinc-400">{r}</div>
                                ))}
                             </div>
                             
                             {/* Transposed Body (Rows 1-12) */}
                             <div className="flex-1 grid grid-rows-12 gap-1">
                                {Array.from({ length: 12 }).map((_, colIndex) => {
                                    const colNum = colIndex + 1;
                                    return (
                                      <div key={colNum} className="grid grid-cols-[auto_repeat(8,1fr)] gap-1">
                                          {/* Row Label (1-12) */}
                                          <div className="w-5 flex items-center justify-center text-xs font-bold text-zinc-400">{colNum}</div>
                                          {/* Cells (A..H for this number) */}
                                          {ROW_HEADERS.map(rowChar => renderPlateCell(rowChar, colNum))}
                                      </div>
                                    )
                                })}
                             </div>
                        </div>
                      ) : (
                        // === MOBILE LAYOUT (Standard) ===
                        // 12 Columns (1-12), 8 Rows (A-H)
                        <div className="w-full max-w-[600px] flex flex-col">
                            {/* Grid Header Labels (1-12) */}
                            <div className="grid grid-cols-[auto_repeat(12,1fr)] gap-1 mb-1 pr-0.5">
                                <div className="w-4"></div>
                                {Array.from({ length: COL_COUNT }).map((_, i) => (
                                    <div key={i} className="text-center text-xs font-bold text-zinc-400">{i + 1}</div>
                                ))}
                            </div>

                            {/* Grid Rows */}
                            <div className="flex-1 grid grid-rows-8 gap-1">
                                {ROW_HEADERS.map((row) => (
                                    <div key={row} className="grid grid-cols-[auto_repeat(12,1fr)] gap-1">
                                        {/* Row Label (A-H) */}
                                        <div className="w-4 flex items-center justify-center text-xs font-bold text-zinc-400">{row}</div>
                                        {/* Cells */}
                                        {Array.from({ length: COL_COUNT }).map((_, colIndex) => renderPlateCell(row, colIndex + 1))}
                                    </div>
                                ))}
                            </div>
                        </div>
                      )}

                  </div>
               </div>
            </div>

            {/* === COL 2: STANDARD CURVE === */}
            <div className="flex flex-col min-h-[600px] lg:h-full lg:min-h-0 bg-white dark:bg-[#09090b] min-w-0 border-r border-zinc-200 dark:border-zinc-800">
               <div className="flex-none h-14 px-4 border-b border-zinc-200 dark:border-zinc-800 flex justify-between items-center bg-zinc-50/50 dark:bg-black/20">
                  <h3 className="font-bold text-base flex items-center gap-2 text-blue-600 dark:text-blue-400">
                     <TrendingUp size={18}/> {t.calibration}
                  </h3>
                  <button 
                     onClick={() => {
                        setSelectionTarget('std');
                        setIsStdSelecting(!isStdSelecting);
                     }}
                     className={`px-3 py-1.5 text-sm font-bold rounded-full border transition-all flex items-center gap-1.5
                        ${isStdSelecting 
                           ? 'bg-blue-100 border-blue-200 text-blue-700 animate-pulse' 
                           : 'bg-white border-zinc-200 text-zinc-600 hover:border-blue-300 hover:text-blue-600'}`}
                  >
                     <MousePointer2 size={14}/> {isStdSelecting ? t.selecting : t.selectRange}
                  </button>
               </div>

               {/* Chart Area */}
               <div className="flex-[1.2] p-4 min-h-[200px] border-b border-zinc-100 dark:border-zinc-800 relative bg-zinc-50/30 dark:bg-zinc-900/10">
                  {hasValidPoints ? (
                     <StandardCurveChart points={stdCurvePoints} slope={fitResult?.slope} intercept={fitResult?.intercept} t={t} />
                  ) : (
                     <div className="absolute inset-0 flex items-center justify-center text-zinc-400 text-sm">
                        {t.noDataSelected}
                     </div>
                  )}
                  
                  {/* Floating Stats Card - With Formula and Color-Coded R2 */}
                  <div className={`absolute top-4 right-4 bg-white/95 dark:bg-zinc-900/95 backdrop-blur border-l-4 rounded-r-xl p-3 shadow-sm flex flex-col gap-1 items-end pointer-events-none z-10 
                    ${fitResult ? (
                       getFitQuality(fitResult.r2) === 'good' ? 'border-l-emerald-500' :
                       getFitQuality(fitResult.r2) === 'warn' ? 'border-l-amber-500' : 'border-l-red-500'
                    ) : 'border-l-zinc-300'}
                  `}>
                      {fitResult ? (
                        <>
                           <div className="font-mono font-bold text-base text-zinc-700 dark:text-zinc-300 whitespace-nowrap">
                             y = {fitResult.slope.toFixed(4)}x {fitResult.intercept >= 0 ? '+' : '-'} {Math.abs(fitResult.intercept).toFixed(4)}
                           </div>
                           <div className="flex items-center gap-2 mt-1">
                                {getFitQuality(fitResult.r2) === 'good' && <CheckCircle size={16} className="text-emerald-500"/>}
                                {getFitQuality(fitResult.r2) === 'warn' && <AlertTriangle size={16} className="text-amber-500"/>}
                                {getFitQuality(fitResult.r2) === 'poor' && <XCircle size={16} className="text-red-500"/>}
                                <div className="font-mono text-sm text-zinc-500 dark:text-zinc-400">
                                    R² = <span className={`font-bold ${
                                        getFitQuality(fitResult.r2) === 'good' ? 'text-emerald-600 dark:text-emerald-400' :
                                        getFitQuality(fitResult.r2) === 'warn' ? 'text-amber-600 dark:text-amber-500' :
                                        'text-red-600 dark:text-red-400'
                                    }`}>{fitResult.r2.toFixed(4)}</span>
                                </div>
                           </div>
                           <div className={`text-xs font-bold uppercase tracking-wider ${
                                getFitQuality(fitResult.r2) === 'good' ? 'text-emerald-600/70' :
                                getFitQuality(fitResult.r2) === 'warn' ? 'text-amber-600/70' :
                                'text-red-600/70'
                           }`}>
                               {getFitQuality(fitResult.r2) === 'good' ? t.fitGood :
                                getFitQuality(fitResult.r2) === 'warn' ? t.fitWarn : t.fitPoor}
                           </div>
                        </>
                      ) : (
                        <div className="text-sm text-zinc-400 italic">No fit data</div>
                      )}
                  </div>
               </div>

               {/* Data Table Area */}
               <div className="flex-1 overflow-y-auto p-0">
                  <div className="grid grid-cols-[32px_1fr_1fr_32px] sticky top-0 bg-zinc-100 dark:bg-zinc-800 text-xs font-bold text-zinc-500 uppercase py-2 px-3 border-b border-zinc-200 dark:border-zinc-700">
                     <div className="text-center">#</div>
                     <div className="text-center">{t.conc} (X)</div>
                     <div className="text-center">{t.od} (Y)</div>
                     <div></div>
                  </div>
                  <div className="divide-y divide-zinc-100 dark:divide-zinc-800">
                     {stdCurvePoints.map((pt, i) => (
                        <div key={pt.id} className="grid grid-cols-[32px_1fr_1fr_32px] py-2 px-3 items-center hover:bg-zinc-50 dark:hover:bg-zinc-800/50 group">
                           <div className="text-center text-sm font-mono text-zinc-400">{i+1}</div>
                           <input 
                              className="w-full text-center bg-transparent text-sm font-mono py-1 focus:bg-blue-50 dark:focus:bg-blue-900/20 rounded outline-none"
                              value={pt.x}
                              onChange={(e) => {
                                 const next = [...stdCurvePoints];
                                 next[i].x = e.target.value;
                                 setStdCurvePoints(next);
                              }}
                           />
                           <input
                              className="w-full text-center bg-transparent text-sm font-mono py-1 font-bold text-blue-600 dark:text-blue-400 focus:bg-blue-50 dark:focus:bg-blue-900/20 rounded outline-none"
                              value={pt.y}
                              placeholder="-"
                              onChange={(e) => {
                                 const next = [...stdCurvePoints];
                                 next[i].y = e.target.value;
                                 setStdCurvePoints(next);
                              }}
                           />
                           <div className="flex justify-center">
                              <button onClick={() => handleDeleteStdPoint(pt.id)} className="text-zinc-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity">
                                 <X size={16}/>
                              </button>
                           </div>
                        </div>
                     ))}
                  </div>
                  <div className="p-3 border-t border-zinc-100 dark:border-zinc-800 flex gap-2">
                     <button onClick={handleAddStdPoint} className="flex-1 py-2 rounded-lg border border-dashed border-zinc-300 dark:border-zinc-700 text-zinc-500 hover:bg-zinc-50 dark:hover:bg-zinc-800 text-sm font-bold flex items-center justify-center gap-2">
                        <Plus size={16}/> {t.addRow}
                     </button>
                     <button onClick={clearStdCurve} className="px-3 py-2 rounded-lg text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors" title="Clear All">
                        <Eraser size={18}/>
                     </button>
                  </div>
               </div>
            </div>

            {/* === COL 3: SAMPLES === */}
            <div className="flex flex-col min-h-[400px] lg:h-full lg:min-h-0 bg-zinc-50 dark:bg-[#050505] min-w-0">
               <div className="flex-none h-14 px-4 border-b border-zinc-200 dark:border-zinc-800 bg-white dark:bg-[#09090b] flex justify-between items-center">
                  <h3 className="font-bold text-base flex items-center gap-2 text-zinc-900 dark:text-zinc-100">
                     <TestTube2 size={18}/> {t.unknowns}
                  </h3>
                  <button 
                     onClick={addUnknownGroup}
                     className="px-4 py-2 bg-zinc-900 dark:bg-white text-white dark:text-black rounded-lg text-sm font-bold flex items-center gap-1.5 hover:opacity-90 active:scale-95 transition-all"
                  >
                     <Plus size={14}/> {t.newGroup}
                  </button>
               </div>

               <div className="flex-1 overflow-y-auto p-4 space-y-4">
                  {unknownGroups.length === 0 ? (
                     <div className="flex flex-col items-center justify-center h-40 text-zinc-400">
                        <TestTube2 size={40} className="mb-2 opacity-50"/>
                        <p className="text-sm">{t.noGroups}</p>
                     </div>
                  ) : (
                     unknownGroups.map((group) => {
                        const color = GROUP_COLORS.find(c => c.name === group.color)!;
                        const isEditing = editingGroupId === group.id;
                        const wellConcs = group.wells.map(w => calculateConc(w.od, fitResult, w.dilution));
                        const stats = calculateStats(wellConcs);

                        return (
                           <div key={group.id} 
                              className={`bg-white dark:bg-[#0a0a0a] rounded-xl border transition-all shadow-sm ${isEditing ? 'border-blue-500 ring-1 ring-blue-500' : 'border-zinc-200 dark:border-zinc-800'}`}
                              onClick={() => { if(!isEditing) { setEditingGroupId(group.id); setSelectionTarget('sample'); } }}
                           >
                              <div className="p-4">
                                 {/* Header */}
                                 <div className="flex items-center gap-2 mb-4">
                                    <div className={`w-3 h-3 rounded-full ${color.bg}`}></div>
                                    <input 
                                       value={group.name}
                                       onChange={(e) => setUnknownGroups(prev => prev.map(g => g.id === group.id ? {...g, name: e.target.value} : g))}
                                       className="font-bold text-base bg-transparent outline-none flex-1 min-w-0"
                                       placeholder={t.sampleName}
                                    />
                                    <button 
                                       onClick={(e) => { e.stopPropagation(); setUnknownGroups(prev => prev.filter(g => g.id !== group.id)); }}
                                       className="text-zinc-300 hover:text-red-500"
                                    >
                                       <X size={16}/>
                                    </button>
                                 </div>

                                 {/* Stats Bar */}
                                 {group.wells.length > 0 && (
                                    <div className="grid grid-cols-3 gap-px bg-zinc-100 dark:bg-zinc-800 rounded-lg overflow-hidden border border-zinc-100 dark:border-zinc-800 mb-4">
                                       <div className="bg-white dark:bg-black/20 p-2 text-center">
                                          <div className="text-[11px] text-zinc-400 uppercase">{t.mean}</div>
                                          <div className={`text-sm font-mono font-bold ${color.text}`}>{stats.mean.toFixed(2)}</div>
                                       </div>
                                       <div className="bg-white dark:bg-black/20 p-2 text-center">
                                          <div className="text-[11px] text-zinc-400 uppercase">{t.sd}</div>
                                          <div className="text-sm font-mono font-bold text-zinc-600 dark:text-zinc-400">{stats.sd.toFixed(2)}</div>
                                       </div>
                                       <div className="bg-white dark:bg-black/20 p-2 text-center">
                                          <div className="text-[11px] text-zinc-400 uppercase">{t.cv}</div>
                                          <div className="text-sm font-mono font-bold text-zinc-600 dark:text-zinc-400">{stats.cv.toFixed(1)}</div>
                                       </div>
                                    </div>
                                 )}

                                 {/* Details (Expanded) */}
                                 {isEditing && (
                                    <div className="space-y-3">
                                       <div className="flex items-center justify-between text-sm text-zinc-500 px-1">
                                          <span>{t.setAllDil}:</span>
                                          <input 
                                             type="number" 
                                             className="w-16 text-right bg-zinc-50 dark:bg-zinc-800 rounded px-2 py-1 outline-none font-mono"
                                             value={group.commonDilution}
                                             onChange={(e) => {
                                                const val = parseFloat(e.target.value) || 1;
                                                setUnknownGroups(prev => prev.map(g => g.id === group.id ? { ...g, commonDilution: val, wells: g.wells.map(w => ({ ...w, dilution: val })) } : g));
                                             }}
                                          />
                                       </div>
                                       
                                       <div className="border-t border-zinc-100 dark:border-zinc-800 pt-2">
                                          {group.wells.map((w, idx) => (
                                             <div key={idx} className="flex items-center gap-2 text-sm py-1.5 group/row">
                                                <div className="w-8 shrink-0 font-mono font-bold text-zinc-400">{w.row}{w.col}</div>
                                                
                                                {/* Dilution */}
                                                <div className="flex-1 min-w-0 bg-zinc-50 dark:bg-zinc-800/50 rounded px-2 py-1 flex items-center justify-between border border-transparent hover:border-zinc-200 dark:hover:border-zinc-700 transition-colors">
                                                   <span className="text-[11px] text-zinc-400 uppercase mr-1 shrink-0">{t.dil}</span>
                                                   <input
                                                      type="number"
                                                      value={w.dilution}
                                                      onChange={(e) => {
                                                          const val = parseFloat(e.target.value) || 1;
                                                          setUnknownGroups(prev => prev.map(g => g.id === group.id ? { ...g, wells: g.wells.map((ww, ii) => ii === idx ? { ...ww, dilution: val } : ww) } : g));
                                                      }}
                                                      className="w-full text-right bg-transparent outline-none font-mono text-sm"
                                                   />
                                                </div>

                                                {/* OD */}
                                                <div className="flex-1 min-w-0 bg-zinc-50 dark:bg-zinc-800/50 rounded px-2 py-1 flex items-center justify-between border border-transparent hover:border-zinc-200 dark:hover:border-zinc-700 transition-colors">
                                                   <span className="text-[11px] text-zinc-400 uppercase mr-1 shrink-0">{t.od}</span>
                                                   <input 
                                                      type="number"
                                                      value={w.od}
                                                      onChange={(e) => {
                                                         const val = parseFloat(e.target.value);
                                                         setUnknownGroups(prev => prev.map(g => g.id === group.id ? { ...g, wells: g.wells.map((ww, ii) => ii === idx ? { ...ww, od: val } : ww) } : g));
                                                      }}
                                                      className="w-full text-right bg-transparent outline-none font-mono text-sm"
                                                   />
                                                </div>

                                                {/* Result */}
                                                <div className="flex-1 min-w-0 bg-blue-50/50 dark:bg-blue-900/10 rounded px-2 py-1 flex items-center justify-between border border-blue-100 dark:border-blue-900/30">
                                                   <span className="text-[11px] text-blue-400/70 uppercase mr-1 shrink-0">{t.conc}</span>
                                                   <span className="font-mono font-bold text-sm text-blue-700 dark:text-blue-300 truncate">
                                                      {wellConcs[idx].toFixed(2)}
                                                   </span>
                                                </div>

                                                <button 
                                                   onClick={() => setUnknownGroups(prev => prev.map(g => g.id === group.id ? { ...g, wells: g.wells.filter((_, i) => i !== idx) } : g))}
                                                   className="shrink-0 text-zinc-300 hover:text-red-500 opacity-0 group-hover/row:opacity-100 w-5 flex justify-center"
                                                >
                                                   <X size={16}/>
                                                </button>
                                             </div>
                                          ))}
                                          <button 
                                             onClick={() => setUnknownGroups(prev => prev.map(g => g.id === group.id ? { ...g, wells: [...g.wells, { row: '?', col: 0, od: 0, dilution: g.commonDilution }] } : g))}
                                             className="w-full mt-2 py-1.5 text-xs font-bold uppercase text-zinc-400 hover:text-blue-500 border border-dashed border-zinc-200 dark:border-zinc-700 rounded hover:bg-blue-50 dark:hover:bg-blue-900/10 transition-colors"
                                          >
                                             {t.addManualWell}
                                          </button>
                                       </div>
                                    </div>
                                 )}
                              </div>
                           </div>
                        );
                     })
                  )}
               </div>
            </div>

          </div>
        )}
      </main>
    </div>
  );
}

const container = document.getElementById('root');
if (container) {
  const root = createRoot(container);
  root.render(<App />);
}