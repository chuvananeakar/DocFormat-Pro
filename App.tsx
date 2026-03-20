import React, { useState } from 'react';
import { Dropzone } from './components/Dropzone';
import { ProcessingLog } from './components/ProcessingLog';
import { processDocx } from './services/docxService';
import { DocumentPreview } from './components/DocumentPreview';
import { ProcessingStatus, ProcessResult, DocxOptions, HeaderType } from './types';
import { 
  FileText, Download, RefreshCw, Sparkles, 
  FileCheck, ShieldCheck, Cpu, LayoutTemplate, 
  Settings2, Zap, ArrowRight, Database, SlidersHorizontal, ChevronDown, ChevronUp, CheckSquare, ListX
} from 'lucide-react';

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<ProcessingStatus>(ProcessingStatus.IDLE);
  const [result, setResult] = useState<ProcessResult | null>(null);
  const [showSettings, setShowSettings] = useState(false);
  
  const [options, setOptions] = useState<DocxOptions>({
    headerType: HeaderType.NONE,
    removeNumbering: false,
    margins: { top: 2, bottom: 2, left: 3, right: 1.5 },
    font: { family: "Times New Roman", sizeNormal: 14, sizeTable: 13 },
    paragraph: { lineSpacing: 1.15, after: 6, indent: 1.27 },
    table: { rowHeight: 0.8 }
  });

  const handleFileSelect = (selectedFile: File) => {
    setFile(selectedFile);
    setResult(null);
    setStatus(ProcessingStatus.IDLE);
  };

  const handleProcess = async () => {
    if (!file) return;

    setStatus(ProcessingStatus.PROCESSING);
    setResult({ success: false, logs: ["Khởi tạo hệ thống xử lý...", "Đang phân tích cấu trúc DOCX..."] });

    // Small delay to allow UI to update
    setTimeout(async () => {
      const res = await processDocx(file, options);
      setResult(res);
      setStatus(res.success ? ProcessingStatus.SUCCESS : ProcessingStatus.ERROR);
    }, 500);
  };

  const handleDownload = () => {
    if (result?.blob && result.fileName) {
      const url = URL.createObjectURL(result.blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = result.fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  };

  const handleReset = () => {
    setFile(null);
    setStatus(ProcessingStatus.IDLE);
    setResult(null);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans flex flex-col relative overflow-hidden selection:bg-blue-100 selection:text-blue-900">
      
      {/* Abstract Tech Background Elements */}
      <div className="absolute top-0 left-0 w-full h-96 bg-gradient-to-b from-blue-50 to-transparent pointer-events-none z-0"></div>
      <div className="absolute top-[-10%] right-[-5%] w-[500px] h-[500px] bg-blue-100/40 rounded-full blur-3xl pointer-events-none z-0"></div>
      <div className="absolute bottom-[-10%] left-[-10%] w-[600px] h-[600px] bg-indigo-100/40 rounded-full blur-3xl pointer-events-none z-0"></div>

      {/* School Top Bar */}
      <div className="relative z-20 bg-gradient-to-r from-slate-900 via-blue-900 to-slate-900 text-white text-center py-2.5 shadow-md">
        <div className="flex items-center justify-center gap-2 text-xs font-bold tracking-widest uppercase">
          <span className="w-1.5 h-1.5 bg-cyan-400 rounded-full animate-pulse"></span>
          TRƯỜNG THCS CHU VĂN AN
          <span className="w-1.5 h-1.5 bg-cyan-400 rounded-full animate-pulse"></span>
        </div>
      </div>

      {/* Header */}
      <header className="relative z-10 bg-white/80 backdrop-blur-md border-b border-white/50 sticky top-0 shadow-sm">
        <div className="max-w-5xl mx-auto px-6 h-18 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3 group cursor-pointer">
            <div className="bg-gradient-to-br from-blue-600 to-indigo-600 p-2 rounded-lg text-white shadow-lg shadow-blue-500/20 transition-transform group-hover:scale-105">
              <Cpu className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-800 tracking-tight leading-none group-hover:text-blue-700 transition-colors">
                DocFormat <span className="bg-clip-text text-transparent bg-gradient-to-r from-blue-600 to-indigo-600">Pro</span>
              </h1>
              <p className="text-[10px] font-semibold text-slate-400 uppercase tracking-wider mt-0.5">Automation System</p>
            </div>
          </div>
          <div className="hidden sm:flex items-center gap-2 text-xs font-semibold text-blue-800 bg-blue-50/80 px-4 py-2 rounded-full border border-blue-100 shadow-sm">
            <ShieldCheck className="w-4 h-4 text-blue-600" />
            Bảo mật Client-Side
          </div>
        </div>
      </header>

      <main className="relative z-10 max-w-4xl mx-auto px-6 py-10 flex-grow w-full">
        
        {/* Intro */}
        <div className="text-center mb-8 space-y-4">
          <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-indigo-50 border border-indigo-100 text-indigo-700 text-[10px] font-bold uppercase tracking-wider mb-2">
            <Sparkles className="w-3 h-3" /> Phiên bản chuyển đổi số 2.0
          </div>
          <h2 className="text-2xl sm:text-3xl md:text-4xl font-black text-slate-800 tracking-tight uppercase">
            <span className="bg-clip-text text-transparent bg-gradient-to-r from-blue-700 via-indigo-700 to-blue-700">CÔNG CỤ CHUẨN HÓA VĂN BẢN HÀNH CHÍNH</span>
          </h2>
          <p className="text-base sm:text-lg text-slate-500 max-w-3xl mx-auto font-medium leading-relaxed">
            Chuẩn hóa tài liệu của bạn ngay lập tức. Tuân thủ các quy tắc định dạng hành chính.
          </p>
        </div>

        {/* Configuration Section */}
        <div className="mb-8">
            <button 
                onClick={() => setShowSettings(!showSettings)}
                className="flex items-center gap-2 text-sm font-bold text-orange-700 bg-orange-50 border-2 border-orange-300 px-6 py-2.5 rounded-full shadow-md hover:bg-orange-100 hover:shadow-lg hover:scale-105 transition-all mx-auto"
            >
                <SlidersHorizontal className="w-4 h-4" />
                <span>Tùy chỉnh thông số</span>
                {showSettings ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
            </button>
            
            {showSettings && (
                <div className="mt-4 bg-white p-6 rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-100 animate-fadeIn">
                    <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-4 flex items-center gap-2">
                        <Settings2 className="w-4 h-4 text-blue-600" /> Cấu hình định dạng
                    </h3>
                    
                    {/* Header Template Option */}
                    <div className="mb-4 pb-4 border-b border-slate-100">
                        <div className="flex flex-col gap-2">
                            <div className="flex items-center gap-2">
                                <LayoutTemplate className="w-4 h-4 text-blue-600" />
                                <span className="text-sm font-bold text-slate-700">Chèn khung Quốc hiệu/Tiêu ngữ mẫu</span>
                            </div>
                            <select 
                                value={options.headerType}
                                onChange={(e) => setOptions({...options, headerType: e.target.value as HeaderType})}
                                className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer text-slate-700"
                            >
                                <option value={HeaderType.NONE}>❌ Không chèn thêm</option>
                                <option value={HeaderType.SCHOOL}>🏫 Văn bản Hành chính (Nhà trường)</option>
                                <option value={HeaderType.PARTY}>⭐ Văn bản Công tác Đảng (Chi bộ)</option>
                                <option value={HeaderType.DEPARTMENT}>📚 Văn bản Nội bộ (Tổ chuyên môn)</option>
                            </select>
                            <p className="text-[10px] text-slate-400">Tự động chèn bảng thông tin cơ quan và Quốc hiệu vào đầu trang theo mẫu đã chọn</p>
                        </div>
                    </div>

                    {/* Remove Numbering Option */}
                    <div className="mb-6 pb-6 border-b border-slate-100">
                        <label className="flex items-center gap-3 cursor-pointer group">
                            <div className={`w-5 h-5 rounded border flex items-center justify-center transition-colors ${options.removeNumbering ? 'bg-orange-500 border-orange-500 text-white' : 'bg-white border-slate-300 group-hover:border-orange-400'}`}>
                                {options.removeNumbering && <ListX className="w-3.5 h-3.5" />}
                            </div>
                            <input 
                                type="checkbox" 
                                className="hidden"
                                checked={options.removeNumbering}
                                onChange={(e) => setOptions({...options, removeNumbering: e.target.checked})}
                            />
                            <div>
                                <span className="text-sm font-bold text-slate-700 group-hover:text-orange-600 transition-colors">Xóa định dạng danh sách tự động (Bullets/Numbering)</span>
                                <p className="text-xs text-slate-400 mt-0.5">Chuyển đổi danh sách 1., 2., • thành văn bản thường để đồng bộ định dạng</p>
                            </div>
                        </label>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        {/* Margins */}
                        <div className="space-y-3">
                            <label className="text-xs font-semibold text-slate-500 uppercase">Lề Trang (cm)</label>
                            <div className="grid grid-cols-2 gap-3">
                                <div>
                                    <span className="text-xs text-slate-400">Trên (Top)</span>
                                    <input 
                                        type="number" step="0.1"
                                        value={options.margins.top}
                                        onChange={(e) => setOptions({...options, margins: {...options.margins, top: parseFloat(e.target.value)}})}
                                        className="w-full mt-1 px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                    />
                                </div>
                                <div>
                                    <span className="text-xs text-slate-400">Dưới (Bottom)</span>
                                    <input 
                                        type="number" step="0.1"
                                        value={options.margins.bottom}
                                        onChange={(e) => setOptions({...options, margins: {...options.margins, bottom: parseFloat(e.target.value)}})}
                                        className="w-full mt-1 px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                    />
                                </div>
                                <div>
                                    <span className="text-xs text-slate-400">Trái (Left)</span>
                                    <input 
                                        type="number" step="0.1"
                                        value={options.margins.left}
                                        onChange={(e) => setOptions({...options, margins: {...options.margins, left: parseFloat(e.target.value)}})}
                                        className="w-full mt-1 px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                    />
                                </div>
                                <div>
                                    <span className="text-xs text-slate-400">Phải (Right)</span>
                                    <input 
                                        type="number" step="0.1"
                                        value={options.margins.right}
                                        onChange={(e) => setOptions({...options, margins: {...options.margins, right: parseFloat(e.target.value)}})}
                                        className="w-full mt-1 px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                    />
                                </div>
                            </div>
                        </div>

                        {/* Font & Paragraph */}
                        <div className="space-y-3">
                            <label className="text-xs font-semibold text-slate-500 uppercase">Font & Đoạn văn</label>
                            <div className="grid grid-cols-2 gap-3">
                                <div>
                                    <span className="text-xs text-slate-400">Cỡ chữ Thường (pt)</span>
                                    <input 
                                        type="number"
                                        value={options.font.sizeNormal}
                                        onChange={(e) => setOptions({...options, font: {...options.font, sizeNormal: parseInt(e.target.value)}})}
                                        className="w-full mt-1 px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                    />
                                </div>
                                <div>
                                    <span className="text-xs text-slate-400">Cỡ chữ Bảng (pt)</span>
                                    <input 
                                        type="number"
                                        value={options.font.sizeTable}
                                        onChange={(e) => setOptions({...options, font: {...options.font, sizeTable: parseInt(e.target.value)}})}
                                        className="w-full mt-1 px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                    />
                                </div>
                                <div>
                                    <span className="text-xs text-slate-400">Dãn dòng (Lines)</span>
                                    <input 
                                        type="number" step="0.05"
                                        value={options.paragraph.lineSpacing}
                                        onChange={(e) => setOptions({...options, paragraph: {...options.paragraph, lineSpacing: parseFloat(e.target.value)}})}
                                        className="w-full mt-1 px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                    />
                                </div>
                                <div>
                                    <span className="text-xs text-slate-400">Thụt đầu dòng (cm)</span>
                                    <input 
                                        type="number" step="0.01"
                                        value={options.paragraph.indent}
                                        onChange={(e) => setOptions({...options, paragraph: {...options.paragraph, indent: parseFloat(e.target.value)}})}
                                        className="w-full mt-1 px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                    />
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            )}
        </div>

        {/* Main Card */}
        <div className="bg-white rounded-2xl shadow-xl shadow-slate-200/50 border border-slate-100 overflow-hidden relative group">
          <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-cyan-400 via-blue-500 to-indigo-500"></div>

          <div className="p-8 space-y-8">
            {/* Step 1: Upload */}
            {!file && (
              <div className="space-y-5 animate-fadeIn">
                <div className="flex items-center gap-3 text-sm font-bold text-slate-400 uppercase tracking-wider">
                  <div className="w-8 h-8 rounded-lg bg-slate-100 text-slate-600 flex items-center justify-center text-xs border border-slate-200 shadow-sm">01</div>
                  <span className="text-slate-500">Tải lên tài liệu</span>
                  <div className="h-px bg-slate-100 flex-grow"></div>
                </div>
                <Dropzone onFileSelect={handleFileSelect} />
              </div>
            )}

            {/* Step 2: File Selected & Actions */}
            {file && status !== ProcessingStatus.SUCCESS && (
              <div className="space-y-6 animate-fadeIn">
                <div className="flex items-center justify-between p-5 bg-gradient-to-br from-slate-50 to-white border border-slate-200 rounded-xl shadow-sm relative overflow-hidden">
                  <div className="absolute left-0 top-0 bottom-0 w-1 bg-blue-500"></div>
                  <div className="flex items-center gap-5 relative z-10">
                    <div className="bg-blue-100/50 p-3 rounded-xl">
                      <FileText className="w-8 h-8 text-blue-600" />
                    </div>
                    <div>
                      <p className="font-bold text-slate-800 text-lg">{file.name}</p>
                      <div className="flex items-center gap-2 mt-1">
                         <span className="text-xs font-mono text-slate-500 bg-slate-100 px-2 py-0.5 rounded">DOCX</span>
                         <span className="text-xs text-slate-400 font-mono">{(file.size / 1024).toFixed(1)} KB</span>
                      </div>
                    </div>
                  </div>
                  <button 
                    onClick={handleReset}
                    className="p-2.5 rounded-lg text-slate-400 hover:bg-red-50 hover:text-red-500 transition-all"
                    title="Hủy bỏ"
                  >
                    <RefreshCw className="w-5 h-5" />
                  </button>
                </div>

                {status === ProcessingStatus.IDLE && (
                  <button
                    onClick={handleProcess}
                    className="w-full py-4 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 text-white rounded-xl font-bold text-lg shadow-lg shadow-blue-500/30 hover:shadow-blue-500/40 transform hover:-translate-y-0.5 transition-all duration-200 flex items-center justify-center gap-3 group"
                  >
                    <Zap className="w-5 h-5 group-hover:text-yellow-300 transition-colors" />
                    <span>Thực hiện Chuẩn hóa ngay</span>
                    <ArrowRight className="w-5 h-5 opacity-70 group-hover:translate-x-1 transition-transform" />
                  </button>
                )}

                {status === ProcessingStatus.PROCESSING && (
                  <button
                    disabled
                    className="w-full py-4 bg-slate-50 text-slate-500 rounded-xl font-semibold text-lg flex items-center justify-center gap-3 cursor-wait border border-slate-200"
                  >
                    <div className="relative">
                       <div className="w-5 h-5 border-2 border-blue-200 border-t-blue-600 rounded-full animate-spin"></div>
                    </div>
                    <span>Đang xử lý dữ liệu...</span>
                  </button>
                )}
              </div>
            )}

            {/* Step 3: Success */}
            {status === ProcessingStatus.SUCCESS && result && (
              <div className="text-center space-y-8 animate-fadeIn py-2">
                 <div className="relative w-24 h-24 mx-auto">
                    <div className="absolute inset-0 bg-green-100 rounded-full animate-ping opacity-20"></div>
                    <div className="relative w-24 h-24 bg-gradient-to-br from-green-50 to-white text-green-600 rounded-full flex items-center justify-center border-4 border-green-50 shadow-inner">
                      <FileCheck className="w-10 h-10 drop-shadow-sm" />
                    </div>
                 </div>
                 
                 <div>
                    <h3 className="text-2xl font-bold text-slate-800 mb-2">Xử lý thành công!</h3>
                    <p className="text-slate-500">Tài liệu đã được chuẩn hóa theo quy định.</p>
                 </div>

                 <div className="flex flex-col sm:flex-row gap-4 justify-center">
                    <button
                      onClick={handleDownload}
                      className="px-8 py-3.5 bg-gradient-to-r from-green-600 to-emerald-600 hover:from-green-700 hover:to-emerald-700 text-white rounded-xl font-bold shadow-lg shadow-green-500/30 flex items-center justify-center gap-2 transform hover:-translate-y-0.5 transition-all"
                    >
                      <Download className="w-5 h-5" />
                      Tải về máy
                    </button>
                    <button
                      onClick={handleReset}
                      className="px-8 py-3.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 hover:border-blue-300 rounded-xl font-semibold transition-all shadow-sm flex items-center justify-center gap-2"
                    >
                      <RefreshCw className="w-4 h-4" />
                      Làm tiếp
                    </button>
                 </div>

                 {/* NEW: Document Preview Injection */}
                 <DocumentPreview blob={result.blob} />
              </div>
            )}

            {/* Logs Area */}
            {result?.logs && result.logs.length > 0 && (
              <div className="pt-6 border-t border-slate-100">
                <ProcessingLog logs={result.logs} />
              </div>
            )}

            {/* Error Display */}
            {status === ProcessingStatus.ERROR && (
              <div className="p-4 bg-red-50 text-red-800 rounded-xl border border-red-200 text-center flex flex-col items-center">
                <p className="font-bold">Đã xảy ra lỗi</p>
                <p className="text-sm opacity-90 mt-1">{result?.error}</p>
                <button onClick={handleReset} className="mt-3 text-sm font-semibold hover:underline text-red-700">Thử lại</button>
              </div>
            )}

          </div>
        </div>

        {/* Footer Info Grid - Tech Cards */}
        <div className="mt-8 grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 hover:shadow-md hover:border-blue-200 transition-all group">
            <div className="w-10 h-10 bg-blue-50 rounded-lg flex items-center justify-center text-blue-600 mb-4 group-hover:bg-blue-600 group-hover:text-white transition-colors">
              <LayoutTemplate className="w-5 h-5" />
            </div>
            <h4 className="font-bold text-slate-800 mb-3 text-sm uppercase tracking-wider">Khổ giấy & Lề</h4>
            <ul className="space-y-2 text-sm text-slate-500">
              <li className="flex justify-between border-b border-slate-50 pb-1"><span>Size:</span> <span className="font-semibold text-slate-700">A4 Portrait</span></li>
              <li className="flex justify-between border-b border-slate-50 pb-1"><span>Margins:</span> <span className="font-semibold text-slate-700">{options.margins.top}, {options.margins.bottom}, {options.margins.left}, {options.margins.right} cm</span></li>
            </ul>
          </div>

          <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 hover:shadow-md hover:border-blue-200 transition-all group">
            <div className="w-10 h-10 bg-indigo-50 rounded-lg flex items-center justify-center text-indigo-600 mb-4 group-hover:bg-indigo-600 group-hover:text-white transition-colors">
              <Settings2 className="w-5 h-5" />
            </div>
            <h4 className="font-bold text-slate-800 mb-3 text-sm uppercase tracking-wider">Đoạn văn & Font</h4>
            <ul className="space-y-2 text-sm text-slate-500">
              <li className="flex justify-between border-b border-slate-50 pb-1"><span>Font:</span> <span className="font-semibold text-slate-700">{options.font.family}</span></li>
              <li className="flex justify-between border-b border-slate-50 pb-1"><span>Size:</span> <span className="font-semibold text-slate-700">{options.font.sizeNormal}pt / Justify</span></li>
              <li className="flex justify-between border-b border-slate-50 pb-1"><span>Line:</span> <span className="font-semibold text-slate-700">{options.paragraph.lineSpacing} / Indent {options.paragraph.indent}</span></li>
            </ul>
          </div>

          <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 hover:shadow-md hover:border-blue-200 transition-all group">
            <div className="w-10 h-10 bg-cyan-50 rounded-lg flex items-center justify-center text-cyan-600 mb-4 group-hover:bg-cyan-600 group-hover:text-white transition-colors">
              <Database className="w-5 h-5" />
            </div>
            <h4 className="font-bold text-slate-800 mb-3 text-sm uppercase tracking-wider">Bảng biểu</h4>
            <ul className="space-y-2 text-sm text-slate-500">
              <li className="flex justify-between border-b border-slate-50 pb-1"><span>Row Height:</span> <span className="font-semibold text-slate-700">Tối thiểu {options.table.rowHeight} cm</span></li>
              <li className="flex justify-between border-b border-slate-50 pb-1"><span>Align:</span> <span className="font-semibold text-slate-700">Center</span></li>
              <li className="flex justify-between border-b border-slate-50 pb-1"><span>Size:</span> <span className="font-semibold text-slate-700">{options.font.sizeTable}pt</span></li>
            </ul>
          </div>
        </div>

      </main>
      
      {/* Design Credit Footer */}
      <footer className="py-6 text-center relative z-10 bg-white/50 border-t border-slate-200 backdrop-blur-sm">
         <p className="text-slate-500 text-sm font-medium">
           <span className="opacity-70">Version: 1.1-2026 &bull; Design by</span> <span className="text-blue-700 font-bold">Lai Cao Dang</span>
         </p>
      </footer>
    </div>
  );
}