import React, { useCallback, useState } from 'react';
import { Upload, AlertCircle, FilePlus } from 'lucide-react';

interface DropzoneProps {
  onFileSelect: (file: File) => void;
  disabled?: boolean;
}

export const Dropzone: React.FC<DropzoneProps> = ({ onFileSelect, disabled }) => {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const validateAndSelect = (file: File | undefined) => {
    setError(null);
    if (!file) return;

    if (!file.name.endsWith('.docx')) {
      setError("Vui lòng tải lên định dạng .docx");
      return;
    }

    onFileSelect(file);
  };

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    if (!disabled) setIsDragging(true);
  }, [disabled]);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (disabled) return;

    const files = Array.from(e.dataTransfer.files) as File[];
    validateAndSelect(files[0]);
  }, [disabled]);

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files ? Array.from(e.target.files) as File[] : [];
    validateAndSelect(files[0]);
  };

  return (
    <div className="w-full group">
      <div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        className={`
          relative rounded-2xl p-10 transition-all duration-300 ease-out text-center cursor-pointer overflow-hidden
          border-2 border-dashed
          ${isDragging 
            ? 'border-blue-500 bg-blue-50/50 scale-[1.01] shadow-lg shadow-blue-100' 
            : 'border-slate-300 hover:border-blue-400 bg-slate-50/50 hover:bg-slate-50'
          }
          ${disabled ? 'opacity-50 cursor-not-allowed grayscale' : ''}
          ${error ? 'border-red-300 bg-red-50/50' : ''}
        `}
      >
        <input
          type="file"
          accept=".docx"
          onChange={handleFileInput}
          disabled={disabled}
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer disabled:cursor-not-allowed z-20"
        />
        
        {/* Decorative Background Icon */}
        <div className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 opacity-[0.03] pointer-events-none z-0">
           <Upload className="w-48 h-48" />
        </div>
        
        <div className="relative z-10 flex flex-col items-center justify-center space-y-4">
          <div className={`p-4 rounded-2xl transition-all duration-300 transform group-hover:scale-110 group-hover:rotate-3 shadow-sm ${
              error 
              ? 'bg-red-100 text-red-600' 
              : 'bg-white text-blue-600 shadow-md shadow-blue-100'
            }`}>
            {error ? (
              <AlertCircle className="w-8 h-8" />
            ) : (
              <FilePlus className="w-8 h-8" />
            )}
          </div>
          
          <div className="space-y-2">
            <h3 className={`text-lg font-bold transition-colors ${error ? 'text-red-700' : 'text-slate-700 group-hover:text-blue-700'}`}>
              {error ? "Định dạng không hợp lệ" : "Kéo thả hoặc chọn tài liệu"}
            </h3>
            <p className="text-slate-400 text-sm max-w-xs mx-auto font-medium">
              {error || "Hỗ trợ định dạng .docx (Microsoft Word)"}
            </p>
          </div>
          
          {!error && !disabled && (
             <div className="pt-2">
                <span className="inline-block px-4 py-1.5 rounded-full bg-blue-100/50 text-blue-700 text-xs font-bold uppercase tracking-wider border border-blue-200 group-hover:bg-blue-600 group-hover:text-white transition-all">
                  Browse Files
                </span>
             </div>
          )}
        </div>
      </div>
    </div>
  );
};