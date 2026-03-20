import React, { useEffect, useRef, useState } from 'react';
import * as docx from 'docx-preview';
import { FileSearch } from 'lucide-react';

interface DocumentPreviewProps {
  blob: Blob;
}

export const DocumentPreview: React.FC<DocumentPreviewProps> = ({ blob }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const [isRendering, setIsRendering] = useState(true);

  useEffect(() => {
    if (containerRef.current && blob) {
      setIsRendering(true);
      containerRef.current.innerHTML = ''; 
      
      docx.renderAsync(blob, containerRef.current, undefined, {
        className: 'docx-preview-page',
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true,
      })
      .then(() => setIsRendering(false))
      .catch(err => {
        console.error("Preview error:", err);
        setIsRendering(false);
      });
    }
  }, [blob]);

  return (
    <div className="mt-8 animate-fadeIn w-full border-t border-slate-100 pt-6">
      <div className="flex items-center justify-center gap-2 text-slate-700 font-bold mb-4">
        <FileSearch className="w-5 h-5 text-blue-600" />
        <span>Bản xem trước tài liệu</span>
      </div>
      <div className="relative w-full rounded-xl border border-slate-200 bg-slate-200/50 overflow-hidden shadow-inner">
        {isRendering && (
          <div className="absolute inset-0 flex flex-col items-center justify-center bg-slate-100/80 backdrop-blur-sm z-10">
             <div className="w-8 h-8 border-4 border-blue-200 border-t-blue-600 rounded-full animate-spin mb-2"></div>
             <p className="text-sm font-semibold text-slate-500">Đang tải bản xem trước...</p>
          </div>
        )}
        <div 
          ref={containerRef} 
          className="w-full h-[500px] overflow-y-auto p-4 flex flex-col items-center"
        />
      </div>
    </div>
  );
};
