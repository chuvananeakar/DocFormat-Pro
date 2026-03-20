import React from 'react';
import { Terminal, Check } from 'lucide-react';

interface ProcessingLogProps {
  logs: string[];
}

export const ProcessingLog: React.FC<ProcessingLogProps> = ({ logs }) => {
  if (logs.length === 0) return null;

  return (
    <div className="bg-slate-900 rounded-lg p-4 font-mono text-xs sm:text-sm text-slate-300 shadow-inner max-h-48 overflow-y-auto w-full">
      <div className="flex items-center gap-2 text-slate-400 mb-2 border-b border-slate-700 pb-2">
        <Terminal className="w-4 h-4" />
        <span className="font-semibold uppercase tracking-wider">System Logs</span>
      </div>
      <div className="space-y-1">
        {logs.map((log, i) => (
          <div key={i} className="flex items-start gap-2 animate-fadeIn">
            <span className="text-green-500 mt-0.5"><Check className="w-3 h-3" /></span>
            <span>{log}</span>
          </div>
        ))}
      </div>
    </div>
  );
};