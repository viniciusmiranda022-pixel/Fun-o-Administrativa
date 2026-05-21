import { Plus } from 'lucide-react';

export default function Topbar() {
  return (
    <header className="h-10 bg-[#1F1F1F] text-white flex items-center px-4 gap-3 shrink-0">
      <span className="font-bold text-sm tracking-tight">-Quest</span>
      <span className="text-[#aaa] text-xs">Security Management Platform – Module Function Administration</span>
      <div className="ml-auto flex items-center gap-3 text-xs text-[#ccc]">
        <span className="flex items-center gap-1.5">
          <span className="w-2 h-2 rounded-full bg-green-500 shrink-0" />
          All Systems Operational
        </span>
        <button className="w-6 h-6 flex items-center justify-center hover:bg-white/10 rounded text-[#aaa]">
          <Plus size={14} />
        </button>
        <div className="flex items-center gap-1.5">
          <span className="w-7 h-7 rounded-full bg-[#0066CC] flex items-center justify-center text-white text-[11px] font-bold select-none">
            VM
          </span>
          <span className="text-[#ccc]">Administrator</span>
        </div>
      </div>
    </header>
  );
}
