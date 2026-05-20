import { Bell, ChevronDown, Menu } from 'lucide-react';

export default function Topbar() {
  return (
    <header className="h-12 bg-[#2F2F2F] text-white flex items-center px-4 gap-3 flex-shrink-0 border-b-2 border-[#0096D6] z-30">
      <Menu size={18} className="text-[#999] cursor-pointer hover:text-white flex-shrink-0" />

      {/* Branding */}
      <div className="flex items-center gap-2 flex-shrink-0">
        <span className="font-bold text-[17px] tracking-tight leading-none">-Quest</span>
        <div className="h-5 w-px bg-[#555]" />
        <span className="text-[13px] font-light text-[#ccc]">
          Security Management Platform – Module Function Administration
        </span>
      </div>

      {/* Right side */}
      <div className="ml-auto flex items-center gap-5 text-xs">
        <span className="flex items-center gap-1.5 text-[#ccc]">
          <span className="w-2 h-2 rounded-full bg-green-400 flex-shrink-0" />
          All Systems Operational
        </span>

        <span className="text-pink-400 text-base leading-none select-none">✦</span>

        <span className="flex items-center gap-1.5 text-[#ccc] cursor-pointer hover:text-white select-none">
          <span className="w-6 h-6 rounded-full bg-[#0078A8] flex items-center justify-center text-[10px] font-bold text-white flex-shrink-0">
            VM
          </span>
          <span className="hidden md:inline">Administrator</span>
          <ChevronDown size={11} />
        </span>

        <Bell size={15} className="text-[#ccc] cursor-pointer hover:text-white" />
      </div>
    </header>
  );
}
