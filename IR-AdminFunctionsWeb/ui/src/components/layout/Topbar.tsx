import { Bell, Info, Menu, ChevronDown } from 'lucide-react';
import { useApp } from '../../context/AppContext';

export default function Topbar() {
  const { selectedTenant } = useApp();

  return (
    <header
      className="h-12 text-white flex items-center px-3 border-b-2 border-[#0096D6] flex-shrink-0 z-20"
      style={{ backgroundColor: '#252525' }}
    >
      {/* Left */}
      <div className="flex items-center gap-3">
        <Menu size={18} className="cursor-pointer opacity-80 hover:opacity-100" />
        <img src="/quest-logo.png" alt="Quest" style={{ height: 22, width: 'auto', display: 'block' }} />
        <div className="h-6 w-px bg-[#555]" />
        <span className="text-[#999] text-xs">Security Management Platform</span>
        <span className="text-[#777] text-xs">-</span>
        <span className="text-[#bbb] text-xs">Module Function Administration</span>
      </div>

      {/* Right */}
      <div className="ml-auto flex items-center gap-4">
        {/* System status */}
        <div className="flex items-center gap-1.5 text-xs text-[#ccc]">
          <span className="w-2 h-2 rounded-full bg-[#28A745] inline-block" />
          All Systems Operational
        </div>

        {/* Gradient Q icon */}
        <div
          className="w-6 h-6 rounded flex items-center justify-center text-white font-bold text-xs"
          style={{ background: 'linear-gradient(135deg, #0096D6, #28A745)' }}
        >
          Q
        </div>

        {/* User */}
        <div className="flex items-center gap-1.5 cursor-pointer hover:opacity-80">
          <div className="w-6 h-6 rounded-full bg-[#0096D6] flex items-center justify-center text-xs font-bold">
            V
          </div>
          <span className="text-xs text-[#ccc]">vinicius.miranda@tisco.com.br</span>
          <ChevronDown size={12} className="text-[#999]" />
        </div>

        {/* Bell */}
        <div className="relative cursor-pointer">
          <Bell size={16} className="text-[#ccc]" />
          <span className="absolute -top-1.5 -right-1.5 bg-[#DC3545] text-white text-[9px] w-3.5 h-3.5 rounded-full flex items-center justify-center font-bold">
            3
          </span>
        </div>

        <Info size={16} className="text-[#ccc] cursor-pointer hover:text-white" />
      </div>
    </header>
  );
}
