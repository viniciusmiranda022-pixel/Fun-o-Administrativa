import { RotateCw } from 'lucide-react';

export default function Sidebar() {
  return (
    <div className="flex h-full">
      <div className="w-[60px] bg-sidebar flex flex-col items-center pt-5">
        <div className="mt-[108px] w-full bg-[#3C3C3C] py-3 relative">
          <div className="absolute left-0 top-0 bottom-0 w-1 bg-[#0096D6]" />
          <div className="flex flex-col items-center text-white">
            <RotateCw size={16} />
            <span className="text-[11px] mt-1">Recover</span>
          </div>
        </div>
      </div>
      <div className="w-[145px] bg-[#404040] text-white pt-5 pl-4">
        <span className="text-xs">Microsoft Entra ID</span>
      </div>
    </div>
  );
}
