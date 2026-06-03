import { useState, useRef, useEffect } from 'react';
import { ChevronDown, AlignJustify } from 'lucide-react';
import { useApp } from '../../context/AppContext';

export default function TenantSelector() {
  const { tenants, selectedTenant, setSelectedTenantId } = useApp();
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen((o) => !o)}
        className="flex items-center gap-1.5 px-2 py-1 border border-[#DEE2E6] bg-white text-xs text-[#333] hover:bg-gray-50"
      >
        <AlignJustify size={12} className="text-[#0078A8]" />
        <span className="font-medium">{selectedTenant?.name ?? 'None'}</span>
        <ChevronDown size={12} className="text-[#666]" />
      </button>

      {open && (
        <div className="absolute right-0 top-full mt-1 w-56 bg-white border border-[#DEE2E6] shadow-lg z-50">
          {tenants.map((t) => (
            <button
              key={t.tenantId}
              onClick={() => { setSelectedTenantId(t.tenantId); setOpen(false); }}
              className={`w-full text-left px-3 py-2 text-xs hover:bg-[#F2F2F2] ${
                t.tenantId === selectedTenant?.tenantId ? 'text-[#0078A8] font-semibold' : 'text-[#333]'
              }`}
            >
              {t.name ?? t.domain ?? t.tenantId}
            </button>
          ))}
          {tenants.length === 0 && (
            <div className="px-3 py-2 text-xs text-[#999]">No tenants configured</div>
          )}
        </div>
      )}
    </div>
  );
}
