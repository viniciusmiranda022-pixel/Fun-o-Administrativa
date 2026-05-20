import { Bell, Info, Menu, ChevronDown, Check } from 'lucide-react';
import { useState, useRef, useEffect } from 'react';
import { useTenant } from '../../context/TenantContext.jsx';
import { useNavigate } from 'react-router-dom';

export default function Topbar() {
  const { tenants, selected, selectTenant } = useTenant();
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  const navigate = useNavigate();

  useEffect(() => {
    const onClick = (e) => {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    };
    document.addEventListener('mousedown', onClick);
    return () => document.removeEventListener('mousedown', onClick);
  }, []);

  return (
    <header className="h-14 bg-[#2F2F2F] text-white flex items-center px-4 border-b-2 border-[#0096D6]">
      <Menu size={18} className="mr-4" />
      <span className="text-2xl font-semibold">Quest</span>
      <div className="mx-4 h-9 w-px bg-[#555]" />
      <span className="text-lg">Security Management Platform</span>
      <div className="ml-auto flex items-center gap-4 text-xs">
        <span className="flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-status-success" /> All Systems Operational
        </span>

        <div className="relative" ref={ref}>
          <button
            onClick={() => setOpen((v) => !v)}
            className="flex items-center gap-2 px-3 py-1.5 rounded bg-[#3C3C3C] hover:bg-[#4A4A4A] text-xs"
          >
            <span className="text-text-secondary uppercase tracking-wide">Tenant:</span>
            <span className="font-semibold text-white">{selected?.name || selected?.tenantId || 'None'}</span>
            <ChevronDown size={12} />
          </button>
          {open && (
            <div className="absolute right-0 mt-1 w-72 bg-white text-[#222] rounded shadow-lg border border-border-light z-50">
              <div className="py-1 max-h-80 overflow-auto">
                {tenants.length === 0 ? (
                  <div className="px-4 py-3 text-xs text-text-secondary">Nenhum tenant configurado.</div>
                ) : (
                  tenants.map((t) => (
                    <button
                      key={t.tenantId}
                      onClick={() => { selectTenant(t.tenantId); setOpen(false); }}
                      className="w-full flex items-center gap-2 px-4 py-2 text-xs hover:bg-gray-50 text-left"
                    >
                      <span className="w-4">
                        {selected?.tenantId === t.tenantId && <Check size={12} className="text-[#0078A8]" />}
                      </span>
                      <div className="flex-1">
                        <div className="font-semibold">{t.name || t.tenantId}</div>
                        {t.domain && <div className="text-[10px] text-text-secondary">{t.domain}</div>}
                      </div>
                    </button>
                  ))
                )}
              </div>
              <div className="border-t border-border-light px-2 py-1">
                <button
                  onClick={() => { setOpen(false); navigate('/tenants'); }}
                  className="w-full text-xs text-[#0078A8] hover:underline px-2 py-1.5 text-left"
                >
                  Gerenciar tenants...
                </button>
              </div>
            </div>
          )}
        </div>

        <Bell size={16} />
        <Info size={16} />
      </div>
    </header>
  );
}
