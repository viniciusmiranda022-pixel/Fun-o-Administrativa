import { useState, useRef, useEffect } from 'react';
import { Outlet, useNavigate } from 'react-router-dom';
import { Filter, ChevronDown, Check } from 'lucide-react';
import Sidebar from './Sidebar.jsx';
import Topbar from './Topbar.jsx';
import TabNav from './TabNav.jsx';
import { useTenant } from '../../context/TenantContext.jsx';

function TenantSelector() {
  const { tenants, selected, selectTenant } = useTenant();
  const navigate = useNavigate();
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  useEffect(() => {
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  return (
    <div className="relative" ref={ref}>
      <button
        onClick={() => setOpen((v) => !v)}
        className="flex items-center gap-1.5 text-xs text-[#333] hover:text-[#0078A8] transition-colors"
      >
        <Filter size={13} className="text-[#0078A8]" />
        <span className="font-medium">{selected?.name || selected?.tenantId || 'None'}</span>
        <ChevronDown size={11} className="text-[#888]" />
      </button>

      {open && (
        <div className="absolute right-0 mt-1.5 w-64 bg-white rounded shadow-lg border border-border-light z-50">
          <div className="py-1 max-h-72 overflow-auto">
            {tenants.length === 0 ? (
              <div className="px-4 py-3 text-xs text-text-secondary">No tenants configured.</div>
            ) : (
              tenants.map((t) => (
                <button
                  key={t.tenantId}
                  onClick={() => { selectTenant(t.tenantId); setOpen(false); }}
                  className="w-full flex items-center gap-2 px-4 py-2 text-xs hover:bg-gray-50 text-left"
                >
                  <span className="w-4 flex-shrink-0">
                    {selected?.tenantId === t.tenantId && <Check size={12} className="text-[#0078A8]" />}
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="font-semibold truncate">{t.name || t.tenantId}</div>
                    {t.domain && <div className="text-[10px] text-text-secondary truncate">{t.domain}</div>}
                  </div>
                </button>
              ))
            )}
          </div>
          <div className="border-t border-border-light px-2 py-1">
            <button
              onClick={() => { setOpen(false); navigate('/tenants'); }}
              className="w-full text-xs text-[#0078A8] hover:underline px-2 py-1 text-left"
            >
              Manage tenants…
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export default function Layout() {
  return (
    <div className="h-screen flex flex-col bg-panel overflow-hidden">
      <Topbar />
      <div className="flex-1 flex overflow-hidden">
        <Sidebar />
        <main className="flex-1 flex flex-col overflow-hidden min-w-0">
          {/* Module header */}
          <div className="flex items-center justify-between px-6 py-2 border-b border-border-light bg-white flex-shrink-0">
            <h1 className="text-sm text-[#0078A8] font-medium">
              Administrative Function Recovery for Microsoft Entra ID
            </h1>
            <TenantSelector />
          </div>

          <TabNav />

          <div className="flex-1 overflow-auto bg-[#F5F6F7]">
            <Outlet />
          </div>
        </main>
      </div>
    </div>
  );
}
