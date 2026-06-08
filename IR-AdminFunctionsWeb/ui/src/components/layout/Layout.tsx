import { Outlet, useLocation } from 'react-router-dom';
import Sidebar from './Sidebar';
import Topbar from './Topbar';
import TabNav from './TabNav';
import TenantSelector from './TenantSelector';

const RECOVER_PATHS = ['/dashboard', '/backups', '/unpacked', '/differences', '/events', '/tasks'];

export default function Layout() {
  const location = useLocation();
  const isRecoverModule = RECOVER_PATHS.some(
    (p) => location.pathname === p || location.pathname.startsWith(p + '/')
  );

  return (
    <div className="h-screen flex flex-col" style={{ backgroundColor: '#F5F5F5' }}>
      <Topbar />
      <div className="flex-1 flex overflow-hidden">
        <Sidebar />
        <main className="flex-1 flex flex-col overflow-hidden bg-white">
          {isRecoverModule ? (
            <>
              {/* Module header bar */}
              <div className="px-5 py-2 border-b border-[#DEE2E6] flex items-center justify-between flex-shrink-0 bg-white">
                <div>
                  <h1 className="text-[15px] font-semibold text-[#1a1a1a] leading-tight">
                    Administrative Function Recovery
                  </h1>
                  <span className="text-xs text-[#0078A8] font-medium">for Microsoft Entra ID</span>
                </div>
                <TenantSelector />
              </div>
              <TabNav />
              <div className="flex-1 overflow-auto" style={{ backgroundColor: '#F5F5F5' }}>
                <Outlet />
              </div>
            </>
          ) : (
            <div className="flex-1 overflow-auto" style={{ backgroundColor: '#F5F5F5' }}>
              <Outlet />
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
