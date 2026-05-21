import { Outlet, useLocation } from 'react-router-dom';
import { SlidersHorizontal } from 'lucide-react';
import Sidebar from './Sidebar.jsx';
import Topbar from './Topbar.jsx';
import TabNav from './TabNav.jsx';

const TAB_ROUTES = ['/dashboard', '/backups', '/unpacked', '/differences', '/events', '/tasks'];

export default function Layout() {
  const { pathname } = useLocation();
  const showTabs = TAB_ROUTES.some((r) => pathname.startsWith(r));

  return (
    <div className="h-screen flex flex-col bg-panel overflow-hidden">
      <Topbar />
      <div className="flex-1 flex overflow-hidden">
        <Sidebar />
        <main className="flex-1 flex flex-col overflow-hidden bg-white">
          <div className="px-5 py-2 border-b border-border-light flex items-center justify-between shrink-0">
            <h1 className="text-sm font-normal text-[#0066CC]">
              Administrative Function Recovery for Microsoft Entra ID
            </h1>
            {showTabs && (
              <button className="flex items-center gap-1.5 text-[11px] text-[#555] border border-border-light rounded px-2 py-1 hover:bg-gray-50">
                <SlidersHorizontal size={11} />
                None
              </button>
            )}
          </div>
          {showTabs && <TabNav />}
          <div className="flex-1 overflow-auto bg-panel">
            <Outlet />
          </div>
        </main>
      </div>
    </div>
  );
}
