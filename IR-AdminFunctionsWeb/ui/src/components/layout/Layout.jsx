import { Outlet } from 'react-router-dom';
import Sidebar from './Sidebar.jsx';
import Topbar from './Topbar.jsx';
import TabNav from './TabNav.jsx';

export default function Layout() {
  return (
    <div className="h-screen flex flex-col bg-panel">
      <Topbar />
      <div className="flex-1 flex overflow-hidden">
        <Sidebar />
        <main className="flex-1 flex flex-col overflow-hidden bg-white">
          <div className="px-6 py-3 border-b border-border-light">
            <h1 className="text-lg text-[#0078A8]">Administrative Function Recovery for Microsoft Entra ID</h1>
          </div>
          <TabNav />
          <div className="flex-1 overflow-auto bg-panel">
            <Outlet />
          </div>
        </main>
      </div>
    </div>
  );
}
