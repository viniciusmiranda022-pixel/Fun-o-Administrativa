import { useEffect, useState } from 'react';
import { Bell, Info, Menu } from 'lucide-react';
import { api } from '../../api/client.js';

export default function Topbar() {
  const [user, setUser] = useState('');

  useEffect(() => {
    api.tenants()
      .then((r) => {
        const t = r?.data?.[0];
        if (t?.appDisplayName) setUser(t.appDisplayName);
      })
      .catch(() => {});
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
        <span>{user || 'operador'}</span>
        <Bell size={16} />
        <Info size={16} />
      </div>
    </header>
  );
}
