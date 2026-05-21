import { createContext, useContext, useEffect, useState } from 'react';
import { api } from '../api/client.js';

const AppContext = createContext(null);

export function AppProvider({ children }) {
  const [tenants, setTenants] = useState([]);
  const [selectedTenantId, setSelectedTenantId] = useState(null);

  useEffect(() => {
    api.tenants()
      .then((r) => {
        const list = r?.data ?? [];
        setTenants(list);
        if (list.length > 0) setSelectedTenantId(list[0].tenantId);
      })
      .catch(() => {
        setTenants([]);
      });
  }, []);

  const selectedTenant = tenants.find((t) => t.tenantId === selectedTenantId) ?? tenants[0] ?? null;

  return (
    <AppContext.Provider value={{ tenants, selectedTenantId, setSelectedTenantId, selectedTenant }}>
      {children}
    </AppContext.Provider>
  );
}

export function useApp() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error('useApp must be used inside AppProvider');
  return ctx;
}
