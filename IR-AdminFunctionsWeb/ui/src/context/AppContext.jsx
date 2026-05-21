import { createContext, useContext, useEffect, useState } from 'react';
import { api } from '../api/client.js';

const MOCK_TENANTS = [
  {
    tenantId: 'ca9b03ea-578e-4277-b684-969fa2a34a9a',
    name: 'Contoso',
    domain: 'M365x24071226.onmicrosoft.com',
    addedAt: '2026-04-16',
    consents: { basic: { granted: true }, restore: { granted: true } }
  }
];

const AppContext = createContext(null);

export function AppProvider({ children }) {
  const [tenants, setTenants] = useState([]);
  const [selectedTenantId, setSelectedTenantId] = useState(null);

  useEffect(() => {
    api.tenants()
      .then((r) => {
        const list = r?.data ?? [];
        if (list.length > 0) {
          setTenants(list);
          setSelectedTenantId(list[0].tenantId);
        } else {
          setTenants(MOCK_TENANTS);
          setSelectedTenantId(MOCK_TENANTS[0].tenantId);
        }
      })
      .catch(() => {
        setTenants(MOCK_TENANTS);
        setSelectedTenantId(MOCK_TENANTS[0].tenantId);
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
