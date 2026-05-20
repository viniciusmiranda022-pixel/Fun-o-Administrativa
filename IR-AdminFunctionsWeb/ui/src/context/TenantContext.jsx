import { createContext, useContext, useEffect, useState, useCallback } from 'react';
import { api } from '../api/client.js';

const TenantContext = createContext(null);

const STORAGE_KEY = 'ir-admin.selected-tenant';

export function TenantProvider({ children }) {
  const [tenants, setTenants] = useState([]);
  const [selectedTenantId, setSelectedTenantId] = useState(() => localStorage.getItem(STORAGE_KEY) || null);
  const [loading, setLoading] = useState(true);

  const reload = useCallback(async () => {
    setLoading(true);
    try {
      const r = await api.tenants();
      const list = r?.data ?? [];
      setTenants(list);
      if (list.length > 0) {
        const exists = list.find((t) => t.tenantId === selectedTenantId);
        if (!exists) {
          setSelectedTenantId(list[0].tenantId);
          localStorage.setItem(STORAGE_KEY, list[0].tenantId);
        }
      } else {
        setSelectedTenantId(null);
        localStorage.removeItem(STORAGE_KEY);
      }
    } catch {
      setTenants([]);
    } finally {
      setLoading(false);
    }
  }, [selectedTenantId]);

  useEffect(() => { reload(); /* eslint-disable-line */ }, []);

  const selectTenant = (tenantId) => {
    setSelectedTenantId(tenantId);
    if (tenantId) localStorage.setItem(STORAGE_KEY, tenantId);
    else localStorage.removeItem(STORAGE_KEY);
  };

  const selected = tenants.find((t) => t.tenantId === selectedTenantId) || null;

  return (
    <TenantContext.Provider value={{ tenants, selected, selectedTenantId, selectTenant, reload, loading }}>
      {children}
    </TenantContext.Provider>
  );
}

export function useTenant() {
  const ctx = useContext(TenantContext);
  if (!ctx) throw new Error('useTenant must be used inside TenantProvider');
  return ctx;
}
