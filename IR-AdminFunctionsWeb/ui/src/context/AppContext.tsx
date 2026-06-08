import { createContext, useContext, useEffect, useState } from 'react';
import type { ReactNode } from 'react';
import { api } from '../api/client';
import type { TenantEntry } from '../types/api';

interface AppContextValue {
  tenants: TenantEntry[];
  selectedTenantId: string | null;
  setSelectedTenantId: (id: string) => void;
  selectedTenant: TenantEntry | null;
}

interface AppProviderProps {
  children: ReactNode;
}

const AppContext = createContext<AppContextValue | null>(null);

export function AppProvider({ children }: AppProviderProps) {
  const [tenants, setTenants] = useState<TenantEntry[]>([]);
  const [selectedTenantId, setSelectedTenantId] = useState<string | null>(null);

  useEffect(() => {
    api.tenants()
      .then((r) => {
        const list: TenantEntry[] = r?.data ?? [];
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

export function useApp(): AppContextValue {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error('useApp must be used inside AppProvider');
  return ctx;
}
