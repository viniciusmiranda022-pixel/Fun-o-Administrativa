import { useEffect, useState } from 'react';
import { api } from '../api/client.js';

function Card({ title, value, subtitle, children }) {
  return (
    <div className="bg-white border border-border-light p-4 min-h-[200px] flex flex-col">
      <div className="text-2xl font-light text-[#222]">{value ?? '—'}</div>
      <div className="text-sm font-semibold text-[#0078A8] mt-1">{title}</div>
      {subtitle && <div className="text-xs text-text-secondary mt-2">{subtitle}</div>}
      {children}
    </div>
  );
}

export default function Dashboard() {
  const [stats, setStats] = useState({});

  useEffect(() => {
    Promise.allSettled([api.backups(), api.tasks(), api.events('Error'), api.tenants()])
      .then(([backups, tasks, errors, tenants]) => {
        setStats({
          backups: backups.status === 'fulfilled' ? (backups.value?.data?.length ?? 0) : 0,
          tasks: tasks.status === 'fulfilled' ? (tasks.value?.data?.items?.length ?? 0) : 0,
          errors: errors.status === 'fulfilled' ? (errors.value?.data?.items?.length ?? 0) : 0,
          tenant: tenants.status === 'fulfilled' ? tenants.value?.data?.[0] : null
        });
      });
  }, []);

  return (
    <div className="p-4 grid grid-cols-1 md:grid-cols-3 gap-3">
      <Card
        title="Tenant protected"
        value={stats.tenant?.displayName ?? 'Unconfigured'}
        subtitle={stats.tenant?.tenantId}
      />
      <Card title="Backups" value={stats.backups} subtitle="Snapshots disponíveis" />
      <Card title="Tasks" value={stats.tasks} subtitle="Histórico recente" />
      <Card title="Unpacked objects" value="—" subtitle="Selecione um backup" />
      <Card title="Differences" value="—" subtitle="Refresh report disponível na aba DIFFERENCES" />
      <Card title="Errors" value={stats.errors} subtitle="Eventos com severity=Error" />
    </div>
  );
}
