import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Shield, ShieldAlert, ShieldCheck, Database, Package, GitCompare,
  ClipboardList, AlertTriangle, Settings, Play, Archive, RefreshCw,
  CheckCircle2, AlertCircle, Loader2, ArrowRight
} from 'lucide-react';
import { api } from '../api/client.js';
import { useTenant } from '../context/TenantContext.jsx';

function ActionButton({ icon: Icon, label, onClick, primary = false, disabled = false }) {
  const cls = primary
    ? 'bg-[#0078A8] text-white hover:bg-[#006090] disabled:opacity-50'
    : 'bg-white border border-border-light text-[#333] hover:bg-gray-50 disabled:opacity-50';
  return (
    <button onClick={onClick} disabled={disabled} className={`flex items-center gap-2 px-3 py-2 text-xs font-semibold rounded ${cls}`}>
      <Icon size={14} /> {label}
    </button>
  );
}

function ProtectionCard({ stats, onConfigure }) {
  const status = stats.protectionStatus;
  const cfg = {
    Protected:       { Icon: ShieldCheck, color: 'text-green-600',  bg: 'bg-green-50',  border: 'border-green-200',  title: 'Tenant is protected' },
    NeedsAttention:  { Icon: ShieldAlert, color: 'text-yellow-600', bg: 'bg-yellow-50', border: 'border-yellow-200', title: 'Atenção necessária' },
    NotProtected:    { Icon: Shield,      color: 'text-red-600',    bg: 'bg-red-50',    border: 'border-red-200',    title: 'Tenant is not protected' }
  }[status] || { Icon: Shield, color: 'text-gray-500', bg: 'bg-gray-50', border: 'border-border-light', title: 'Status desconhecido' };

  const { Icon } = cfg;
  const fmt = (d) => d ? new Date(d).toLocaleString('pt-BR') : '—';

  return (
    <div className={`col-span-1 md:col-span-2 ${cfg.bg} ${cfg.border} border rounded-lg p-5`}>
      <div className="flex items-start gap-3">
        <Icon className={cfg.color} size={28} />
        <div className="flex-1">
          <h2 className={`text-base font-semibold ${cfg.color}`}>{cfg.title}</h2>
          {stats.tenant ? (
            <>
              <p className="text-sm text-[#333] mt-0.5 mb-3">
                {stats.tenant.name || stats.tenant.tenantId}
                {stats.tenant.domain && <span className="text-text-secondary"> · {stats.tenant.domain}</span>}
              </p>
              <dl className="grid grid-cols-2 gap-x-6 gap-y-1.5 text-xs">
                <div className="flex justify-between"><dt className="text-text-secondary">Schedule:</dt>
                  <dd className="font-semibold">{stats.tenant.backupConfig?.scheduleEnabled ? 'Enabled' : 'Disabled'}</dd></div>
                <div className="flex justify-between"><dt className="text-text-secondary">Retention:</dt>
                  <dd className="font-semibold">{stats.tenant.backupConfig?.retentionDays ?? 90} dias</dd></div>
                <div className="flex justify-between"><dt className="text-text-secondary">Último backup:</dt>
                  <dd className="font-semibold">{fmt(stats.lastBackupAt)}</dd></div>
                <div className="flex justify-between"><dt className="text-text-secondary">Backups:</dt>
                  <dd className="font-semibold">{stats.backupsCount}</dd></div>
                <div className="flex justify-between"><dt className="text-text-secondary">Roles protegidas:</dt>
                  <dd className="font-semibold">{stats.protectedRoleDefinitionsCount ?? '—'}</dd></div>
                <div className="flex justify-between"><dt className="text-text-secondary">Assignments:</dt>
                  <dd className="font-semibold">{stats.protectedRoleAssignmentsCount ?? '—'}</dd></div>
                <div className="flex justify-between col-span-2 mt-1 pt-2 border-t border-border-light">
                  <dt className="text-text-secondary">Consentimentos:</dt>
                  <dd className="flex gap-3">
                    <span className={`flex items-center gap-1 ${stats.consentBasicGranted ? 'text-green-700' : 'text-yellow-700'}`}>
                      {stats.consentBasicGranted ? <CheckCircle2 size={12} /> : <AlertCircle size={12} />} Basic
                    </span>
                    <span className={`flex items-center gap-1 ${stats.consentRestoreGranted ? 'text-green-700' : 'text-yellow-700'}`}>
                      {stats.consentRestoreGranted ? <CheckCircle2 size={12} /> : <AlertCircle size={12} />} Restore
                    </span>
                  </dd>
                </div>
              </dl>
              <button onClick={onConfigure}
                className="mt-3 text-xs text-[#0078A8] hover:underline flex items-center gap-1">
                <Settings size={12} /> Configurar backup
              </button>
            </>
          ) : (
            <p className="text-sm text-[#333] mt-1">Nenhum tenant configurado.</p>
          )}
        </div>
      </div>
    </div>
  );
}

function StatCard({ icon: Icon, title, value, subtitle, color = 'text-[#0078A8]', onClick }) {
  return (
    <button
      onClick={onClick}
      className="bg-white border border-border-light rounded-lg p-4 text-left hover:border-[#0078A8] hover:shadow-sm transition flex flex-col"
    >
      <div className="flex items-center justify-between mb-2">
        <Icon className={color} size={18} />
        <ArrowRight className="text-gray-300" size={14} />
      </div>
      <div className="text-2xl font-light text-[#222]">{value ?? '—'}</div>
      <div className={`text-sm font-semibold ${color} mt-0.5`}>{title}</div>
      {subtitle && <div className="text-xs text-text-secondary mt-1">{subtitle}</div>}
    </button>
  );
}

function CreateBackupModal({ tenant, onClose, onConfirm, busy }) {
  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-md mx-4">
        <div className="flex items-center justify-between px-6 py-4 border-b border-border-light">
          <h2 className="text-base font-semibold text-[#222]">Create backup</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl">×</button>
        </div>
        <div className="px-6 py-5 text-sm text-[#333]">
          Do you want to back up administrative functions for <strong>{tenant?.name || tenant?.tenantId}</strong>?
          <p className="text-xs text-text-secondary mt-2">
            Uma task de Administrative Function Backup será criada e aparecerá em Tasks.
          </p>
        </div>
        <div className="flex justify-end gap-2 px-6 py-3 border-t border-border-light bg-gray-50">
          <button onClick={onClose} disabled={busy}
            className="px-4 py-2 text-sm border border-border-light rounded hover:bg-white disabled:opacity-60">Cancel</button>
          <button onClick={onConfirm} disabled={busy}
            className="px-4 py-2 text-sm bg-[#0078A8] text-white rounded hover:bg-[#006090] disabled:opacity-60 flex items-center gap-2">
            {busy && <Loader2 size={14} className="animate-spin" />}
            Create
          </button>
        </div>
      </div>
    </div>
  );
}

export default function Dashboard() {
  const navigate = useNavigate();
  const { selected, tenants, loading: tLoading } = useTenant();
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(true);
  const [notice, setNotice] = useState(null);
  const [busy, setBusy] = useState(false);
  const [showCreate, setShowCreate] = useState(false);

  const load = async () => {
    setLoading(true);
    try {
      const r = await api.dashboard(selected?.tenantId);
      setStats(r?.data || null);
    } catch (e) {
      setStats(null);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { if (!tLoading) load(); /* eslint-disable-line */ }, [selected?.tenantId, tLoading]);

  const createBackup = async () => {
    if (!selected) return;
    setBusy(true);
    try {
      const r = await api.runBackupFor(selected.tenantId);
      setShowCreate(false);
      setNotice({ kind: 'info', message: 'The Administrative Function Backup task has been created.', jobId: r?.data?.id });
      load();
    } catch (e) {
      setNotice({ kind: 'error', message: e.message });
    } finally {
      setBusy(false);
    }
  };

  const refreshDifferences = async () => {
    if (!selected) return;
    setBusy(true);
    try {
      await api.refreshDifferences(selected.tenantId, stats?.lastBackupId);
      setNotice({ kind: 'info', message: 'Refresh Differences task has been created.' });
      load();
    } catch (e) {
      setNotice({ kind: 'error', message: e.message });
    } finally {
      setBusy(false);
    }
  };

  if (tLoading || loading) {
    return <div className="p-8 text-text-secondary flex items-center gap-2"><Loader2 className="animate-spin" size={14} /> Carregando...</div>;
  }

  if (tenants.length === 0) {
    return (
      <div className="p-8 max-w-2xl mx-auto">
        <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-5">
          <h3 className="font-semibold text-yellow-900">Nenhum tenant configurado</h3>
          <p className="text-sm text-yellow-800 mt-1 mb-3">Adicione um Microsoft 365 tenant para começar a proteger funções administrativas.</p>
          <button onClick={() => navigate('/tenants')} className="px-3 py-1.5 bg-yellow-700 text-white rounded text-sm">
            Ir para Tenants
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Dashboard</h2>
          <p className="text-xs text-text-secondary mt-0.5">
            Painel de proteção de funções administrativas — {selected?.name || selected?.tenantId}
          </p>
        </div>
        <button onClick={load} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1">
          <RefreshCw size={12} /> Atualizar
        </button>
      </div>

      <div className="flex flex-wrap gap-2 mb-5">
        <ActionButton icon={Database}   label="Manage Backups"      onClick={() => navigate('/manage-backups')} />
        <ActionButton icon={ShieldCheck} label="Manage Restores"     onClick={() => navigate('/manage-restores')} />
        <ActionButton icon={Play}        label="Create Backup"        primary onClick={() => setShowCreate(true)} disabled={busy || !selected} />
        <ActionButton icon={Archive}     label="Unpack Backup"        onClick={() => navigate('/unpack')} />
        <ActionButton icon={RefreshCw}   label="Refresh Differences"  onClick={refreshDifferences} disabled={busy || !selected} />
      </div>

      {notice && (
        <div className={`mb-4 px-3 py-2 rounded text-xs flex items-center justify-between ${
          notice.kind === 'error' ? 'bg-red-50 border border-red-200 text-red-800' : 'bg-blue-50 border border-blue-200 text-blue-900'
        }`}>
          <span>{notice.message}</span>
          <div className="flex items-center gap-2">
            {notice.jobId && (
              <button onClick={() => navigate('/tasks')} className="underline">View details</button>
            )}
            <button onClick={() => setNotice(null)} className="text-gray-500 hover:text-gray-700">×</button>
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <ProtectionCard stats={stats || {}} onConfigure={() => navigate('/manage-backups')} />

        <StatCard icon={Database}   title="Backups"           value={stats?.backupsCount ?? 0}
                  subtitle="Snapshots disponíveis" onClick={() => navigate('/backups')} />
        <StatCard icon={ClipboardList} title="Tasks"           value={stats?.tasksTotalCount ?? 0}
                  subtitle={`${stats?.tasksRunningCount ?? 0} em execução`} onClick={() => navigate('/tasks')} />
        <StatCard icon={Package}    title="Unpacked"          value={stats?.unpackedBackupsCount ?? 0}
                  subtitle="Backups descompactados" onClick={() => navigate('/unpacked')} />
        <StatCard icon={GitCompare} title="Differences"       value={stats?.differencesCount ?? '—'}
                  subtitle="Comparações pendentes" onClick={() => navigate('/differences')} />
        <StatCard icon={AlertTriangle} title="Errors"          value={stats?.errorsCount ?? 0}
                  subtitle={`${stats?.warningsCount ?? 0} avisos`} color="text-red-600"
                  onClick={() => navigate('/events')} />
      </div>

      {showCreate && (
        <CreateBackupModal tenant={selected} busy={busy} onClose={() => setShowCreate(false)} onConfirm={createBackup} />
      )}
    </div>
  );
}
