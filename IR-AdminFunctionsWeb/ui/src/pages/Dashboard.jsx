import { useEffect, useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  ShieldCheck, ShieldAlert, Shield, Database, Package, GitCompare,
  ClipboardList, AlertCircle, AlertTriangle, Loader2, CheckCircle2,
  Settings, Play, Archive, RefreshCw, RotateCcw, ArrowRight
} from 'lucide-react';
import { api } from '../api/client.js';
import { useTenant } from '../context/TenantContext.jsx';
import ManageBackupsModal  from '../components/modals/ManageBackupsModal.jsx';
import ManageRestoresModal from '../components/modals/ManageRestoresModal.jsx';
import UnpackBackupModal   from '../components/modals/UnpackBackupModal.jsx';

/* ─── Mini bar chart (SVG) ─────────────────────────────────────────── */
function BackupBarChart({ backups }) {
  const N = 30;
  const days = useMemo(() => {
    const arr = [];
    const now = new Date();
    for (let i = N - 1; i >= 0; i--) {
      const d = new Date(now);
      d.setDate(d.getDate() - i);
      arr.push({ key: d.toISOString().split('T')[0], count: 0, label: d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) });
    }
    backups.forEach((b) => {
      const ds = (b.collectedAt || b.startedAt || b.modifiedAt || '').split('T')[0];
      const bucket = arr.find((a) => a.key === ds);
      if (bucket) bucket.count++;
    });
    return arr;
  }, [backups]);

  const max = Math.max(...days.map((d) => d.count), 1);
  const bW = 9, gap = 3, H = 64;
  const totalW = N * (bW + gap);

  return (
    <svg width="100%" viewBox={`0 0 ${totalW} ${H + 14}`} className="mt-1">
      {days.map((day, i) => {
        const h = day.count > 0 ? Math.max(Math.round((day.count / max) * H), 3) : 0;
        const x = i * (bW + gap);
        const showLabel = i === 0 || i === Math.floor(N / 4) || i === Math.floor(N / 2) || i === Math.floor((3 * N) / 4) || i === N - 1;
        return (
          <g key={day.key}>
            <rect x={x} y={H - h} width={bW} height={h || 0} fill="#5CB85C" rx={1} />
            {h === 0 && <rect x={x} y={H - 1} width={bW} height={1} fill="#ddd" />}
            {showLabel && (
              <text x={x + bW / 2} y={H + 11} textAnchor="middle" fontSize={6} fill="#999" fontFamily="sans-serif">
                {day.label}
              </text>
            )}
          </g>
        );
      })}
    </svg>
  );
}

/* ─── Donut chart (SVG) ─────────────────────────────────────────────── */
function DonutChart({ segments, total }) {
  const R = 34, r = 22, cx = 42, cy = 42;
  if (!total || segments.every((s) => s.value === 0)) {
    return (
      <svg width={84} height={84} viewBox="0 0 84 84">
        <circle cx={cx} cy={cy} r={R} fill="none" stroke="#E5E7EB" strokeWidth={R - r} />
      </svg>
    );
  }
  let angle = -Math.PI / 2;
  const paths = [];
  segments.filter((s) => s.value > 0).forEach((seg) => {
    const sweep = (seg.value / total) * 2 * Math.PI;
    const x1 = cx + R * Math.cos(angle), y1 = cy + R * Math.sin(angle);
    angle += sweep;
    const x2 = cx + R * Math.cos(angle), y2 = cy + R * Math.sin(angle);
    const xi1 = cx + r * Math.cos(angle), yi1 = cy + r * Math.sin(angle);
    const xi2 = cx + r * Math.cos(angle - sweep), yi2 = cy + r * Math.sin(angle - sweep);
    const lg = sweep > Math.PI ? 1 : 0;
    paths.push({ d: `M${x1},${y1} A${R},${R},0,${lg},1,${x2},${y2} L${xi1},${yi1} A${r},${r},0,${lg},0,${xi2},${yi2} Z`, color: seg.color });
  });
  return (
    <svg width={84} height={84} viewBox="0 0 84 84">
      {paths.map((p, i) => <path key={i} d={p.d} fill={p.color} />)}
    </svg>
  );
}

/* ─── Action button ─────────────────────────────────────────────────── */
function ActionBtn({ Icon, label, onClick, disabled = false }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className="flex items-center gap-1.5 px-3 py-1.5 text-[11px] font-semibold uppercase tracking-wide border border-[#C8C8C8] bg-white text-[#333] hover:bg-gray-50 hover:border-[#0078A8] hover:text-[#0078A8] disabled:opacity-40 disabled:cursor-not-allowed transition-colors rounded-sm"
    >
      <Icon size={13} /> {label}
    </button>
  );
}

/* ─── Protection card ───────────────────────────────────────────────── */
function ProtectionCard({ stats, tenant, onConfigure }) {
  const status = stats?.protectionStatus;
  const cfg = {
    Protected:      { Icon: ShieldCheck, textCls: 'text-green-700', title: 'Tenant is protected' },
    NeedsAttention: { Icon: ShieldAlert,  textCls: 'text-yellow-700', title: 'Attention needed' },
    NotProtected:   { Icon: Shield,       textCls: 'text-red-600',    title: 'Tenant is not protected' },
  }[status] || { Icon: Shield, textCls: 'text-gray-500', title: 'Unknown status' };
  const { Icon } = cfg;
  const bc = tenant?.backupConfig || {};
  const fmt = (d) => d ? new Date(d).toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'short' }) : '—';

  const rows = [
    ['Name',                           tenant?.name || tenant?.tenantId || '—'],
    ['Backup schedule',                 bc.scheduleEnabled ? 'Enabled' : 'Disabled'],
    ['Retention',                       bc.retentionDays ?? 90],
    ['Back up custom role definitions', bc.backupRoleDefinitions !== false ? 'Enabled' : 'Disabled'],
    ['Back up role permissions',        bc.backupRolePermissions !== false ? 'Enabled' : 'Disabled'],
    ['Back up role assignments',        bc.backupRoleAssignments !== false ? 'Enabled' : 'Disabled'],
    ['Back up assignment scopes',       bc.backupAssignmentScopes !== false ? 'Enabled' : 'Disabled'],
  ];

  return (
    <div className="bg-white border border-border-light rounded p-4 flex flex-col row-span-2">
      <div className="flex items-center gap-2 mb-3">
        <Icon size={18} className={cfg.textCls} />
        <h3 className={`text-sm font-semibold ${cfg.textCls}`}>{cfg.title}</h3>
      </div>

      {tenant ? (
        <>
          <table className="w-full text-xs">
            <tbody>
              {rows.map(([label, val]) => (
                <tr key={label} className="border-b border-gray-100">
                  <td className="py-1.5 text-text-secondary pr-3 whitespace-nowrap">{label}</td>
                  <td className="py-1.5 text-right font-medium text-[#333]">{val}</td>
                </tr>
              ))}
              <tr>
                <td className="py-1.5 text-text-secondary">Basic consent</td>
                <td className="py-1.5 text-right">
                  <span className={`inline-flex items-center gap-1 font-medium ${stats?.consentBasicGranted ? 'text-green-700' : 'text-yellow-600'}`}>
                    {stats?.consentBasicGranted ? <CheckCircle2 size={11} /> : <AlertCircle size={11} />}
                    {stats?.consentBasicGranted ? 'Granted' : 'Not granted'}
                  </span>
                </td>
              </tr>
              <tr>
                <td className="py-1.5 text-text-secondary">Restore consent</td>
                <td className="py-1.5 text-right">
                  <span className={`inline-flex items-center gap-1 font-medium ${stats?.consentRestoreGranted ? 'text-green-700' : 'text-yellow-600'}`}>
                    {stats?.consentRestoreGranted ? <CheckCircle2 size={11} /> : <AlertCircle size={11} />}
                    {stats?.consentRestoreGranted ? 'Granted' : 'Not granted'}
                  </span>
                </td>
              </tr>
            </tbody>
          </table>

          <button onClick={onConfigure} className="mt-auto pt-3 text-xs text-[#0078A8] hover:underline flex items-center gap-1 self-start">
            <Settings size={11} /> Configure backups
          </button>
        </>
      ) : (
        <div className="text-xs text-text-secondary mt-2">
          Configure backups to protect administrative functions.
        </div>
      )}
    </div>
  );
}

/* ─── Backups card ──────────────────────────────────────────────────── */
function BackupsCard({ count, backups, onClick }) {
  return (
    <div className="bg-white border border-border-light rounded p-4 flex flex-col">
      <h3 className="text-sm font-semibold text-[#222]">{count} backups</h3>
      <div className="flex-1">
        <BackupBarChart backups={backups} />
      </div>
      <button onClick={onClick} className="mt-2 text-xs text-[#0078A8] hover:underline font-semibold self-start uppercase tracking-wide">
        Show all
      </button>
    </div>
  );
}

/* ─── Tasks card ────────────────────────────────────────────────────── */
function TasksCard({ count, tasks, onClick }) {
  return (
    <div className="bg-white border border-border-light rounded p-4 flex flex-col">
      <h3 className="text-sm font-semibold text-[#222] mb-2">{count} tasks</h3>
      <div className="flex-1 space-y-1 min-h-[60px]">
        {tasks.slice(0, 5).map((t, i) => (
          <div key={i} className="flex items-center gap-2 text-xs">
            {t.status === 'Completed' ? (
              <CheckCircle2 size={11} className="text-green-600 flex-shrink-0" />
            ) : t.status === 'Failed' ? (
              <AlertCircle size={11} className="text-red-500 flex-shrink-0" />
            ) : (
              <Loader2 size={11} className="text-blue-500 animate-spin flex-shrink-0" />
            )}
            <span className="truncate text-[#333]">{t.kind || t.name || t.type || '—'}</span>
          </div>
        ))}
        {tasks.length === 0 && <p className="text-xs text-text-secondary">No tasks.</p>}
      </div>
      <button onClick={onClick} className="mt-2 text-xs text-[#0078A8] hover:underline font-semibold self-start uppercase tracking-wide">
        Show all
      </button>
    </div>
  );
}

/* ─── Unpacked objects card ─────────────────────────────────────────── */
function UnpackedCard({ count, stats, onClick }) {
  const segs = [
    { value: stats?.protectedRoleDefinitionsCount || Math.round(count * 0.3), color: '#5C6BC0', label: 'Roles' },
    { value: stats?.protectedRoleAssignmentsCount || Math.round(count * 0.4), color: '#26A69A', label: 'Assignments' },
    { value: Math.round(count * 0.2), color: '#FFA726', label: 'Permissions' },
    { value: Math.round(count * 0.1), color: '#EF5350', label: 'Scopes' },
  ];
  const total = segs.reduce((a, b) => a + b.value, 0) || count;

  return (
    <div className="bg-white border border-border-light rounded p-4 flex flex-col">
      <h3 className="text-sm font-semibold text-[#222] mb-2">{count} unpacked objects</h3>
      <div className="flex items-center gap-4">
        <DonutChart segments={segs} total={total} />
        <div className="flex-1 space-y-1">
          {segs.map((s) => (
            <div key={s.label} className="flex items-center gap-1.5 text-[11px] text-[#333]">
              <span className="w-2 h-2 rounded-sm flex-shrink-0" style={{ background: s.color }} />
              {s.label}
            </div>
          ))}
        </div>
      </div>
      <button onClick={onClick} className="mt-2 text-xs text-[#0078A8] hover:underline font-semibold self-start uppercase tracking-wide">
        Show all
      </button>
    </div>
  );
}

/* ─── Differences card ──────────────────────────────────────────────── */
function DifferencesCard({ count, diffs, onClick }) {
  const changeTypes = [
    { label: 'Link added',       color: 'text-blue-600'  },
    { label: 'New object',       color: 'text-green-600' },
    { label: 'Object hard deleted', color: 'text-red-600' },
    { label: 'Permission changed',  color: 'text-yellow-600' },
  ];

  return (
    <div className="bg-white border border-border-light rounded p-4 flex flex-col">
      <h3 className="text-sm font-semibold text-[#222] mb-2">{count} differences</h3>
      <div className="flex-1 space-y-1 min-h-[60px]">
        {diffs.length > 0 ? diffs.slice(0, 4).map((d, i) => (
          <div key={i} className="flex items-center justify-between text-xs">
            <span className="text-[#333] truncate">{d.label || d.RoleName || d.name}</span>
            {d.count != null && <span className="font-semibold text-[#333] ml-2">{d.count}</span>}
          </div>
        )) : changeTypes.map((ct, i) => (
          <div key={i} className="flex items-center justify-between text-xs">
            <span className={`${ct.color} truncate`}>{ct.label}</span>
          </div>
        ))}
        {diffs.length === 0 && count === 0 && <p className="text-xs text-text-secondary">No differences found.</p>}
      </div>
      <button onClick={onClick} className="mt-2 text-xs text-[#0078A8] hover:underline font-semibold self-start uppercase tracking-wide">
        Show all
      </button>
    </div>
  );
}

/* ─── Errors card ───────────────────────────────────────────────────── */
function ErrorsCard({ count, errors, onClick }) {
  return (
    <div className="bg-white border border-border-light rounded p-4 flex flex-col">
      <h3 className="text-sm font-semibold text-[#222] mb-2">{count} errors</h3>
      <div className="flex-1 space-y-1">
        {errors.slice(0, 5).map((e, i) => (
          <div key={i} className="flex items-start gap-1.5 text-xs">
            <AlertCircle size={11} className="text-red-500 flex-shrink-0 mt-0.5" />
            <span className="text-[#333] truncate">{e.message || e.Message}</span>
          </div>
        ))}
        {errors.length === 0 && <p className="text-xs text-text-secondary">No errors.</p>}
      </div>
      <button onClick={onClick} className="mt-2 text-xs text-[#0078A8] hover:underline font-semibold self-start uppercase tracking-wide">
        Show all
      </button>
    </div>
  );
}

/* ─── Create Backup modal ───────────────────────────────────────────── */
function CreateBackupModal({ tenant, busy, onClose, onConfirm }) {
  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center">
      <div className="bg-white rounded shadow-xl w-full max-w-md mx-4">
        <div className="flex items-center justify-between px-6 py-4 border-b border-border-light">
          <h2 className="text-sm font-semibold text-[#222]">Create backup</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl leading-none">×</button>
        </div>
        <div className="px-6 py-5 text-sm text-[#333]">
          Do you want to back up administrative functions for{' '}
          <strong>{tenant?.name || tenant?.tenantId}</strong>?
          <p className="text-xs text-text-secondary mt-2">
            An Administrative Function Backup task will be created and appear in Tasks.
          </p>
        </div>
        <div className="flex justify-end gap-2 px-6 py-3 border-t border-border-light bg-gray-50">
          <button onClick={onClose} disabled={busy}
            className="px-4 py-1.5 text-sm border border-border-light rounded hover:bg-white disabled:opacity-60">
            Cancel
          </button>
          <button onClick={onConfirm} disabled={busy}
            className="px-4 py-1.5 text-sm bg-[#0078A8] text-white rounded hover:bg-[#006090] disabled:opacity-60 flex items-center gap-2">
            {busy && <Loader2 size={13} className="animate-spin" />} Create
          </button>
        </div>
      </div>
    </div>
  );
}

/* ─── No tenant placeholder ─────────────────────────────────────────── */
function NoTenant({ onAdd }) {
  return (
    <div className="flex-1 flex items-center justify-center p-12">
      <div className="text-center max-w-sm">
        <Shield size={48} className="mx-auto mb-4 text-[#C8C8C8]" />
        <h2 className="text-base font-semibold text-[#333] mb-1">Tenant is not configured</h2>
        <p className="text-sm text-text-secondary mb-4">
          Add a Microsoft Entra ID tenant to start protecting administrative functions.
        </p>
        <button onClick={onAdd}
          className="px-4 py-2 bg-[#0078A8] text-white text-sm rounded hover:bg-[#006090]">
          Add Tenant
        </button>
      </div>
    </div>
  );
}

/* ─── Main Dashboard ────────────────────────────────────────────────── */
export default function Dashboard() {
  const navigate = useNavigate();
  const { selected, tenants, loading: tLoading } = useTenant();
  const [stats, setStats]       = useState(null);
  const [backups, setBackups]   = useState([]);
  const [jobs, setJobs]         = useState([]);
  const [errors, setErrors]     = useState([]);
  const [loading, setLoading]   = useState(true);
  const [notice, setNotice]     = useState(null);
  const [busy, setBusy]         = useState(false);
  const [showCreate, setShowCreate] = useState(false);
  const [activeModal, setActiveModal] = useState(null);

  const load = async () => {
    setLoading(true);
    try {
      const [statsR, backupsR, jobsR, eventsR] = await Promise.allSettled([
        api.dashboard(selected?.tenantId),
        api.backups(),
        api.jobs(),
        api.events('Error'),
      ]);
      if (statsR.status === 'fulfilled')   setStats(statsR.value?.data || null);
      if (backupsR.status === 'fulfilled') setBackups(backupsR.value?.data ?? []);
      if (jobsR.status === 'fulfilled')    setJobs(jobsR.value?.data ?? []);
      if (eventsR.status === 'fulfilled')  setErrors(eventsR.value?.data?.items ?? []);
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
      setNotice({ kind: 'success', msg: 'The Administrative Function Backup task has been created.', jobId: r?.data?.id });
      load();
    } catch (e) {
      setNotice({ kind: 'error', msg: e.message });
    } finally {
      setBusy(false);
    }
  };

  const refreshDiffs = async () => {
    if (!selected) return;
    setBusy(true);
    try {
      await api.refreshDifferences(selected.tenantId, stats?.lastBackupId);
      setNotice({ kind: 'success', msg: 'Refresh Differences task has been created.' });
    } catch (e) {
      setNotice({ kind: 'error', msg: e.message });
    } finally {
      setBusy(false);
    }
  };

  const tenantBackups = selected
    ? backups.filter((b) => !b.tenantId || b.tenantId === selected.tenantId)
    : backups;

  if (tLoading || loading) {
    return (
      <div className="p-8 text-text-secondary flex items-center gap-2 text-sm">
        <Loader2 className="animate-spin" size={16} /> Loading…
      </div>
    );
  }

  if (tenants.length === 0) {
    return <NoTenant onAdd={() => navigate('/tenants')} />;
  }

  return (
    <div className="p-5">
      {/* Action buttons */}
      <div className="flex flex-wrap items-center gap-2 mb-4">
        <ActionBtn Icon={Database}    label="Manage Backups"     onClick={() => setActiveModal('manageBackups')} />
        <ActionBtn Icon={RotateCcw}   label="Manage Restores"    onClick={() => setActiveModal('manageRestores')} />
        <ActionBtn Icon={Play}        label="Create Backup"      onClick={() => setShowCreate(true)} disabled={busy || !selected} />
        <ActionBtn Icon={Archive}     label="Unpack Backup"      onClick={() => setActiveModal('unpackBackup')} />
        <ActionBtn Icon={RefreshCw}   label="Refresh Differences" onClick={refreshDiffs} disabled={busy || !selected} />
      </div>

      {/* Notice bar */}
      {notice && (
        <div className={`mb-4 px-4 py-2.5 rounded text-xs flex items-center justify-between ${
          notice.kind === 'error' ? 'bg-red-50 border border-red-200 text-red-800' : 'bg-green-50 border border-green-200 text-green-900'
        }`}>
          <span>{notice.msg}</span>
          <div className="flex items-center gap-3 ml-4">
            {notice.jobId && (
              <button onClick={() => navigate('/tasks')} className="underline font-semibold whitespace-nowrap">
                View details
              </button>
            )}
            <button onClick={() => setNotice(null)} className="hover:opacity-70">×</button>
          </div>
        </div>
      )}

      {/* Card grid — mirrors ODR layout */}
      <div className="grid grid-cols-3 gap-4" style={{ gridTemplateRows: 'auto auto auto' }}>
        {/* Col 1: Protection (row-span-2) */}
        <ProtectionCard
          stats={stats}
          tenant={stats?.tenant || selected}
          onConfigure={() => setActiveModal('manageBackups')}
        />

        {/* Col 2, row 1: Backups with bar chart */}
        <BackupsCard
          count={stats?.backupsCount ?? tenantBackups.length}
          backups={tenantBackups}
          onClick={() => navigate('/backups')}
        />

        {/* Col 3, row 1: Tasks */}
        <TasksCard
          count={stats?.tasksTotalCount ?? jobs.length}
          tasks={jobs.slice(0, 5)}
          onClick={() => navigate('/tasks')}
        />

        {/* Col 2, row 2: Unpacked Objects with donut */}
        <UnpackedCard
          count={stats?.unpackedBackupsCount ?? 0}
          stats={stats}
          onClick={() => navigate('/unpacked')}
        />

        {/* Col 3, row 2: Differences */}
        <DifferencesCard
          count={stats?.differencesCount ?? 0}
          diffs={[]}
          onClick={() => navigate('/differences')}
        />

        {/* Row 3: Errors — full width */}
        <div className="col-span-3">
          <ErrorsCard
            count={stats?.errorsCount ?? errors.length}
            errors={errors.slice(0, 5)}
            onClick={() => navigate('/events')}
          />
        </div>
      </div>

      {showCreate && (
        <CreateBackupModal
          tenant={selected}
          busy={busy}
          onClose={() => setShowCreate(false)}
          onConfirm={createBackup}
        />
      )}

      {activeModal === 'manageBackups' && (
        <ManageBackupsModal onClose={() => setActiveModal(null)} />
      )}

      {activeModal === 'manageRestores' && (
        <ManageRestoresModal onClose={() => setActiveModal(null)} />
      )}

      {activeModal === 'unpackBackup' && (
        <UnpackBackupModal
          onClose={() => setActiveModal(null)}
          onUnpacked={(jobId) => {
            setActiveModal(null);
            setNotice({ kind: 'success', msg: 'Unpack task created.', jobId });
            load();
          }}
        />
      )}
    </div>
  );
}
