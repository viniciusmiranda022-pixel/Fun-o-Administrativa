import { useState, useEffect, useMemo } from 'react';
import type { ReactNode } from 'react';
import { buildBars } from '../utils/chartUtils';
import { useNavigate } from 'react-router-dom';
import {
  Settings, ArrowDownToLine, RotateCcw, Package,
  CheckCircle2, X, Check, AlertCircle,
} from 'lucide-react';
import { useApp } from '../context/AppContext';
import Modal from '../components/common/Modal';
import BackupUnpackingModal from '../components/BackupUnpackingModal';
import { api } from '../api/client';
import type { BackupSnapshot, TenantEntry, TaskEntry, EventEntry } from '../types/api';

interface ToolBtnProps {
  icon: LucideIcon;
  label: string;
  onClick?: () => void;
  disabled?: boolean;
}

/* ── Toolbar button ── */
function ToolBtn({ icon: Icon, label, onClick, disabled }: ToolBtnProps) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      className="flex items-center gap-1.5 px-3 py-2 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2] disabled:opacity-40 disabled:cursor-not-allowed"
    >
      <Icon size={13} />
      {label}
    </button>
  );
}

interface Bar {
  label: string;
  h: number;
}

interface BarChartProps {
  bars?: Bar[];
  fullWidth?: boolean;
}

/* ── Simple bar chart ── */
function BarChart({ bars = [], fullWidth = false }: BarChartProps) {
  return (
    <div className={`flex items-end gap-1 ${fullWidth ? 'w-full' : 'flex-1'} px-1 pt-2`} style={{ height: fullWidth ? 100 : 80 }}>
      {bars.map((b, i) => (
        <div key={i} className="flex-1 flex flex-col items-center gap-0.5">
          <div
            className="w-full rounded-sm"
            style={{ height: `${b.h}%`, backgroundColor: '#6CB33F', minHeight: 2 }}
          />
          {fullWidth && <span className="text-[9px] text-[#888] mt-0.5">{b.label}</span>}
        </div>
      ))}
    </div>
  );
}

interface DonutCounts {
  definitions: number;
  assignments: number;
}

/* ── SVG Donut chart from unpacked objects counts ── */
function buildDonutPaths(counts: DonutCounts | undefined) {
  const segments = [
    { label: 'Role Definitions', count: counts?.definitions ?? 0, color: '#0096D6' },
    { label: 'Role Assignments', count: counts?.assignments ?? 0, color: '#6CB33F' },
  ];
  const total = segments.reduce((s, x) => s + x.count, 0);
  const cx = 60; const cy = 60; const r = 48; const ir = 30;
  let angle = -Math.PI / 2;
  const paths = segments.filter(s => s.count > 0).map((seg) => {
    const sweep = (seg.count / total) * 2 * Math.PI;
    const x1 = cx + r * Math.cos(angle);
    const y1 = cy + r * Math.sin(angle);
    const x2 = cx + r * Math.cos(angle + sweep);
    const y2 = cy + r * Math.sin(angle + sweep);
    const ix1 = cx + ir * Math.cos(angle);
    const iy1 = cy + ir * Math.sin(angle);
    const ix2 = cx + ir * Math.cos(angle + sweep);
    const iy2 = cy + ir * Math.sin(angle + sweep);
    const large = sweep > Math.PI ? 1 : 0;
    const d = `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} L ${ix2} ${iy2} A ${ir} ${ir} 0 ${large} 0 ${ix1} ${iy1} Z`;
    angle += sweep;
    return { d, color: seg.color };
  });
  return { segments, total, paths };
}

function DonutChart({ counts }: { counts: DonutCounts }) {
  const { segments, total, paths } = useMemo(() => buildDonutPaths(counts), [counts?.definitions, counts?.assignments]);

  if (total === 0) {
    return (
      <div className="flex items-center gap-4 mt-2">
        <svg width="120" height="120" viewBox="0 0 120 120">
          <circle cx="60" cy="60" r="48" fill="#F2F2F2" />
          <circle cx="60" cy="60" r="30" fill="white" />
          <text x="60" y="57" textAnchor="middle" fontSize="11" fill="#888">0</text>
          <text x="60" y="68" textAnchor="middle" fontSize="7" fill="#888">objects</text>
        </svg>
        <div className="text-xs text-[#888]">No unpacked objects yet.</div>
      </div>
    );
  }

  return (
    <div className="flex items-center gap-4 mt-2">
      <svg width="120" height="120" viewBox="0 0 120 120">
        {paths.map((p, i) => (
          <path key={i} d={p.d} fill={p.color} />
        ))}
        <text x="60" y="57" textAnchor="middle" fontSize="14" fontWeight="bold" fill="#333">{total}</text>
        <text x="60" y="68" textAnchor="middle" fontSize="7" fill="#888">objects</text>
      </svg>
      <div className="flex flex-col gap-1">
        {segments.map((s) => (
          <div key={s.label} className="flex items-center gap-1.5 text-[10px] text-[#555]">
            <span className="w-2.5 h-2.5 rounded-sm flex-shrink-0" style={{ backgroundColor: s.color }} />
            {s.label} ({s.count})
          </div>
        ))}
      </div>
    </div>
  );
}

interface CardProps {
  title: string;
  icon?: ReactNode;
  showAll?: boolean;
  onShowAll?: () => void;
  children: ReactNode;
}

/* ── Card shell ── */
function Card({ title, icon, showAll, onShowAll, children }: CardProps) {
  return (
    <div className="bg-white border border-[#DEE2E6] shadow-sm flex flex-col min-h-[220px]">
      <div className="px-4 pt-3 pb-2 border-b border-[#DEE2E6] flex items-center gap-2">
        {icon}
        <span className="text-sm font-semibold text-[#222]">{title}</span>
      </div>
      <div className="flex-1 px-4 py-3 overflow-hidden">{children}</div>
      {showAll && (
        <div className="px-4 pb-3">
          <button onClick={onShowAll} className="text-xs text-[#0078A8] hover:underline font-semibold">
            SHOW ALL
          </button>
        </div>
      )}
    </div>
  );
}

interface InfoRowProps {
  label: string;
  value: ReactNode;
  valueColor?: string;
}

/* ── Info row in protected card ── */
function InfoRow({ label, value, valueColor = '#0078A8' }: InfoRowProps) {
  return (
    <div className="flex items-start justify-between text-xs py-0.5">
      <span className="text-[#555]">{label}:</span>
      <span className="font-medium ml-2" style={{ color: valueColor }}>{value}</span>
    </div>
  );
}

/* ═══════════════════════ MODALS ═══════════════════════ */

interface BackupOpts {
  roleDefs: boolean;
  rolePerms: boolean;
  roleAssignments: boolean;
  scopes: boolean;
  principals: boolean;
}

function ConfigureBackupModal({ tenant, onClose }: { tenant: TenantEntry | null; onClose: () => void }) {
  const [schedule, setSchedule] = useState('Enabled');
  const [retention, setRetention] = useState(1825);
  const [opts, setOpts] = useState<BackupOpts>({
    roleDefs: true, rolePerms: true, roleAssignments: true,
    scopes: true, principals: false,
  });

  function toggleOpt(k: keyof BackupOpts) {
    setOpts((o) => ({ ...o, [k]: !o[k] }));
  }

  return (
    <Modal title="Configure backup" onClose={onClose} maxWidth="max-w-lg">
      <div className="px-6 py-4 space-y-4">
        <div className="flex items-center gap-3">
          <label className="text-xs text-[#333] w-40">Schedule</label>
          <select
            value={schedule}
            onChange={(e) => setSchedule(e.target.value)}
            className="border border-[#DEE2E6] px-2 py-1 text-xs bg-white"
          >
            <option>Enabled</option>
            <option>Disabled</option>
          </select>
        </div>
        <div className="flex items-center gap-3">
          <label className="text-xs text-[#333] w-40">Retention policy (days)</label>
          <input
            type="number"
            value={retention}
            onChange={(e) => setRetention(Number(e.target.value))}
            className="border border-[#DEE2E6] px-2 py-1 text-xs w-24"
          />
        </div>
        <div className="space-y-2">
          {([
            ['roleDefs', 'Back up custom administrative role definitions'],
            ['rolePerms', 'Back up administrative role permissions'],
            ['roleAssignments', 'Back up administrative role assignments'],
            ['scopes', 'Back up administrative assignment scopes'],
            ['principals', 'Back up referenced principals metadata'],
          ] as [keyof BackupOpts, string][]).map(([k, label]) => (
            <label key={k} className="flex items-center gap-2 text-xs text-[#333] cursor-pointer">
              <input type="checkbox" checked={opts[k]} onChange={() => toggleOpt(k)} />
              {label}
            </label>
          ))}
        </div>
        <div className="flex justify-end gap-2 pt-2 border-t border-[#DEE2E6]">
          <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
          <button onClick={onClose} className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090]">Save</button>
        </div>
      </div>
    </Modal>
  );
}

function ManageBackupsModal({ onClose, tenants }: { onClose: () => void; tenants: TenantEntry[] }) {
  const [configuringTenant, setConfiguringTenant] = useState<TenantEntry | null>(null);
  const rows = tenants.map((t) => ({
    ...t,
    schedule: 'Enabled',
    retention: 1825,
    advanced: 'Enabled',
    consent: t.consents?.basic?.granted ? 'Granted' : 'Not Granted',
  }));

  return (
    <>
      <Modal title="Manage Backups" onClose={onClose} maxWidth="max-w-3xl">
        <div className="flex" style={{ minHeight: 360 }}>
          <div className="w-36 border-r border-[#DEE2E6] bg-[#FAFAFA] p-3 flex-shrink-0">
            <div className="border-l-4 border-[#E67E22] pl-2">
              <span className="text-xs font-semibold text-[#333]">Tenants</span>
            </div>
          </div>
          <div className="flex-1 flex flex-col p-4">
            <div className="mb-3">
              <input placeholder="Search tenants..." className="border border-[#DEE2E6] px-3 py-1.5 text-xs w-full" />
            </div>
            <div className="flex-1 overflow-auto">
              {rows.length === 0 ? (
                <div className="text-xs text-[#888] p-4">No tenants configured.</div>
              ) : (
                <table className="w-full text-xs">
                  <thead>
                    <tr className="bg-[#F2F2F2]">
                      {['Tenant', 'Schedule', 'Retention', 'Advanced data', 'Consent', 'Action'].map((h) => (
                        <th key={h} className="text-left px-2 py-1.5 border-b border-[#DEE2E6] font-semibold text-[#333]">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((r) => (
                      <tr key={r.tenantId} className="border-b border-[#DEE2E6] hover:bg-[#F9FAFB]">
                        <td className="px-2 py-1.5">{r.name}</td>
                        <td className="px-2 py-1.5 text-[#0078A8]">{r.schedule}</td>
                        <td className="px-2 py-1.5 text-[#0078A8]">{r.retention}</td>
                        <td className="px-2 py-1.5 text-[#0078A8]">{r.advanced}</td>
                        <td className="px-2 py-1.5 text-[#28A745]">{r.consent}</td>
                        <td className="px-2 py-1.5">
                          <button
                            onClick={() => setConfiguringTenant(r)}
                            className="px-2 py-0.5 border border-[#DEE2E6] text-[#0078A8] hover:bg-[#F2F2F2] text-xs"
                          >
                            EDIT
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
            <div className="flex justify-end pt-3 border-t border-[#DEE2E6] mt-3">
              <button onClick={onClose} className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090]">Finish</button>
            </div>
          </div>
        </div>
      </Modal>
      {configuringTenant && (
        <ConfigureBackupModal tenant={configuringTenant} onClose={() => setConfiguringTenant(null)} />
      )}
    </>
  );
}

function ManageRestoresModal({ onClose, tenants }: { onClose: () => void; tenants: TenantEntry[] }) {
  const rows = tenants.map((t) => ({
    ...t,
    basic: t.consents?.basic?.granted ? 'Granted' : 'Not Granted',
    restore: t.consents?.restore?.granted ? 'Granted' : 'Not Granted',
    status: (t.consents?.basic?.granted && t.consents?.restore?.granted) ? 'All Granted' : 'Incomplete',
  }));

  return (
    <Modal title="Manage Restores" onClose={onClose} maxWidth="max-w-2xl">
      <div className="p-4">
        {rows.length === 0 ? (
          <div className="text-xs text-[#888] p-4">No tenants configured.</div>
        ) : (
          <table className="w-full text-xs">
            <thead>
              <tr className="bg-[#F2F2F2]">
                {['Tenant', 'Basic Consent', 'Restore Consent', 'Consent Status'].map((h) => (
                  <th key={h} className="text-left px-3 py-2 border-b border-[#DEE2E6] font-semibold text-[#333]">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.tenantId} className="border-b border-[#DEE2E6]">
                  <td className="px-3 py-2">{r.name}</td>
                  <td className="px-3 py-2">
                    <span className={r.consents?.basic?.granted ? 'text-[#28A745]' : 'text-[#DC3545]'}>
                      {r.consents?.basic?.granted ? '✓ ' : '✗ '}{r.basic}
                    </span>
                  </td>
                  <td className="px-3 py-2">
                    <span className={r.consents?.restore?.granted ? 'text-[#28A745]' : 'text-[#DC3545]'}>
                      {r.consents?.restore?.granted ? '✓ ' : '✗ '}{r.restore}
                    </span>
                  </td>
                  <td className="px-3 py-2 font-medium" style={{ color: r.status === 'All Granted' ? '#28A745' : '#DC3545' }}>{r.status}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        <div className="flex justify-end mt-4">
          <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
        </div>
      </div>
    </Modal>
  );
}

function CreateBackupModal({ onClose, tenant }: { onClose: () => void; tenant: TenantEntry | null }) {
  const navigate = useNavigate();
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function handleCreate() {
    setSubmitting(true);
    setError(null);
    try {
      await api.runBackup();
      onClose();
      navigate('/tasks');
    } catch (err) {
      setError((err instanceof Error ? err.message : null) || 'Failed to start backup');
      setSubmitting(false);
    }
  }

  return (
    <Modal title="Create Backup" onClose={onClose} maxWidth="max-w-md">
      <div className="px-6 py-5">
        {error ? (
          <div className="space-y-3">
            <div className="flex items-center gap-2 text-[#DC3545] text-sm">
              <X size={18} />
              {error}
            </div>
            <div className="flex justify-end gap-2">
              <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Close</button>
            </div>
          </div>
        ) : (
          <>
            <p className="text-sm text-[#333] mb-2">
              Do you want to back up the administrative functions for this tenant?
            </p>
            <p className="text-sm font-semibold text-[#0078A8] mb-5">{tenant?.name ?? 'Unknown'}</p>
            <div className="flex justify-end gap-2">
              <button onClick={onClose} disabled={submitting} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50 disabled:opacity-50">Cancel</button>
              <button onClick={handleCreate} disabled={submitting} className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090] disabled:opacity-50">
                {submitting ? 'Starting...' : 'Create'}
              </button>
            </div>
          </>
        )}
      </div>
    </Modal>
  );
}

/* ═══════════════════════ MAIN PAGE ═══════════════════════ */
export default function Dashboard() {
  const navigate = useNavigate();
  const { selectedTenant, tenants } = useApp();

  const [modal, setModal] = useState<string | null>(null);
  const [backups, setBackups] = useState<BackupSnapshot[]>([]);
  const [tasks, setTasks] = useState<TaskEntry[]>([]);
  const [events, setEvents] = useState<EventEntry[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    setLoading(true);
    Promise.all([
      api.backups().catch(() => null),
      api.tasks().catch(() => null),
      api.events('Error').catch(() => null),
    ]).then(([b, t, e]) => {
      setBackups(b?.data ?? []);
      setTasks(t?.data?.items ?? []);
      setEvents(e?.data?.items ?? []);
    }).finally(() => setLoading(false));
  }, []);

  const backupCount = backups.length;
  const taskCount = tasks.length;
  const errorCount = events.filter(e => (e.severity || '').toLowerCase() === 'error').length;
  const bars = useMemo(() => buildBars(backups, 10), [backups]);
  const firstBar = bars[0];
  const lastBar = bars[bars.length - 1];

  const tenantList = selectedTenant ? [selectedTenant] : tenants;

  return (
    <div className="flex flex-col">
      {/* ── Toolbar ── */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-2 flex-shrink-0">
        <ToolBtn icon={Settings} label="MANAGE BACKUPS" onClick={() => setModal('manageBackups')} />
        <ToolBtn icon={ArrowDownToLine} label="MANAGE RESTORES" onClick={() => setModal('manageRestores')} />
        <ToolBtn icon={RotateCcw} label="CREATE BACKUP" onClick={() => setModal('createBackup')} disabled={!selectedTenant} />
        <ToolBtn icon={Package} label="UNPACK BACKUP" onClick={() => setModal('unpack')} disabled={backups.length === 0} />
      </div>

      {/* ── Cards grid ── */}
      <div className="p-4 grid grid-cols-3 gap-3">

        {/* Card 1: Tenant protected */}
        <Card
          title={selectedTenant ? 'Tenant is protected' : 'No tenant configured'}
          icon={<CheckCircle2 size={16} className={selectedTenant ? 'text-[#28A745]' : 'text-[#FFC107]'} />}
        >
          {selectedTenant ? (
            <div className="space-y-0.5">
              <InfoRow label="Name" value={
                <button className="text-[#0078A8] hover:underline" onClick={() => navigate('/tenants')}>
                  {selectedTenant.name}
                </button>
              } />
              <InfoRow label="Backup schedule" value="Enabled" />
              <InfoRow label="Retention" value="1825" />
              <InfoRow label="Back up role definitions" value="Enabled" />
              <InfoRow label="Back up role assignments" value="Enabled" />
              <InfoRow label="Basic consent" value={selectedTenant.consents?.basic?.granted ? 'Granted' : 'Not Granted'}
                valueColor={selectedTenant.consents?.basic?.granted ? '#28A745' : '#DC3545'} />
              <InfoRow label="Restore consent" value={selectedTenant.consents?.restore?.granted ? 'Granted' : 'Not Granted'}
                valueColor={selectedTenant.consents?.restore?.granted ? '#28A745' : '#DC3545'} />
            </div>
          ) : (
            <div className="flex items-center gap-2 text-xs text-[#FFC107] mt-2">
              <AlertCircle size={14} />
              <span>Add a tenant to enable backups.</span>
            </div>
          )}
        </Card>

        {/* Card 2: Backups */}
        <Card title={`${backupCount} backup${backupCount !== 1 ? 's' : ''}`} showAll onShowAll={() => navigate('/backups')}>
          <BarChart bars={bars} />
          <div className="flex justify-between text-[10px] text-[#999] mt-1 px-1">
            <span>{firstBar?.label}</span>
            <span>{lastBar?.label}</span>
          </div>
        </Card>

        {/* Card 3: Tasks */}
        <Card title={`${taskCount} task${taskCount !== 1 ? 's' : ''}`} showAll onShowAll={() => navigate('/tasks')}>
          <div className="space-y-1.5 mt-1">
            {tasks.length === 0 ? (
              <span className="text-xs text-[#888]">No tasks yet.</span>
            ) : tasks.slice(0, 5).map((t, i) => (
              <div key={i} className="flex items-start gap-1.5 text-xs text-[#555]">
                <Check size={12} className={`mt-0.5 flex-shrink-0 ${t.status === 'Failed' ? 'text-[#DC3545]' : 'text-[#28A745]'}`} />
                <span className="truncate">{t.name || t.type}</span>
              </div>
            ))}
          </div>
        </Card>

        {/* Card 4: Unpacked objects (role defs + assignments from latest backup) */}
        <Card title={`${(backups[0]?.roleDefinitionsCount ?? 0) + (backups[0]?.roleAssignmentsCount ?? 0)} unpacked objects`} showAll onShowAll={() => navigate('/unpacked')}>
          <DonutChart counts={{
            definitions: backups[0]?.roleDefinitionsCount ?? 0,
            assignments: backups[0]?.roleAssignmentsCount ?? 0,
          }} />
        </Card>

        {/* Card 5: Differences */}
        <Card title="Differences" showAll onShowAll={() => navigate('/differences')}>
          <div className="text-xs text-[#888] mt-2">
            Run a comparison to see differences between backup and current tenant state.
          </div>
        </Card>

        {/* Card 6: Errors */}
        <Card title={`${errorCount} error${errorCount !== 1 ? 's' : ''}`} showAll onShowAll={() => navigate('/events')}>
          <div className="space-y-1.5 mt-1">
            {events.length === 0 ? (
              <span className="text-xs text-[#888]">No errors logged.</span>
            ) : events.slice(0, 5).map((e, i) => (
              <div key={i} className="flex items-start gap-1.5 text-xs text-[#555]">
                <X size={12} className="text-[#DC3545] mt-0.5 flex-shrink-0" />
                <span className="line-clamp-2">{e.message || e.description}</span>
              </div>
            ))}
          </div>
        </Card>

        {/* Card 7: Consent Status */}
        <Card title="Consent Status" icon={<AlertCircle size={16} className="text-[#0096D6]" />}>
          <div className="space-y-3 mt-2">
            <div className="flex items-center gap-2 text-xs">
              <CheckCircle2 size={14} className={selectedTenant?.consents?.basic?.granted ? 'text-[#28A745]' : 'text-[#DC3545]'} />
              <span className="text-[#333]">Basic consent:</span>
              <span className="font-semibold" style={{ color: selectedTenant?.consents?.basic?.granted ? '#28A745' : '#DC3545' }}>
                {selectedTenant ? (selectedTenant.consents?.basic?.granted ? 'Granted' : 'Not Granted') : '—'}
              </span>
            </div>
            <div className="flex items-center gap-2 text-xs">
              <CheckCircle2 size={14} className={selectedTenant?.consents?.restore?.granted ? 'text-[#28A745]' : 'text-[#DC3545]'} />
              <span className="text-[#333]">Restore consent:</span>
              <span className="font-semibold" style={{ color: selectedTenant?.consents?.restore?.granted ? '#28A745' : '#DC3545' }}>
                {selectedTenant ? (selectedTenant.consents?.restore?.granted ? 'Granted' : 'Not Granted') : '—'}
              </span>
            </div>
            {!selectedTenant && (
              <div className="flex items-center gap-2 text-xs text-[#FFC107]">
                <AlertCircle size={14} />
                <span>No tenant configured. Add a tenant to enable backups.</span>
              </div>
            )}
          </div>
        </Card>

      </div>

      {/* ── Modals ── */}
      {modal === 'manageBackups' && (
        <ManageBackupsModal onClose={() => setModal(null)} tenants={tenantList} />
      )}
      {modal === 'manageRestores' && (
        <ManageRestoresModal onClose={() => setModal(null)} tenants={tenantList} />
      )}
      {modal === 'createBackup' && (
        <CreateBackupModal onClose={() => setModal(null)} tenant={selectedTenant} />
      )}
      {modal === 'unpack' && (
        <BackupUnpackingModal onClose={() => setModal(null)} backups={backups} />
      )}
    </div>
  );
}
