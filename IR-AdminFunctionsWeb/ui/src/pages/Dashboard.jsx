import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  Settings, ArrowDownToLine, RotateCcw, Package,
  CheckCircle2, X, ChevronDown, Check, AlertCircle
} from 'lucide-react';
import { useApp } from '../context/AppContext.jsx';
import Modal from '../components/common/Modal.jsx';

/* ── Toolbar button ── */
function ToolBtn({ icon: Icon, label, onClick, disabled }) {
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

/* ── Simple bar chart ── */
function BarChart({ fullWidth = false }) {
  const bars = [
    { label: 'Apr 16', h: 60 }, { label: 'Apr 20', h: 45 }, { label: 'Apr 24', h: 70 },
    { label: 'Apr 28', h: 55 }, { label: 'May 2',  h: 80 }, { label: 'May 6',  h: 65 },
    { label: 'May 10', h: 50 }, { label: 'May 14', h: 90 }, { label: 'May 18', h: 75 },
    { label: 'May 20', h: 40 },
  ];
  return (
    <div className={`flex items-end gap-1 ${fullWidth ? 'w-full' : 'flex-1'} px-1 pt-2`} style={{ height: fullWidth ? 100 : 80 }}>
      {bars.map((b, i) => (
        <div key={i} className="flex-1 flex flex-col items-center gap-0.5">
          <div
            className="w-full rounded-sm"
            style={{ height: `${b.h}%`, backgroundColor: '#6CB33F' }}
          />
          {fullWidth && <span className="text-[9px] text-[#888] mt-0.5">{b.label}</span>}
        </div>
      ))}
    </div>
  );
}

/* ── SVG Donut chart ── */
function DonutChart() {
  const segments = [
    { label: 'Custom Roles', count: 32, color: '#0096D6' },
    { label: 'Role Permissions', count: 128, color: '#6CB33F' },
    { label: 'Role Assignments', count: 97, color: '#FFC107' },
    { label: 'Assignment Scopes', count: 55, color: '#17A2B8' },
    { label: 'Referenced Principals', count: 25, color: '#DC3545' },
  ];
  const total = segments.reduce((s, x) => s + x.count, 0);
  const cx = 60; const cy = 60; const r = 48; const ir = 30;
  let angle = -Math.PI / 2;
  const paths = segments.map((seg) => {
    const pct = seg.count / total;
    const sweep = pct * 2 * Math.PI;
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

/* ── Card shell ── */
function Card({ title, icon, showAll, onShowAll, children }) {
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

/* ── Info row in protected card ── */
function InfoRow({ label, value, valueColor = '#0078A8' }) {
  return (
    <div className="flex items-start justify-between text-xs py-0.5">
      <span className="text-[#555]">{label}:</span>
      <span className="font-medium ml-2" style={{ color: valueColor }}>{value}</span>
    </div>
  );
}

/* ═══════════════════════ MODALS ═══════════════════════ */

/* Configure Backup Modal (nested) */
function ConfigureBackupModal({ tenant, onClose }) {
  const [schedule, setSchedule] = useState('Enabled');
  const [retention, setRetention] = useState(1825);
  const [opts, setOpts] = useState({
    roleDefs: true, rolePerms: true, roleAssignments: true,
    scopes: true, principals: false,
  });

  function toggleOpt(k) {
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
          {[
            ['roleDefs', 'Back up custom administrative role definitions'],
            ['rolePerms', 'Back up administrative role permissions'],
            ['roleAssignments', 'Back up administrative role assignments'],
            ['scopes', 'Back up administrative assignment scopes'],
            ['principals', 'Back up referenced principals metadata'],
          ].map(([k, label]) => (
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

/* Manage Backups Modal */
function ManageBackupsModal({ onClose, tenants }) {
  const [configuringTenant, setConfiguringTenant] = useState(null);

  const rows = tenants.map((t) => ({
    ...t,
    schedule: 'Enabled',
    retention: 1825,
    advanced: 'Enabled',
    consent: 'Granted',
  }));

  return (
    <>
      <Modal title="Manage Backups" onClose={onClose} maxWidth="max-w-3xl">
        <div className="flex" style={{ minHeight: 360 }}>
          {/* Left panel */}
          <div className="w-36 border-r border-[#DEE2E6] bg-[#FAFAFA] p-3 flex-shrink-0">
            <div className="border-l-4 border-[#E67E22] pl-2">
              <span className="text-xs font-semibold text-[#333]">Tenants</span>
            </div>
          </div>
          {/* Right panel */}
          <div className="flex-1 flex flex-col p-4">
            <div className="mb-3">
              <input
                placeholder="Search tenants..."
                className="border border-[#DEE2E6] px-3 py-1.5 text-xs w-full"
              />
            </div>
            <div className="flex-1 overflow-auto">
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

/* Manage Restores Modal */
function ManageRestoresModal({ onClose, tenants }) {
  const rows = tenants.map((t) => ({
    ...t,
    basic: t.consents?.basic?.granted ? 'Granted' : 'Not Granted',
    restore: t.consents?.restore?.granted ? 'Granted' : 'Not Granted',
    status: (t.consents?.basic?.granted && t.consents?.restore?.granted) ? 'All Granted' : 'Incomplete',
  }));

  return (
    <Modal title="Manage Restores" onClose={onClose} maxWidth="max-w-2xl">
      <div className="p-4">
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
                <td className="px-3 py-2 font-medium text-[#28A745]">{r.status}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="flex justify-end mt-4">
          <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
        </div>
      </div>
    </Modal>
  );
}

/* Create Backup Modal */
function CreateBackupModal({ onClose, tenant }) {
  const [done, setDone] = useState(false);

  function handleCreate() {
    setDone(true);
    setTimeout(onClose, 2000);
  }

  return (
    <Modal title="Create Backup" onClose={onClose} maxWidth="max-w-md">
      <div className="px-6 py-5">
        {done ? (
          <div className="flex items-center gap-2 text-[#28A745] text-sm font-semibold">
            <CheckCircle2 size={18} />
            Backup started successfully!
          </div>
        ) : (
          <>
            <p className="text-sm text-[#333] mb-2">
              Do you want to back up administrative functions for this tenant?
            </p>
            <p className="text-sm font-semibold text-[#0078A8] mb-5">{tenant?.name ?? 'Unknown'}</p>
            <div className="flex justify-end gap-2">
              <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
              <button onClick={handleCreate} className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090]">
                Create
              </button>
            </div>
          </>
        )}
      </div>
    </Modal>
  );
}

/* Backup Unpacking Modal */
function BackupUnpackingModal({ onClose }) {
  const [opts, setOpts] = useState({ clear: true, diff: true, validate: false });
  const mockBackups = [
    { id: '2026-05-20T18:00:00Z', tenant: 'Contoso', label: 'Today at 6:00 PM' },
    { id: '2026-05-20T13:12:00Z', tenant: 'Contoso', label: 'Today at 1:12 PM' },
    { id: '2026-05-19T15:00:16Z', tenant: 'Contoso', label: 'Yesterday at 3:00 PM' },
  ];
  const [selected, setSelected] = useState(null);

  function toggle(k) { setOpts((o) => ({ ...o, [k]: !o[k] })); }

  return (
    <Modal title="Administrative Function Backup Unpacking" onClose={onClose} maxWidth="max-w-2xl">
      <div className="px-6 py-4 space-y-4">
        <div className="space-y-2">
          {[
            ['clear', 'Clear administrative function objects from previously unpacked backups'],
            ['diff', 'Perform administrative function differences operation during unpack'],
            ['validate', 'Validate backup integrity before unpack'],
          ].map(([k, label]) => (
            <label key={k} className="flex items-center gap-2 text-xs text-[#333] cursor-pointer">
              <input type="checkbox" checked={opts[k]} onChange={() => toggle(k)} />
              {label}
            </label>
          ))}
        </div>
        <div>
          <p className="text-xs font-semibold text-[#333] mb-2">Select a backup to unpack:</p>
          <table className="w-full text-xs border border-[#DEE2E6]">
            <thead>
              <tr className="bg-[#F2F2F2]">
                <th className="w-8 px-2 py-1.5 border-b border-[#DEE2E6]"></th>
                <th className="text-left px-3 py-1.5 border-b border-[#DEE2E6] font-semibold text-[#333]">Created</th>
                <th className="text-left px-3 py-1.5 border-b border-[#DEE2E6] font-semibold text-[#333]">Tenant</th>
              </tr>
            </thead>
            <tbody>
              {mockBackups.map((b) => (
                <tr
                  key={b.id}
                  className={`border-b border-[#DEE2E6] cursor-pointer ${selected === b.id ? 'bg-[#E6F4FA]' : 'hover:bg-[#F9FAFB]'}`}
                  onClick={() => setSelected(b.id)}
                >
                  <td className="px-2 py-1.5 text-center">
                    <input type="radio" checked={selected === b.id} onChange={() => setSelected(b.id)} />
                  </td>
                  <td className="px-3 py-1.5">{b.label}</td>
                  <td className="px-3 py-1.5">{b.tenant}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="flex justify-end gap-2 pt-2 border-t border-[#DEE2E6]">
          <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
          <button
            disabled={!selected}
            onClick={onClose}
            className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090] disabled:opacity-40"
          >
            Unpack
          </button>
        </div>
      </div>
    </Modal>
  );
}

/* ═══════════════════════ MAIN PAGE ═══════════════════════ */
export default function Dashboard() {
  const navigate = useNavigate();
  const { selectedTenant } = useApp();

  const [modal, setModal] = useState(null); // 'manageBackups' | 'manageRestores' | 'createBackup' | 'unpack'
  const tenants = selectedTenant ? [selectedTenant] : [];

  const tasks = [
    'Unpack backup 2026-05-19 15:00:16',
    'Backup tenant Contoso',
    'Diff restore objects',
    'Backup tenant Contoso',
    'Unpack backup 2026-04-28 06:00:12',
  ];

  const errors = [
    'Quest Recovery Function - Administrative Roles Backup failed to register',
    'Service principal consent validation failed for Contoso',
    'Role assignment scope could not be resolved',
    'Referenced principal not found in directory',
    'Backup integrity check failed: manifest hash mismatch',
  ];

  const diffs = [
    { label: 'Custom role deleted', count: 1 },
    { label: 'Role permission removed', count: 2 },
    { label: 'Role assignment added', count: 3 },
    { label: 'Referenced principal not found', count: 2 },
  ];

  return (
    <div className="flex flex-col">
      {/* ── Toolbar ── */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-2 flex-shrink-0">
        <ToolBtn icon={Settings} label="MANAGE BACKUPS" onClick={() => setModal('manageBackups')} />
        <ToolBtn icon={ArrowDownToLine} label="MANAGE RESTORES" onClick={() => setModal('manageRestores')} />
        <ToolBtn icon={RotateCcw} label="CREATE BACKUP" onClick={() => setModal('createBackup')} />
        <ToolBtn icon={Package} label="UNPACK BACKUP" onClick={() => setModal('unpack')} />
      </div>

      {/* ── Cards grid ── */}
      <div className="p-4 grid grid-cols-3 gap-3">

        {/* Card 1: Tenant protected */}
        <Card
          title={selectedTenant ? 'Tenant is protected' : 'No tenant selected'}
          icon={<CheckCircle2 size={16} className="text-[#28A745]" />}
        >
          <div className="space-y-0.5">
            <InfoRow label="Name" value={
              <button className="text-[#0078A8] hover:underline" onClick={() => navigate('/tenants')}>
                {selectedTenant?.name ?? '—'}
              </button>
            } />
            <InfoRow label="Backup schedule" value="Enabled" />
            <InfoRow label="Retention" value="1825" />
            <InfoRow label="Back up custom role definitions" value="Enabled" />
            <InfoRow label="Back up role permissions" value="Enabled" />
            <InfoRow label="Back up role assignments" value="Enabled" />
            <InfoRow label="Back up assignment scopes" value="Enabled" />
            <InfoRow label="Basic consent" value={selectedTenant?.consents?.basic?.granted ? 'Granted' : 'Not Granted'}
              valueColor={selectedTenant?.consents?.basic?.granted ? '#28A745' : '#DC3545'} />
            <InfoRow label="Restore consent" value={selectedTenant?.consents?.restore?.granted ? 'Granted' : 'Not Granted'}
              valueColor={selectedTenant?.consents?.restore?.granted ? '#28A745' : '#DC3545'} />
          </div>
        </Card>

        {/* Card 2: 149 backups */}
        <Card title="149 backups" showAll onShowAll={() => navigate('/backups')}>
          <BarChart />
          <div className="flex justify-between text-[10px] text-[#999] mt-1 px-1">
            <span>Apr 16</span>
            <span>May 20</span>
          </div>
        </Card>

        {/* Card 3: 55 tasks */}
        <Card title="55 tasks" showAll onShowAll={() => navigate('/tasks')}>
          <div className="space-y-1.5 mt-1">
            {tasks.map((t, i) => (
              <div key={i} className="flex items-start gap-1.5 text-xs text-[#555]">
                <Check size={12} className="text-[#28A745] mt-0.5 flex-shrink-0" />
                <span className="truncate">{t}</span>
              </div>
            ))}
          </div>
        </Card>

        {/* Card 4: 337 unpacked objects */}
        <Card title="337 unpacked objects" showAll onShowAll={() => navigate('/unpacked')}>
          <DonutChart />
        </Card>

        {/* Card 5: 8 differences */}
        <Card title="8 differences" showAll onShowAll={() => navigate('/differences')}>
          <div className="space-y-2 mt-1">
            {diffs.map((d, i) => (
              <div key={i} className="flex items-center justify-between text-xs">
                <span className="text-[#555]">{d.label}</span>
                <span className="font-semibold text-[#333]">{d.count}</span>
              </div>
            ))}
          </div>
        </Card>

        {/* Card 6: 94 errors */}
        <Card title="94 errors" showAll onShowAll={() => navigate('/events')}>
          <div className="space-y-1.5 mt-1">
            {errors.map((e, i) => (
              <div key={i} className="flex items-start gap-1.5 text-xs text-[#555]">
                <X size={12} className="text-[#DC3545] mt-0.5 flex-shrink-0" />
                <span className="line-clamp-2">{e}</span>
              </div>
            ))}
          </div>
        </Card>

        {/* Card 7: Consent Status */}
        <Card title="Consent Status" icon={<AlertCircle size={16} className="text-[#0096D6]" />}>
          <div className="space-y-3 mt-2">
            <div className="flex items-center gap-2 text-xs">
              <CheckCircle2 size={14} className="text-[#28A745]" />
              <span className="text-[#333]">Basic consent:</span>
              <span className="text-[#28A745] font-semibold">
                {selectedTenant?.consents?.basic?.granted ? 'Granted' : 'Not Granted'}
              </span>
            </div>
            <div className="flex items-center gap-2 text-xs">
              <CheckCircle2 size={14} className="text-[#28A745]" />
              <span className="text-[#333]">Restore consent:</span>
              <span className="text-[#28A745] font-semibold">
                {selectedTenant?.consents?.restore?.granted ? 'Granted' : 'Not Granted'}
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
        <ManageBackupsModal onClose={() => setModal(null)} tenants={tenants} />
      )}
      {modal === 'manageRestores' && (
        <ManageRestoresModal onClose={() => setModal(null)} tenants={tenants} />
      )}
      {modal === 'createBackup' && (
        <CreateBackupModal onClose={() => setModal(null)} tenant={selectedTenant} />
      )}
      {modal === 'unpack' && (
        <BackupUnpackingModal onClose={() => setModal(null)} />
      )}
    </div>
  );
}
