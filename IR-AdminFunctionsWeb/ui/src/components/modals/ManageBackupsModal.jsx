import { useEffect, useState } from 'react';
import { Edit2, Play, RefreshCw, CheckCircle2, AlertCircle, Loader2, Save, X } from 'lucide-react';
import { api } from '../../api/client.js';

/* ─── Configure Backup (secondary modal) ─────────────────────────────── */
function ConfigureBackupModal({ tenant, onClose, onSaved }) {
  const init = tenant.backupConfig || {};
  const [cfg, setCfg] = useState({
    scheduleEnabled:      init.scheduleEnabled ?? false,
    retentionDays:        init.retentionDays ?? 90,
    backupRoleDefinitions:  init.backupRoleDefinitions !== false,
    backupRolePermissions:  init.backupRolePermissions !== false,
    backupRoleAssignments:  init.backupRoleAssignments !== false,
    backupAssignmentScopes: init.backupAssignmentScopes !== false,
    includeBuiltInRoles:    init.includeBuiltInRoles ?? false,
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);

  const save = async () => {
    setSaving(true);
    setError(null);
    try {
      await api.updateTenantBackupConfig(tenant.tenantId, cfg);
      onSaved();
    } catch (e) {
      setError(e.message);
    } finally {
      setSaving(false);
    }
  };

  const Opt = ({ k, label, sub }) => (
    <label className="flex items-start gap-2 text-xs cursor-pointer select-none">
      <input type="checkbox" checked={cfg[k]}
        onChange={(e) => setCfg({ ...cfg, [k]: e.target.checked })} className="mt-0.5 flex-shrink-0" />
      <div>
        <div className="text-[#333] font-medium">{label}</div>
        {sub && <div className="text-text-secondary mt-0.5">{sub}</div>}
      </div>
    </label>
  );

  return (
    /* z-[60] so it appears above the Manage Backups modal (z-50) */
    <div className="fixed inset-0 bg-black/30 z-[60] flex items-center justify-center">
      <div className="bg-white shadow-2xl w-full max-w-[520px] mx-4 max-h-[90vh] flex flex-col rounded-sm">
        <div className="flex items-center justify-between px-5 py-3.5 border-b border-border-light">
          <div>
            <h3 className="text-sm font-semibold text-[#222]">Configure backup</h3>
            <p className="text-[11px] text-text-secondary mt-0.5">{tenant.name || tenant.tenantId}</p>
          </div>
          <button onClick={onClose} className="text-[#0078A8] hover:text-blue-800 font-bold text-lg leading-none">×</button>
        </div>

        <div className="px-5 py-4 overflow-y-auto flex-1 space-y-4 text-sm">
          {/* Schedule */}
          <div>
            <label className="block text-xs font-medium text-[#333] mb-1">Schedule</label>
            <select value={cfg.scheduleEnabled ? 'enabled' : 'disabled'}
              onChange={(e) => setCfg({ ...cfg, scheduleEnabled: e.target.value === 'enabled' })}
              className="border border-border-light rounded-sm px-2 py-1.5 text-xs w-48">
              <option value="enabled">Enabled</option>
              <option value="disabled">Disabled</option>
            </select>
          </div>

          {/* Retention */}
          <div>
            <p className="text-xs text-text-secondary mb-2">
              The retention period specifies how long backups are retained. Changes take effect on the next backup.
            </p>
            <label className="block text-xs font-medium text-[#333] mb-1">Retention policy (days)</label>
            <input type="number" min={1} max={1825} value={cfg.retentionDays}
              onChange={(e) => setCfg({ ...cfg, retentionDays: parseInt(e.target.value, 10) || 90 })}
              className="border border-border-light rounded-sm px-2 py-1.5 text-xs w-32" />
          </div>

          {/* Backup options */}
          <div>
            <h4 className="text-xs font-semibold text-[#333] mb-1">Backup options</h4>
            <p className="text-xs text-text-secondary mb-3">
              Some administrative role objects and relationships require specific Microsoft Graph permissions.
              Select the advanced administrative function data to include in the backup.
            </p>
            <div className="space-y-2.5">
              <Opt k="backupRoleDefinitions"  label="Back up custom administrative role definitions"
                sub="Saves name, description, state and allowedResourceActions of each custom role." />
              <Opt k="backupRolePermissions"  label="Back up administrative role permissions"
                sub="Saves the full permission set (allowed/excluded) within each role." />
              <Opt k="backupRoleAssignments"  label="Back up administrative role assignments"
                sub="Saves who is assigned to each role (users, groups, service principals)." />
              <Opt k="backupAssignmentScopes" label="Back up administrative assignment scopes"
                sub="Saves directory scope, administrative unit scope and app scope per assignment." />
              <Opt k="includeBuiltInRoles"    label="Back up referenced principals metadata"
                sub="Saves minimal metadata of referenced principals to support restore resolution." />
            </div>
          </div>

          {error && <div className="bg-red-50 border border-red-200 text-red-800 text-xs px-3 py-2 rounded">{error}</div>}
        </div>

        <div className="flex items-center justify-end gap-2 px-5 py-3 border-t border-border-light bg-gray-50">
          <button onClick={onClose} disabled={saving}
            className="px-4 py-1.5 text-sm border border-[#0078A8] text-[#0078A8] rounded-sm hover:bg-blue-50 disabled:opacity-50">
            Cancel
          </button>
          <button onClick={save} disabled={saving}
            className="px-4 py-1.5 text-sm bg-[#0078A8] text-white rounded-sm hover:bg-[#006090] disabled:opacity-50 flex items-center gap-1.5">
            {saving && <Loader2 size={13} className="animate-spin" />} Save
          </button>
        </div>
      </div>
    </div>
  );
}

/* ─── Manage Backups modal ────────────────────────────────────────────── */
export default function ManageBackupsModal({ onClose }) {
  const [tenants, setTenants] = useState([]);
  const [dashboards, setDashboards] = useState({});
  const [selectedId, setSelectedId] = useState(null);
  const [editing, setEditing] = useState(null);
  const [running, setRunning] = useState(null);
  const [loading, setLoading] = useState(true);

  const load = async () => {
    setLoading(true);
    try {
      const r = await api.tenants();
      const list = r?.data ?? [];
      setTenants(list);
      if (list.length > 0 && !selectedId) setSelectedId(list[0].tenantId);
      const stats = {};
      await Promise.all(list.map(async (t) => {
        try { stats[t.tenantId] = (await api.dashboard(t.tenantId))?.data || null; }
        catch { stats[t.tenantId] = null; }
      }));
      setDashboards(stats);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []); /* eslint-disable-line */

  const runNow = async (tenantId) => {
    setRunning(tenantId);
    try { await api.runBackupFor(tenantId); }
    finally { setRunning(null); }
  };

  const selectedTenant = tenants.find((t) => t.tenantId === selectedId);

  const ConsentDot = ({ ok }) => ok
    ? <CheckCircle2 size={14} className="text-green-600" />
    : <AlertCircle size={14} className="text-yellow-500" />;

  return (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center">
      <div className="bg-white shadow-2xl w-full max-w-[700px] mx-4 max-h-[80vh] flex flex-col rounded-sm">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-3.5 border-b border-border-light">
          <h2 className="text-sm font-semibold text-[#222]">Manage backups</h2>
          <button onClick={onClose} className="text-[#0078A8] hover:text-blue-800 font-bold text-lg leading-none">×</button>
        </div>

        <div className="flex flex-1 overflow-hidden">
          {/* Left menu */}
          <div className="w-[120px] border-r border-border-light flex-shrink-0 pt-3">
            <button className="relative w-full text-left px-4 py-2 text-xs font-medium text-[#333]">
              <span className="absolute left-0 top-0 bottom-0 w-[3px] bg-orange-500" />
              Tenants
            </button>
          </div>

          {/* Right area */}
          <div className="flex-1 flex flex-col min-w-0">
            {/* Toolbar */}
            <div className="flex items-center gap-2 px-4 py-2 border-b border-border-light">
              <button
                disabled={!selectedId}
                onClick={() => { const t = tenants.find(x => x.tenantId === selectedId); if (t) setEditing(t); }}
                className="flex items-center gap-1 text-xs text-[#0078A8] hover:underline disabled:opacity-40 disabled:no-underline"
              >
                <Edit2 size={12} /> EDIT
              </button>
              {selectedId && (
                <button
                  onClick={() => runNow(selectedId)}
                  disabled={running === selectedId}
                  className="flex items-center gap-1 text-xs text-[#0078A8] hover:underline disabled:opacity-40 ml-3"
                >
                  {running === selectedId ? <Loader2 size={12} className="animate-spin" /> : <Play size={12} />} RUN NOW
                </button>
              )}
              <span className="ml-auto text-xs text-text-secondary">{tenants.length} result{tenants.length !== 1 ? 's' : ''}</span>
            </div>

            {/* Grid */}
            <div className="flex-1 overflow-auto">
              {loading ? (
                <div className="text-center py-8 text-text-secondary text-xs flex items-center justify-center gap-2">
                  <Loader2 size={13} className="animate-spin" /> Loading…
                </div>
              ) : (
                <table className="w-full text-xs">
                  <thead className="bg-[#F0F0F0] sticky top-0">
                    <tr>
                      <th className="text-left px-4 py-2 font-semibold text-[#444]">Tenant</th>
                      <th className="text-left px-4 py-2 font-semibold text-[#444]">Schedule</th>
                      <th className="text-left px-4 py-2 font-semibold text-[#444]">Retention</th>
                      <th className="text-left px-4 py-2 font-semibold text-[#444]">Last backup</th>
                      <th className="text-left px-4 py-2 font-semibold text-[#444]">Consent</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tenants.map((t) => {
                      const sel = t.tenantId === selectedId;
                      const stats = dashboards[t.tenantId];
                      const last = stats?.lastBackupAt ? new Date(stats.lastBackupAt).toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'short' }) : '—';
                      return (
                        <tr key={t.tenantId}
                          onClick={() => setSelectedId(t.tenantId)}
                          className={`border-b border-[#E8E8E8] cursor-pointer ${sel ? 'bg-[#EBF4FF]' : 'hover:bg-gray-50'}`}
                        >
                          <td className="px-4 py-2.5">
                            <div className="font-medium text-[#222]">{t.name || t.tenantId}</div>
                            {t.domain && <div className="text-[10px] text-text-secondary">{t.domain}</div>}
                          </td>
                          <td className="px-4 py-2.5">{t.backupConfig?.scheduleEnabled ? 'Enabled' : 'Disabled'}</td>
                          <td className="px-4 py-2.5">{t.backupConfig?.retentionDays ?? 90}</td>
                          <td className="px-4 py-2.5">{last}</td>
                          <td className="px-4 py-2.5">
                            <ConsentDot ok={t.consents?.basic?.granted} />
                          </td>
                        </tr>
                      );
                    })}
                    {tenants.length === 0 && (
                      <tr><td colSpan={5} className="px-4 py-6 text-center text-text-secondary">No tenants configured.</td></tr>
                    )}
                  </tbody>
                </table>
              )}
            </div>

            {/* Footer */}
            <div className="flex items-center justify-between px-4 py-2.5 border-t border-border-light bg-gray-50">
              <span className="text-[11px] text-text-secondary">1–{tenants.length} of {tenants.length} Results</span>
              <button onClick={onClose}
                className="px-5 py-1.5 text-sm bg-[#0078A8] text-white rounded-sm hover:bg-[#006090]">
                Finish
              </button>
            </div>
          </div>
        </div>
      </div>

      {editing && (
        <ConfigureBackupModal
          tenant={editing}
          onClose={() => setEditing(null)}
          onSaved={() => { setEditing(null); load(); }}
        />
      )}
    </div>
  );
}
