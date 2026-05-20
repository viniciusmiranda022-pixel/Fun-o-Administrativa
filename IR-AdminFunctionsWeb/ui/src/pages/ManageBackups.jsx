import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import {
  ArrowLeft, Edit2, Play, RefreshCw, CheckCircle2, AlertCircle, Clock,
  Database, Loader2, X, Save
} from 'lucide-react';
import { api } from '../api/client.js';

function ConfigureBackupModal({ tenant, onClose, onSaved }) {
  const initial = tenant.backupConfig || {};
  const [cfg, setCfg] = useState({
    scheduleEnabled: initial.scheduleEnabled ?? false,
    retentionDays: initial.retentionDays ?? 90,
    frequency: initial.frequency ?? 'Daily',
    timeOfDay: initial.timeOfDay ?? '02:00',
    backupRoleDefinitions: initial.backupRoleDefinitions ?? true,
    backupRolePermissions: initial.backupRolePermissions ?? true,
    backupRoleAssignments: initial.backupRoleAssignments ?? true,
    backupAssignmentScopes: initial.backupAssignmentScopes ?? true,
    includeBuiltInRoles: initial.includeBuiltInRoles ?? false
  });
  const [runNow, setRunNow] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);

  const save = async () => {
    setSaving(true);
    setError(null);
    try {
      await api.updateTenantBackupConfig(tenant.tenantId, cfg);
      if (runNow) {
        await api.runBackupFor(tenant.tenantId);
      }
      onSaved(runNow);
    } catch (e) {
      setError(e.message);
    } finally {
      setSaving(false);
    }
  };

  const Check = ({ k, label, sub }) => (
    <label className="flex items-start gap-2 text-sm cursor-pointer">
      <input type="checkbox" checked={cfg[k]} onChange={(e) => setCfg({ ...cfg, [k]: e.target.checked })}
        className="mt-0.5" />
      <div>
        <div className="text-[#333]">{label}</div>
        {sub && <div className="text-xs text-text-secondary">{sub}</div>}
      </div>
    </label>
  );

  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-2xl mx-4 max-h-[90vh] flex flex-col">
        <div className="flex items-center justify-between px-6 py-4 border-b border-border-light">
          <div>
            <h2 className="text-base font-semibold text-[#222]">Configure Administrative Function Backup</h2>
            <p className="text-xs text-text-secondary mt-0.5">{tenant.name || tenant.tenantId}</p>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700"><X size={20} /></button>
        </div>

        <div className="px-6 py-4 overflow-y-auto space-y-5">
          <section>
            <h3 className="text-sm font-semibold text-[#222] mb-2">Schedule</h3>
            <label className="flex items-center gap-2 text-sm mb-3 cursor-pointer">
              <input type="checkbox" checked={cfg.scheduleEnabled}
                onChange={(e) => setCfg({ ...cfg, scheduleEnabled: e.target.checked })} />
              <span>Habilitar backup agendado</span>
            </label>
            <div className="grid grid-cols-2 gap-3">
              <label className="block text-xs">
                <span className="text-text-secondary">Frequência</span>
                <select disabled={!cfg.scheduleEnabled} value={cfg.frequency}
                  onChange={(e) => setCfg({ ...cfg, frequency: e.target.value })}
                  className="mt-1 w-full border border-border-light rounded px-2 py-1.5 text-sm disabled:opacity-50">
                  <option value="Daily">Diário</option>
                  <option value="Weekly">Semanal</option>
                  <option value="Manual">Manual</option>
                </select>
              </label>
              <label className="block text-xs">
                <span className="text-text-secondary">Horário (HH:mm)</span>
                <input type="time" disabled={!cfg.scheduleEnabled} value={cfg.timeOfDay}
                  onChange={(e) => setCfg({ ...cfg, timeOfDay: e.target.value })}
                  className="mt-1 w-full border border-border-light rounded px-2 py-1.5 text-sm disabled:opacity-50" />
              </label>
            </div>
          </section>

          <section>
            <h3 className="text-sm font-semibold text-[#222] mb-2">Retenção</h3>
            <label className="block text-xs">
              <span className="text-text-secondary">Manter backups por</span>
              <select value={cfg.retentionDays}
                onChange={(e) => setCfg({ ...cfg, retentionDays: parseInt(e.target.value, 10) })}
                className="mt-1 w-full border border-border-light rounded px-2 py-1.5 text-sm">
                <option value={30}>30 dias</option>
                <option value={90}>90 dias</option>
                <option value={180}>180 dias</option>
                <option value={365}>1 ano (365 dias)</option>
                <option value={1825}>5 anos (1825 dias)</option>
              </select>
            </label>
          </section>

          <section>
            <h3 className="text-sm font-semibold text-[#222] mb-2">Escopo do backup</h3>
            <div className="space-y-2">
              <Check k="backupRoleDefinitions" label="Back up custom role definitions"
                sub="Salva nome, descrição, estado e permissões das funções administrativas customizadas." />
              <Check k="backupRolePermissions" label="Back up role permissions"
                sub="Salva as permissões incluídas dentro de cada função." />
              <Check k="backupRoleAssignments" label="Back up role assignments"
                sub="Salva quem está atribuído a cada função (usuários, grupos, service principals)." />
              <Check k="backupAssignmentScopes" label="Back up assignment scope"
                sub="Salva o escopo da atribuição para permitir restauração precisa." />
              <Check k="includeBuiltInRoles" label="Incluir built-in roles"
                sub="Por padrão somente custom roles. Habilite se precisar histórico de atribuições em roles built-in." />
            </div>
          </section>

          <section>
            <label className="flex items-center gap-2 text-sm cursor-pointer">
              <input type="checkbox" checked={runNow} onChange={(e) => setRunNow(e.target.checked)} />
              <span>Executar backup imediatamente após salvar</span>
            </label>
          </section>

          {error && <div className="bg-red-50 border border-red-200 text-red-800 text-xs px-3 py-2 rounded">{error}</div>}
        </div>

        <div className="flex items-center justify-end gap-2 px-6 py-3 border-t border-border-light bg-gray-50">
          <button onClick={onClose} disabled={saving}
            className="px-4 py-2 text-sm border border-border-light rounded hover:bg-white disabled:opacity-60">Cancelar</button>
          <button onClick={save} disabled={saving}
            className="px-4 py-2 text-sm bg-[#0078A8] text-white rounded hover:bg-[#006090] disabled:opacity-60 flex items-center gap-2">
            {saving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
            Salvar
          </button>
        </div>
      </div>
    </div>
  );
}

function ScheduleBadge({ enabled }) {
  return enabled ? (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-green-700 bg-green-50">
      <CheckCircle2 size={11} /> Enabled
    </span>
  ) : (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-gray-600 bg-gray-100">
      <Clock size={11} /> Disabled
    </span>
  );
}

export default function ManageBackups() {
  const navigate = useNavigate();
  const [tenants, setTenants] = useState([]);
  const [dashboards, setDashboards] = useState({});
  const [editing, setEditing] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [notice, setNotice] = useState(null);

  const load = async () => {
    setLoading(true);
    setError(null);
    try {
      const r = await api.tenants();
      const list = r?.data ?? [];
      setTenants(list);
      const stats = {};
      await Promise.all(list.map(async (t) => {
        try {
          const d = await api.dashboard(t.tenantId);
          stats[t.tenantId] = d?.data || null;
        } catch { stats[t.tenantId] = null; }
      }));
      setDashboards(stats);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const onSaved = (didRunNow) => {
    setEditing(null);
    setNotice(didRunNow ? 'Configuração salva. Backup iniciado.' : 'Configuração salva.');
    load();
  };

  const runNow = async (tenantId) => {
    try {
      await api.runBackupFor(tenantId);
      setNotice('Backup task created. Acompanhe em Tasks.');
      load();
    } catch (e) { setError(e.message); }
  };

  return (
    <div className="p-6">
      <button onClick={() => navigate('/dashboard')} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 mb-3">
        <ArrowLeft size={12} /> Voltar ao Dashboard
      </button>

      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Manage Backups</h2>
          <p className="text-xs text-text-secondary mt-0.5">Configure backup de funções administrativas por tenant.</p>
        </div>
        <button onClick={load} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1">
          <RefreshCw size={12} /> Atualizar
        </button>
      </div>

      {error && <div className="bg-red-50 border border-red-200 text-red-800 text-sm px-3 py-2 rounded mb-3">{error}</div>}
      {notice && (
        <div className="bg-blue-50 border border-blue-200 text-blue-900 text-xs px-3 py-2 rounded mb-3 flex justify-between">
          <span>{notice}</span>
          <button onClick={() => setNotice(null)} className="text-gray-500">×</button>
        </div>
      )}

      {loading ? (
        <div className="text-text-secondary text-sm flex items-center gap-2"><Loader2 className="animate-spin" size={14} /> Carregando...</div>
      ) : tenants.length === 0 ? (
        <div className="text-center py-16 text-text-secondary">
          <Database size={32} className="mx-auto mb-3 opacity-30" />
          <p className="text-sm">Nenhum tenant configurado.</p>
          <button onClick={() => navigate('/tenants')} className="mt-2 text-xs text-[#0078A8] hover:underline">Adicionar tenant</button>
        </div>
      ) : (
        <div className="bg-white border border-border-light rounded-lg overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-xs text-text-secondary">
              <tr>
                <th className="text-left px-4 py-2 font-semibold">Tenant</th>
                <th className="text-left px-4 py-2 font-semibold">Schedule</th>
                <th className="text-left px-4 py-2 font-semibold">Retenção</th>
                <th className="text-left px-4 py-2 font-semibold">Último backup</th>
                <th className="text-left px-4 py-2 font-semibold">Consent Basic</th>
                <th className="text-right px-4 py-2 font-semibold">Ações</th>
              </tr>
            </thead>
            <tbody>
              {tenants.map((t) => {
                const stats = dashboards[t.tenantId];
                const last = stats?.lastBackupAt ? new Date(stats.lastBackupAt).toLocaleString('pt-BR') : '—';
                return (
                  <tr key={t.tenantId} className="border-t border-border-light hover:bg-gray-50">
                    <td className="px-4 py-3">
                      <div className="font-semibold text-[#222]">{t.name || t.tenantId}</div>
                      {t.domain && <div className="text-[11px] text-text-secondary">{t.domain}</div>}
                    </td>
                    <td className="px-4 py-3"><ScheduleBadge enabled={t.backupConfig?.scheduleEnabled} /></td>
                    <td className="px-4 py-3 text-xs">{t.backupConfig?.retentionDays ?? 90} dias</td>
                    <td className="px-4 py-3 text-xs">{last}</td>
                    <td className="px-4 py-3">
                      {t.consents?.basic?.granted ? (
                        <span className="inline-flex items-center gap-1 text-xs text-green-700">
                          <CheckCircle2 size={12} /> Granted
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 text-xs text-yellow-700">
                          <AlertCircle size={12} /> Not Granted
                        </span>
                      )}
                    </td>
                    <td className="px-4 py-3 text-right">
                      <div className="flex justify-end gap-2">
                        <button onClick={() => runNow(t.tenantId)}
                          disabled={!t.consents?.basic?.granted}
                          title={t.consents?.basic?.granted ? 'Executar backup agora' : 'Conceda consent Basic primeiro'}
                          className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 disabled:opacity-40 disabled:no-underline disabled:cursor-not-allowed">
                          <Play size={12} /> Run now
                        </button>
                        <button onClick={() => setEditing(t)}
                          className="text-xs text-[#0078A8] hover:underline flex items-center gap-1">
                          <Edit2 size={12} /> Edit
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {editing && (
        <ConfigureBackupModal tenant={editing} onClose={() => setEditing(null)} onSaved={onSaved} />
      )}
    </div>
  );
}
