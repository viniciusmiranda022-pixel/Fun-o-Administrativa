import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { RefreshCw, Play, Archive, CheckCircle2, AlertCircle, Clock, Loader2 } from 'lucide-react';
import { api } from '../api/client.js';
import { useTenant } from '../context/TenantContext.jsx';

function StatusPill({ status, error }) {
  if (error) return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-red-700 bg-red-50">
      <AlertCircle size={11} /> Falhou
    </span>
  );
  if (!status || status === 'Completed') return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-green-700 bg-green-50">
      <CheckCircle2 size={11} /> Concluído
    </span>
  );
  if (status === 'Running') return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-blue-700 bg-blue-50">
      <Loader2 size={11} className="animate-spin" /> Em execução
    </span>
  );
  return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-gray-600 bg-gray-100">
      <Clock size={11} /> {status}
    </span>
  );
}

function UnpackedBadge({ unpacked }) {
  if (unpacked) return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold text-purple-700 bg-purple-50">
      <Archive size={11} /> Descompactado
    </span>
  );
  return <span className="text-xs text-text-secondary">—</span>;
}

function fmt(v) { return v != null ? v.toLocaleString('pt-BR') : '—'; }
function fmtBytes(bytes) {
  if (!bytes) return '—';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}
function fmtDate(d) { return d ? new Date(d).toLocaleString('pt-BR') : '—'; }
function fmtDuration(startedAt, finishedAt) {
  if (!startedAt || !finishedAt) return '—';
  const s = Math.round((new Date(finishedAt) - new Date(startedAt)) / 1000);
  if (s < 60) return `${s}s`;
  return `${Math.floor(s / 60)}m ${s % 60}s`;
}

export default function Backups() {
  const navigate = useNavigate();
  const { selectedTenantId } = useTenant();
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [runningId, setRunningId] = useState(null);

  const reload = async () => {
    setLoading(true);
    setError(null);
    try {
      const resp = await api.backups();
      setRows(resp?.data ?? []);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { reload(); }, []);

  const handleRunBackup = async () => {
    setError(null);
    try {
      const resp = await api.runBackupFor(selectedTenantId);
      const job = resp?.data;
      setRunningId(job?.id ?? 'running');
      navigate('/tasks');
    } catch (e) {
      setError(e.message);
    }
  };

  const handleUnpack = (backup) => {
    navigate('/unpack', { state: { backupId: backup.id } });
  };

  const displayed = selectedTenantId
    ? rows.filter((r) => !r.tenantId || r.tenantId === selectedTenantId)
    : rows;

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-4">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Backups</h2>
          <p className="text-xs text-text-secondary mt-0.5">
            Snapshots coletados das funções administrativas.
            {selectedTenantId && ` Filtrando por tenant: ${selectedTenantId}`}
          </p>
        </div>
        <div className="flex gap-2">
          <button onClick={reload} className="text-xs flex items-center gap-1 px-3 py-1.5 border border-border-light rounded bg-white hover:bg-gray-50">
            <RefreshCw size={12} /> Atualizar
          </button>
          <button onClick={handleRunBackup} disabled={!!runningId}
            className="text-xs flex items-center gap-1 px-3 py-1.5 bg-[#0078A8] text-white rounded hover:bg-[#006090] disabled:opacity-50">
            {runningId ? <Loader2 size={12} className="animate-spin" /> : <Play size={12} />}
            Executar backup
          </button>
        </div>
      </div>

      {error && <div className="bg-red-50 border border-red-200 text-red-800 text-xs px-3 py-2 rounded mb-3">{error}</div>}

      <div className="bg-white border border-border-light rounded-lg overflow-hidden">
        {loading ? (
          <div className="text-center py-12 text-text-secondary text-sm">Carregando backups...</div>
        ) : displayed.length === 0 ? (
          <div className="text-center py-12 text-text-secondary text-sm">Nenhum backup encontrado.</div>
        ) : (
          <table className="w-full text-sm">
            <thead className="bg-gray-50 text-xs text-text-secondary">
              <tr>
                <th className="text-left px-4 py-2 font-semibold">Tenant</th>
                <th className="text-left px-4 py-2 font-semibold">Data / Hora</th>
                <th className="text-left px-4 py-2 font-semibold">Duração</th>
                <th className="text-left px-4 py-2 font-semibold">Status</th>
                <th className="text-right px-4 py-2 font-semibold">Funções custom</th>
                <th className="text-right px-4 py-2 font-semibold">Permissões</th>
                <th className="text-right px-4 py-2 font-semibold">Atribuições</th>
                <th className="text-right px-4 py-2 font-semibold">Principals</th>
                <th className="text-right px-4 py-2 font-semibold">Tamanho</th>
                <th className="text-left px-4 py-2 font-semibold">Descompactado</th>
                <th className="px-4 py-2"></th>
              </tr>
            </thead>
            <tbody>
              {displayed.map((b, idx) => (
                <tr key={b.id ?? idx} className="border-t border-border-light hover:bg-gray-50">
                  <td className="px-4 py-2 text-xs font-mono text-text-secondary">
                    {b.tenantId ? b.tenantId.substring(0, 8) + '…' : '—'}
                  </td>
                  <td className="px-4 py-2 text-xs font-mono">
                    {fmtDate(b.collectedAt || b.startedAt || b.modifiedAt)}
                  </td>
                  <td className="px-4 py-2 text-xs">
                    {fmtDuration(b.startedAt, b.finishedAt)}
                  </td>
                  <td className="px-4 py-2">
                    <StatusPill status={b.status} error={b.manifestError} />
                  </td>
                  <td className="px-4 py-2 text-xs text-right font-mono">
                    {fmt(b.roleDefinitionsCount ?? b.customRoleDefinitionsCount)}
                  </td>
                  <td className="px-4 py-2 text-xs text-right font-mono">
                    {fmt(b.rolePermissionsCount ?? b.permissionsCount)}
                  </td>
                  <td className="px-4 py-2 text-xs text-right font-mono">
                    {fmt(b.roleAssignmentsCount ?? b.assignmentsCount)}
                  </td>
                  <td className="px-4 py-2 text-xs text-right font-mono">
                    {fmt(b.principalsCount ?? b.principalsReferenced)}
                  </td>
                  <td className="px-4 py-2 text-xs text-right font-mono">
                    {fmtBytes(b.sizeBytes ?? b.size)}
                    {b.sizeDisplay && !b.sizeBytes ? <span>{b.sizeDisplay}</span> : null}
                  </td>
                  <td className="px-4 py-2">
                    <UnpackedBadge unpacked={b.isUnpacked ?? b.unpacked} />
                  </td>
                  <td className="px-4 py-2 text-right">
                    <button onClick={() => handleUnpack(b)}
                      className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 justify-end whitespace-nowrap">
                      <Archive size={11} /> Descompactar
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}
