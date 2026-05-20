import { useEffect, useState } from 'react';
import { Loader2, Play } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';
import StatusBadge from '../components/common/StatusBadge.jsx';
import { api, pollJob } from '../api/client.js';

const columns = [
  { key: 'id', header: 'Backup timestamp', width: '20%' },
  { key: 'tenantId', header: 'Tenant', width: '20%', render: (r) => r.tenantId ?? '—' },
  {
    key: 'collectedAt',
    header: 'Collected at',
    width: '18%',
    render: (r) => r.collectedAt ?? r.modifiedAt
  },
  { key: 'roleDefinitionsCount', header: 'Roles', width: '8%', render: (r) => r.roleDefinitionsCount ?? '—' },
  { key: 'roleAssignmentsCount', header: 'Assignments', width: '10%', render: (r) => r.roleAssignmentsCount ?? '—' },
  { key: 'sizeDisplay', header: 'Size', width: '8%' },
  {
    key: 'status',
    header: 'Status',
    width: '10%',
    render: (r) => <StatusBadge value={r.manifestError ? 'Failed' : 'Completed'} />
  }
];

export default function Backups() {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [runState, setRunState] = useState(null);

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

  useEffect(() => {
    reload();
  }, []);

  const handleRun = async () => {
    setRunState({ status: 'starting' });
    try {
      const resp = await api.runBackup();
      const job = resp?.data;
      setRunState({ status: 'running', jobId: job.id });
      const finished = await pollJob(job.id, {
        onTick: (j) => setRunState({ status: j.status, jobId: job.id })
      });
      setRunState({ status: finished.status, jobId: job.id, result: finished });
      reload();
    } catch (e) {
      setRunState({ status: 'Failed', error: e.message });
    }
  };

  return (
    <div className="p-4 space-y-3">
      <div className="flex items-center justify-between bg-white border border-border-light px-4 py-3">
        <h2 className="text-sm font-semibold text-[#333]">Backups</h2>
        <button
          onClick={handleRun}
          disabled={runState?.status === 'Running' || runState?.status === 'starting'}
          className="flex items-center gap-2 bg-accent hover:bg-accent-hover disabled:opacity-50 text-white text-xs px-3 py-1.5"
        >
          {runState?.status === 'Running' || runState?.status === 'starting' ? (
            <Loader2 size={14} className="animate-spin" />
          ) : (
            <Play size={14} />
          )}
          Run backup
        </button>
      </div>

      {runState && (
        <div className="bg-white border border-border-light px-4 py-2 text-xs">
          Job <code>{runState.jobId ?? '—'}</code>: <StatusBadge value={runState.status} />
          {runState.error && <span className="text-status-error ml-2">{runState.error}</span>}
        </div>
      )}

      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-2 text-xs">{error}</div>
      )}

      {loading ? (
        <div className="bg-white border border-border-light px-4 py-8 text-center text-xs text-text-secondary">
          Carregando backups…
        </div>
      ) : (
        <DataTable columns={columns} rows={rows} emptyMessage="Nenhum backup encontrado." />
      )}
    </div>
  );
}
