import { useEffect, useState } from 'react';
import { Loader2, RefreshCw } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';
import StatusBadge from '../components/common/StatusBadge.jsx';
import { api, pollJob } from '../api/client.js';

const columns = [
  { key: 'RoleName', header: 'Name', width: '25%' },
  { key: 'RoleType', header: 'Type', width: '10%' },
  { key: 'Status', header: 'Status', width: '12%', render: (r) => <StatusBadge value={r.Status} /> },
  { key: 'DefinitionChanged', header: 'Definition', width: '12%' },
  { key: 'AssignmentsChanged', header: 'Assignments', width: '12%' },
  { key: 'DifferenceCount', header: 'Diffs', width: '8%' },
  { key: 'Summary', header: 'Summary', width: '21%' }
];

export default function Differences() {
  const [rows, setRows] = useState([]);
  const [generatedAt, setGeneratedAt] = useState(null);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState(null);

  const loadCached = () => {
    api.compareResults()
      .then((r) => {
        const items = Array.isArray(r?.data?.data) ? r.data.data : [];
        setRows(items);
        setGeneratedAt(r?.data?.generatedAt);
      })
      .catch((e) => setError(e.message));
  };

  useEffect(loadCached, []);

  const handleRefresh = async () => {
    setRunning(true);
    setError(null);
    try {
      const resp = await api.runCompare();
      const job = resp?.data;
      const finished = await pollJob(job.id);
      if (finished.status === 'Failed') {
        setError(finished.error || 'Compare falhou');
      } else {
        loadCached();
      }
    } catch (e) {
      setError(e.message);
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className="p-4 space-y-3">
      <div className="bg-white border border-border-light px-4 py-3 flex items-center justify-between">
        <div className="text-xs text-text-secondary">
          {generatedAt ? `Gerado em ${new Date(generatedAt).toLocaleString()}` : 'Sem comparação ainda'}
        </div>
        <button
          onClick={handleRefresh}
          disabled={running}
          className="flex items-center gap-2 bg-accent hover:bg-accent-hover disabled:opacity-50 text-white text-xs px-3 py-1.5"
        >
          {running ? <Loader2 size={14} className="animate-spin" /> : <RefreshCw size={14} />}
          Refresh report
        </button>
      </div>
      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-2 text-xs">{error}</div>
      )}
      <DataTable columns={columns} rows={rows} emptyMessage="No items to display" />
    </div>
  );
}
