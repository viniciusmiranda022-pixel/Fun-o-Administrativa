import { useState, useEffect } from 'react';
import { RotateCcw, Download, RefreshCw, Columns, Search, FileText } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';
import { api, pollJob } from '../api/client.js';

function DiffCell({ row }) {
  if (row.diffType === 'Removed') {
    return <span className="text-[#DC3545] line-through text-xs">{row.diffText}</span>;
  }
  if (row.diffType === 'New') {
    return <span className="text-[#17A2B8] text-xs">→ {row.diffText}</span>;
  }
  if (row.diffType === 'Changed') {
    return <span className="text-[#FFC107] text-xs">~ {row.diffText}</span>;
  }
  return <span className="text-xs text-[#28A745]">{row.diffText}</span>;
}

const columns = [
  {
    key: 'name', header: 'Role Name', sortable: true, width: '22%',
    render: (r) => (
      <div className="flex items-center gap-1.5">
        <FileText size={12} className="text-[#0096D6] flex-shrink-0" />
        <span className="text-[#0078A8] truncate text-xs">{r.name}</span>
      </div>
    )
  },
  { key: 'roleType', header: 'Role Type', width: '10%' },
  {
    key: 'status', header: 'Status', width: '11%',
    render: (r) => (
      <span className={`text-xs font-medium ${r.diffType === 'Removed' ? 'text-[#DC3545]' : r.diffType === 'New' ? 'text-[#17A2B8]' : r.diffType === 'Changed' ? 'text-[#FFC107]' : 'text-[#28A745]'}`}>
        {r.diffType}
      </span>
    )
  },
  { key: 'defChanged', header: 'Definition Changed', width: '13%' },
  { key: 'assignChanged', header: 'Assignments Changed', width: '14%' },
  { key: 'diffCount', header: '# Differences', width: '10%' },
  {
    key: 'summary', header: 'Summary', width: '20%',
    render: (r) => <DiffCell row={r} />
  },
];

function mapResult(item) {
  return {
    id: item.RoleName || item.roleName,
    name: item.RoleName || item.roleName || '—',
    roleType: item.RoleType || item.roleType || '—',
    diffType: item.Status || item.status || 'Equal',
    defChanged: item.DefinitionChanged || item.definitionChanged || 'No',
    assignChanged: item.AssignmentsChanged || item.assignmentsChanged || 'No',
    diffCount: item.DifferenceCount ?? item.differenceCount ?? 0,
    diffText: item.Summary || item.summary || '—',
  };
}

export default function Differences() {
  const [view, setView] = useState('list');
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [rows, setRows] = useState([]);
  const [generatedAt, setGeneratedAt] = useState(null);
  const [loading, setLoading] = useState(true);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState(null);

  function loadResults() {
    setLoading(true);
    api.compareResults()
      .then((r) => {
        if (!r?.success) { setRows([]); return; }
        const data = r?.data;
        if (!data) { setRows([]); return; }
        setGeneratedAt(data.generatedAt);
        const items = Array.isArray(data.data) ? data.data : [];
        setRows(items.map(mapResult));
      })
      .catch(() => setRows([]))
      .finally(() => setLoading(false));
  }

  useEffect(() => { loadResults(); }, []);

  async function handleRunCompare() {
    setRunning(true);
    setError(null);
    try {
      const resp = await api.runCompare();
      const job = resp?.data;
      if (!job?.id) throw new Error('No job id returned');
      await pollJob(job.id, {
        intervalMs: 3000,
        timeoutMs: 300000,
      });
      loadResults();
    } catch (err) {
      setError(err.message);
    } finally {
      setRunning(false);
    }
  }

  const nonEqual = rows.filter((r) => r.diffType !== 'Equal');
  const filtered = rows.filter((r) => {
    if (statusFilter && r.diffType !== statusFilter) return false;
    if (search && !r.name.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  const statusOptions = [...new Set(rows.map(r => r.diffType))];

  return (
    <div className="flex flex-col h-full">
      {/* Toggle + Filter bar */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-3 flex-shrink-0 flex-wrap">
        <div className="flex border border-[#DEE2E6] overflow-hidden text-xs">
          {['list', 'differences'].map((v) => (
            <button
              key={v}
              onClick={() => setView(v)}
              className={`px-3 py-1 capitalize ${view === v ? 'bg-[#0078A8] text-white' : 'bg-white text-[#555] hover:bg-[#F2F2F2]'}`}
            >
              {v === 'list' ? 'List View' : 'Differences'}
            </button>
          ))}
        </div>
        {generatedAt && (
          <span className="text-xs text-[#888]">
            Last compare: {new Date(generatedAt).toLocaleString()}
          </span>
        )}
        <div className="flex items-center gap-1 text-xs">
          <span className="text-[#555]">Status:</span>
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white"
          >
            <option value="">Any</option>
            {statusOptions.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
      </div>

      <div className="p-4 space-y-3 flex-1 overflow-auto">
        {/* Action toolbar */}
        <div className="bg-white border border-[#DEE2E6] px-3 py-2 flex items-center gap-2">
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] opacity-40 cursor-not-allowed">
            <RotateCcw size={13} />
            RESTORE
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Download size={13} />
            EXPORT
          </button>
          <button
            onClick={handleRunCompare}
            disabled={running}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2] disabled:opacity-40"
          >
            <RefreshCw size={13} className={running ? 'animate-spin' : ''} />
            {running ? 'RUNNING…' : 'RUN COMPARE'}
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Columns size={13} />
            EDIT COLUMNS
          </button>
          <div className="ml-auto flex items-center gap-2">
            <span className="text-xs text-[#666]">
              {loading ? 'Loading…' : `${filtered.length} role${filtered.length !== 1 ? 's' : ''} (${nonEqual.length} differ)`}
            </span>
            <div className="flex items-center border border-[#DEE2E6] overflow-hidden">
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search..."
                className="px-2 py-1 text-xs outline-none w-40"
              />
              <button className="px-2 py-1 bg-[#F2F2F2] border-l border-[#DEE2E6] text-[#0078A8]">
                <Search size={13} />
              </button>
            </div>
          </div>
        </div>

        {error && (
          <div className="bg-white border border-[#DC3545] p-3 text-xs text-[#DC3545]">
            Compare failed: {error}
          </div>
        )}

        {loading ? (
          <div className="text-xs text-[#888] p-4">Loading results…</div>
        ) : rows.length === 0 ? (
          <div className="bg-white border border-[#DEE2E6] p-8 text-center text-xs text-[#888]">
            No compare results yet. Click "RUN COMPARE" to compare the latest backup against the current tenant state.
          </div>
        ) : (
          <DataTable columns={columns} rows={filtered} totalCount={filtered.length} />
        )}
      </div>
    </div>
  );
}
