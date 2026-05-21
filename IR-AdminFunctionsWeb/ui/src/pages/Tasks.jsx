import { useState, useEffect } from 'react';
import { CheckCircle2, Columns, Search, X, Clock, AlertTriangle, RefreshCw } from 'lucide-react';
import { api } from '../api/client.js';

function StatusIcon({ status }) {
  const s = (status || '').toLowerCase();
  if (s === 'completed') return <CheckCircle2 size={13} className="text-[#28A745]" />;
  if (s === 'failed') return <X size={13} className="text-[#DC3545]" />;
  if (s === 'warning') return <AlertTriangle size={13} className="text-[#FFC107]" />;
  if (s === 'running' || s === 'queued') return <Clock size={13} className="text-[#17A2B8]" />;
  return <Clock size={13} className="text-[#17A2B8]" />;
}

function TaskDetailPanel({ task, onClose }) {
  const statusColor = (s) => {
    const v = (s || '').toLowerCase();
    if (v === 'completed') return '#28A745';
    if (v === 'failed') return '#DC3545';
    if (v === 'warning') return '#FFC107';
    return '#17A2B8';
  };

  return (
    <div className="w-72 border-l border-[#DEE2E6] bg-white flex flex-col flex-shrink-0 text-xs overflow-auto">
      <div className="flex items-center justify-between px-4 py-3 border-b border-[#DEE2E6] bg-[#F8F8F8]">
        <span className="font-semibold text-[#333] text-sm">Task Details</span>
        <button onClick={onClose} className="text-[#999] hover:text-[#333] text-lg font-bold leading-none">×</button>
      </div>
      <div className="p-4 space-y-3">
        <div>
          <div className="text-[#888] mb-0.5">Name</div>
          <div className="font-medium text-[#333]">{task.name}</div>
        </div>
        <div>
          <div className="text-[#888] mb-0.5">Type</div>
          <div className="text-[#333]">{task.type}</div>
        </div>
        <div>
          <div className="text-[#888] mb-0.5">Status</div>
          <div className="flex items-center gap-1.5">
            <StatusIcon status={task.status} />
            <span className="font-semibold" style={{ color: statusColor(task.status) }}>{task.status}</span>
          </div>
        </div>
        <div>
          <div className="text-[#888] mb-0.5">Operation</div>
          <div className="text-[#333] break-words">{task.operation}</div>
        </div>
        <div>
          <div className="text-[#888] mb-0.5">Time</div>
          <div className="text-[#333]">{task.modified}</div>
        </div>
      </div>
    </div>
  );
}

function mapTask(t) {
  const d = new Date(t.time || t.Time || '');
  return {
    id: `${t.name}-${t.time}`,
    name: t.name || t.Name || t.type || '—',
    status: t.status || t.Status || '—',
    type: t.type || t.Type || '—',
    modified: isNaN(d) ? '—' : d.toLocaleString(),
    operation: t.operation || t.Operation || '—',
  };
}

const COLUMNS = [
  {
    key: 'status', header: '', width: '3%',
    render: (r) => <StatusIcon status={r.status} />
  },
  {
    key: 'name', header: 'Name', sortable: true, width: '30%',
    render: (r) => <span className="text-[#0078A8] hover:underline cursor-pointer">{r.name}</span>
  },
  {
    key: 'status2', header: 'Status', width: '12%',
    render: (r) => {
      const color = (s) => {
        const v = (s || '').toLowerCase();
        if (v === 'completed') return '#28A745';
        if (v === 'failed') return '#DC3545';
        if (v === 'warning') return '#FFC107';
        return '#17A2B8';
      };
      return <span className="text-xs font-semibold" style={{ color: color(r.status) }}>{r.status}</span>;
    }
  },
  { key: 'type', header: 'Type', sortable: true, width: '12%' },
  { key: 'modified', header: 'Time', sortable: true, width: '18%' },
  { key: 'operation', header: 'Operation', width: '25%' },
];

function sinceFromFilter(filter) {
  const now = new Date();
  if (filter === 'today') { const d = new Date(now); d.setHours(0, 0, 0, 0); return d.toISOString(); }
  if (filter === '7days') return new Date(now - 7 * 86400000).toISOString();
  if (filter === '30days') return new Date(now - 30 * 86400000).toISOString();
  return undefined;
}

export default function Tasks() {
  const [search, setSearch] = useState('');
  const [dateFilter, setDateFilter] = useState('30days');
  const [selectedTask, setSelectedTask] = useState(null);
  const [tasks, setTasks] = useState([]);
  const [loading, setLoading] = useState(true);

  function loadTasks(filter) {
    setLoading(true);
    api.tasks(sinceFromFilter(filter ?? dateFilter))
      .then((r) => setTasks((r?.data?.items ?? []).map(mapTask)))
      .catch(() => setTasks([]))
      .finally(() => setLoading(false));
  }

  useEffect(() => { loadTasks(dateFilter); }, [dateFilter]);

  const rows = tasks.filter(
    (r) => !search || r.name.toLowerCase().includes(search.toLowerCase()) || r.type.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="flex h-full overflow-hidden">
      <div className="flex-1 flex flex-col min-w-0">
        <div className="p-4 space-y-3 flex-1 overflow-auto">
          {/* Toolbar */}
          <div className="bg-white border border-[#DEE2E6] px-3 py-2 flex items-center gap-2 flex-wrap">
            <button
              onClick={() => loadTasks(dateFilter)}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]"
            >
              <RefreshCw size={13} />
              REFRESH
            </button>
            <div className="flex border border-[#DEE2E6] overflow-hidden text-xs">
              {[
                { id: 'today', label: 'Today' },
                { id: '7days', label: '7 days' },
                { id: '30days', label: '30 days' },
              ].map((opt) => (
                <button
                  key={opt.id}
                  onClick={() => setDateFilter(opt.id)}
                  className={`px-3 py-1 ${dateFilter === opt.id ? 'bg-[#E6F4FA] text-[#0078A8] border-[#0096D6]' : 'bg-white text-[#555] hover:bg-[#F2F2F2]'}`}
                >
                  {opt.label}
                </button>
              ))}
            </div>
            <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
              <Columns size={13} />
              EDIT COLUMNS
            </button>
            <div className="ml-auto flex items-center gap-2">
              <span className="text-xs text-[#666]">
                {loading ? 'Loading…' : `${rows.length} task${rows.length !== 1 ? 's' : ''}`}
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

          {loading ? (
            <div className="text-xs text-[#888] p-4">Loading tasks…</div>
          ) : rows.length === 0 ? (
            <div className="bg-white border border-[#DEE2E6] p-8 text-center text-xs text-[#888]">
              No tasks yet. Tasks are created when backup, unpack, compare, or restore operations run.
            </div>
          ) : (
            <div className="bg-white border border-[#DEE2E6]">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-[#F2F2F2] border-b border-[#DEE2E6]">
                    {COLUMNS.map((col) => (
                      <th key={col.key} className="text-left px-3 py-2 font-semibold text-[#333]" style={{ width: col.width }}>
                        {col.header}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row) => (
                    <tr
                      key={row.id}
                      onClick={() => setSelectedTask(row)}
                      className={`border-b border-[#DEE2E6] cursor-pointer ${selectedTask?.id === row.id ? 'bg-[#E6F4FA]' : 'hover:bg-[#F9FAFB]'}`}
                    >
                      {COLUMNS.map((col) => (
                        <td key={col.key} className="px-3 py-2">
                          {col.render ? col.render(row) : row[col.key]}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>

      {selectedTask && (
        <TaskDetailPanel task={selectedTask} onClose={() => setSelectedTask(null)} />
      )}
    </div>
  );
}
