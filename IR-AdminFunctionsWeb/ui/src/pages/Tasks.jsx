import { useState } from 'react';
import { CheckCircle2, Columns, Search, X, Clock, AlertTriangle } from 'lucide-react';

const MOCK_TASKS = [
  { id: 1, name: 'Unpack backup 2026-05-19 15:00:16', status: 'Completed', type: 'Unpacking', modified: 'Today 1:22 PM', created: 'Today 1:20 PM', operation: 'Unpacked 337 objects', progress: 100 },
  { id: 2, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: 'Today 1:12 PM', created: 'Today 1:12 PM', operation: 'Backed up 32 role objects', progress: 100 },
  { id: 3, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: '05/19/2026 3:05 PM', created: '05/19/2026 3:00 PM', operation: 'Backed up 32 role objects', progress: 100 },
  { id: 4, name: 'Unpack backup 2026-05-18 06:00:05', status: 'Completed', type: 'Unpacking', modified: '05/18/2026 6:05 AM', created: '05/18/2026 6:00 AM', operation: 'Unpacked 335 objects', progress: 100 },
  { id: 5, name: 'Diff restore objects', status: 'Completed', type: 'Restore', modified: '05/06/2026 3:20 PM', created: '05/06/2026 3:14 PM', operation: 'Restored 3 objects', progress: 100 },
  { id: 6, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: '05/06/2026 6:10 AM', created: '05/06/2026 6:00 AM', operation: 'Backed up 32 role objects', progress: 100 },
  { id: 7, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: '05/05/2026 6:08 AM', created: '05/05/2026 6:00 AM', operation: 'Backed up 31 role objects', progress: 100 },
  { id: 8, name: 'Backup tenant Contoso', status: 'Failed', type: 'Backup', modified: '05/04/2026 6:05 AM', created: '05/04/2026 6:00 AM', operation: 'Access denied', progress: 0 },
  { id: 9, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: '05/03/2026 6:07 AM', created: '05/03/2026 6:00 AM', operation: 'Backed up 31 role objects', progress: 100 },
  { id: 10, name: 'Unpack backup 2026-05-02 06:00:00', status: 'Completed', type: 'Unpacking', modified: '05/02/2026 6:06 AM', created: '05/02/2026 6:00 AM', operation: 'Unpacked 330 objects', progress: 100 },
  { id: 11, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: '05/02/2026 6:08 AM', created: '05/02/2026 6:00 AM', operation: 'Backed up 31 role objects', progress: 100 },
  { id: 12, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: '05/01/2026 6:05 AM', created: '05/01/2026 6:00 AM', operation: 'Backed up 31 role objects', progress: 100 },
  { id: 13, name: 'Backup tenant Contoso', status: 'Warning', type: 'Backup', modified: '04/30/2026 6:06 AM', created: '04/30/2026 6:00 AM', operation: 'Backed up with warnings', progress: 100 },
  { id: 14, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: '04/29/2026 6:05 AM', created: '04/29/2026 6:00 AM', operation: 'Backed up 30 role objects', progress: 100 },
  { id: 15, name: 'Backup tenant Contoso', status: 'Completed', type: 'Backup', modified: '04/28/2026 6:05 AM', created: '04/28/2026 6:00 AM', operation: 'Backed up 30 role objects', progress: 100 },
  ...Array.from({ length: 40 }, (_, i) => ({
    id: 16 + i,
    name: `Backup tenant Contoso`,
    status: i % 10 === 3 ? 'Failed' : i % 8 === 5 ? 'Warning' : 'Completed',
    type: i % 5 === 0 ? 'Unpacking' : 'Backup',
    modified: `04/${String(27 - Math.floor(i / 2)).padStart(2, '0')}/2026 6:05 AM`,
    created: `04/${String(27 - Math.floor(i / 2)).padStart(2, '0')}/2026 6:00 AM`,
    operation: 'Backed up 30 role objects',
    progress: i % 10 === 3 ? 0 : 100,
  })),
];

function StatusIcon({ status }) {
  if (status === 'Completed') return <CheckCircle2 size={13} className="text-[#28A745]" />;
  if (status === 'Failed') return <X size={13} className="text-[#DC3545]" />;
  if (status === 'Warning') return <AlertTriangle size={13} className="text-[#FFC107]" />;
  return <Clock size={13} className="text-[#17A2B8]" />;
}

function TaskDetailPanel({ task, onClose }) {
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
            <span className={
              task.status === 'Completed' ? 'text-[#28A745] font-semibold' :
              task.status === 'Failed' ? 'text-[#DC3545] font-semibold' :
              task.status === 'Warning' ? 'text-[#FFC107] font-semibold' :
              'text-[#17A2B8] font-semibold'
            }>{task.status}</span>
          </div>
        </div>
        <div>
          <div className="text-[#888] mb-0.5">Last operation</div>
          <div className="text-[#333]">{task.operation}</div>
        </div>
        <div>
          <div className="text-[#888] mb-0.5">Created</div>
          <div className="text-[#333]">{task.created}</div>
        </div>
        <div>
          <div className="text-[#888] mb-0.5">Modified</div>
          <div className="text-[#333]">{task.modified}</div>
        </div>
        {task.type === 'Unpacking' && (
          <>
            <div>
              <div className="text-[#888] mb-0.5">Processed objects</div>
              <div className="text-[#333]">337</div>
            </div>
            <div className="grid grid-cols-3 gap-2">
              <div className="text-center p-2 border border-[#DEE2E6] rounded">
                <div className="text-[#28A745] font-bold text-base">337</div>
                <div className="text-[9px] text-[#888]">Success</div>
              </div>
              <div className="text-center p-2 border border-[#DEE2E6] rounded">
                <div className="text-[#FFC107] font-bold text-base">0</div>
                <div className="text-[9px] text-[#888]">Warning</div>
              </div>
              <div className="text-center p-2 border border-[#DEE2E6] rounded">
                <div className="text-[#DC3545] font-bold text-base">0</div>
                <div className="text-[9px] text-[#888]">Error</div>
              </div>
            </div>
          </>
        )}
        <div className="pt-2 border-t border-[#DEE2E6]">
          <button className="text-[#0078A8] hover:underline text-xs">View Events →</button>
        </div>
      </div>
    </div>
  );
}

const columns = [
  {
    key: 'status', header: '', width: '3%',
    render: (r) => <StatusIcon status={r.status} />
  },
  {
    key: 'name', header: 'Name', sortable: true, width: '28%',
    render: (r) => <span className="text-[#0078A8] hover:underline cursor-pointer">{r.name}</span>
  },
  {
    key: 'status2', header: 'Status', width: '10%',
    render: (r) => (
      <span className={`text-xs font-semibold ${r.status === 'Completed' ? 'text-[#28A745]' : r.status === 'Failed' ? 'text-[#DC3545]' : 'text-[#FFC107]'}`}>
        {r.status}
      </span>
    )
  },
  { key: 'type', header: 'Type', sortable: true, width: '10%' },
  { key: 'modified', header: 'Modified', sortable: true, width: '14%' },
  { key: 'created', header: 'Created', sortable: true, width: '14%' },
  { key: 'operation', header: 'Operation', width: '17%' },
  {
    key: 'progress', header: 'Progress', width: '4%',
    render: (r) => <span className="text-xs">{r.progress}%</span>
  },
];

export default function Tasks() {
  const [search, setSearch] = useState('');
  const [selectedTask, setSelectedTask] = useState(null);

  const rows = MOCK_TASKS.filter(
    (r) => !search || r.name.toLowerCase().includes(search.toLowerCase()) || r.type.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="flex h-full overflow-hidden">
      <div className="flex-1 flex flex-col min-w-0">
        <div className="p-4 space-y-3 flex-1 overflow-auto">
          {/* Toolbar */}
          <div className="bg-white border border-[#DEE2E6] px-3 py-2 flex items-center gap-2">
            <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
              <Columns size={13} />
              EDIT COLUMNS
            </button>
            <div className="ml-auto flex items-center gap-2">
              <span className="text-xs text-[#666]">55 tasks</span>
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

          <DataTable
            columns={columns}
            rows={rows}
            totalCount={55}
            onRowClick={(row) => setSelectedTask(row)}
            selectedRows={selectedTask ? [selectedTask.id] : []}
          />
        </div>
      </div>

      {selectedTask && (
        <TaskDetailPanel task={selectedTask} onClose={() => setSelectedTask(null)} />
      )}
    </div>
  );
}
