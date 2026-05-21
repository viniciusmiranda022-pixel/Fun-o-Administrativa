import { useState } from 'react';
import { Download, CheckSquare, Columns, Search, Info, AlertTriangle, XCircle } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';

const MOCK_EVENTS = [
  { id: 1, time: 'Today at 1:22 PM', severity: 'Info', description: 'The Differences report has been updated successfully.', objectName: '—', taskName: 'Unpack backup 2026-05-19 15:00:16' },
  { id: 2, time: 'Today at 1:21 PM', severity: 'Info', description: 'The Differences report has been updated successfully.', objectName: '—', taskName: 'Unpack backup 2026-05-19 15:00:16' },
  { id: 3, time: 'Today at 1:21 PM', severity: 'Info', description: 'Update of the Differences report has been started.', objectName: '—', taskName: 'Unpack backup 2026-05-19 15:00:16' },
  { id: 4, time: 'Today at 1:20 PM', severity: 'Info', description: 'Unpacked backup: 2026-05-19T15:00:16.889866Z', objectName: '—', taskName: 'Unpack backup 2026-05-19 15:00:16' },
  { id: 5, time: 'Today at 1:20 PM', severity: 'Info', description: 'Task has been started (Started by user)', objectName: '—', taskName: 'Unpack backup 2026-05-19 15:00:16' },
  { id: 6, time: 'Today at 1:12 PM', severity: 'Info', description: 'Task has been started (Started by user)', objectName: '—', taskName: 'Backup tenant Contoso' },
  { id: 7, time: 'Today at 1:11 PM', severity: 'Info', description: 'Backup completed: 32 role objects backed up successfully.', objectName: '—', taskName: 'Backup tenant Contoso' },
  { id: 8, time: 'Today at 1:10 PM', severity: 'Info', description: 'Backup task initiated for tenant Contoso.', objectName: 'Contoso', taskName: 'Backup tenant Contoso' },
  { id: 9, time: '05/19/2026 3:05 PM', severity: 'Info', description: 'Unpack operation completed: 337 objects unpacked.', objectName: '—', taskName: 'Unpack backup 2026-05-19 15:00:16' },
  { id: 10, time: '05/19/2026 3:00 PM', severity: 'Info', description: 'Task has been started (Scheduled)', objectName: '—', taskName: 'Backup tenant Contoso' },
  { id: 11, time: '05/06/2026 3:19 PM', severity: 'Info', description: 'Directory role Agent ID Administrator was assigned to group Helpdesk Managers.', objectName: 'Helpdesk Managers', taskName: 'Diff restore objects' },
  { id: 12, time: '05/06/2026 3:18 PM', severity: 'Info', description: 'Object attributes were restored successfully.', objectName: '—', taskName: 'Diff restore objects' },
  { id: 13, time: '05/06/2026 3:18 PM', severity: 'Info', description: 'Group was re-created in directory.', objectName: 'Helpdesk Managers', taskName: 'Diff restore objects' },
  { id: 14, time: '05/06/2026 3:15 PM', severity: 'Warning', description: 'Referenced principal could not be found; assignment skipped.', objectName: '—', taskName: 'Diff restore objects' },
  { id: 15, time: '05/06/2026 3:14 PM', severity: 'Error', description: 'Quest Recovery Function - Administrative Roles Backup failed: access denied.', objectName: '—', taskName: 'Backup tenant Contoso' },
  ...Array.from({ length: 385 }, (_, i) => ({
    id: 16 + i,
    time: `05/${String(Math.floor(i / 20) + 1).padStart(2, '0')}/2026 ${Math.floor(Math.random() * 12) + 1}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')} PM`,
    severity: ['Info', 'Info', 'Info', 'Warning', 'Error'][i % 5],
    description: ['Backup completed successfully.', 'Differences report updated.', 'Task started by scheduler.', 'Tenant consent check passed.', 'Role assignment synced.'][i % 5],
    objectName: '—',
    taskName: ['Backup tenant Contoso', 'Unpack backup', 'Diff restore objects'][i % 3],
  })),
];

function SeverityIcon({ severity }) {
  if (severity === 'Error') return <XCircle size={13} className="text-[#DC3545]" />;
  if (severity === 'Warning') return <AlertTriangle size={13} className="text-[#FFC107]" />;
  return <Info size={13} className="text-[#17A2B8]" />;
}

const columns = [
  {
    key: 'severity', header: '', width: '3%',
    render: (r) => <SeverityIcon severity={r.severity} />
  },
  { key: 'time', header: 'Time ▼', sortable: true, width: '14%' },
  {
    key: 'description', header: 'Description', width: '40%',
    render: (r) => <span className="line-clamp-2 text-xs">{r.description}</span>
  },
  { key: 'objectName', header: 'Object Name', width: '16%' },
  {
    key: 'taskName', header: 'Task Name', width: '27%',
    render: (r) => <span className="text-[#0078A8] hover:underline cursor-pointer text-xs">{r.taskName}</span>
  },
];

function MiniHistogram() {
  const bars = Array.from({ length: 30 }, (_, i) => ({
    h: Math.floor(Math.random() * 70) + 20,
    color: i === 28 || i === 29 ? '#17A2B8' : i % 7 === 0 ? '#DC3545' : '#17A2B8',
  }));
  return (
    <div className="bg-white border border-[#DEE2E6] px-4 pt-2 pb-1.5" style={{ height: 52 }}>
      <div className="flex items-end gap-0.5 h-8">
        {bars.map((b, i) => (
          <div key={i} className="flex-1 rounded-sm" style={{ height: `${b.h}%`, backgroundColor: b.color }} />
        ))}
      </div>
      <div className="flex justify-between text-[9px] text-[#999] mt-0.5">
        <span>Apr 22</span><span>Apr 29</span><span>May 6</span><span>May 13</span><span>May 20</span>
      </div>
    </div>
  );
}

export default function Events() {
  const [severity, setSeverity] = useState('');
  const [relevance, setRelevance] = useState('');
  const [dateFilter, setDateFilter] = useState('30days');
  const [search, setSearch] = useState('');

  const rows = MOCK_EVENTS.slice(0, 400).filter((r) => {
    if (severity && r.severity !== severity) return false;
    if (search && !r.description.toLowerCase().includes(search.toLowerCase()) && !r.taskName.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  return (
    <div className="flex flex-col h-full">
      {/* Filter bar */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-3 flex-shrink-0 flex-wrap">
        <div className="flex items-center gap-1 text-xs">
          <span className="text-[#555]">Severity:</span>
          <select
            value={severity}
            onChange={(e) => setSeverity(e.target.value)}
            className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white"
          >
            <option value="">Any</option>
            <option value="Info">Information</option>
            <option value="Warning">Warning</option>
            <option value="Error">Error</option>
          </select>
        </div>
        <div className="flex items-center gap-1 text-xs">
          <span className="text-[#555]">Relevance:</span>
          <select
            value={relevance}
            onChange={(e) => setRelevance(e.target.value)}
            className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white"
          >
            <option value="">Any</option>
          </select>
        </div>
        <div className="flex items-center gap-0.5">
          {[
            { id: 'today', label: 'Today' },
            { id: '7days', label: '7 days' },
            { id: '30days', label: '30 days' },
          ].map((opt) => (
            <button
              key={opt.id}
              onClick={() => setDateFilter(opt.id)}
              className={`px-3 py-1 text-xs border border-[#DEE2E6] ${dateFilter === opt.id ? 'bg-[#E6F4FA] text-[#0078A8] border-[#0096D6]' : 'bg-white text-[#555] hover:bg-[#F2F2F2]'}`}
            >
              {opt.label}
            </button>
          ))}
          <button className="px-3 py-1 text-xs border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2] ml-1">
            Custom range
          </button>
        </div>
      </div>

      <div className="p-4 space-y-3 flex-1 overflow-auto">
        <MiniHistogram />

        {/* Action toolbar */}
        <div className="bg-white border border-[#DEE2E6] px-3 py-2 flex items-center gap-2">
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Download size={13} />
            EXPORT
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <CheckSquare size={13} />
            ACKNOWLEDGE
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Columns size={13} />
            EDIT COLUMNS
          </button>
          <div className="ml-auto flex items-center gap-2">
            <span className="text-xs text-[#666]">612 events</span>
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

        <DataTable columns={columns} rows={rows} totalCount={612} showCheckboxes={true} />
      </div>
    </div>
  );
}
