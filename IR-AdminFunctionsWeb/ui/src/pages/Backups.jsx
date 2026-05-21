import { useState } from 'react';
import { Package, Columns, Search } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';

/* ── Mock data ── */
const MOCK_BACKUPS = Array.from({ length: 15 }, (_, i) => {
  const daysAgo = i < 2 ? 0 : Math.floor(i / 2);
  const hours = [18, 13, 6, 0, 18, 13, 6, 0, 18, 13, 6, 0, 18, 13, 6][i];
  const mins = i % 3 === 1 ? 12 : 0;
  const label = daysAgo === 0
    ? `Today at ${hours}:${String(mins).padStart(2, '0')} ${hours >= 12 ? 'PM' : 'AM'}`
    : daysAgo === 1
    ? `Yesterday at ${hours}:${String(mins).padStart(2, '0')} ${hours >= 12 ? 'PM' : 'AM'}`
    : `${daysAgo} days ago at ${hours}:00 ${hours >= 12 ? 'PM' : 'AM'}`;

  return {
    id: `backup-${i}`,
    time: label,
    tenant: 'Contoso',
    duration: i === 0 ? 'a few seconds' : `${Math.floor(Math.random() * 10) + 1}s`,
    customRoles: 32,
    rolePerms: 128,
    roleAssignments: 97,
    scopes: 55,
    refPrincipals: 25,
    size: `${(Math.random() * 2 + 0.5).toFixed(1)} MB`,
    status: 'Completed',
    unpacked: i === 0 || i === 1 ? 'Yes' : '—',
  };
});

const columns = [
  { key: 'time', header: 'Time', sortable: true, width: '15%' },
  { key: 'tenant', header: 'Tenant', sortable: true, width: '10%' },
  { key: 'duration', header: 'Duration', width: '8%' },
  { key: 'customRoles', header: 'Custom Roles', width: '9%' },
  { key: 'rolePerms', header: 'Role Permissions', width: '11%' },
  { key: 'roleAssignments', header: 'Role Assignments', width: '11%' },
  { key: 'scopes', header: 'Assignment Scopes', width: '12%' },
  { key: 'refPrincipals', header: 'Referenced Principals', width: '14%' },
  { key: 'size', header: 'Backup Size', width: '8%' },
  {
    key: 'status', header: 'Status', width: '8%',
    render: (r) => (
      <span className={`inline-flex items-center gap-1 text-xs font-semibold ${r.status === 'Completed' ? 'text-[#28A745]' : 'text-[#DC3545]'}`}>
        {r.status}
      </span>
    )
  },
  { key: 'unpacked', header: 'Unpacked', width: '6%' },
];

/* Simple mini bar chart */
function MiniBarChart() {
  const bars = [60, 45, 70, 55, 80, 65, 50, 90, 75, 40, 85, 60, 95, 70, 50, 65, 80, 55, 40, 75, 60, 45, 90, 70, 55, 80, 65, 50, 40, 85, 70, 90, 55, 65, 40];
  const labels = ['Apr 16', '', '', '', 'Apr 23', '', '', '', 'Apr 30', '', '', '', 'May 7', '', '', '', 'May 14', '', '', '', 'May 20'];
  return (
    <div className="bg-white border border-[#DEE2E6] px-4 pt-3 pb-2">
      <div className="flex items-end gap-0.5" style={{ height: 60 }}>
        {bars.map((h, i) => (
          <div key={i} className="flex-1 flex flex-col justify-end">
            <div className="w-full rounded-sm" style={{ height: `${h}%`, backgroundColor: '#6CB33F' }} />
          </div>
        ))}
      </div>
      <div className="flex justify-between text-[9px] text-[#999] mt-1 px-0">
        {['Apr 16', 'Apr 23', 'Apr 30', 'May 7', 'May 14', 'May 20'].map((l, i) => (
          <span key={i}>{l}</span>
        ))}
      </div>
    </div>
  );
}

export default function Backups() {
  const [search, setSearch] = useState('');
  const [dateFilter, setDateFilter] = useState('30days');

  const rows = MOCK_BACKUPS.filter(
    (r) => !search || r.time.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="flex flex-col h-full">
      {/* Filter bar */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-3 flex-shrink-0">
        <span className="text-xs text-[#555]">Status:</span>
        <select className="border border-[#DEE2E6] px-2 py-1 text-xs bg-white">
          <option>Any</option>
          <option>Completed</option>
          <option>Failed</option>
        </select>
        <div className="flex items-center gap-0.5 ml-3">
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
        {/* Mini bar chart */}
        <MiniBarChart />

        {/* Action toolbar */}
        <div className="bg-white border border-[#DEE2E6] px-3 py-2 flex items-center gap-2">
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Package size={13} />
            UNPACK
          </button>
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Columns size={13} />
            EDIT COLUMNS
          </button>
          <div className="ml-auto flex items-center gap-2">
            <span className="text-xs text-[#666]">150 objects</span>
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

        <DataTable columns={columns} rows={rows} totalCount={150} />
      </div>
    </div>
  );
}
