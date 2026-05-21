import { useState } from 'react';
import { RotateCcw, Download, Columns, Search, FileText } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';

const MOCK_OBJECTS = [
  { id: 1, name: 'Global Reader Extension', type: 'Custom Role', backupDate: 'Yesterday at 12:00', roleId: 'abc123-4d5e-6f78', description: 'Can read all resources', permissions: '15', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 2, name: 'Helpdesk Manager Plus', type: 'Custom Role', backupDate: 'Yesterday at 12:00', roleId: 'def456-7a8b-9c01', description: 'Can manage help desk', permissions: '8', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 3, name: 'Security Auditor Pro', type: 'Custom Role', backupDate: 'Yesterday at 12:00', roleId: 'ghi789-0b1c-2d34', description: 'Security audit capabilities', permissions: '22', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 4, name: 'Read Users', type: 'Role Permission', backupDate: 'Yesterday at 12:00', roleId: 'abc123-4d5e-6f78', description: 'microsoft.directory/users/read', permissions: '1', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 5, name: 'Read Groups', type: 'Role Permission', backupDate: 'Yesterday at 12:00', roleId: 'abc123-4d5e-6f78', description: 'microsoft.directory/groups/read', permissions: '1', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 6, name: 'John Doe → Global Reader Extension', type: 'Role Assignment', backupDate: 'Yesterday at 12:00', roleId: 'abc123-4d5e-6f78', description: '—', permissions: '—', assignedTo: 'John Doe', principalType: 'User', scope: '/', status: 'Active' },
  { id: 7, name: 'Helpdesk Group → Helpdesk Manager Plus', type: 'Role Assignment', backupDate: 'Yesterday at 12:00', roleId: 'def456-7a8b-9c01', description: '—', permissions: '—', assignedTo: 'Helpdesk Group', principalType: 'Group', scope: '/', status: 'Active' },
  { id: 8, name: 'AU-North / Global Reader Extension', type: 'Assignment Scope', backupDate: 'Yesterday at 12:00', roleId: 'abc123-4d5e-6f78', description: '—', permissions: '—', assignedTo: '—', principalType: '—', scope: 'AU-North', status: 'Active' },
  { id: 9, name: 'Compliance Reader Custom', type: 'Custom Role', backupDate: 'Yesterday at 12:00', roleId: 'jkl012-3m4n-5o67', description: 'Read compliance data', permissions: '6', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 10, name: 'DevOps Lead Role', type: 'Custom Role', backupDate: 'Yesterday at 12:00', roleId: 'pqr345-6s7t-8u90', description: 'Manage DevOps pipelines', permissions: '12', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 11, name: 'Read Applications', type: 'Role Permission', backupDate: 'Yesterday at 12:00', roleId: 'def456-7a8b-9c01', description: 'microsoft.directory/applications/read', permissions: '1', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 12, name: 'ServiceAccount1 → Security Auditor Pro', type: 'Role Assignment', backupDate: 'Yesterday at 12:00', roleId: 'ghi789-0b1c-2d34', description: '—', permissions: '—', assignedTo: 'ServiceAccount1', principalType: 'ServicePrincipal', scope: '/', status: 'Active' },
  { id: 13, name: 'Network Operator Plus', type: 'Custom Role', backupDate: 'Yesterday at 12:00', roleId: 'vwx678-9y0z-1a23', description: 'Manage network resources', permissions: '9', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  { id: 14, name: 'AU-South / Compliance Reader Custom', type: 'Assignment Scope', backupDate: 'Yesterday at 12:00', roleId: 'jkl012-3m4n-5o67', description: '—', permissions: '—', assignedTo: '—', principalType: '—', scope: 'AU-South', status: 'Active' },
  { id: 15, name: 'Finance Admin Role', type: 'Custom Role', backupDate: 'Yesterday at 12:00', roleId: 'bcd901-2e3f-4g56', description: 'Finance department admin', permissions: '18', assignedTo: '—', principalType: '—', scope: '/', status: 'Active' },
  ...Array.from({ length: 20 }, (_, i) => ({
    id: 16 + i,
    name: `Role Object ${16 + i}`,
    type: ['Custom Role', 'Role Permission', 'Role Assignment', 'Assignment Scope'][i % 4],
    backupDate: 'Yesterday at 12:00',
    roleId: `id-${i}`,
    description: '—',
    permissions: String(Math.floor(Math.random() * 20)),
    assignedTo: '—',
    principalType: '—',
    scope: '/',
    status: 'Active',
  })),
];

const columns = [
  {
    key: 'name', header: 'Name ▲', sortable: true, width: '20%',
    render: (r) => (
      <div className="flex items-center gap-1.5">
        <FileText size={12} className="text-[#0096D6] flex-shrink-0" />
        <span className="text-[#0078A8] hover:underline cursor-pointer truncate">{r.name}</span>
      </div>
    )
  },
  { key: 'type', header: 'Object Type', sortable: true, width: '12%' },
  { key: 'backupDate', header: 'Backup Date', sortable: true, width: '13%' },
  { key: 'roleId', header: 'Role Definition ID', width: '12%', render: (r) => <span className="font-mono text-[10px] text-[#666] truncate block max-w-[120px]">{r.roleId}</span> },
  { key: 'description', header: 'Description', width: '16%', render: (r) => <span className="truncate block max-w-[150px]">{r.description}</span> },
  { key: 'permissions', header: 'Permissions', width: '8%' },
  { key: 'assignedTo', header: 'Assigned To', width: '10%' },
  { key: 'principalType', header: 'Principal Type', width: '9%' },
  { key: 'scope', header: 'Scope', width: '6%' },
  {
    key: 'status', header: 'Status', width: '6%',
    render: (r) => <span className="text-[#28A745] font-semibold text-xs">{r.status}</span>
  },
];

export default function UnpackedObjects() {
  const [view, setView] = useState('list');
  const [search, setSearch] = useState('');
  const [filters, setFilters] = useState({ backup: '', type: '', roleType: '', assignType: '', principalType: '', scopeType: '' });

  const rows = MOCK_OBJECTS.filter(
    (r) => !search || r.name.toLowerCase().includes(search.toLowerCase()) || r.type.toLowerCase().includes(search.toLowerCase())
  );

  function setFilter(k, v) { setFilters((f) => ({ ...f, [k]: v })); }

  const filterDefs = [
    { key: 'backup', label: 'Backup' },
    { key: 'type', label: 'Type' },
    { key: 'roleType', label: 'Role Type' },
    { key: 'assignType', label: 'Assignment Type' },
    { key: 'principalType', label: 'Principal Type' },
    { key: 'scopeType', label: 'Scope Type' },
  ];

  return (
    <div className="flex flex-col h-full">
      {/* Toggle + Filter bar */}
      <div className="bg-white border-b border-[#DEE2E6] px-4 py-2 flex items-center gap-3 flex-shrink-0">
        <div className="flex border border-[#DEE2E6] overflow-hidden text-xs">
          {['list', 'objects'].map((v) => (
            <button
              key={v}
              onClick={() => setView(v)}
              className={`px-3 py-1 capitalize ${view === v ? 'bg-[#0078A8] text-white' : 'bg-white text-[#555] hover:bg-[#F2F2F2]'}`}
            >
              {v === 'list' ? 'List View' : 'Objects'}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-2 ml-2 flex-wrap">
          {filterDefs.map(({ key, label }) => (
            <div key={key} className="flex items-center gap-1 text-xs">
              <span className="text-[#555]">{label}:</span>
              <select
                value={filters[key]}
                onChange={(e) => setFilter(key, e.target.value)}
                className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white"
              >
                <option value="">Any</option>
              </select>
            </div>
          ))}
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
          <button className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] hover:bg-[#F2F2F2]">
            <Columns size={13} />
            EDIT COLUMNS
          </button>
          <div className="ml-auto flex items-center gap-2">
            <span className="text-xs text-[#666]">337 objects</span>
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

        <DataTable columns={columns} rows={rows} totalCount={337} />
      </div>
    </div>
  );
}
