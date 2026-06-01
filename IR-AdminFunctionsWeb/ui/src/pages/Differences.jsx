import { useState, useEffect } from 'react';
import { RotateCcw, Download, RefreshCw, Columns, Search, FileText, AlertTriangle } from 'lucide-react';
import DataTable from '../components/common/DataTable.jsx';
import Modal from '../components/common/Modal.jsx';
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

function extractRestoreError(errors) {
  if (!errors?.length) return 'Failed';
  const metaRe = /^(No linha|At line|\+ |[ \t]*\+ Category|[ \t]*\+ FullyQualified)/;
  const useful = errors.filter(l => !metaRe.test(l)).join(' ');
  const marker = 'Restore failed:';
  const idx = useful.indexOf(marker);
  if (idx !== -1) return useful.slice(idx + marker.length).trim();
  return useful.replace(/^[A-Za-z]:\\[^:]+\.ps1\s*:\s*/i, '').trim() || 'Failed';
}

// selectedRows: array of { name, diffType, roleType }
// targetCandidates: array of role names (custom roles that exist in the tenant)
function RestoreModal({ selectedRows, targetCandidates, backupId, onClose, onDone }) {
  const [phase, setPhase] = useState('confirm'); // confirm | running | done
  const [progress, setProgress] = useState({ current: 0, total: selectedRows.length, log: [] });
  const [error, setError] = useState(null);
  const [targets, setTargets] = useState({}); // { roleName: currentRoleName }

  const newRoles = selectedRows.filter(r => r.diffType === 'New');
  const restorableRows = selectedRows.filter(r => r.diffType !== 'New');
  const removedRows = selectedRows.filter(r => r.diffType === 'Removed');

  async function handleApply() {
    if (!backupId) {
      setError('BackupId not available. Run a new RUN COMPARE to reload the results.');
      return;
    }
    setPhase('running');
    const toRestore = restorableRows.length > 0 ? restorableRows : selectedRows;
    const total = toRestore.length;
    const log = [];
    for (let i = 0; i < total; i++) {
      const { name: roleName } = toRestore[i];
      const currentRoleName = targets[roleName] || undefined;
      const label = currentRoleName ? `${currentRoleName} → ${roleName}` : roleName;
      setProgress({ current: i + 1, total, log: [...log, `Restoring: ${label}…`] });
      try {
        const resp = await api.applyRestore({ backupId, roleName, currentRoleName });
        const job = resp?.data;
        if (!job?.id) throw new Error('No job id returned');
        const result = await pollJob(job.id, { intervalMs: 3000, timeoutMs: 300000 });
        const scriptFailed = result?.status === 'Failed' || result?.result?.success === false;
        if (scriptFailed) {
          const errMsg = result?.error || extractRestoreError(result?.result?.errors) || 'Failed';
          log.push(`✗ ${label}: ${errMsg}`);
        } else {
          log.push(`✓ ${label}: restored successfully`);
        }
      } catch (err) {
        log.push(`✗ ${label}: ${err.message}`);
      }
      setProgress({ current: i + 1, total, log: [...log] });
    }
    setPhase('done');
  }

  return (
    <Modal title="Restore Selected Roles" onClose={phase !== 'running' ? onClose : undefined} maxWidth="max-w-lg">
      <div className="p-5 space-y-4">
        {phase === 'confirm' && (
          <>
            <div className="flex items-start gap-2 bg-[#FFF8E1] border border-[#FFC107] p-3 text-xs text-[#856404]">
              <AlertTriangle size={14} className="flex-shrink-0 mt-0.5" />
              <span>
                This operation will restore {restorableRows.length > 0 ? restorableRows.length : selectedRows.length} role(s) to the backup state. This action cannot be automatically undone.
              </span>
            </div>
            {newRoles.length > 0 && (
              <div className="flex items-start gap-2 bg-[#FFF3CD] border border-[#FFC107] p-3 text-xs text-[#856404]">
                <AlertTriangle size={14} className="flex-shrink-0 mt-0.5" />
                <span>
                  <strong>{newRoles.length} role(s) with status "New"</strong> were skipped: these roles only exist in the current tenant and have no snapshot in the backup to restore.
                  {newRoles.length > 0 && (
                    <span className="block mt-1 font-mono">{newRoles.map(r => r.name).join(', ')}</span>
                  )}
                </span>
              </div>
            )}
            {!backupId && (
              <div className="text-xs text-[#DC3545]">
                Warning: BackupId not available. Run a new RUN COMPARE so that the results include the backup ID.
              </div>
            )}
            {restorableRows.length === 0 && newRoles.length > 0 && (
              <div className="text-xs text-[#DC3545]">
                No selected role can be restored. Select roles with status "Removed" or "Changed".
              </div>
            )}
            {removedRows.length > 0 && (
              <div className="flex items-start gap-2 bg-[#E7F3F8] border border-[#0078A8] p-3 text-xs text-[#0078A8]">
                <AlertTriangle size={14} className="flex-shrink-0 mt-0.5" />
                <span>
                  For <strong>Removed</strong> roles: if the role was renamed in the tenant, choose its current name under "Apply over" — the backup will be applied to that role, renaming it back. Leave as "Create new role" if it was truly deleted.
                </span>
              </div>
            )}
            <div>
              <p className="text-xs font-semibold text-[#333] mb-1">Selected roles:</p>
              <ul className="space-y-1.5 max-h-56 overflow-auto">
                {selectedRows.map((r) => (
                  <li key={r.name} className="text-xs text-[#555]">
                    <div className="flex items-center gap-1.5">
                      <RotateCcw size={11} className={r.diffType === 'New' ? 'text-[#FFC107] flex-shrink-0' : 'text-[#0078A8] flex-shrink-0'} />
                      <span className={r.diffType === 'New' ? 'line-through text-[#888]' : ''}>{r.name}</span>
                      <span className="text-[#888]">— {r.diffType}</span>
                      {r.diffType === 'New' && <span className="text-[#888] italic">(skipped)</span>}
                    </div>
                    {r.diffType === 'Removed' && (
                      <div className="ml-5 mt-1 flex items-center gap-1.5 flex-wrap">
                        <span className="text-[#888]">Apply over:</span>
                        <select
                          value={targets[r.name] || ''}
                          onChange={(e) => setTargets((t) => ({ ...t, [r.name]: e.target.value }))}
                          className="border border-[#DEE2E6] px-1.5 py-0.5 text-xs bg-white max-w-[14rem]"
                        >
                          <option value="">Create new role</option>
                          {targetCandidates.map((c) => (
                            <option key={c} value={c}>{c}</option>
                          ))}
                        </select>
                      </div>
                    )}
                  </li>
                ))}
              </ul>
            </div>
            {error && <p className="text-xs text-[#DC3545]">{error}</p>}
            <div className="flex justify-end gap-2 pt-1">
              <button
                onClick={onClose}
                className="px-4 py-1.5 text-xs border border-[#DEE2E6] bg-white text-[#555] hover:bg-[#F2F2F2]"
              >
                CANCEL
              </button>
              <button
                onClick={handleApply}
                disabled={!backupId || restorableRows.length === 0}
                className="px-4 py-1.5 text-xs font-semibold bg-[#0078A8] text-white hover:bg-[#005f87] disabled:opacity-40 disabled:cursor-not-allowed"
              >
                RESTORE
              </button>
            </div>
          </>
        )}

        {phase === 'running' && (
          <div className="space-y-3">
            <p className="text-xs text-[#555]">
              Restoring {progress.current}/{progress.total}…
            </p>
            <div className="w-full bg-[#E9ECEF] h-1.5">
              <div
                className="bg-[#0078A8] h-1.5 transition-all"
                style={{ width: `${(progress.current / progress.total) * 100}%` }}
              />
            </div>
            <ul className="space-y-0.5 max-h-48 overflow-auto">
              {progress.log.map((line, i) => (
                <li key={i} className="text-xs text-[#555] font-mono">{line}</li>
              ))}
            </ul>
          </div>
        )}

        {phase === 'done' && (
          <div className="space-y-3">
            <p className="text-xs font-semibold text-[#28A745]">Operation completed.</p>
            <ul className="space-y-0.5 max-h-48 overflow-auto">
              {progress.log.map((line, i) => (
                <li key={i} className={`text-xs font-mono ${line.startsWith('✗') ? 'text-[#DC3545]' : 'text-[#28A745]'}`}>{line}</li>
              ))}
            </ul>
            <div className="flex justify-end">
              <button
                onClick={onDone}
                className="px-4 py-1.5 text-xs font-semibold bg-[#0078A8] text-white hover:bg-[#005f87]"
              >
                CLOSE
              </button>
            </div>
          </div>
        )}
      </div>
    </Modal>
  );
}

export default function Differences() {
  const [view, setView] = useState('list');
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [rows, setRows] = useState([]);
  const [generatedAt, setGeneratedAt] = useState(null);
  const [backupId, setBackupId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [running, setRunning] = useState(false);
  const [error, setError] = useState(null);
  const [selectedIds, setSelectedIds] = useState([]);
  const [showRestoreModal, setShowRestoreModal] = useState(false);

  function loadResults() {
    setLoading(true);
    api.compareResults()
      .then((r) => {
        if (!r?.success) { setRows([]); return; }
        const data = r?.data;
        if (!data) { setRows([]); return; }
        setGeneratedAt(data.generatedAt);
        setBackupId(data.backupId ?? null);
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
    setSelectedIds([]);
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

  const selectedRoleNames = selectedIds
    .map((id) => rows.find((r) => r.id === id))
    .filter(Boolean)
    .map((r) => ({ name: r.name, diffType: r.diffType, roleType: r.roleType }));

  const targetCandidates = rows
    .filter((r) => r.roleType === 'Custom' && r.diffType !== 'Removed')
    .map((r) => r.name)
    .sort();

  return (
    <div className="flex flex-col h-full">
      {showRestoreModal && (
        <RestoreModal
          selectedRows={selectedRoleNames}
          targetCandidates={targetCandidates}
          backupId={backupId}
          onClose={() => setShowRestoreModal(false)}
          onDone={() => { setShowRestoreModal(false); setSelectedIds([]); loadResults(); }}
        />
      )}

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
          <button
            onClick={() => setShowRestoreModal(true)}
            disabled={selectedIds.length === 0}
            className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold border border-[#DEE2E6] bg-white text-[#0078A8] ${
              selectedIds.length === 0
                ? 'opacity-40 cursor-not-allowed'
                : 'hover:bg-[#F2F2F2] cursor-pointer'
            }`}
          >
            <RotateCcw size={13} />
            RESTORE{selectedIds.length > 0 ? ` (${selectedIds.length})` : ''}
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
          <DataTable
            columns={columns}
            rows={filtered}
            totalCount={filtered.length}
            onSelectionChange={setSelectedIds}
          />
        )}
      </div>
    </div>
  );
}
