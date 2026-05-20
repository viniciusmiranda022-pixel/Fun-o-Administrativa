import { useEffect, useState } from 'react';
import { Loader2, Archive } from 'lucide-react';
import { api } from '../../api/client.js';
import { useTenant } from '../../context/TenantContext.jsx';

/* ─── Backup Unpacking modal ─────────────────────────────────────────── */
export default function UnpackBackupModal({ onClose, onUnpacked }) {
  const { selected } = useTenant();
  const [backups, setBackups]   = useState([]);
  const [selectedId, setSelectedId] = useState(null);
  const [loading, setLoading]   = useState(true);
  const [unpacking, setUnpacking] = useState(false);
  const [error, setError]       = useState(null);
  const [opts, setOpts] = useState({
    clearPrevious:      true,
    computeDifferences: true,
    validateIntegrity:  false,
  });

  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        const r = await api.backups();
        const list = (r?.data ?? []).filter((b) =>
          !selected?.tenantId || !b.tenantId || b.tenantId === selected.tenantId);
        setBackups(list);
        if (list.length > 0) setSelectedId(list[0].id);
      } finally {
        setLoading(false);
      }
    })();
  }, [selected?.tenantId]); /* eslint-disable-line */

  const handleUnpack = async () => {
    if (!selectedId) return;
    setUnpacking(true);
    setError(null);
    try {
      const r = await api.unpackBackup(selectedId, opts);
      onUnpacked?.(r?.data?.id, selectedId);
      onClose();
    } catch (e) {
      setError(e.message);
    } finally {
      setUnpacking(false);
    }
  };

  const Opt = ({ k, label }) => (
    <label className="flex items-start gap-2 text-xs cursor-pointer select-none">
      <input type="checkbox" checked={opts[k]}
        onChange={(e) => setOpts({ ...opts, [k]: e.target.checked })} className="mt-0.5 flex-shrink-0" />
      <span className="text-[#333]">{label}</span>
    </label>
  );

  const fmtDate = (b) => {
    const d = b.collectedAt || b.startedAt || b.modifiedAt;
    if (!d) return '—';
    const date = new Date(d);
    const now = new Date();
    const diff = now - date;
    const mins = Math.floor(diff / 60000);
    const hrs  = Math.floor(diff / 3600000);
    if (mins < 1)  return 'just now';
    if (hrs < 24)  return `Today at ${date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true })}`;
    if (hrs < 48)  return `Yesterday at ${date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true })}`;
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  return (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center">
      <div className="bg-white shadow-2xl w-full max-w-[640px] mx-4 max-h-[85vh] flex flex-col rounded-sm">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-3.5 border-b border-border-light">
          <h2 className="text-sm font-semibold text-[#222]">Backup Unpacking</h2>
          <button onClick={onClose} className="text-[#0078A8] hover:text-blue-800 font-bold text-lg leading-none">×</button>
        </div>

        <div className="flex-1 overflow-y-auto px-6 py-4 space-y-4">
          {/* Options */}
          <div className="space-y-2">
            <Opt k="clearPrevious"      label="Clear objects from previously unpacked backups" />
            <Opt k="computeDifferences" label="Perform differences operation during unpack" />
          </div>

          <div className="border-t border-border-light pt-3">
            <h3 className="text-xs font-semibold text-[#333] mb-2">Select backup to unpack</h3>

            {error && (
              <div className="bg-red-50 border border-red-200 text-red-800 text-xs px-3 py-2 rounded mb-2">{error}</div>
            )}

            {loading ? (
              <div className="text-center py-8 text-xs text-text-secondary flex items-center justify-center gap-2">
                <Loader2 size={13} className="animate-spin" /> Loading backups…
              </div>
            ) : backups.length === 0 ? (
              <div className="text-center py-8 text-xs text-text-secondary">
                <Archive size={24} className="mx-auto mb-2 opacity-30" />
                No backups available for this tenant.
              </div>
            ) : (
              <div className="border border-border-light rounded-sm overflow-hidden">
                <table className="w-full text-xs">
                  <thead className="bg-[#F0F0F0] sticky top-0">
                    <tr>
                      <th className="text-left px-4 py-2 font-semibold text-[#444]">Created</th>
                      <th className="text-left px-4 py-2 font-semibold text-[#444]">Tenant</th>
                      <th className="text-right px-4 py-2 font-semibold text-[#444]">Roles</th>
                      <th className="text-right px-4 py-2 font-semibold text-[#444]">Size</th>
                      <th className="text-left px-4 py-2 font-semibold text-[#444]">Status</th>
                    </tr>
                  </thead>
                  <tbody className="max-h-[260px] overflow-y-auto block">
                    {backups.slice(0, 25).map((b) => {
                      const sel = b.id === selectedId;
                      return (
                        <tr key={b.id}
                          onClick={() => setSelectedId(b.id)}
                          className={`border-b border-[#E8E8E8] cursor-pointer table-row ${sel ? 'bg-[#EBF4FF]' : 'hover:bg-gray-50'}`}
                        >
                          <td className="px-4 py-2">{fmtDate(b)}</td>
                          <td className="px-4 py-2 text-text-secondary">
                            {b.tenantId ? b.tenantId.substring(0, 12) + '…' : selected?.name || '—'}
                          </td>
                          <td className="px-4 py-2 text-right">{b.roleDefinitionsCount ?? '—'}</td>
                          <td className="px-4 py-2 text-right">{b.sizeDisplay || '—'}</td>
                          <td className="px-4 py-2">{b.manifestError ? 'Failed' : 'Completed'}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <div className="px-4 py-1.5 bg-gray-50 border-t border-border-light text-[11px] text-text-secondary flex items-center justify-between">
                  <span>1–{Math.min(backups.length, 25)} of {backups.length} Results</span>
                  <span>Show: 25</span>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="flex items-center justify-end gap-2 px-6 py-3 border-t border-border-light bg-gray-50">
          <button onClick={onClose} disabled={unpacking}
            className="px-4 py-1.5 text-sm border border-[#0078A8] text-[#0078A8] rounded-sm hover:bg-blue-50 disabled:opacity-50">
            Cancel
          </button>
          <button onClick={handleUnpack} disabled={unpacking || !selectedId || loading}
            className="px-4 py-1.5 text-sm bg-[#0078A8] text-white rounded-sm hover:bg-[#006090] disabled:opacity-50 flex items-center gap-1.5">
            {unpacking && <Loader2 size={13} className="animate-spin" />}
            Unpack
          </button>
        </div>
      </div>
    </div>
  );
}
