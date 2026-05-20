import { useEffect, useState } from 'react';
import { CheckCircle2, AlertCircle, ShieldCheck, ShieldAlert, ExternalLink, Eye, Loader2, X } from 'lucide-react';
import { api } from '../../api/client.js';

function ConsentBadge({ granted }) {
  return granted
    ? <span className="inline-flex items-center gap-1 text-xs font-semibold text-green-700"><CheckCircle2 size={12} /> Granted</span>
    : <span className="inline-flex items-center gap-1 text-xs font-semibold text-yellow-600"><AlertCircle size={12} /> Not Granted</span>;
}

function DetailsPanel({ tenant, onClose }) {
  const basic   = tenant.consents?.basic;
  const restore = tenant.consents?.restore;
  const fmt = (d) => d ? new Date(d).toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' }) : '—';

  return (
    <div className="fixed inset-y-0 right-0 w-[400px] bg-white border-l border-border-light shadow-2xl z-[70] overflow-y-auto">
      <div className="flex items-center justify-between px-5 py-4 border-b border-border-light">
        <h3 className="text-sm font-semibold text-[#222]">Consent Details</h3>
        <button onClick={onClose} className="text-[#0078A8] font-bold text-lg leading-none">×</button>
      </div>
      <div className="p-5 space-y-4 text-xs">
        <div>
          <div className="text-text-secondary uppercase text-[10px] tracking-wide mb-1">Tenant</div>
          <div className="font-semibold text-[#222]">{tenant.name || tenant.tenantId}</div>
          {tenant.domain && <div className="text-text-secondary">{tenant.domain}</div>}
        </div>

        <div className="border border-border-light rounded p-3">
          <div className="flex items-center justify-between mb-2">
            <span className="font-semibold text-[#222]">Administrative Function Recovery – Basic</span>
            <ConsentBadge granted={basic?.granted} />
          </div>
          <p className="text-text-secondary mb-3">Read-only. Required for backup, comparison and reporting.</p>
          <dl className="space-y-1">
            <div className="flex justify-between"><dt className="text-text-secondary">Permission:</dt><dd className="font-mono">RoleManagement.Read.Directory</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Granted at:</dt><dd>{fmt(basic?.grantedAt)}</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Granted by:</dt><dd>{basic?.grantedByUser ?? '—'}</dd></div>
          </dl>
        </div>

        <div className="border border-orange-200 rounded p-3 bg-orange-50/30">
          <div className="flex items-center justify-between mb-2">
            <span className="font-semibold text-[#222] flex items-center gap-1.5">
              Administrative Function Recovery – Restore
              <span className="text-[9px] uppercase font-bold text-orange-700 bg-orange-100 px-1.5 py-0.5 rounded">Sensitive</span>
            </span>
            <ConsentBadge granted={restore?.granted} />
          </div>
          <p className="text-text-secondary mb-3">Write access. Required to restore custom roles, permissions and assignments. Modifies Entra ID RBAC.</p>
          <dl className="space-y-1">
            <div className="flex justify-between"><dt className="text-text-secondary">Permission:</dt><dd className="font-mono">RoleManagement.ReadWrite.Directory</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Granted at:</dt><dd>{fmt(restore?.grantedAt)}</dd></div>
            <div className="flex justify-between"><dt className="text-text-secondary">Granted by:</dt><dd>{restore?.grantedByUser ?? '—'}</dd></div>
          </dl>
        </div>
      </div>
    </div>
  );
}

/* ─── Manage Restores modal ───────────────────────────────────────────── */
export default function ManageRestoresModal({ onClose }) {
  const [tenants, setTenants]   = useState([]);
  const [loading, setLoading]   = useState(true);
  const [selectedId, setSelectedId] = useState(null);
  const [detail, setDetail]     = useState(null);
  const [granting, setGranting] = useState(null);

  const load = async () => {
    setLoading(true);
    try {
      const r = await api.tenants();
      const list = r?.data ?? [];
      setTenants(list);
      if (list.length > 0 && !selectedId) setSelectedId(list[0].tenantId);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
    const onMsg = (ev) => { if (ev?.data?.type === 'oauth-complete') load(); };
    window.addEventListener('message', onMsg);
    return () => window.removeEventListener('message', onMsg);
  }, []); /* eslint-disable-line */

  const startGrant = async (tenantId) => {
    setGranting(tenantId);
    try {
      const r = await api.oauthStart(tenantId);
      const w = window.open(r.data.url, 'oauth-consent', 'width=700,height=800');
      if (!w) window.location.href = r.data.url;
    } catch (e) {
      console.error(e);
    } finally {
      setGranting(null);
    }
  };

  const status = (t) => {
    const b = t.consents?.basic?.granted;
    const r = t.consents?.restore?.granted;
    if (b && r) return { label: 'Fully authorized', Icon: ShieldCheck, cls: 'text-green-700' };
    if (b)      return { label: 'Backup only',      Icon: ShieldAlert,  cls: 'text-yellow-600' };
    return             { label: 'Not authorized',   Icon: ShieldAlert,  cls: 'text-red-600' };
  };

  return (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center">
      <div className="bg-white shadow-2xl w-full max-w-[680px] mx-4 max-h-[80vh] flex flex-col rounded-sm">
        {/* Header */}
        <div className="flex items-center justify-between px-6 py-3.5 border-b border-border-light">
          <h2 className="text-sm font-semibold text-[#222]">Manage Restore</h2>
          <button onClick={onClose} className="text-[#0078A8] hover:text-blue-800 font-bold text-lg leading-none">×</button>
        </div>

        {/* Grid */}
        <div className="flex-1 overflow-auto">
          {loading ? (
            <div className="text-center py-8 text-xs text-text-secondary flex items-center justify-center gap-2">
              <Loader2 size={13} className="animate-spin" /> Loading…
            </div>
          ) : (
            <table className="w-full text-xs">
              <thead className="bg-[#F0F0F0] sticky top-0">
                <tr>
                  <th className="text-left px-5 py-2 font-semibold text-[#444]">Tenant</th>
                  <th className="text-left px-4 py-2 font-semibold text-[#444]">Basic Consent</th>
                  <th className="text-left px-4 py-2 font-semibold text-[#444]">Restore Consent</th>
                  <th className="text-left px-4 py-2 font-semibold text-[#444]">Consent Status</th>
                  <th className="text-right px-4 py-2 font-semibold text-[#444]">Actions</th>
                </tr>
              </thead>
              <tbody>
                {tenants.map((t) => {
                  const sel = t.tenantId === selectedId;
                  const s = status(t);
                  const SIcon = s.Icon;
                  return (
                    <tr key={t.tenantId}
                      onClick={() => setSelectedId(t.tenantId)}
                      className={`border-b border-[#E8E8E8] cursor-pointer ${sel ? 'bg-[#EBF4FF]' : 'hover:bg-gray-50'}`}
                    >
                      <td className="px-5 py-2.5">
                        <div className="font-medium text-[#222]">{t.name || t.tenantId}</div>
                        {t.domain && <div className="text-[10px] text-text-secondary">{t.domain}</div>}
                      </td>
                      <td className="px-4 py-2.5"><ConsentBadge granted={t.consents?.basic?.granted} /></td>
                      <td className="px-4 py-2.5"><ConsentBadge granted={t.consents?.restore?.granted} /></td>
                      <td className="px-4 py-2.5">
                        <span className={`flex items-center gap-1 font-semibold ${s.cls}`}>
                          <SIcon size={12} /> {s.label}
                        </span>
                      </td>
                      <td className="px-4 py-2.5">
                        <div className="flex items-center justify-end gap-3">
                          <button onClick={(e) => { e.stopPropagation(); setDetail(t); }}
                            className="flex items-center gap-1 text-[#0078A8] hover:underline">
                            <Eye size={11} /> View Details
                          </button>
                          <button onClick={(e) => { e.stopPropagation(); startGrant(t.tenantId); }}
                            disabled={granting === t.tenantId}
                            className="flex items-center gap-1 text-[#0078A8] hover:underline disabled:opacity-40">
                            {granting === t.tenantId ? <Loader2 size={11} className="animate-spin" /> : <ExternalLink size={11} />}
                            {(t.consents?.basic?.granted || t.consents?.restore?.granted) ? 'Re-Grant' : 'Grant Consent'}
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
                {tenants.length === 0 && (
                  <tr><td colSpan={5} className="px-5 py-6 text-center text-text-secondary">No tenants configured.</td></tr>
                )}
              </tbody>
            </table>
          )}
        </div>

        {/* Footer */}
        <div className="flex items-center justify-between px-5 py-2.5 border-t border-border-light bg-gray-50">
          <span className="text-[11px] text-text-secondary">1–{tenants.length} of {tenants.length} Results</span>
          <button onClick={onClose}
            className="px-4 py-1.5 text-sm border border-[#0078A8] text-[#0078A8] rounded-sm hover:bg-blue-50">
            Cancel
          </button>
        </div>
      </div>

      {detail && <DetailsPanel tenant={detail} onClose={() => setDetail(null)} />}
    </div>
  );
}
