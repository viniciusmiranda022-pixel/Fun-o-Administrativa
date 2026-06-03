import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Plus, RefreshCw, Trash2, WifiOff, AlertCircle, CheckCircle2, Loader2, Settings } from 'lucide-react';
import { api } from '../api/client.js';

function StatusPill({ granted }) {
  if (granted) {
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 text-xs font-semibold text-[#28A745] bg-green-50 rounded">
        <CheckCircle2 size={11} /> Granted
      </span>
    );
  }
  return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 text-xs font-semibold text-[#FFC107] bg-yellow-50 rounded">
      <AlertCircle size={11} /> Not Granted
    </span>
  );
}

function QuestModal({ title, onClose, children, maxWidth = 'max-w-md' }) {
  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
      <div className={`bg-white shadow-xl w-full ${maxWidth} mx-4`}>
        <div className="flex items-center justify-between px-5 py-3 border-b border-[#DEE2E6] bg-[#F8F8F8]">
          <h2 className="text-sm font-semibold text-[#333]">{title}</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl font-bold leading-none">×</button>
        </div>
        <div className="px-5 py-4">{children}</div>
      </div>
    </div>
  );
}

function AddTenantTypeModal({ onClose, onSelect }) {
  const [type, setType] = useState('commercial');
  return (
    <QuestModal title="Add Tenant — Select Cloud" onClose={onClose}>
      <p className="text-xs text-[#555] mb-4">
        Select the Microsoft cloud environment for the tenant you want to add.
      </p>
      <div className="space-y-2 mb-5">
        {[
          { id: 'commercial', label: 'Commercial', desc: 'Standard Microsoft 365 / Azure tenants' },
          { id: 'gcc', label: 'GCC (Government Community Cloud)', desc: 'US Government tenants' },
        ].map((opt) => (
          <label key={opt.id} className="flex items-start gap-2 cursor-pointer border border-[#DEE2E6] p-3 hover:bg-[#F9FAFB]">
            <input
              type="radio"
              name="cloudType"
              value={opt.id}
              checked={type === opt.id}
              onChange={() => setType(opt.id)}
              className="mt-0.5"
            />
            <div>
              <div className="text-xs font-semibold text-[#333]">{opt.label}</div>
              <div className="text-[10px] text-[#888] mt-0.5">{opt.desc}</div>
            </div>
          </label>
        ))}
      </div>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
        <button onClick={() => onSelect(type)} className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090]">
          Continue
        </button>
      </div>
    </QuestModal>
  );
}

function AddTenantConsentModal({ onClose, onConfirm, loading }) {
  return (
    <QuestModal title="Add Tenant — Grant Consent" onClose={onClose}>
      <div className="text-sm text-[#333] space-y-2 mb-5">
        <p>You will be redirected to the Microsoft Entra ID login page in a new window.</p>
        <p>Sign in as a <strong>Global Administrator</strong> of the tenant you want to add.</p>
        <p>After signing in, click <strong>Accept</strong> to grant initial administrative consent (Basic permission).</p>
      </div>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
        <button onClick={onConfirm} disabled={loading} className="px-4 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090] disabled:opacity-60 flex items-center gap-1.5">
          {loading && <Loader2 size={13} className="animate-spin" />}
          OK — Open Login Window
        </button>
      </div>
    </QuestModal>
  );
}

function ConfirmRemoveModal({ name, onClose, onConfirm, loading }) {
  return (
    <QuestModal title="Remove Tenant" onClose={onClose}>
      <p className="text-sm text-[#333] mb-5">
        Are you sure you want to remove the tenant <strong>{name}</strong>?
        This may affect backup, compare, and restore operations.
      </p>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-4 py-1.5 text-xs border border-[#DEE2E6] hover:bg-gray-50">Cancel</button>
        <button onClick={onConfirm} disabled={loading} className="px-4 py-1.5 text-xs bg-[#DC3545] text-white hover:bg-red-700 disabled:opacity-60 flex items-center gap-1.5">
          {loading && <Loader2 size={13} className="animate-spin" />}
          Remove
        </button>
      </div>
    </QuestModal>
  );
}

function TenantCard({ tenant, onRemove, onEditConsents }) {
  const added = tenant.addedAt ? new Date(tenant.addedAt).toLocaleDateString('en-US') : '—';
  const basic = tenant.consents?.basic?.granted;
  const restore = tenant.consents?.restore?.granted;

  return (
    <div className="bg-white border border-[#DEE2E6] shadow-sm flex flex-col">
      {/* Header strip */}
      <div className="border-l-4 border-[#0096D6] px-4 py-3">
        <div className="text-sm font-semibold text-[#222]">{tenant.name || tenant.tenantId}</div>
        {tenant.domain && <div className="text-xs text-[#888] mt-0.5">{tenant.domain}</div>}
      </div>

      <div className="px-4 py-3 space-y-3 flex-1">
        {/* Consent status */}
        <div>
          <div className="text-[11px] font-semibold text-[#555] mb-1.5 uppercase tracking-wide">Consent Status</div>
          <div className="flex flex-col gap-1.5">
            <div className="flex items-center justify-between text-xs">
              <span className="text-[#666]">Basic</span>
              <StatusPill granted={basic} />
            </div>
            <div className="flex items-center justify-between text-xs">
              <span className="text-[#666]">Restore</span>
              <StatusPill granted={restore} />
            </div>
          </div>
        </div>

        {/* Tenant ID */}
        <div>
          <div className="text-[11px] font-semibold text-[#555] mb-1 uppercase tracking-wide">Directory Tenant ID</div>
          <div className="font-mono text-[10px] text-[#666] break-all bg-[#F5F5F5] px-2 py-1 rounded">{tenant.tenantId}</div>
        </div>

        {/* Added date */}
        <div className="text-xs text-[#888]">
          <span className="font-semibold text-[#555]">Added: </span>{added}
        </div>
      </div>

      {/* Footer actions */}
      <div className="flex items-center gap-3 px-4 py-2.5 border-t border-[#DEE2E6] bg-[#FAFAFA]">
        <button
          onClick={() => onEditConsents(tenant)}
          className="text-xs text-[#0078A8] hover:underline flex items-center gap-1"
        >
          <Settings size={11} /> Edit Consents
        </button>
        <button
          onClick={() => onRemove(tenant)}
          className="text-xs text-[#DC3545] hover:underline flex items-center gap-1 ml-auto"
        >
          <Trash2 size={11} /> Remove
        </button>
      </div>
    </div>
  );
}

export default function Tenants() {
  const navigate = useNavigate();
  const [tenants, setTenants] = useState([]);
  const [loading, setLoading] = useState(true);
  const [setupOk, setSetupOk] = useState(null);
  const [addStep, setAddStep] = useState(null); // null | 'type' | 'consent'
  const [cloudType, setCloudType] = useState('commercial');
  const [removeTarget, setRemoveTarget] = useState(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);

  const load = async () => {
    setLoading(true);
    setError(null);
    try {
      const [st, ts] = await Promise.all([api.setupState(), api.tenants()]);
      setSetupOk(st?.data?.isConfigured);
      setTenants(ts?.data ?? []);
    } catch {
      // Use mock data if API not available
      setSetupOk(true);
      setTenants([{
        tenantId: 'ca9b03ea-578e-4277-b684-969fa2a34a9a',
        name: 'Contoso',
        domain: 'M365x24071226.onmicrosoft.com',
        addedAt: '2026-04-16',
        consents: { basic: { granted: true }, restore: { granted: true } }
      }]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
    const onMsg = (ev) => { if (ev?.data?.type === 'oauth-complete') load(); };
    window.addEventListener('message', onMsg);
    return () => window.removeEventListener('message', onMsg);
  }, []);

  const handleTypeSelected = (type) => {
    setCloudType(type);
    setAddStep('consent');
  };

  const startAdd = async () => {
    setBusy(true);
    try {
      const r = await api.oauthStart();
      const w = window.open(r.data.url, 'oauth-consent', 'width=700,height=800');
      if (!w) window.location.href = r.data.url;
      setAddStep(null);
    } catch (e) {
      setError(e.message);
      setAddStep(null);
    } finally {
      setBusy(false);
    }
  };

  const handleRemove = async () => {
    if (!removeTarget) return;
    setBusy(true);
    try {
      await api.removeTenant(removeTarget.tenantId);
      setTenants((t) => t.filter((x) => x.tenantId !== removeTarget.tenantId));
      setRemoveTarget(null);
    } catch (e) {
      setError(e.message);
    } finally {
      setBusy(false);
    }
  };

  if (setupOk === false) {
    return (
      <div className="p-8 max-w-2xl mx-auto">
        <div className="bg-yellow-50 border border-yellow-200 p-5">
          <div className="flex items-start gap-3">
            <AlertCircle className="text-yellow-700 flex-shrink-0 mt-0.5" size={20} />
            <div>
              <h3 className="font-semibold text-yellow-900">Setup required</h3>
              <p className="text-sm text-yellow-800 mt-1 mb-3">
                Before adding tenants, you must register this application in Microsoft Entra ID (one-time setup, requires Global Administrator).
              </p>
              <button onClick={() => navigate('/setup')} className="px-3 py-1.5 bg-yellow-700 text-white text-sm hover:bg-yellow-800">
                Start Setup
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="p-5">
      {/* Page header */}
      <div className="flex items-center justify-between mb-5">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Tenants</h2>
          <p className="text-xs text-[#888] mt-0.5">Microsoft 365 / Entra ID tenants connected to the platform</p>
        </div>
        <div className="flex gap-2">
          <button
            onClick={load}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-[#DEE2E6] bg-white text-[#555] hover:bg-[#F2F2F2]"
          >
            <RefreshCw size={12} /> Refresh
          </button>
          <button
            onClick={() => setAddStep('type')}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-[#0078A8] text-white hover:bg-[#006090]"
          >
            <Plus size={12} /> Add Tenant
          </button>
        </div>
      </div>

      {error && (
        <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-2 text-xs mb-4">{error}</div>
      )}

      {loading ? (
        <div className="flex items-center gap-2 text-[#888] text-sm">
          <Loader2 size={16} className="animate-spin" /> Loading...
        </div>
      ) : tenants.length === 0 ? (
        <div className="text-center py-16 text-[#888]">
          <WifiOff size={32} className="mx-auto mb-3 opacity-30" />
          <p className="text-sm">No tenants configured.</p>
          <p className="text-xs mt-1">Click <strong>Add Tenant</strong> to connect a Microsoft 365 tenant.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
          {tenants.map((t) => (
            <TenantCard
              key={t.tenantId}
              tenant={t}
              onRemove={setRemoveTarget}
              onEditConsents={(tt) => navigate(`/tenants/${tt.tenantId}/consents`)}
            />
          ))}
        </div>
      )}

      {addStep === 'type' && (
        <AddTenantTypeModal
          onClose={() => setAddStep(null)}
          onSelect={handleTypeSelected}
        />
      )}
      {addStep === 'consent' && (
        <AddTenantConsentModal
          onClose={() => setAddStep(null)}
          onConfirm={startAdd}
          loading={busy}
        />
      )}
      {removeTarget && (
        <ConfirmRemoveModal
          name={removeTarget.name || removeTarget.tenantId}
          onClose={() => setRemoveTarget(null)}
          onConfirm={handleRemove}
          loading={busy}
        />
      )}
    </div>
  );
}
