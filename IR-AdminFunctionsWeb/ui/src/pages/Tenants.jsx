import { useEffect, useState } from 'react';
import { Plus, RefreshCw, Trash2, Wifi, WifiOff, AlertCircle, CheckCircle2, Loader2 } from 'lucide-react';
import { api } from '../api/client.js';

function StatusBadge({ status }) {
  const map = {
    Connected:    { icon: CheckCircle2, cls: 'text-green-600 bg-green-50',   label: 'Connected' },
    Unconfigured: { icon: AlertCircle,  cls: 'text-yellow-600 bg-yellow-50', label: 'Unconfigured' },
    Error:        { icon: WifiOff,      cls: 'text-red-600 bg-red-50',       label: 'Error' },
  };
  const s = map[status] ?? map.Unconfigured;
  const Icon = s.icon;
  return (
    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-semibold ${s.cls}`}>
      <Icon size={12} /> {s.label}
    </span>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-lg mx-4">
        <div className="flex items-center justify-between px-6 py-4 border-b border-border-light">
          <h2 className="text-base font-semibold text-[#222]">{title}</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl font-bold leading-none">×</button>
        </div>
        <div className="px-6 py-4">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, name, value, onChange, placeholder, required }) {
  return (
    <div className="mb-4">
      <label className="block text-xs font-semibold text-[#333] mb-1">
        {label}{required && <span className="text-red-500 ml-0.5">*</span>}
      </label>
      <input
        name={name}
        value={value}
        onChange={onChange}
        placeholder={placeholder}
        className="w-full border border-border-light rounded px-3 py-2 text-sm focus:outline-none focus:border-[#0078A8]"
      />
    </div>
  );
}

function AddTenantModal({ onClose, onAdded }) {
  const [form, setForm] = useState({ name: '', tenantId: '', clientId: '', certificateThumbprint: '', domain: '' });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState(null);

  const handle = (e) => setForm((f) => ({ ...f, [e.target.name]: e.target.value }));

  const submit = async (e) => {
    e.preventDefault();
    setSaving(true);
    setError(null);
    try {
      const resp = await api.addTenant(form);
      if (!resp.success) throw new Error(resp.error);
      onAdded(resp.data);
    } catch (err) {
      setError(err.message);
    } finally {
      setSaving(false);
    }
  };

  return (
    <Modal title="Add Tenant" onClose={onClose}>
      <form onSubmit={submit}>
        <Field label="Name" name="name" value={form.name} onChange={handle} placeholder="Contoso" required />
        <Field label="Directory Tenant ID" name="tenantId" value={form.tenantId} onChange={handle} placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" required />
        <Field label="Client ID (App Registration)" name="clientId" value={form.clientId} onChange={handle} placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" required />
        <Field label="Certificate Thumbprint" name="certificateThumbprint" value={form.certificateThumbprint} onChange={handle} placeholder="40-char hex" required />
        <Field label="Domain (opcional)" name="domain" value={form.domain} onChange={handle} placeholder="contoso.onmicrosoft.com" />
        {error && <p className="text-red-600 text-xs mb-3">{error}</p>}
        <div className="flex justify-end gap-2 pt-2">
          <button type="button" onClick={onClose} className="px-4 py-2 text-sm border border-border-light rounded hover:bg-gray-50">
            Cancelar
          </button>
          <button type="submit" disabled={saving} className="px-4 py-2 text-sm bg-[#0078A8] text-white rounded hover:bg-[#006090] disabled:opacity-60 flex items-center gap-2">
            {saving && <Loader2 size={14} className="animate-spin" />}
            {saving ? 'Salvando...' : 'Add Tenant'}
          </button>
        </div>
      </form>
    </Modal>
  );
}

function ConfirmModal({ name, onClose, onConfirm, loading }) {
  return (
    <Modal title="Remover Tenant" onClose={onClose}>
      <p className="text-sm text-[#333] mb-4">
        Tem certeza que deseja remover o tenant <strong>{name}</strong>?
        Isso pode impactar operações de backup, comparação e restore configuradas para esse tenant.
      </p>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-4 py-2 text-sm border border-border-light rounded hover:bg-gray-50">
          Cancelar
        </button>
        <button onClick={onConfirm} disabled={loading} className="px-4 py-2 text-sm bg-red-600 text-white rounded hover:bg-red-700 disabled:opacity-60 flex items-center gap-2">
          {loading && <Loader2 size={14} className="animate-spin" />}
          Remover
        </button>
      </div>
    </Modal>
  );
}

function TenantCard({ tenant, onRemove }) {
  const [status, setStatus] = useState(null);
  const [testing, setTesting] = useState(false);

  const testConnection = async () => {
    setTesting(true);
    setStatus(null);
    try {
      const resp = await api.tenantStatus(tenant.tenantId);
      setStatus(resp?.data ?? { status: 'Error', message: 'Sem resposta' });
    } catch (err) {
      setStatus({ status: 'Error', message: err.message });
    } finally {
      setTesting(false);
    }
  };

  const added = tenant.addedAt ? new Date(tenant.addedAt).toLocaleDateString('pt-BR') : '—';

  return (
    <div className="bg-white border border-border-light rounded-lg p-5 flex flex-col gap-3">
      <div>
        <div className="text-base font-semibold text-[#222]">{tenant.name || tenant.tenantId}</div>
        {tenant.domain && <div className="text-xs text-text-secondary mt-0.5">{tenant.domain}</div>}
      </div>

      <div className="text-xs text-text-secondary space-y-1">
        <div><span className="font-semibold text-[#333]">Directory Tenant ID</span></div>
        <div className="font-mono text-[11px] text-[#555] break-all">{tenant.tenantId}</div>
      </div>

      <div className="text-xs">
        <span className="font-semibold text-[#333]">Status</span>
        <div className="mt-1">
          {tenant.isConfigured
            ? <StatusBadge status={status?.status ?? 'Connected'} />
            : <StatusBadge status="Unconfigured" />}
          {status?.message && (
            <div className="text-[11px] text-text-secondary mt-1 break-all">{status.message}</div>
          )}
        </div>
      </div>

      <div className="text-xs text-text-secondary">
        <span className="font-semibold text-[#333]">Adicionado em: </span>{added}
      </div>

      <div className="flex items-center gap-3 pt-1 border-t border-border-light">
        <button
          onClick={testConnection}
          disabled={testing}
          className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 disabled:opacity-60"
        >
          {testing ? <Loader2 size={12} className="animate-spin" /> : <Wifi size={12} />}
          {testing ? 'Testando...' : 'Test Connection'}
        </button>
        <button
          onClick={onRemove}
          className="text-xs text-red-600 hover:underline flex items-center gap-1 ml-auto"
        >
          <Trash2 size={12} /> Remove
        </button>
      </div>
    </div>
  );
}

export default function Tenants() {
  const [tenants, setTenants] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showAdd, setShowAdd] = useState(false);
  const [removeTarget, setRemoveTarget] = useState(null);
  const [removing, setRemoving] = useState(false);
  const [error, setError] = useState(null);

  const load = async () => {
    setLoading(true);
    setError(null);
    try {
      const resp = await api.tenants();
      setTenants(resp?.data ?? []);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const handleAdded = (tenant) => {
    setTenants((t) => [...t, tenant]);
    setShowAdd(false);
  };

  const handleRemove = async () => {
    if (!removeTarget) return;
    setRemoving(true);
    try {
      await api.removeTenant(removeTarget.tenantId);
      setTenants((t) => t.filter((x) => x.tenantId !== removeTarget.tenantId));
      setRemoveTarget(null);
    } catch (e) {
      setError(e.message);
    } finally {
      setRemoving(false);
    }
  };

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-5">
        <div>
          <h2 className="text-base font-semibold text-[#222]">Tenants</h2>
          <p className="text-xs text-text-secondary mt-0.5">Microsoft 365 / Entra ID tenants conectados à plataforma</p>
        </div>
        <div className="flex gap-2">
          <button onClick={load} className="flex items-center gap-1.5 px-3 py-1.5 text-xs border border-border-light rounded hover:bg-gray-50">
            <RefreshCw size={13} /> Atualizar
          </button>
          <button
            onClick={() => setShowAdd(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs bg-[#0078A8] text-white rounded hover:bg-[#006090]"
          >
            <Plus size={13} /> Add Tenant
          </button>
        </div>
      </div>

      {error && <div className="text-red-600 text-sm mb-4">{error}</div>}

      {loading ? (
        <div className="flex items-center gap-2 text-text-secondary text-sm">
          <Loader2 size={16} className="animate-spin" /> Carregando...
        </div>
      ) : tenants.length === 0 ? (
        <div className="text-center py-16 text-text-secondary">
          <WifiOff size={32} className="mx-auto mb-3 opacity-30" />
          <p className="text-sm">Nenhum tenant configurado.</p>
          <p className="text-xs mt-1">Clique em <strong>Add Tenant</strong> para conectar um Microsoft 365 tenant.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
          {tenants.map((t) => (
            <TenantCard key={t.id} tenant={t} onRemove={() => setRemoveTarget(t)} />
          ))}
        </div>
      )}

      {showAdd && <AddTenantModal onClose={() => setShowAdd(false)} onAdded={handleAdded} />}
      {removeTarget && (
        <ConfirmModal
          name={removeTarget.name || removeTarget.tenantId}
          onClose={() => setRemoveTarget(null)}
          onConfirm={handleRemove}
          loading={removing}
        />
      )}
    </div>
  );
}
