import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Plus, RefreshCw, Trash2, WifiOff, AlertCircle, CheckCircle2, Loader2, Settings } from 'lucide-react';
import { api } from '../api/client.js';

function StatusPill({ status }) {
  const map = {
    Connected:    { icon: CheckCircle2, cls: 'text-green-600 bg-green-50',   label: 'Connected' },
    Unconfigured: { icon: AlertCircle,  cls: 'text-yellow-600 bg-yellow-50', label: 'Unconfigured' },
    Error:        { icon: WifiOff,      cls: 'text-red-600 bg-red-50',       label: 'Error' }
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
      <div className="bg-white rounded-lg shadow-xl w-full max-w-md mx-4">
        <div className="flex items-center justify-between px-6 py-4 border-b border-border-light">
          <h2 className="text-base font-semibold text-[#222]">{title}</h2>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700 text-xl font-bold leading-none">×</button>
        </div>
        <div className="px-6 py-4">{children}</div>
      </div>
    </div>
  );
}

function AddTenantPrompt({ onClose, onConfirm, loading }) {
  return (
    <Modal title="Add Tenant" onClose={onClose}>
      <div className="text-sm text-[#333] space-y-3 mb-4">
        <p>Você será redirecionado para a página de login do Microsoft Entra ID em uma nova janela.</p>
        <p>É necessário autenticar como <strong>Global Administrator</strong> do tenant que será adicionado.</p>
        <p>Após o login, clique em <strong>Aceitar</strong> para conceder o consentimento administrativo inicial (permissão Basic).</p>
      </div>
      <div className="flex justify-end gap-2">
        <button onClick={onClose} className="px-4 py-2 text-sm border border-border-light rounded hover:bg-gray-50">
          Cancelar
        </button>
        <button onClick={onConfirm} disabled={loading} className="px-4 py-2 text-sm bg-[#0078A8] text-white rounded hover:bg-[#006090] disabled:opacity-60 flex items-center gap-2">
          {loading && <Loader2 size={14} className="animate-spin" />}
          OK
        </button>
      </div>
    </Modal>
  );
}

function ConfirmRemove({ name, onClose, onConfirm, loading }) {
  return (
    <Modal title="Remover Tenant" onClose={onClose}>
      <p className="text-sm text-[#333] mb-4">
        Tem certeza que deseja remover o tenant <strong>{name}</strong>?
        Isso pode impactar operações de backup, comparação e restore configuradas.
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

function TenantCard({ tenant, onRemove, onEditConsents }) {
  const added = tenant.addedAt ? new Date(tenant.addedAt).toLocaleDateString('pt-BR') : '—';
  const basic = tenant.consents?.basic?.granted;
  const restore = tenant.consents?.restore?.granted;
  const granted = (basic ? 1 : 0) + (restore ? 1 : 0);
  const notGranted = 2 - granted;

  return (
    <div className="bg-white border border-border-light rounded-lg p-5 flex flex-col gap-3">
      <div>
        <div className="text-base font-semibold text-[#222]">{tenant.name || tenant.tenantId}</div>
        {tenant.domain && <div className="text-xs text-text-secondary mt-0.5">{tenant.domain}</div>}
      </div>

      <div className="text-xs">
        <div className="font-semibold text-[#333] mb-1">Consent Status</div>
        <div className="space-y-1">
          <div className="flex items-center gap-1 text-green-700">
            <CheckCircle2 size={12} /> {granted} Granted
          </div>
          <div className="flex items-center gap-1 text-yellow-700">
            <AlertCircle size={12} /> {notGranted} Not Granted
          </div>
        </div>
      </div>

      <div className="border-t border-border-light pt-2">
        <div className="text-xs font-semibold text-[#333] mb-1">Directory Tenant ID</div>
        <div className="font-mono text-[11px] text-[#555] break-all">{tenant.tenantId}</div>
      </div>

      <div className="text-xs text-text-secondary">
        <span className="font-semibold text-[#333]">Last Updated: </span>{added}
      </div>

      <div className="flex items-center gap-3 pt-2 border-t border-border-light">
        <button
          onClick={() => onEditConsents(tenant)}
          className="text-xs text-[#0078A8] hover:underline flex items-center gap-1"
        >
          <Settings size={12} /> Edit Consents
        </button>
        <button
          onClick={() => onRemove(tenant)}
          className="text-xs text-red-600 hover:underline flex items-center gap-1 ml-auto"
        >
          <Trash2 size={12} /> Remove
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
  const [showAdd, setShowAdd] = useState(false);
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
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    load();
    const onMsg = (ev) => {
      if (ev?.data?.type === 'oauth-complete') load();
    };
    window.addEventListener('message', onMsg);
    return () => window.removeEventListener('message', onMsg);
  }, []);

  const startAdd = async () => {
    setBusy(true);
    try {
      const r = await api.oauthStart();
      const w = window.open(r.data.url, 'oauth-consent', 'width=700,height=800');
      if (!w) window.location.href = r.data.url;
      setShowAdd(false);
    } catch (e) {
      setError(e.message);
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
        <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-5">
          <div className="flex items-start gap-3">
            <AlertCircle className="text-yellow-700 flex-shrink-0 mt-0.5" size={20} />
            <div>
              <h3 className="font-semibold text-yellow-900">Setup necessário</h3>
              <p className="text-sm text-yellow-800 mt-1 mb-3">
                Antes de adicionar tenants, é necessário registrar esta aplicação no Microsoft Entra ID
                (uma única vez, com uma conta Global Administrator).
              </p>
              <button onClick={() => navigate('/setup')} className="px-3 py-1.5 bg-yellow-700 text-white rounded text-sm hover:bg-yellow-800">
                Iniciar Setup
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

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
            <TenantCard
              key={t.id}
              tenant={t}
              onRemove={setRemoveTarget}
              onEditConsents={(tt) => navigate(`/tenants/${tt.tenantId}/consents`)}
            />
          ))}
        </div>
      )}

      {showAdd && <AddTenantPrompt onClose={() => setShowAdd(false)} onConfirm={startAdd} loading={busy} />}
      {removeTarget && (
        <ConfirmRemove
          name={removeTarget.name || removeTarget.tenantId}
          onClose={() => setRemoveTarget(null)}
          onConfirm={handleRemove}
          loading={busy}
        />
      )}
    </div>
  );
}
