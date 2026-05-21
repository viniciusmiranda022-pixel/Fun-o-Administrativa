import { useEffect, useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { ArrowLeft, CheckCircle2, AlertCircle, RefreshCw, ShieldAlert, Loader2, ExternalLink, ChevronDown, ChevronRight, Save } from 'lucide-react';
import { api } from '../api/client.js';

function ConsentCard({ block, onGrant, granting, sensitive }) {
  const [showDetails, setShowDetails] = useState(false);

  const StatusIcon = block.granted ? CheckCircle2 : AlertCircle;
  const statusColor = block.granted ? 'text-green-600' : 'text-yellow-600';
  const statusLabel = block.granted ? 'Granted' : 'Not Granted';

  return (
    <div className={`bg-white border rounded-lg p-5 ${sensitive ? 'border-orange-300' : 'border-border-light'}`}>
      <div className="flex items-start justify-between mb-3">
        <div>
          <div className="flex items-center gap-2">
            <h3 className="text-base font-semibold text-[#222]">{block.title}</h3>
            {sensitive && (
              <span className="text-[10px] uppercase tracking-wide font-semibold text-orange-700 bg-orange-100 px-2 py-0.5 rounded">
                Sensível
              </span>
            )}
          </div>
          <p className="text-xs text-text-secondary mt-1">{block.description}</p>
        </div>
        <div className={`flex items-center gap-1 text-sm font-semibold ${statusColor}`}>
          <StatusIcon size={16} /> {statusLabel}
        </div>
      </div>

      {sensitive && (
        <div className="bg-orange-50 border border-orange-200 rounded p-2.5 text-xs text-orange-900 mb-3 flex gap-2">
          <ShieldAlert size={14} className="flex-shrink-0 mt-0.5" />
          <span>
            Esta permissão permite alteração no RBAC do Microsoft Entra ID. Conceda apenas quando o restore de funções administrativas for necessário.
          </span>
        </div>
      )}

      <div className="text-xs space-y-2 mb-3">
        <div>
          <span className="font-semibold text-[#333]">Operações habilitadas: </span>
          {block.operations.join(', ')}
        </div>
        {block.grantedAt && (
          <div>
            <span className="font-semibold text-[#333]">Concedido em: </span>
            {new Date(block.grantedAt).toLocaleString('pt-BR')}
          </div>
        )}
      </div>

      <button
        onClick={() => setShowDetails((v) => !v)}
        className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 mb-3"
      >
        {showDetails ? <ChevronDown size={12} /> : <ChevronRight size={12} />} View Details
      </button>

      {showDetails && (
        <div className="bg-gray-50 rounded p-3 mb-3 text-xs space-y-1">
          <div>
            <span className="font-semibold">Permission: </span>
            <code className="font-mono text-[11px] bg-white px-1.5 py-0.5 rounded border border-border-light">{block.permission}</code>
          </div>
          <div>
            <span className="font-semibold">Type: </span>
            Application
          </div>
          <div>
            <span className="font-semibold">Resource: </span>
            Microsoft Graph
          </div>
        </div>
      )}

      <button
        onClick={() => onGrant(block)}
        disabled={granting}
        className={`text-xs px-3 py-1.5 rounded flex items-center gap-1.5 ${
          block.granted
            ? 'border border-border-light text-[#333] hover:bg-gray-50'
            : 'bg-[#0078A8] text-white hover:bg-[#006090]'
        } disabled:opacity-60`}
      >
        {granting ? <Loader2 size={12} className="animate-spin" /> : <ExternalLink size={12} />}
        {block.granted ? 'Re-Grant Consent' : 'Grant Consent'}
      </button>
    </div>
  );
}

function CertificateSection({ tenantId, currentThumbprint }) {
  const [thumbprint, setThumbprint] = useState(currentThumbprint || '');
  const [saving, setSaving] = useState(false);
  const [saved, setSaved] = useState(false);
  const [saveError, setSaveError] = useState(null);

  async function handleSave() {
    setSaving(true);
    setSaved(false);
    setSaveError(null);
    try {
      await api.updateTenantThumbprint(tenantId, thumbprint.trim());
      setSaved(true);
      setTimeout(() => setSaved(false), 3000);
    } catch (e) {
      setSaveError(e.message || 'Erro ao salvar');
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="bg-white border border-[#DEE2E6] rounded-lg p-5 mt-4">
      <h3 className="text-sm font-semibold text-[#222] mb-1">Autenticação por Certificado</h3>
      <p className="text-xs text-[#666] mb-3">
        O script de backup usa autenticação via certificado. Informe o thumbprint do certificado registrado no Azure AD para este aplicativo.
      </p>
      <div className="flex gap-2 items-center">
        <input
          type="text"
          value={thumbprint}
          onChange={(e) => setThumbprint(e.target.value)}
          placeholder="Ex: A1B2C3D4E5F6..."
          className="flex-1 border border-[#DEE2E6] px-3 py-1.5 text-xs font-mono outline-none focus:border-[#0078A8]"
        />
        <button
          onClick={handleSave}
          disabled={saving}
          className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold bg-[#0078A8] text-white hover:bg-[#006090] disabled:opacity-50"
        >
          {saving ? <Loader2 size={12} className="animate-spin" /> : <Save size={12} />}
          Salvar
        </button>
      </div>
      {saved && <p className="text-xs text-[#28A745] mt-1.5">Thumbprint salvo com sucesso.</p>}
      {saveError && <p className="text-xs text-[#DC3545] mt-1.5">{saveError}</p>}
    </div>
  );
}

export default function TenantConsents() {
  const { tenantId } = useParams();
  const navigate = useNavigate();
  const [consents, setConsents] = useState(null);
  const [tenant, setTenant] = useState(null);
  const [loading, setLoading] = useState(true);
  const [granting, setGranting] = useState(false);
  const [error, setError] = useState(null);

  const load = async () => {
    setLoading(true);
    setError(null);
    try {
      const [c, t] = await Promise.all([
        api.tenantConsents(tenantId),
        api.tenants()
      ]);
      setConsents(c.data);
      setTenant((t.data || []).find((x) => x.tenantId === tenantId));
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
  }, [tenantId]);

  const handleGrant = async () => {
    setGranting(true);
    try {
      const r = await api.oauthStart(tenantId);
      const w = window.open(r.data.url, 'oauth-consent', 'width=700,height=800');
      if (!w) window.location.href = r.data.url;
    } catch (e) {
      setError(e.message);
    } finally {
      setGranting(false);
    }
  };

  if (loading) return <div className="p-6"><Loader2 className="animate-spin" /></div>;
  if (error) return <div className="p-6 text-red-600">{error}</div>;
  if (!consents) return null;

  const blocks = [
    {
      key: 'basic',
      title: 'Basic',
      ...consents.basic,
      operations: consents.basic.operations
    },
    {
      key: 'restore',
      title: 'Restore',
      ...consents.restore,
      operations: consents.restore.operations
    }
  ];

  return (
    <div className="p-6 max-w-4xl">
      <button onClick={() => navigate('/tenants')} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1 mb-3">
        <ArrowLeft size={12} /> Voltar para Tenants
      </button>

      <div className="mb-6">
        <h2 className="text-base font-semibold text-[#222]">Edit Consents</h2>
        <p className="text-xs text-text-secondary mt-0.5">
          {tenant?.name || tenantId} — Administrative Function Recovery
        </p>
      </div>

      <div className="bg-blue-50 border border-blue-200 rounded p-3 text-xs text-blue-900 mb-5">
        Permissões são concedidas por categoria. <strong>Basic</strong> habilita backup e comparação (leitura).{' '}
        <strong>Restore</strong> habilita recuperação (leitura + escrita). Ao clicar em <em>Grant</em>, você será redirecionado
        para o login Microsoft e precisará autenticar como Global Administrator.
      </div>

      <div className="flex justify-end mb-3">
        <button onClick={load} className="text-xs text-[#0078A8] hover:underline flex items-center gap-1">
          <RefreshCw size={12} /> Atualizar status
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {blocks.map((b) => (
          <ConsentCard
            key={b.key}
            block={b}
            onGrant={handleGrant}
            granting={granting}
            sensitive={consents.restore.sensitive && b.key === 'restore'}
          />
        ))}
      </div>

      <CertificateSection tenantId={tenantId} currentThumbprint={tenant?.certificateThumbprint} />
    </div>
  );
}
