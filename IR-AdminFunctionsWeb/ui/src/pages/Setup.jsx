import { useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, CheckCircle2, AlertCircle, ExternalLink, Copy } from 'lucide-react';
import { api } from '../api/client.js';

export default function Setup() {
  const navigate = useNavigate();
  const [state, setState] = useState(null);
  const [session, setSession] = useState(null);
  const [error, setError] = useState(null);
  const [starting, setStarting] = useState(false);

  useEffect(() => {
    api.setupState().then((r) => setState(r.data)).catch((e) => setError(e.message));
  }, []);

  useEffect(() => {
    if (!session?.sessionId) return;
    const timer = setInterval(async () => {
      try {
        const r = await api.setupStatus(session.sessionId);
        const s = r.data;
        setSession((prev) => ({ ...prev, ...s }));
        if (s.status === 'Completed') {
          clearInterval(timer);
          setTimeout(() => navigate('/tenants'), 2000);
        }
        if (s.status === 'Failed') clearInterval(timer);
      } catch (e) { /* keep polling */ }
    }, 2000);
    return () => clearInterval(timer);
  }, [session?.sessionId, navigate]);

  const startSetup = async () => {
    setStarting(true);
    setError(null);
    try {
      const r = await api.setupStart();
      setSession({ sessionId: r.data.sessionId, ...r.data });
    } catch (e) {
      setError(e.message);
    } finally {
      setStarting(false);
    }
  };

  const copyCode = () => {
    if (session?.userCode) navigator.clipboard.writeText(session.userCode);
  };

  if (!state) return <div className="p-8"><Loader2 className="animate-spin" /></div>;

  if (state.isConfigured && !session) {
    return (
      <div className="p-8 max-w-2xl mx-auto">
        <div className="bg-white border border-border-light rounded-lg p-6">
          <div className="flex items-center gap-3 mb-4">
            <CheckCircle2 className="text-green-600" size={24} />
            <h2 className="text-base font-semibold">Aplicação já configurada</h2>
          </div>
          <dl className="text-sm space-y-2 mb-4">
            <div className="flex gap-2"><dt className="font-semibold text-[#333] w-40">Display Name:</dt><dd>{state.displayName}</dd></div>
            <div className="flex gap-2"><dt className="font-semibold text-[#333] w-40">Client ID:</dt><dd className="font-mono text-xs">{state.clientId}</dd></div>
            <div className="flex gap-2"><dt className="font-semibold text-[#333] w-40">Tenant Bootstrap:</dt><dd className="font-mono text-xs">{state.bootstrapTenantId}</dd></div>
            <div className="flex gap-2"><dt className="font-semibold text-[#333] w-40">Criado em:</dt><dd>{state.createdAt ? new Date(state.createdAt).toLocaleString('pt-BR') : '—'}</dd></div>
          </dl>
          <button onClick={() => navigate('/tenants')} className="px-4 py-2 bg-[#0078A8] text-white rounded text-sm">
            Ir para Tenants
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="p-8 max-w-2xl mx-auto">
      <div className="bg-white border border-border-light rounded-lg p-6">
        <h2 className="text-base font-semibold text-[#222] mb-1">Initial Setup</h2>
        <p className="text-sm text-text-secondary mb-6">
          Antes de adicionar tenants, é necessário registrar esta aplicação como App Multi-Tenant no Microsoft Entra ID.
          O processo é feito uma única vez por um Global Administrator.
        </p>

        {!session && (
          <>
            <div className="bg-blue-50 border border-blue-200 rounded p-4 mb-4 text-sm">
              <strong className="text-blue-900">O que vai acontecer:</strong>
              <ol className="list-decimal ml-5 mt-2 space-y-1 text-blue-800">
                <li>Você verá um código curto (formato XXX-XXX-XXX).</li>
                <li>Abre <code>https://microsoft.com/devicelogin</code> em qualquer navegador e digita o código.</li>
                <li>Autentica com uma conta <strong>Global Administrator</strong> do seu tenant.</li>
                <li>Aceita as permissões (Application.ReadWrite.All para criar o app).</li>
                <li>Esta tela atualiza sozinha quando o app é registrado.</li>
              </ol>
            </div>

            {error && <div className="text-red-600 text-sm mb-4">{error}</div>}

            <button
              onClick={startSetup}
              disabled={starting}
              className="px-4 py-2 bg-[#0078A8] text-white rounded text-sm hover:bg-[#006090] disabled:opacity-60 flex items-center gap-2"
            >
              {starting && <Loader2 size={14} className="animate-spin" />}
              Iniciar Setup
            </button>
          </>
        )}

        {session && session.status !== 'Completed' && session.status !== 'Failed' && (
          <div>
            <div className="text-sm font-semibold mb-2">{session.message || 'Aguardando...'}</div>

            {session.userCode ? (
              <div className="bg-gray-50 border-2 border-dashed border-gray-300 rounded p-6 text-center mb-4">
                <div className="text-xs text-text-secondary mb-2">Seu código de autenticação:</div>
                <div className="text-3xl font-mono font-bold tracking-widest text-[#0078A8] mb-3">{session.userCode}</div>
                <button onClick={copyCode} className="text-xs text-[#0078A8] hover:underline inline-flex items-center gap-1 mr-3">
                  <Copy size={12} /> Copiar código
                </button>
                <a
                  href={session.verificationUrl}
                  target="_blank"
                  rel="noreferrer"
                  className="text-xs text-[#0078A8] hover:underline inline-flex items-center gap-1"
                >
                  <ExternalLink size={12} /> Abrir microsoft.com/devicelogin
                </a>
              </div>
            ) : (
              <div className="flex items-center gap-2 text-sm text-text-secondary mb-4">
                <Loader2 className="animate-spin" size={14} /> Solicitando código à Microsoft...
              </div>
            )}

            {session.status === 'Registering' && (
              <div className="flex items-center gap-2 text-sm text-green-700">
                <Loader2 className="animate-spin" size={14} /> Usuário autenticado. Registrando aplicação...
              </div>
            )}
          </div>
        )}

        {session?.status === 'Completed' && (
          <div className="bg-green-50 border border-green-200 rounded p-4">
            <div className="flex items-center gap-2 text-green-800 font-semibold mb-1">
              <CheckCircle2 size={18} /> Setup concluído!
            </div>
            <p className="text-sm text-green-700">Client ID: <code className="font-mono">{session.clientId}</code></p>
            <p className="text-xs text-green-600 mt-2">Redirecionando para a página de Tenants...</p>
          </div>
        )}

        {session?.status === 'Failed' && (
          <div className="bg-red-50 border border-red-200 rounded p-4">
            <div className="flex items-center gap-2 text-red-800 font-semibold mb-1">
              <AlertCircle size={18} /> Setup falhou
            </div>
            <p className="text-sm text-red-700">{session.message}</p>
            <button onClick={() => { setSession(null); setError(null); }} className="mt-3 text-xs text-[#0078A8] hover:underline">
              Tentar novamente
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
