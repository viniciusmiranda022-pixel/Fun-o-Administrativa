import axios from 'axios';

const client = axios.create({
  baseURL: '/api',
  timeout: 30000,
  headers: { 'Content-Type': 'application/json' }
});

client.interceptors.response.use(
  (resp) => resp,
  (err) => {
    const message = err.response?.data?.error || err.message || 'Erro desconhecido';
    return Promise.reject(new Error(message));
  }
);

export const api = {
  health: () => client.get('/health').then((r) => r.data),
  settings: () => client.get('/settings').then((r) => r.data),
  tenants: () => client.get('/tenants').then((r) => r.data),
  backups: () => client.get('/backups').then((r) => r.data),
  backup: (id) => client.get(`/backups/${id}`).then((r) => r.data),
  runBackup: () => client.post('/backups/run').then((r) => r.data),
  unpacked: (backupId) => client.get('/unpacked-objects', { params: { backupId } }).then((r) => r.data),
  runCompare: (backupId) => client.post('/compare/run', { backupId }).then((r) => r.data),
  compareResults: () => client.get('/compare/results').then((r) => r.data),
  previewRestore: (payload) => client.post('/restore/preview', payload).then((r) => r.data),
  applyRestore: (payload) => client.post('/restore/apply', payload).then((r) => r.data),
  tasks: () => client.get('/tasks').then((r) => r.data),
  events: (severity) => client.get('/events', { params: severity ? { severity } : {} }).then((r) => r.data),
  job: (id) => client.get(`/jobs/${id}`).then((r) => r.data),
  jobs: () => client.get('/jobs').then((r) => r.data)
};

export async function pollJob(id, { intervalMs = 2000, timeoutMs = 600000, onTick } = {}) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const resp = await api.job(id);
    const job = resp?.data;
    if (onTick) onTick(job);
    if (job?.status === 'Completed' || job?.status === 'Failed') {
      return job;
    }
    await new Promise((r) => setTimeout(r, intervalMs));
  }
  throw new Error(`Job ${id} excedeu timeout de polling`);
}
