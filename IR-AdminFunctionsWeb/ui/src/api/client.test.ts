import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock axios before importing client
vi.mock('axios', () => ({
  default: {
    create: () => ({
      get: vi.fn(),
      post: vi.fn(),
      delete: vi.fn(),
      put: vi.fn(),
      interceptors: {
        response: { use: vi.fn() },
      },
    }),
  },
}));

describe('pollJob', () => {
  it('returns job when status is Completed', async () => {
    // Since axios is mocked, we test pollJob logic in isolation
    // by creating a minimal mock api.job function
    const completedJob = { id: 'abc', status: 'Completed', kind: 'backup' };

    // Use dynamic import to get the pollJob after mocks are set up
    const { pollJob, api } = await import('./client');

    // Spy on api.job
    const spy = vi.spyOn(api, 'job').mockResolvedValue({ data: completedJob, success: true });

    const result = await pollJob('abc', { intervalMs: 10, timeoutMs: 1000 });
    expect(result.status).toBe('Completed');
    spy.mockRestore();
  });

  it('throws when timeout is exceeded', async () => {
    const runningJob = { id: 'xyz', status: 'Running', kind: 'backup' };
    const { pollJob, api } = await import('./client');

    const spy = vi.spyOn(api, 'job').mockResolvedValue({ data: runningJob, success: true });

    await expect(
      pollJob('xyz', { intervalMs: 10, timeoutMs: 50 })
    ).rejects.toThrow(/timeout/i);

    spy.mockRestore();
  });
});
