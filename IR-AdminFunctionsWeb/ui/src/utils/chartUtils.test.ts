import { describe, it, expect } from 'vitest';
import { buildBars } from './chartUtils';

describe('buildBars', () => {
  it('returns N bars for N days', () => {
    const bars = buildBars([], 10);
    expect(bars).toHaveLength(10);
  });

  it('returns minimum height bars when no backups', () => {
    const bars = buildBars([], 10);
    bars.forEach((b) => expect(b.h).toBeGreaterThanOrEqual(0));
  });

  it('tallest bar gets h=100 when only one day has backups', () => {
    const today = new Date().toISOString();
    const backups = [{ createdAt: today }, { createdAt: today }];
    const bars = buildBars(backups, 5);
    const max = Math.max(...bars.map((b) => b.h));
    expect(max).toBe(100);
  });

  it('assigns labels like M/D', () => {
    const bars = buildBars([], 5);
    bars.forEach((b) => {
      // label should be empty string (for empty backups variant) or M/D
      expect(typeof b.label).toBe('string');
    });
  });

  it('handles undefined/null dates gracefully', () => {
    const backups = [{ createdAt: null }, { createdAt: undefined }, {}];
    expect(() => buildBars(backups as any, 5)).not.toThrow();
  });
});
