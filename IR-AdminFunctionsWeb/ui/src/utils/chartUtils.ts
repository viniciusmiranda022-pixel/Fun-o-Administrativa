import type { BackupSnapshot } from '../types/api';

interface Bar {
  label: string;
  h: number;
}

/**
 * Builds a bar-chart data array from backup timestamps.
 */
export function buildBars(backups: BackupSnapshot[], days = 10): Bar[] {
  if (!backups || backups.length === 0) {
    return Array.from({ length: days }, () => ({ label: '', h: 0 }));
  }

  const now = new Date();
  const buckets = Array.from({ length: days }, (_, i) => {
    const d = new Date(now);
    d.setDate(d.getDate() - (days - 1 - i));
    return { label: `${d.getMonth() + 1}/${d.getDate()}`, count: 0, date: d.toDateString() };
  });

  backups.forEach((b) => {
    const bDate = new Date(b.createdAt || b.collectedAt || '').toDateString();
    const bucket = buckets.find((bk) => bk.date === bDate);
    if (bucket) bucket.count++;
  });

  const max = Math.max(...buckets.map((b) => b.count), 1);
  return buckets.map((b) => ({ label: b.label, h: Math.round((b.count / max) * 100) || 2 }));
}
