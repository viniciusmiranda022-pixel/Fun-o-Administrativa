/**
 * Builds a bar-chart data array from backup timestamps.
 * @param {Array} backups - Array of backup objects with createdAt/collectedAt
 * @param {number} days - Number of days to cover (default 10)
 * @returns {Array<{label: string, h: number}>}
 */
export function buildBars(backups, days = 10) {
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
