/** Truncate `text` to `max` characters, appending an ellipsis if cut. */
export function truncate(text: string, max: number): string {
  if (text.length <= max) {
    return text;
  }
  return `${text.slice(0, Math.max(0, max - 1))}…`;
}

/** Render an ISO date as a human relative string ("today", "3 days ago", …). */
export function formatRelativeDate(iso: string): string {
  const then: Date = new Date(iso);
  const now: Date = new Date();
  const diffMs: number = now.getTime() - then.getTime();
  const days: number = Math.floor(diffMs / 86_400_000);

  if (days <= 0) { return 'today'; }
  if (days === 1) { return 'yesterday'; }
  if (days < 30) { return `${days} days ago`; }

  const months: number = Math.floor(days / 30);
  if (months < 12) { return `${months} month${months === 1 ? '' : 's'} ago`; }

  const years: number = Math.floor(months / 12);
  return `${years} year${years === 1 ? '' : 's'} ago`;
}