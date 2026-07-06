/**
 * Pure helpers for the Priority Field Customizer, extracted so they can be
 * unit-tested without the SharePoint runtime (no sp-* imports here).
 */

export type PriorityClass = 'high' | 'low' | 'normal';

/**
 * Maps a raw cell value to a priority class by case-insensitive substring match:
 * "high" -> high, "low" -> low, anything else -> normal.
 */
export function getPriorityClass(value: string | undefined): PriorityClass {
  const lower: string = (value ?? '').toLowerCase();
  if (lower.indexOf('high') >= 0) return 'high';
  if (lower.indexOf('low') >= 0) return 'low';
  return 'normal';
}

/**
 * Builds the badge text shown in the cell: optional prefix + the raw value.
 */
export function buildBadgeText(value: string | undefined, prefix?: string): string {
  const p: string = prefix ? `${prefix}: ` : '';
  return `${p}${value ?? ''}`;
}