/**
 * Pure helpers for the Actions Command Set, extracted so they can be
 * unit-tested without the SharePoint runtime (no sp-* imports here).
 */

/** The "Export selection" command is visible only when at least one row is selected. */
export function shouldShowExport(selectedCount: number): boolean {
  return selectedCount > 0;
}

/** Builds the message shown by the "Notify count" command. */
export function buildNotifyMessage(prefix: string | undefined, selectedCount: number): string {
  const p: string = prefix ?? 'Selected';
  return `${p}: ${selectedCount} row(s).`;
}

/** Builds the message shown by the "Export selection" command (demo). */
export function buildExportMessage(selectedCount: number): string {
  return `Export: ${selectedCount} row(s) selected (demo).`;
}