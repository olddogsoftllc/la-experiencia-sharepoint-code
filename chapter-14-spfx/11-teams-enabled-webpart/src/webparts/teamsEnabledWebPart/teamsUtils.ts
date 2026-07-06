/**
 * Pure helpers for the Teams-Enabled Web Part, extracted so they can be
 * unit-tested without the SharePoint / Teams runtime (no sp-* imports here).
 */

export type ThemeName = 'dark' | 'default';

/** Maps a SharePoint theme inversion flag to a theme name. */
export function resolveThemeName(isInverted: boolean): ThemeName {
  return isInverted ? 'dark' : 'default';
}

/** Suffix that tells the user where the web part is running. */
export function hostRunningSuffix(hasTeamsContext: boolean): string {
  return hasTeamsContext ? ' (running inside Microsoft Teams)' : ' (running in SharePoint)';
}