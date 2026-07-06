/**
 * Pure helpers for the Footer Application Customizer, extracted so they can be
 * unit-tested without the SharePoint runtime (no sp-* imports here).
 */

export const DEFAULT_FOOTER_TEXT: string = 'Powered by SPFx';

export interface IFooterBarStyle {
  backgroundColor: string;
  color: string;
  padding: string;
  fontSize: string;
}

/** Returns the configured footer text, falling back to the default when empty. */
export function resolveFooterText(text: string | undefined): string {
  return text && text.length > 0 ? text : DEFAULT_FOOTER_TEXT;
}

/** Returns the inline style for the footer bar (brand blue, white text). */
export function footerBarStyle(): IFooterBarStyle {
  return {
    backgroundColor: '#0078d4',
    color: '#ffffff',
    padding: '6px 16px',
    fontSize: '12px'
  };
}