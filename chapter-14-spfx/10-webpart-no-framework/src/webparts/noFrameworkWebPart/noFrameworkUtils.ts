/**
 * Pure helper for the No-Framework Web Part, extracted so the markup layout
 * can be unit-tested without the SharePoint runtime (no sp-* imports here).
 */

export interface IMarkupOptions {
  containerClass: string;
  teamsClass: string;
  darkClass: string;
  btnClass: string;
}

/** Builds the web part's HTML markup (vanilla TS — no React). */
export function buildMarkup(heading: string, site: string, user: string, opts: IMarkupOptions): string {
  return `
      <section class="${opts.containerClass} ${opts.teamsClass} ${opts.darkClass}">
        <h2>${heading}</h2>
        <p>Site: <strong>${site}</strong></p>
        <p>Hello, <strong>${user}</strong></p>
        <button class="${opts.btnClass}">Reload</button>
      </section>`;
}