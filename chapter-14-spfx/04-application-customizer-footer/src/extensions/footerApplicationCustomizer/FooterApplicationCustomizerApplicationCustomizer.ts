import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  type PlaceholderContent,
  type PlaceholderName
} from '@microsoft/sp-application-base';

import * as strings from 'FooterApplicationCustomizerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'FooterApplicationCustomizerApplicationCustomizer';

export interface IFooterApplicationCustomizerApplicationCustomizerProperties {
  /** Text shown in the global footer bar. */
  footerText: string;
}

/**
 * Application Customizer that injects a footer bar into the Bottom placeholder
 * of every modern page where the extension is active.
 *
 * Covers the book's cap14 "Tipos de Extensiones" (Application Customizer).
 */
export default class FooterApplicationCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IFooterApplicationCustomizerApplicationCustomizerProperties> {

  private _bottomPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Placeholders may not be ready immediately; re-render when they change.
    this.context.placeholderProvider.changedEvent.add(this, this._onPlaceholdersChanged);
    this._renderFooter();

    return Promise.resolve();
  }

  private _onPlaceholdersChanged = (): void => {
    this._renderFooter();
  };

  private _renderFooter(): void {
    if (!this.context.placeholderProvider.placeholderNames.includes(PlaceholderName.Bottom)) {
      return;
    }

    this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
    if (!this._bottomPlaceholder) {
      return;
    }

    const footerText: string = this.properties.footerText ?? 'Powered by SPFx';

    const bar: HTMLDivElement = document.createElement('div');
    bar.className = 'footer-bar';
    bar.style.backgroundColor = '#0078d4';
    bar.style.color = '#ffffff';
    bar.style.padding = '6px 16px';
    bar.style.fontSize = '12px';
    bar.textContent = footerText;

    this._bottomPlaceholder.domElement.innerHTML = '';
    this._bottomPlaceholder.domElement.appendChild(bar);
  }
}