import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './NoFrameworkWebPartWebPart.module.scss';
import * as strings from 'NoFrameworkWebPartWebPartStrings';
import { buildMarkup } from './noFrameworkUtils';

export interface INoFrameworkWebPartWebPartProps {
  heading: string;
}

/**
 * Web Part built with vanilla TypeScript (no React, no UI framework).
 * Renders straight into this.domElement and cleans up its own listeners.
 *
 * Covers the book's chapter 14 "Web Part without React (No Framework)".
 */
export default class NoFrameworkWebPartWebPart extends BaseClientSideWebPart<INoFrameworkWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;

  protected onInit(): Promise<void> {
    this.properties.heading = this.properties.heading ?? 'No Framework Web Part';
    return Promise.resolve();
  }

  public render(): void {
    const site: string = this.context.pageContext.web.title;
    const user: string = this.context.pageContext.user.displayName;
    const teamsClass: string = !!this.context.sdks.microsoftTeams ? styles.teams : '';
    const darkClass: string = this._isDarkTheme ? styles.dark : '';

    this.domElement.innerHTML = buildMarkup(this.properties.heading, site, user, {
      containerClass: styles.noFrameworkWebPart,
      teamsClass,
      darkClass,
      btnClass: styles.btn
    });

    const btn: HTMLButtonElement | null = this.domElement.querySelector(`.${styles.btn}`);
    btn?.addEventListener('click', this._onReload);
  }

  private _onReload = (): void => {
    this.render();
  };

  protected onDispose(): void {
    // Without React YOU must remove listeners, or you leak them on unmount.
    const btn: HTMLButtonElement | null = this.domElement.querySelector(`.${styles.btn}`);
    btn?.removeEventListener('click', this._onReload);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || undefined);
      this.domElement.style.setProperty('--link', semanticColors.link || undefined);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || undefined);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('heading', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}