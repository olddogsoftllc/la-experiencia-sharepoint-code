import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TeamsEnabledWebPartWebPartStrings';
import TeamsEnabledWebPart from './components/TeamsEnabledWebPart';
import { ITeamsEnabledWebPartProps } from './components/ITeamsEnabledWebPartProps';

export interface ITeamsEnabledWebPartWebPartProps {
  title: string;
}

export default class TeamsEnabledWebPartWebPart extends BaseClientSideWebPart<ITeamsEnabledWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _host: string = 'SharePoint';
  private _theme: string = 'default';

  protected onInit(): Promise<void> {
    this.properties.title = this.properties.title ?? 'Teams-Enabled Web Part';

    if (this.context.sdks.microsoftTeams) {
      // Running in Teams / Office / Outlook — read the host + theme from Teams JS.
      this.context.sdks.microsoftTeams.teamsJs.app.getContext().then((ctx) => {
        this._host = ctx.app.host.name;
        this._theme = ctx.app.theme ?? 'default';
        this.render();
      }).catch(() => {
        this._host = 'Teams';
      });
    }

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<ITeamsEnabledWebPartProps> = React.createElement(
      TeamsEnabledWebPart,
      {
        title: this.properties.title,
        host: this._host,
        theme: this._theme,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
    this._theme = currentTheme.isInverted ? 'dark' : 'default';
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  value: this.properties.title
                })
              ]
            }
          ]
        }
      ]
    };
  }
}