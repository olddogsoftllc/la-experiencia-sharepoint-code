import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { AadHttpClient } from '@microsoft/sp-http';

import * as strings from 'OwnApiWebPartWebPartStrings';
import OwnApiWebPart from './components/OwnApiWebPart';
import { IOwnApiWebPartProps } from './components/IOwnApiWebPartProps';
import { OwnApiService } from './services/OwnApiService';

export interface IOwnApiWebPartWebPartProps {
  apiId: string;   // App ID URI of the target API (e.g. api://contoso-api)
  apiUrl: string;  // Full endpoint URL to call
}

export default class OwnApiWebPartWebPart extends BaseClientSideWebPart<IOwnApiWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _apiService: OwnApiService | undefined;

  protected async onInit(): Promise<void> {
    this.properties.apiId = this.properties.apiId ?? 'api://your-api-app-id-uri';
    this.properties.apiUrl = this.properties.apiUrl ?? 'https://your-api.azurewebsites.net/api/data';
    await this._ensureService();
  }

  private async _ensureService(): Promise<void> {
    if (!this.properties.apiId) {
      return;
    }
    const client: AadHttpClient = await this.context.aadHttpClientFactory
      .getClient(this.properties.apiId);
    this._apiService = new OwnApiService(client, this.properties.apiUrl);
  }

  public render(): void {
    const element: React.ReactElement<IOwnApiWebPartProps> = React.createElement(
      OwnApiWebPart,
      {
        apiService: this._apiService,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'apiId' || propertyPath === 'apiUrl') {
      // Re-create the service with the new endpoint, then re-render.
      this._ensureService().then(() => this.render()).catch(() => { /* surfaced in UI */ });
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    this._isDarkTheme = !!currentTheme.isInverted;
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
                PropertyPaneTextField('apiId', {
                  label: strings.ApiIdFieldLabel,
                  value: this.properties.apiId
                }),
                PropertyPaneTextField('apiUrl', {
                  label: strings.ApiUrlFieldLabel,
                  value: this.properties.apiUrl
                })
              ]
            }
          ]
        }
      ]
    };
  }
}