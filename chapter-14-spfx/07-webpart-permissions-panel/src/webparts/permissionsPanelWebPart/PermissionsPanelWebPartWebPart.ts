import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'PermissionsPanelWebPartWebPartStrings';
import PermissionsPanelWebPart from './components/PermissionsPanelWebPart';
import { IPermissionsPanelWebPartProps } from './components/IPermissionsPanelWebPartProps';
import { GraphService } from './services/GraphService';
import { TelemetryClient } from './services/TelemetryClient';

export interface IPermissionsPanelWebPartWebPartProps {
  siteId: string;
}

export default class PermissionsPanelWebPartWebPart extends BaseClientSideWebPart<IPermissionsPanelWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _graphService: GraphService | undefined;

  protected async onInit(): Promise<void> {
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
    // Instrumentation key intentionally left empty here; wire it from a config
    // list / Property Pane in production (see chapter 14 "Error handling and telemetry").
    const telemetry: TelemetryClient = new TelemetryClient();
    this._graphService = new GraphService(client, telemetry);

    this.properties.siteId = this.properties.siteId ?? this.context.pageContext.site.id.toString();
  }

  public render(): void {
    const element: React.ReactElement<IPermissionsPanelWebPartProps> = React.createElement(
      PermissionsPanelWebPart,
      {
        graphService: this._graphService as GraphService,
        siteId: this.properties.siteId,
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
                PropertyPaneTextField('siteId', {
                  label: strings.SiteIdFieldLabel,
                  value: this.properties.siteId,
                  description: 'Graph site id (defaults to the current site).'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}