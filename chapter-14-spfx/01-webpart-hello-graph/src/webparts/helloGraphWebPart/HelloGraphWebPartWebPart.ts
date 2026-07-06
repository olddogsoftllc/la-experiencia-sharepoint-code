import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'HelloGraphWebPartWebPartStrings';
import HelloGraphWebPart from './components/HelloGraphWebPart';
import { IHelloGraphWebPartProps } from './components/IHelloGraphWebPartProps';
import { GraphService } from './services/GraphService';

export interface IHelloGraphWebPartWebPartProps {
  query: string;
  maxResults: number;
}

export default class HelloGraphWebPartWebPart extends BaseClientSideWebPart<IHelloGraphWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _graphService: GraphService | undefined;

  protected async onInit(): Promise<void> {
    const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
    this._graphService = new GraphService(client);

    // Sensible defaults so the Web Part shows something on first add.
    this.properties.query = this.properties.query ?? 'book';
    this.properties.maxResults = this.properties.maxResults ?? 10;
  }

  public render(): void {
    const element: React.ReactElement<IHelloGraphWebPartProps> = React.createElement(
      HelloGraphWebPart,
      {
        graphService: this._graphService as GraphService,
        query: this.properties.query,
        maxResults: this.properties.maxResults,
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
                PropertyPaneTextField('query', {
                  label: strings.QueryFieldLabel,
                  value: this.properties.query
                }),
                PropertyPaneSlider('maxResults', {
                  label: strings.MaxResultsFieldLabel,
                  min: 1,
                  max: 50,
                  value: this.properties.maxResults
                })
              ]
            }
          ]
        }
      ]
    };
  }
}