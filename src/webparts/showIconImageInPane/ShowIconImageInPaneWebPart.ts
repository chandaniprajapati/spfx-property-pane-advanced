import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ShowIconImageInPaneWebPartStrings';
import ShowIconImageInPane from './components/ShowIconImageInPane';
import { IShowIconImageInPaneProps } from './components/IShowIconImageInPaneProps';

export interface IShowIconImageInPaneWebPartProps {
  description: string;
}

export default class ShowIconImageInPaneWebPart extends BaseClientSideWebPart<IShowIconImageInPaneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IShowIconImageInPaneProps> = React.createElement(
      ShowIconImageInPane,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const imgBarChart: string = require('./assets/bar-chart.svg');
    const imgPieChart: string = require('./assets/pie-chart.svg');
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneChoiceGroup('imgChartType', {
                  label: "Images",
                  options: [
                    {
                      key: 'Column',
                      text: 'Column chart',
                      imageSrc: imgBarChart,
                    },
                    {
                      key: 'Pie',
                      text: 'Pie chart',
                      imageSrc: imgPieChart,
                    }
                  ]
                }),
                PropertyPaneChoiceGroup('iconChartType', {
                  label: "Icons",
                  options: [
                    {
                      key: 'Column',
                      text: 'Column chart',
                      iconProps: {
                        officeFabricIconFontName: 'BarChart4'
                      }
                    },
                    {
                      key: 'Pie',
                      text: 'Pie chart',
                      iconProps: {
                        officeFabricIconFontName: 'PieDouble'
                      }
                    }
                  ]
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
