import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'StepsPropertyPaneWebPartStrings';
import StepsPropertyPane from './components/StepsPropertyPane';
import { IStepsPropertyPaneProps } from './components/IStepsPropertyPaneProps';

export interface IStepsPropertyPaneWebPartProps {
  description: string;
}

export default class StepsPropertyPaneWebPart extends BaseClientSideWebPart<IStepsPropertyPaneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IStepsPropertyPaneProps> = React.createElement(
      StepsPropertyPane,
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
    return {
      pages: [
        {
          header: {
            description: "First page."
          },
          displayGroupsAsAccordion: false,
          groups: [
            {
              groupName: "Group 1",
              groupFields: [
                PropertyPaneTextField('property', {
                  label: "Property 1"
                }),
                PropertyPaneTextField('property2', {
                  label: "Property 2"
                })
              ]
            },
            {
              groupName: "Group 2",
              groupFields: [
                PropertyPaneTextField('property3', {
                  label: "Property 3"
                }),
                PropertyPaneTextField('property4', {
                  label: "Property 4"
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Second page."
          },
          displayGroupsAsAccordion: false,
          groups: [
            {
              groupName: "Group 3",
              groupFields: [
                PropertyPaneTextField('property5', {
                  label: "Property 5"
                }),
                PropertyPaneTextField('property6', {
                  label: "Property 6"
                })
              ]
            },
            {
              groupName: "Group 4",
              groupFields: [
                PropertyPaneTextField('property7', {
                  label: "Property 7"
                }),
                PropertyPaneTextField('property8', {
                  label: "Property 8"
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Third page."
          },
          displayGroupsAsAccordion: false,
          groups: [
            {
              groupName: "Group 5",
              groupFields: [
                PropertyPaneTextField('property9', {
                  label: "Property 9"
                }),
                PropertyPaneTextField('property10', {
                  label: "Property 10"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
