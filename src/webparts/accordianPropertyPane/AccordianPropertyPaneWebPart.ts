import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AccordianPropertyPaneWebPartStrings';
import AccordianPropertyPane from './components/AccordianPropertyPane';
import { IAccordianPropertyPaneProps } from './components/IAccordianPropertyPaneProps';

export interface IAccordianPropertyPaneWebPartProps {
  description: string;
}

export default class AccordianPropertyPaneWebPart extends BaseClientSideWebPart<IAccordianPropertyPaneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IAccordianPropertyPaneProps> = React.createElement(
      AccordianPropertyPane,
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
            description: "An accordion property pane demo"
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "Group 1",
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('demoProperty1', {
                  label: "Property 1"
                }),
                PropertyPaneTextField('demoProperty2', {
                  label: "Property 2"
                })
              ]
            },
            {
              groupName: "Group 2",
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('demoProperty3', {
                  label: "Property 3"
                }),
                PropertyPaneTextField('demoProperty4', {
                  label: "Property 4"
                })
              ]
            },
            {
              groupName: "Group 3",
              isCollapsed: true,
              groupFields: [
                PropertyPaneTextField('demoProperty5', {
                  label: "Property 5"
                }),
                PropertyPaneTextField('demoProperty6', {
                  label: "Property 6"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
