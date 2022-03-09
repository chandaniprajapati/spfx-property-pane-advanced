import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import * as strings from 'DynamicPropertyPaneWebPartStrings';
import DynamicPropertyPane from './components/DynamicPropertyPane';
import { IDynamicPropertyPaneProps } from './components/IDynamicPropertyPaneProps';

export interface IDynamicPropertyPaneWebPartProps {
  description: string;
  textOrImageType: string;
  simpleText: string;
  filePickerResult: IFilePickerResult;
}

export default class DynamicPropertyPaneWebPart extends BaseClientSideWebPart<IDynamicPropertyPaneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IDynamicPropertyPaneProps> = React.createElement(
      DynamicPropertyPane,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        textOrImageType: this.properties.textOrImageType,
        simpleText: this.properties.simpleText,
        filePickerResult: this.properties.filePickerResult,
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

    let textControl: any ;
    let imageSourceControl: any;

    if (this.properties.textOrImageType === "Text") {
      textControl = PropertyPaneTextField('simpleText', {
        label: "Text",
        placeholder: "Enter Text"
      });
    }
    else {
      imageSourceControl = PropertyFieldFilePicker('filePicker', {
        context: this.context as any,
        filePickerResult: this.properties.filePickerResult,
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        properties: this.properties,
        onSave: (e: IFilePickerResult) => { console.log(e); this.properties.filePickerResult = e; },
        onChanged: (e: IFilePickerResult) => { console.log(e); this.properties.filePickerResult = e; },
        key: "filePickerId",
        buttonLabel: "File Picker",
        label: "File Picker",
      })
    }

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
                PropertyPaneChoiceGroup('textOrImageType', {
                  label: 'Image/Text',
                  options: [{
                    key: 'Text',
                    text: 'Text',
                    checked: true
                  },
                  {
                    key: 'Image',
                    text: 'Image',
                  }
                  ]
                }),
                textControl,
                imageSourceControl
              ]
            }
          ]
        }
      ]
    };
  }
}
