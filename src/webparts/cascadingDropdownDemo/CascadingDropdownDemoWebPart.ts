import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { sp } from '@pnp/sp';
import { SPService } from '../../common/service/SPService';

import * as strings from 'CascadingDropdownDemoWebPartStrings';
import CascadingDropdownDemo from './components/CascadingDropdownDemo';
import { ICascadingDropdownDemoProps } from './components/ICascadingDropdownDemoProps';

export interface ICascadingDropdownDemoWebPartProps {
  description: string;
  lists: string;
  fields: string[];
}

export default class CascadingDropdownDemoWebPart extends BaseClientSideWebPart<ICascadingDropdownDemoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _services: SPService = null;
  private _listFields: IPropertyPaneDropdownOption[] = [];

  public async getListFields() {
    if (this.properties.lists) {
      let allFields = await this._services.getFields(this.properties.lists);
      (this._listFields as []).length = 0;
      this._listFields.push(...allFields.map(field => ({ key: field.InternalName, text: field.Title })));
    }
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
      this._services = new SPService(this.context);
      this.getListFields = this.getListFields.bind(this);
    });
  }

  private listConfigurationChanged(propertyPath: string, oldValue: any, newValue: any) {
    console.log("LIST FIELDS:", this._listFields);
    if (propertyPath === 'lists' && newValue) {
      this.properties.fields = [];
      this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }

  public render(): void {
    this.getListFields();
    const element: React.ReactElement<ICascadingDropdownDemoProps> = React.createElement(
      CascadingDropdownDemo,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        list: this.properties.lists,
        fields: this.properties.fields
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  baseTemplate: 100,
                  onPropertyChange: this.listConfigurationChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: null,
                  key: 'listPickerFieldId',
                }),
                PropertyFieldMultiSelect('fields', {
                  key: 'multiSelect',
                  label: "Multi select list fields",
                  options: this._listFields,
                  selectedKeys: this.properties.fields
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
