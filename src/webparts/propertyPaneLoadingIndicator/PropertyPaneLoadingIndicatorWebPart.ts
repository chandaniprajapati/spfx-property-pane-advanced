import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { sp } from '@pnp/sp';
import { SPService } from '../../common/service/SPService';
import * as strings from 'PropertyPaneLoadingIndicatorWebPartStrings';
import PropertyPaneLoadingIndicator from './components/PropertyPaneLoadingIndicator';
import { IPropertyPaneLoadingIndicatorProps } from './components/IPropertyPaneLoadingIndicatorProps';

export interface IPropertyPaneLoadingIndicatorWebPartProps {
  description: string;
  fields: string[];
}

export default class PropertyPaneLoadingIndicatorWebPart extends BaseClientSideWebPart<IPropertyPaneLoadingIndicatorWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private showLoadingIndicator: boolean = true;
  private _services: SPService = null;
  private _listFields: any[] = [];
  private allFields: any;

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
      this._services = new SPService(this.context);
    });
  }

  public render(): void {
    const element: React.ReactElement<IPropertyPaneLoadingIndicatorProps> = React.createElement(
      PropertyPaneLoadingIndicator,
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

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    let allFields = await this._services.getFields('64573b08-321c-4107-a5e1-6c3b5a1ba607');
    (this._listFields as []).length = 0;
    this._listFields.push(...allFields.map(field => ({ key: field.InternalName, text: field.Title })));
    this.showLoadingIndicator = false;
    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      showLoadingIndicator: this.showLoadingIndicator,
      loadingIndicatorDelayTime: 5000,
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
                PropertyPaneDropdown('listFields', {
                  label: "List Fields",
                  options: this._listFields
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
