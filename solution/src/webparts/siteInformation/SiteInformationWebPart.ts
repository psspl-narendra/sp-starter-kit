import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IWebPartPropertiesMetadata,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

import * as strings from 'SiteInformationWebPartStrings';
import SiteInformation from './components/SiteInformation';
import { ISiteInformationProps } from './components/ISiteInformationProps';
import { ISiteInformationWebPartProps } from './ISiteInformationWebPartProps';

export default class SiteInformationWebPart extends BaseClientSideWebPart<ISiteInformationWebPartProps> {

  private propertyFieldTermPicker;
  private propertyFieldPeoplePicker;
  private principalType;

  public onInit(): Promise<void> {

    return super.onInit().then(async (_) => {
      
      //chunk shared by all web parts
      const { sp } = await import(
        /* webpackChunkName: 'pnp-sp' */
        "@pnp/sp");

      // initialize the PnP JS library
      sp.setup({
        spfxContext: this.context
      });

      // initialize the Site Title property reading the current site title via PnP JS
      if (!this.properties.siteTitle) {
        sp.web.select("Title").get().then((r: any) => {
          this.properties.siteTitle = r.Title;
        });
      }
    });
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    // makes the properties of the web part searchable with the SPO search engine
    return {
      'siteTitle': { isSearchablePlainText: true },
      'siteContact': { isSearchablePlainText: true },
      'siteOrganization': { isSearchablePlainText: true }
    };
  }

  public render(): void {
    const element: React.ReactElement<ISiteInformationProps> = React.createElement(
      SiteInformation,
      {
        siteTitle: this.properties.siteTitle,
        siteContactLogin: (this.properties.siteContact && this.properties.siteContact.length > 0) ?
          this.properties.siteContact[0].login : "",
        siteContactEmail: (this.properties.siteContact && this.properties.siteContact.length > 0) ?
          this.properties.siteContact[0].email: null,
        siteContactFullName: (this.properties.siteContact && this.properties.siteContact.length > 0) ?
          this.properties.siteContact[0].fullName: null,
        siteContactImageUrl: (this.properties.siteContact && this.properties.siteContact.length > 0) ?
          this.properties.siteContact[0].imageUrl: null,
        siteOrganization: (this.properties.siteOrganization && this.properties.siteOrganization.length > 0) ?
          this.properties.siteOrganization[0].name : "",
        needsConfiguration: this.needsConfiguration(),
        configureHandler: () => {
          this.context.propertyPane.open();
        },
        errorHandler: (errorMessage: string) => {
          if (this.displayMode === DisplayMode.Edit) {
            this.context.statusRenderer.renderError(this.domElement, errorMessage);
          } else {
            // nothing to do, if we are not in edit Mode
          }
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  
  //executes only before property pane is loaded.
  protected async loadPropertyPaneResources(): Promise<void> {
    // import additional controls/components

    const { PropertyFieldTermPicker } = await import (
      /* webpackChunkName: 'pnp-propcontrols-termpicker' */
      '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker'
    );
    const { PropertyFieldPeoplePicker, PrincipalType } = await import (
      /* webpackChunkName: 'pnp-propcontrols-peoplepicker' */
      '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker'
    );

    this.propertyFieldTermPicker = PropertyFieldTermPicker;
    this.propertyFieldPeoplePicker = PropertyFieldPeoplePicker;
    this.principalType = PrincipalType;
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
                PropertyPaneTextField('siteTitle', {
                  label: strings.SiteTitleFieldLabel
                }),
                this.propertyFieldPeoplePicker('siteContact', {
                  label: strings.SiteContactFieldLabel,
                  initialData: this.properties.siteContact,
                  allowDuplicate: false,
                  multiSelect: false,
                  principalType: [this.principalType.Users],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'siteContactId'
                }),
                this.propertyFieldTermPicker('siteOrganization', {
                  label: strings.SiteOrganizationFieldLabel,
                  panelTitle: strings.SiteOrganizationPanelTitle,
                  initialValues: this.properties.siteOrganization,
                  allowMultipleSelections: false,
                  excludeSystemGroup: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  limitByGroupNameOrID: 'PnPTermSets',
                  limitByTermsetNameOrID: 'PnP-Organizations',
                  key: 'siteOrganizationId'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  // method to refresh any error after properties configuration
  protected onAfterPropertyPaneChangesApplied(): void {
    this.context.statusRenderer.clearError(this.domElement);
  }

  // method to determine if the web part has to be configured
  private needsConfiguration(): boolean {
    // as long as we don't have the stock symbol, we need configuration
    return ((!this.properties.siteTitle ||
      this.properties.siteTitle.length === 0) ||
      (!this.properties.siteContact) ||
      (!this.properties.siteOrganization));
  }
}
