import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';

import * as strings from 'RecentlyVisitedSitesWebPartStrings';
import { RecentlyVisitedSites, IRecentlyVisitedSitesProps } from './components';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IRecentlyVisitedSitesWebPartProps {
  title: string;
}

export default class RecentlyVisitedSitesWebPart extends BaseClientSideWebPart<IRecentlyVisitedSitesWebPartProps> {
  private graphClient: MSGraphClient;

  public onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          this.graphClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  public render(): void {
    const element: React.ReactElement<IRecentlyVisitedSitesProps> = React.createElement(
      RecentlyVisitedSites,
      {
        title: this.properties.title,
        graphClient: this.graphClient,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
