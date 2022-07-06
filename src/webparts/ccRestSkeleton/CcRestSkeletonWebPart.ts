import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {
  HttpClient,
  HttpClientResponse
} from '@microsoft/sp-http';

import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CcRestSkeletonWebPartStrings';
import CcRestSkeleton from './components/CcRestSkeleton';
import { ICcRestSkeletonProps } from './components/ICcRestSkeletonProps';

export interface ICcRestSkeletonWebPartProps {
  description: string;
}

export default class CcRestSkeletonWebPart extends BaseClientSideWebPart<ICcRestSkeletonWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    if (!this.renderedOnce) {
      this._getAPIResult()
        .then(response => {
          const element: React.ReactElement<ICcRestSkeletonProps > = React.createElement(
            CcRestSkeleton,
            {
              APIResult : response,
              description: this.properties.description,
              isDarkTheme: this._isDarkTheme,
              environmentMessage: this._environmentMessage,
              hasTeamsContext: !!this.context.sdks.microsoftTeams,
              userLoginName: this.context.pageContext.user.loginName,
              userEmail: this.context.pageContext.user.email,
              userDisplayName: this.context.pageContext.user.displayName
            }
          );

          ReactDom.render(element, this.domElement);
        });
    }
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

  private _getAPIResult(): Promise<any> {
    /*
    return this.context.httpClient.get(
      `https://geek-jokes.sameerkumar.website/api`,
      HttpClient.configurations.v1
    )
    */

    return this.context.httpClient.get(
      `https://prod-05.northcentralus.logic.azure.com/workflows/d39d62a5b544437199449debc22f2127/triggers/manual/paths/invoke/user/` + this.context.pageContext.user.loginName + `?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=xkKS0wqSX2o7-OcjpZ09EmqyRIqayP0XHvTtmvfaZIs`,
      HttpClient.configurations.v1
    )


    .then((response: HttpClientResponse) => {
      return response.text();
    })
    .then(textResponse => {
      return textResponse;
    }) as Promise<any>;
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
