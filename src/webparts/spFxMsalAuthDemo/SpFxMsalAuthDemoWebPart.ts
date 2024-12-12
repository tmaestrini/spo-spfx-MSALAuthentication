import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpFxMsalAuthDemoWebPartStrings';
import { SpFxMsalAuthDemo } from './components/SpFxMsalAuthDemo';
import { ISpFxMsalAuthDemoProps } from './components/ISpFxMsalAuthDemoProps';


export interface ISpFxMsalAuthDemoWebPartProps {
  applicationID: string;
  redirectUri: string;
  tenantUrl: string;
}

export default class SpFxMsalAuthDemoWebPart extends BaseClientSideWebPart<ISpFxMsalAuthDemoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpFxMsalAuthDemoProps> = React.createElement(
      SpFxMsalAuthDemo,
      {
        applicationID: this.properties.applicationID,
        redirectUri: this.properties.redirectUri,
        tenantUrl: this.properties.tenantUrl,
        httpClient: this.context.httpClient,
        userMail: this.context.pageContext.user.email,
      }
    );

    ReactDom.render(element, this.domElement);
  }


  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
                PropertyPaneTextField('applicationID', {
                  label: strings.ApplicationIDFieldLabel
                }),
                PropertyPaneTextField('redirectUri', {
                  label: strings.RedirectUriFieldLabel
                }),
                PropertyPaneTextField('tenantUrl', {
                  label: strings.TenantUrlFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
