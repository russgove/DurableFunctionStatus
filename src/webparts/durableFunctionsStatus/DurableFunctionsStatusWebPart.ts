import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DurableFunctionsStatusWebPartStrings';
import DurableFunctionsStatus from './components/DurableFunctionsStatus';
import { IDurableFunctionsStatusProps } from './components/IDurableFunctionsStatusProps';

export interface IDurableFunctionsStatusWebPartProps {
  baseUrl: string;
  taskHub:string;
  systemKey:string;
  orchestrationNames:string;
}

export default class DurableFunctionsStatusWebPart extends BaseClientSideWebPart<IDurableFunctionsStatusWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDurableFunctionsStatusProps> = React.createElement(
      DurableFunctionsStatus,
      {
        baseUrl: this.properties.baseUrl,
        taskHub: this.properties.taskHub,
        systemKey: this.properties.systemKey,
        httpClient:this.context.httpClient,
        orchestrationNames:this.properties.orchestrationNames.split('\n')
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
return Promise.resolve();
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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
                PropertyPaneTextField('baseUrl', {
                  label: strings.BaseUrlFieldLabel
                }),
                PropertyPaneTextField('taskHub', {
                  label: strings.TaskHubFieldLabel
                }),
                PropertyPaneTextField('systemKey', {
                  label: strings.SystemKeyFieldLabel
                }),
                PropertyPaneTextField('orchestrationNames', {
                  label: strings.OrchestrationNamesFieldLabel,multiline:true
                }),
                
                

              ]
            }
          ]
        }
      ]
    };
  }
}
