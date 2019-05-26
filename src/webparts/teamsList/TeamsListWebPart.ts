import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import * as microsoftTeams from '@microsoft/teams-js';
import * as strings from 'TeamsListWebPartStrings';
import TeamsList from './components/TeamsList';
import { ITeamsListProps } from './components/ITeamsListProps';
import { MSGraphClient } from '@microsoft/sp-http';

export interface ITeamsListWebPartProps {
  description: string;
}
export default class TeamsListWebPart extends BaseClientSideWebPart<ITeamsListWebPartProps> {
  private _teamsContext: microsoftTeams.Context;
  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }
  public render(): void {
    this.context.msGraphClientFactory.getClient()
    .then((client: MSGraphClient): void => {  
    const element: React.ReactElement<ITeamsListProps > = React.createElement(
      TeamsList,
      {
        graphClient: client,
        userEmail:this.context.pageContext.user.email,
        description:this.properties.description
      }
    );
    ReactDom.render(element, this.domElement);
  });
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
