import * as microsoftTeams from '@microsoft/teams-js';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
} from '@microsoft/sp-property-pane';

import * as strings from 'MyFirstTeamsTabWebPartStrings';
import MyFirstTeamsTab from './components/MyFirstTeamsTab';
import { IMyFirstTeamsTabProps } from './components/IMyFirstTeamsTabProps';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/Callout';
import { PropertyFieldChoiceGroupWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldChoiceGroupWithCallout';


export interface IMyFirstTeamsTabWebPartProps {
  description: string;
  choiceGroupWithCalloutValue: string;
  x: number;
  y: number;
  resultStack: string[];
}

export default class MyFirstTeamsTabWebPart extends BaseClientSideWebPart<IMyFirstTeamsTabWebPartProps> {

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

    let title: string = '';
    let subtitle: string = '';
    let siteTabTitle: string = '';

    if (this._teamsContext) {
      title = "Welcome to Teams!";
      subtitle = "Twórz własne tabs!";
      siteTabTitle = "Witamy w Team: " + this._teamsContext.teamName;
    }
    else
    {
      title = "Welcome to SharePoint!";
      subtitle = "Powiekszaj swoje doświadczenie w SharePoint!";
      siteTabTitle = "Witamy na tenant: " + this.context.pageContext.web.title;
    }


    const element: React.ReactElement<IMyFirstTeamsTabProps > = React.createElement(
      MyFirstTeamsTab,
      {
        description: this.properties.description,
        resultStack: this.addToResults()
      }
    );

  ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected addToResults(): string[] {
    let sum: string;
    switch(this.properties.choiceGroupWithCalloutValue) { 
      case strings.PlusOperation: { 
         sum = (this.properties.x + this.properties.y)+'';
         break; 
      } 
      case strings.MinusOperation: { 
        sum = (this.properties.x - this.properties.y)+'';
        break; 
      }
      case strings.DevisionOperation: { 
        if (this.properties.y !== 0) {
          sum = (this.properties.x/this.properties.y)+'';  
        } else {
          sum = strings.ErrorDivisionMessage;
        }
        break; 
      }

      case strings.RemOfDiv: { 
        if (this.properties.y !== 0) {
          sum = (this.properties.x%this.properties.y)+'';  
        } else {
          sum = strings.ErrorDivisionMessage;
        }
        break;
      }

      case strings.MultiOperation: {
        sum = (this.properties.x * this.properties.y)+'';
        break; 
      }
      case strings.MultiOperation: {
        sum = (this.properties.x * this.properties.y)+'';
        break; 
      }
      case strings.PiValue: {
        sum = '3,14';
        break; 
      }
      default: { 
         sum = '0';
         break; 
      } 
   } 
    this.properties.resultStack.push(sum);
    return this.properties.resultStack;
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
                PropertyFieldNumber(strings.XFieldKey, {
                  key: strings.XFieldKey,
                  label: strings.XFieldLabel,
                  description: strings.XFieldLabel,
                  disabled: false
                }),
                PropertyFieldNumber(strings.YFieldKey, {
                  key: strings.YFieldKey,
                  label: strings.YFieldLabel,
                  description: strings.YFieldLabel,
                  disabled: false
                }),
                PropertyFieldChoiceGroupWithCallout(strings.ChoiceGroupWithCalloutValue, {
                  calloutContent: React.createElement('div', {}, 'Select operation'),
                  calloutTrigger: CalloutTriggers.Hover,
                  key: strings.ChoiceGroupWithCalloutFieldId,
                  label: strings.ChoiceGroupLabel,
                  options: [{
                    key: strings.PlusOperation,
                    text: strings.PlusOperation,
                    checked: this.properties.choiceGroupWithCalloutValue === strings.PlusOperation
                  }, {
                    key: strings.MinusOperation,
                    text: strings.MinusOperation,
                    checked: this.properties.choiceGroupWithCalloutValue === strings.MinusOperation
                  }, {
                    key: strings.DevisionOperation,
                    text: strings.DevisionOperation,
                    checked: this.properties.choiceGroupWithCalloutValue === strings.DevisionOperation
                  },  {
                    key: strings.RemOfDiv,
                    text: strings.RemOfDiv,
                    checked: this.properties.choiceGroupWithCalloutValue === strings.RemOfDiv
                  }, {
                    key: strings.MultiOperation,
                    text: strings.MultiOperation,
                    checked: this.properties.choiceGroupWithCalloutValue === strings.MultiOperation
                  }, {
                    key: strings.PiValue,
                    text: strings.PiValue,
                    checked: this.properties.choiceGroupWithCalloutValue === strings.PiValue
                  }]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}