import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'SynopsysEventsWebPartStrings';
import SynopsysEvents from './components/SynopsysEvents';
import { ISynopsysEventsProps } from './components/ISynopsysEventsProps';

import { sp } from '@pnp/sp';

export interface ISynopsysEventsWebPartProps {
  webpartTitle: string;
  webpartLabel: string;
  siteURL: string;
  seeAllURL: string;
  eventListName: string;
  context?: WebPartContext;
}

export default class SynopsysEventsWebPart extends BaseClientSideWebPart<ISynopsysEventsWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropDownDisabled: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      if(this.properties.webpartTitle == undefined){
        this.properties.webpartTitle = "Events";
      }
      if(this.properties.webpartLabel == undefined){
        this.properties.webpartLabel = "See All";
      }
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    
    this.context.propertyPane.refresh();  

    this.onDispose();
    this.render();
    
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>(
      (
        resolve: (options: IPropertyPaneDropdownOption[]) => void,
        reject: (error: any) => void
      ) => {
        sp.setup({
          sp: {
            baseUrl: this.context.pageContext.web.absoluteUrl
          }
        });
        sp.web.lists.filter("BaseTemplate eq 106").get().then(listOptions => {
          let lists = [];
          listOptions.forEach(list => {
            if (!list.Hidden) {
              lists.push({
                key: list.Title,
                text: list.Title
              });
            }
          });
          resolve(lists);
        });

      });
  }
  protected onPropertyPaneConfigurationStart(): void {
    this.listsDropDownDisabled = !this.lists;
    if (this.lists) {
      return;
    }
    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      'lists'
    );
    this.loadLists().then(
      (listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.listsDropDownDisabled = false;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      }
    );
  }

  public render(): void {
    const element: React.ReactElement<ISynopsysEventsProps> = React.createElement(
      SynopsysEvents,
      {
        webpartTitle: this.properties.webpartTitle,
        webpartLabel: this.properties.webpartLabel,
        siteURL: this.context.pageContext.web.absoluteUrl,
        seeAllURL: this.properties.seeAllURL,
        eventListName: this.properties.eventListName,
        context: this.context
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
         
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: "Webpart Title",
                  multiline: false,
                  placeholder: "Enter webpart title",
                }),
                PropertyPaneTextField('webpartLabel', {
                  label: "See All Label",
                  multiline: false,
                  placeholder: "Enter Label",
                }),
                PropertyPaneTextField('seeAllURL', {
                  label: "See All URL",
                  multiline: false,
                  placeholder: "Enter Label",
                }),
                PropertyPaneDropdown('eventListName', {
                  label: "Select a List",
                  options: this.lists,
                  disabled: this.listsDropDownDisabled,
                  selectedKey : ' ',
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
