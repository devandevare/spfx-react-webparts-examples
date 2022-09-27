import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'DocumentSearchWebPartStrings';
import DocumentSearch from './components/DocumentSearch';
import { IDocumentSearchProps } from './components/IDocumentSearchProps';
import { sp } from '@pnp/sp';

export interface IDocumentSearchWebPartProps {
  context: WebPartContext;
  //listTitle: string;
  webpartTitle: string;
  toggleAddNew:boolean;
}

export default class DocumentSearchWebPart extends BaseClientSideWebPart<IDocumentSearchWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropDownDisabled: boolean = false;

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      if(this.properties.webpartTitle == undefined){
        this.properties.webpartTitle = "Document Search";
      }
      if(this.properties.toggleAddNew == undefined){
        this.properties.toggleAddNew = false;
      }
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if(propertyPath == 'toggleAddNew' || propertyPath == 'webpartTitle'){
      this.context.propertyPane.refresh();
      //this.properties.toggleAddNew == newValue;
    }else {
    this.context.propertyPane.refresh();  

    this.onDispose();
    this.render();
    }
   
  }
  /*
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
        sp.web.lists.get().then(listOptions => {
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
    )
  }
 */

  public render(): void {
    const element: React.ReactElement<IDocumentSearchProps> = React.createElement(
      DocumentSearch,
      {
        //listTitle: this.properties.listTitle,
        context: this.context,
        webpartTitle: this.properties.webpartTitle,
        toggleAddNew: this.properties.toggleAddNew
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
              groupName: "Custom Properties",
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: "Webpart Title",
                  multiline: false,
                  placeholder: "Enter webpart title",
      
                }),               
                PropertyPaneToggle('toggleAddNew',{
                  label: 'Show/Hide Add New Button',
                  checked:false
                })
                
              ]
            }
          ]
        }
      ]
    };
  }
}
