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

import * as strings from 'PictureGalleryWebPartStrings';
import PictureGallery from './components/PictureGallery';
import { IPictureGalleryProps } from './components/IPictureGalleryProps';
import { sp } from '@pnp/sp';

export interface IPictureGalleryWebPartProps {
  webpartTitle: string;
  webpartLabel: string;
  listTitle: string;
  context: WebPartContext;
  siteUrl: string;
  seeAllURL: string;
  showIndex: boolean;
  showBullets: boolean;
  infinite: boolean;
  showThumbnails: boolean;
  showFullscreenButton: boolean;
  //showGalleryFullscreenButton: boolean;
  showPlayButton: boolean;
  //showGalleryPlayButton: boolean;
  showNav: boolean;
  isRTL: boolean;
  slideDuration: number;
  slideInterval: number;
  slideOnThumbnailOver: boolean;
  thumbnailPosition: any;
  webpartType: any;
  useWindowKeyDown: boolean;
}

export default class PictureGalleryWebPart extends BaseClientSideWebPart<IPictureGalleryWebPartProps> {
  private lists: IPropertyPaneDropdownOption[];
  private listsDropDownDisabled: boolean = false;

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      if (this.properties.webpartTitle == undefined) {
        this.properties.webpartTitle = "Picture Gallery";
      }
      if (this.properties.webpartLabel == undefined) {
        this.properties.webpartLabel = "See All";
      }
      /*
      if (this.properties.infinite == undefined) {
        this.properties.infinite = true;
      }
      if (this.properties.isRTL == undefined) {
        this.properties.isRTL = false;
      }
      if (this.properties.showBullets == undefined) {
        this.properties.showBullets = false;
      }
      if (this.properties.showFullscreenButton == undefined) {
        this.properties.showFullscreenButton = true;
      }
      
      if (this.properties.showIndex == undefined) {
        this.properties.showIndex = false;
      }
      if (this.properties.showNav == undefined) {
        this.properties.showNav = true;
      }
      if (this.properties.showPlayButton == undefined) {
        this.properties.showPlayButton = true;
      }
      if (this.properties.showThumbnails == undefined) {
        this.properties.showThumbnails = true;
      }
      if (this.properties.useWindowKeyDown == undefined) {
        this.properties.useWindowKeyDown = false;
      }

      if (this.properties.thumbnailPosition == undefined || this.properties.thumbnailPosition == "") {
        this.properties.thumbnailPosition = "bottom";
      }
      if (this.properties.slideDuration == undefined) {
        this.properties.slideDuration = 450;
      }
      if (this.properties.slideInterval == undefined) {
        this.properties.slideInterval = 2000;
      }
      */
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
    );
  }

  public render(): void {
    const element: React.ReactElement<IPictureGalleryProps> = React.createElement(
      PictureGallery,
      {
        webpartTitle: this.properties.webpartTitle,
        webpartLabel: this.properties.webpartLabel,
        listTitle: this.properties.listTitle,
        context: this.context,
        siteURL: this.context.pageContext.web.absoluteUrl,
        seeAllURL: this.properties.seeAllURL,
        showIndex: this.properties.showIndex,
        showBullets: this.properties.showBullets,
        infinite: this.properties.infinite,
        showThumbnails: this.properties.showThumbnails,
        showFullscreenButton: this.properties.showFullscreenButton,
        //showGalleryFullscreenButton: this.properties.showGalleryFullscreenButton,
        showPlayButton: this.properties.showPlayButton,
        //showGalleryPlayButton: this.properties.showGalleryPlayButton,
        showNav: this.properties.showNav,
        isRTL: this.properties.isRTL,
        slideDuration: this.properties.slideDuration,
        slideInterval: this.properties.slideInterval,
        slideOnThumbnailOver: this.properties.slideOnThumbnailOver,
        thumbnailPosition: this.properties.thumbnailPosition,
        useWindowKeyDown: this.properties.useWindowKeyDown,
        webpartType: this.properties.webpartType
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
                PropertyPaneDropdown('webpartType', {
                  label: 'Webpart Type',
                  options: [

                    { key: 'PictureGallery', text: 'Picture Gallery' },

                    { key: 'VideoGallery', text: 'Video Gallery' }

                  ]
                }),
                PropertyPaneTextField('webpartTitle', {
                  label: "Webpart Title",
                  multiline: false,
                  placeholder: "Enter webpart title",
                }),
                PropertyPaneTextField('webpartLabel', {
                  label: "See All Label",
                  multiline: false,
                  placeholder: "Enter label",
                }),
                PropertyPaneTextField('seeAllURL', {
                  label: "See All URL",
                  multiline: false,
                  placeholder: "Enter label",
                }),
                PropertyPaneDropdown('listTitle', {
                  label: "Select a List",
                  options: this.lists,
                  disabled: this.listsDropDownDisabled,
                  selectedKey: ' ',
                }),
                PropertyPaneToggle('showIndex', {
                  label: 'Image Counter',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('showBullets', {
                  label: 'Bullets',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('infinite', {
                  label: 'Automatic Scrolling',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('showThumbnails', {
                  label: 'Image Thumbnails',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('showFullscreenButton', {
                  label: 'Fullscreen Button',
                  onText: 'On',
                  offText: 'Off'
                }),
                /*PropertyPaneToggle('showGalleryFullscreenButton', {
                  label: 'Gallery Fullscreen Button',
                  onText: 'On',
                  offText: 'Off'
                }),*/
                PropertyPaneToggle('showPlayButton', {
                  label: 'Play Button',
                  onText: 'On',
                  offText: 'Off'
                }),                
                PropertyPaneTextField('slideInterval', {
                  label: "Play Slide Duration",
                  multiline: false,
                  placeholder: "Enter Play Slide Duration",
                }),
              /*  PropertyPaneToggle('showGalleryPlayButton', {
                  label: 'Gallery Play Button',
                  onText: 'On',
                  offText: 'Off'
                }),*/
                PropertyPaneToggle('showNav', {
                  label: 'Navigation Buttons',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('slideDuration', {
                  label: "Navigation Slide Duration",
                  multiline: false,
                  placeholder: "Enter Navigation Slide Duration",
                }),
                PropertyPaneToggle('isRTL', {
                  label: 'Is Right To Left',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('slideOnThumbnailOver', {
                  label: 'Slide On Thumbnail Over',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('useWindowKeyDown', {
                  label: 'Use Window Key Down',
                  onText: 'On',
                  offText: 'Off'
                }),
              
                // PropertyPaneTextField('thumbnailPosition', {
                //   label: "Thumbnail Position",
                //   multiline: false,
                //   placeholder: "Enter thumbnail position",
                // }),
                PropertyPaneDropdown('thumbnailPosition', {
                  label: 'Thumbnail Position',
                  options: [

                    { key: 'top', text: 'top' },

                    { key: 'bottom', text: 'bottom' },

                    { key: 'left', text: 'left' },

                    { key: 'right', text: 'right' }

                  ]
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
