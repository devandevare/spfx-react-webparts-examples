import * as React from 'react';
import styles from './LeftNavigation.module.scss';
import { ILeftNavigationProps } from './ILeftNavigationProps';
import dataService from "../../../Common/DataService";
import CONSTANTS from "../../../Common/Constants";
import { IStackTokens, PrimaryButton, Stack } from 'office-ui-fabric-react'; //'@fluentui/react';
//import { PrimaryButton } from 'office-ui-fabric-react';
import { RxJsEventEmitter } from '../../RxJsEventEmitter/RxJsEventEmitter';
import commonMethods from "../../../Common/CommonMethods";
import * as $ from 'jquery';


let selectedRegion: any = [];
let selectedProgramType: any = [];
export interface ILeftNavigationState {
  leftNavigationItems: any[];
  error: string;
  
}

const outerStackTokens: IStackTokens = { childrenGap: 5, padding: 10 };

export interface IEventData {
  sharedRegion: any;
  sharedProgramType: any;
  sharedLeftNavigation: string;
  sharedSiteLocationMetdataAvailable: boolean;
}

const commonService = new dataService();
const commonMethod = new commonMethods();
export default class LeftNavigation extends React.Component<ILeftNavigationProps, ILeftNavigationState> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  constructor(props: ILeftNavigationProps, state: ILeftNavigationState) {
    super(props);
    this.state = {
      leftNavigationItems: [],
      error: '',
      
    };
    this.eventEmitter.on("shareData", this.receiveData.bind(this));
  }

  private receiveData(data: IEventData) {

    selectedRegion = data.sharedRegion;
    selectedProgramType = data.sharedProgramType;
  }

  public dropDownValidation = async (listTitle: string) => {
    if (listTitle == undefined || listTitle == ' ') {
      this.setState({
        error: CONSTANTS.SYS_CONFIG.SELECT_LIST
      });
    }
    else {

      //Check all required fields available in selected list.

      let isValidListColumns = await commonMethod.isValidListColumns(listTitle, CONSTANTS.LIST_VALIDATION_COLUMNS.LEFT_NAVIGATION);
      // alert("isValidListColumns:- " + isValidListColumns);
      //Check all required fields available in selected list.
      if (isValidListColumns) {
        this.GetLeftNavigation();
      } else {
        this.setState({
          error: CONSTANTS.SYS_CONFIG.LEFT_NAVIGATION_LIST_NOT_MATCH
        });
      }

    }
  }

  public componentDidMount(): void {
    this.dropDownValidation(this.props.listTitle);
    
   
  }


  private openDocumentSearch = (id: number, title: string, libraryName: string, libraryTitle: string, isSiteLocationMetadataAvailable: string): void => {

    let documentSearchURL = this.props.context.pageContext.web.absoluteUrl + CONSTANTS.SYS_CONFIG.DOCUMENT_SEARCH_PAGE + "?lnav=" + id + "&slmda=" + isSiteLocationMetadataAvailable;
    let homePageURL = this.props.context.pageContext.web.absoluteUrl + CONSTANTS.SYS_CONFIG.HOME_PAGE;
    let siteLocationMetadataAvailable: boolean = isSiteLocationMetadataAvailable.toString().toLowerCase() == "true" ? true : false;
    let queryStringParameters = new URLSearchParams(window.location.search);
    let documentSearchQueryString: string = "";

    if (selectedRegion.text != undefined && selectedRegion.text != "Select") {
      documentSearchQueryString += "&rgn=" + selectedRegion.key;
    } else {
      documentSearchQueryString += "&rgn=Select";
    }
    if (selectedProgramType.text != undefined && selectedProgramType.text != "Select") {
      documentSearchQueryString += "&pty=" + selectedProgramType.key;
    } else {
      documentSearchQueryString += "&pty=Select";
    }

    let currentPage = window.location.pathname.split('/')[(window.location.pathname.split('/').length - 1)];

    documentSearchURL = documentSearchURL + documentSearchQueryString;
    let documentSearchPage = CONSTANTS.SYS_CONFIG.DOCUMENT_SEARCH_PAGE.split('/')[(CONSTANTS.SYS_CONFIG.DOCUMENT_SEARCH_PAGE.split('/').length - 1)];
    if (currentPage.toLocaleLowerCase() == documentSearchPage.toLocaleLowerCase()) {

      if (CONSTANTS.SYS_CONFIG.LEFT_NAVIGATION_HOME_TITLE.toLocaleLowerCase() == title.toLocaleLowerCase()) {
        window.open(homePageURL, "_self");
      } else {
        this.sendData(queryStringParameters.get("rgn"), queryStringParameters.get("pty"), id.toString(), siteLocationMetadataAvailable);
      }

    } else {
      //if click on Home left navigation then skip
      if (CONSTANTS.SYS_CONFIG.LEFT_NAVIGATION_HOME_TITLE.toLocaleLowerCase() != title.toLocaleLowerCase()) {
        window.open(documentSearchURL, "_blank");
      }
    }

  }

  private sendData(selectedRegionValue: string, selectedProgramTypeValue: string, leftNavigationId: string, siteLocationMeatdataAvailable: boolean): void {
    var eventBody = {
      sharedRegion: selectedRegionValue,
      sharedProgramType: selectedProgramTypeValue,
      sharedLeftNavigation: leftNavigationId,
      sharedSiteLocationMetdataAvailable: siteLocationMeatdataAvailable
    } as IEventData;

    let leftNavigationItems = this.state.leftNavigationItems;
    leftNavigationItems.forEach(async (leftNavigationItem: any, index: number) => {
      if (leftNavigationItem.Id == leftNavigationId) {
        leftNavigationItem.IsActiveLeftNav = true;
      } else {
        leftNavigationItem.IsActiveLeftNav = false;
      }
    });
    this.setState({
      leftNavigationItems: leftNavigationItems
    });
    this.eventEmitter.emit("shareData", eventBody);
  }

  private GetLeftNavigation = (): void => {

    let leftNavItems: any[] = [];
    let queryStringParameters = new URLSearchParams(window.location.search);
    let queryStringLeftNav: string = "";
    let currentPage = window.location.pathname.split('/')[(window.location.pathname.split('/').length - 1)];


    let documentSearchPage = CONSTANTS.SYS_CONFIG.DOCUMENT_SEARCH_PAGE.split('/')[(CONSTANTS.SYS_CONFIG.DOCUMENT_SEARCH_PAGE.split('/').length - 1)];

    leftNavItems.push({
      Title: CONSTANTS.SYS_CONFIG.LEFT_NAVIGATION_HOME_TITLE,
      Id: 0,
      LibraryName: "",
      LibraryTitle: "",
      IsSiteLocationMetadataAvailable: false,
      IsActiveLeftNav: currentPage.toLocaleLowerCase() != documentSearchPage.toLocaleLowerCase() ? true : false,
      Icon: this.props.context.pageContext.site.absoluteUrl + CONSTANTS.ICONS.LEFT_NAVIGATION_HOME_ICON,
      ActiveIcon: this.props.context.pageContext.site.absoluteUrl + CONSTANTS.ICONS.LEFT_NAVIGATION_HOME_ACTIVEICON
    });


    commonService.getLeftNavigationConfigurationData(CONSTANTS.LIST_NAME.LEFT_NAVIGATION_CONFIGURATION_LIST, CONSTANTS.SELECTCOLUMNS.LEFT_NAVIGATION_COLS, CONSTANTS.FILTERCONDITION.LEFT_NAVIGATION_QUERY, CONSTANTS.ORDERBY.LEFT_NAVIGATION).then((data: any) => {

      if (data.length > 0) {
        data.forEach(async (leftNavigationItem: any, index: number) => {

          if (queryStringParameters.get("lnav") == leftNavigationItem.Id) {
            //alert(queryStringParameters.get("lnav"));
            queryStringLeftNav = queryStringParameters.get("lnav");
          }
          leftNavItems.push({
            Title: leftNavigationItem.Title,
            Id: leftNavigationItem.Id,
            LibraryName: leftNavigationItem.LibraryName,
            LibraryTitle: leftNavigationItem.LibraryTitle,
            IsSiteLocationMetadataAvailable: leftNavigationItem.IsSiteLocationMetadataAvailable,
            IsActiveLeftNav: queryStringLeftNav == leftNavigationItem.Id ? true : false,
            Icon: leftNavigationItem.Icon != null ? JSON.parse(leftNavigationItem.Icon).serverUrl + JSON.parse(leftNavigationItem.Icon).serverRelativeUrl : this.props.context.pageContext.site.absoluteUrl + CONSTANTS.ICONS.LEFT_NAVIGATION_DEFAULT_ICON,
            ActiveIcon: leftNavigationItem.ActiveIcon != null ? JSON.parse(leftNavigationItem.ActiveIcon).serverUrl + JSON.parse(leftNavigationItem.ActiveIcon).serverRelativeUrl : this.props.context.pageContext.site.absoluteUrl + CONSTANTS.ICONS.LEFT_NAVIGATION_DEFAULT_ACTIVEICON
            //URL: this.props.context.pageContext.web.absoluteUrl + 
          });
        });
      }

      this.setState({
        leftNavigationItems: leftNavItems
      });
      //return PracticeOptions;
    });
    //Set global level css change on page
    this.setGlobalCssChanges();
  }

  private setGlobalCssChanges = (): void => {

    


    $(".ms-SPLegacyFabricBlock").css('background-color', '#1a458a');
    
    $(".CanvasSection").attr('style', 'background-color: #1a458a; padding: 0px !important;');
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1)").attr('style', 'margin-left: 21%; padding-top:24px !important;');
    //text editor style    
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) p").attr('style', 'color:#ffffff !important; font-family: roboto !important; margin-left:40px;');
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) p span").attr('style', 'color:#ffffff !important; font-family: roboto !important; margin-left:40px;');
    //$(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) div").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) div h1").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) div h1").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) div h2").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) div h3").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) div h4").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".Canvas--withLayout > div > div:nth-child(1) > div > div:nth-child(1) div h5").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    
    //////////////////
    $(".Canvas--withLayout > div > div:nth-child(2) > div > div:nth-child(1)").css('width', '22%');
    $(".Canvas--withLayout > div > div:nth-child(2) > div > div:nth-child(1)").attr('style', 'padding: 0px !important; width: 22%; position: absolute; top: 0; left: 0; background:#0e2851; height: 100%;');
    $(".Canvas--withLayout > div > div:nth-child(2) > div > div:nth-child(1)").css('padding-left', '0px !important');
    $(".Canvas--withLayout > div > div:nth-child(2) > div > div:nth-child(1) > div").css('padding-left', '0px !important');
    $(".Canvas--withLayout > div > div:nth-child(2) > div > div:nth-child(1) > div").css('margin-bottom', '0px');
    $(".Canvas--withLayout > div > div:nth-child(2) > div > div:nth-child(1) > div").css('margin-top', '0px');
    
    $(".ControlZone").attr('style', 'margin-top: 0px; margin-bottom: 0px; padding: 0px !important;');
    //New code
    //$(".Canvas--withLayout > div > div:nth-child(2) > div > div:nth-child(2)").attr('style', 'margin-left: 23%; padding-left: 38px; padding-right: 0px; background-color: #1a458a;  width: 39% !important;');
    $(".Canvas--withLayout > div > div:nth-child(2) > div > div:nth-child(2)").attr('style', 'margin-left: 23%; padding-left: 28px; padding-right: 0px; background-color: #1a458a;  width: 38% !important; padding-bottom: 50px; ');
    $(".Canvas--withLayout>div>div:nth-child(2)>div>div:nth-child(3)").attr('style', 'padding-left: 30px; padding-right: 0px; background-color: #1a458a; width: 39% !important;');
    $(".Canvas--withLayout>div>div:nth-child(2)>div>div:nth-child(2)>div:nth-child(1)").attr('style', 'margin-top: 0px; padding: 0px !important;');
    $(".Canvas--withLayout>div>div:nth-child(2)>div>div:nth-child(3)>div:nth-child(1)").attr('style', 'margin-top: 0px; padding: 0px !important;');
    $(".Canvas--withLayout > div > div").attr('style', 'padding-left: 0px; background-color: #1a458a;');
    $(".webPartContainer").attr('style', 'background: #f0f0f7; display: none;');
    
    
    //$(".Canvas--withLayout > div > div > div > div:nth-child(2) span").attr('style', 'color: white !important;');
   /*
    $("#c57442ee-6e3b-43a4-b0cc-fafbb023de3f > div.ControlZone--position > div.ControlZone-control > div > div.oneLineWidth.rte--inline.rte--inline-update.uniformSpacingForElements.blockQuoteFont.rte--edit.disable-link.cke_editable.rteEmphasis.root-83.cke_editable_inline.cke_contents_ltr.cke_show_borders p > span").attr('style', 'color:white !important;');
    $("#c57442ee-6e3b-43a4-b0cc-fafbb023de3f > div.ControlZone--position > div.ControlZone-control > div > div.oneLineWidth.rte--inline.rte--inline-update.uniformSpacingForElements.blockQuoteFont.rte--edit.disable-link.cke_editable.rteEmphasis.root-83.cke_editable_inline.cke_contents_ltr.cke_show_borders p").attr('style', 'color:white !important;');
    $(".ControlZone-control > div > div.oneLineWidth.rte--inline.rte--inline-update.uniformSpacingForElements.blockQuoteFont.rte--edit.disable-link.cke_editable.rteEmphasis.root-83.cke_editable_inline.cke_contents_ltr.cke_show_borders p > span").attr('style', 'color: white !important;');
    $(".ControlZone-control > div > div.oneLineWidth.rte--inline.rte--inline-update.uniformSpacingForElements.blockQuoteFont.rte--edit.disable-link.cke_editable.rteEmphasis.root-83.cke_editable_inline.cke_contents_ltr.cke_show_borders p").attr('style', 'color: white !important;');
    $(".ControlZone-control > div > div.oneLineWidth.rte--inline.rte--inline-update.uniformSpacingForElements.blockQuoteFont.rte--edit.disable-link.cke_editable.rteEmphasis.root-83.cke_editable_inline.cke_contents_ltr.cke_show_borders h1").attr('style', 'color: white !important;');
    $(".ControlZone-control > div > div.oneLineWidth.rte--inline.rte--inline-update.uniformSpacingForElements.blockQuoteFont.rte--edit.disable-link.cke_editable.rteEmphasis.root-83.cke_editable_inline.cke_contents_ltr.cke_show_borders h2").attr('style', 'color: white !important;');
    $(".ControlZone-control > div > div.oneLineWidth.rte--inline.rte--inline-update.uniformSpacingForElements.blockQuoteFont.rte--edit.disable-link.cke_editable.rteEmphasis.root-83.cke_editable_inline.cke_contents_ltr.cke_show_borders h3").attr('style', 'color: white !important;');
    */

    /*
    $(".cke_editable h1").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".cke_editable h2").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".cke_editable h3").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".cke_editable h4").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".cke_editable h5").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".cke_editable p").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".cke_editable p span").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');

    $(".rte--edit h1").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".rte--edit h2").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".rte--edit h3").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".rte--edit h4").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".rte--edit p").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
    $(".rte--edit p span").attr('style', 'color:#ffffff !important; font-family: roboto !important;margin-left:40px;');
*/
    $(".rte--edit").attr('style', 'background: white !important;');

    
    
    
  }

  public render(): React.ReactElement<ILeftNavigationProps> {
    return (
      
      <div className={styles.leftNavigation} >

     
        <Stack tokens={outerStackTokens} className={styles.stackclass}>
          <React.Fragment>
            {this.state.leftNavigationItems.map((item: any, i: number) => {

              return (
                <Stack.Item grow className={styles.leftNav}>
                  {/*<PrimaryButton className={item.IsActiveLeftNav == true ? styles.activeButton : styles.navigation_button} onClick={() => this.openDocumentSearch(item.Id, item.Title, item.LibraryName, item.LibraryTitle, item.IsSiteLocationMetadataAvailable)} text={item.Title} />*/}
                  <div className={item.IsActiveLeftNav == true ? styles.activeLeftNav : styles.LeftNav}>
                    <div className={`${styles.leftNavInnerDiv} ${styles.leftNavIconDiv} `}>
                      <img className={styles.leftNaveIcon}
                        onClick={() => this.openDocumentSearch(item.Id, item.Title, item.LibraryName, item.LibraryTitle, item.IsSiteLocationMetadataAvailable)}
                        src={item.IsActiveLeftNav == true ? item.ActiveIcon : item.Icon}></img>
                    </div>
                    <div className={`${styles.leftNavInnerDiv} ${styles.leftNavLinkDiv} `}>
                      <a className={styles.leftNavTitle} onClick={() => this.openDocumentSearch(item.Id, item.Title, item.LibraryName, item.LibraryTitle, item.IsSiteLocationMetadataAvailable)} >{item.Title}</a>
                    </div>
                    <div className={`${styles.leftNavActiveDiv} ${item.IsActiveLeftNav == true ? styles.leftNavActiveFirstDiv : styles.leftNavFirstDiv} `}>
                      &nbsp;
                    </div>
                  </div>
                </Stack.Item>
              );

            })
            }
            <Stack.Item grow className={styles.leftNavHUBWatermarkLogoStackItem}>
              <img className={styles.leftNavHUBWatermarkLogo}
                src={this.props.context.pageContext.site.absoluteUrl + CONSTANTS.ICONS.LEFT_NAVIGATION_HUBWATERMARKLOGO}></img>
            </Stack.Item>
          </React.Fragment>

          <div className={styles.errorDiv}>
            {this.state.error}
          </div>
        </Stack>
     
      </div>
      
    );
  }
}
