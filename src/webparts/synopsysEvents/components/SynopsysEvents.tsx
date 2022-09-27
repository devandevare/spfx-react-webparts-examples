import * as React from 'react';
import styles from './SynopsysEvents.module.scss';
import { ISynopsysEventsProps } from './ISynopsysEventsProps';
import dataService from '../../../Common/DataService';
import CONSTANTS from '../../../Common/Constants';
import { RxJsEventEmitter } from '../../RxJsEventEmitter/RxJsEventEmitter';
import { Checkbox, ComboBox, DatePicker, DefaultButton, IComboBox, IComboBoxOption, IconButton, IIconProps, IStackProps, IStackTokens, Label, MessageBar, MessageBarType, Modal, PrimaryButton, Stack, TextField } from 'office-ui-fabric-react'; //'@fluentui/react';
import * as _ from 'lodash';
import commonMethods from "../../../Common/CommonMethods";
//import Dropzone from 'react-dropzone';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/DateTimePicker';

let g_selectedRegion: string = "";
let g_selectedProgramType: string = "";

export interface ISynopsysEventsStates {
  eventData: any[];
  seeAllURL: string;
  newFormURL: string;
  error: string;
  regionOptions: any[];
  programTypeOptions: any[];
  selectedRegionValue: any;
  selectedProgramTypeValue: any;
  selectedSiteLocationValue: any;
  errorMessage: string;
  displayNoDataMassage: boolean;
  isCurrentUserPresentInGroup: boolean;
  ownersGroup: string;
  membersGroup: string;
  isModalOpen: boolean;
  regionErrorMessage: string;
  programTypeErrorMessage: string;
  siteLocationErrorMessage: string;
  titleErrorMessage: string;
  descriptionErrorMessage: string;
  locationErrorMessage: string;
  categoryErrorMessage: string;
  showMessageBar: boolean;
  messageType?: MessageBarType;    
  message?: string; 
  siteLocationOptions: any;
  siteLocationData: any[];
  siteLocationColl: any[];
  Title: any;
  Location: any;
  Description: any;
  isMarkAsImportant: boolean;
  categoryOptions: any[];
  selectedCategoryValue: any;
  eventStartDate: Date;
  eventEndDate: Date;
  isAllDayEvent: boolean;
}

export interface IEventData {
  sharedRegion: any;
  sharedProgramType: any;
}

const verticalStackProps: IStackProps = {  
  styles: { root: { overflow: 'hidden', width: '100%' } },  
  tokens: { childrenGap: 20 }  
}; 
const cancelIcon: IIconProps = { iconName: 'Cancel' };
const outerStackTokens: IStackTokens = { childrenGap: 5, padding: 10 };
const innerStackTokens: IStackTokens = { childrenGap: 10 };
const commonService = new dataService();
const commonMethod = new commonMethods();
export default class SynopsysEvents extends React.Component<ISynopsysEventsProps, ISynopsysEventsStates> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  public constructor(props: ISynopsysEventsProps, state: ISynopsysEventsStates) {
    super(props);
    this.state = {
      eventData: [],
      seeAllURL: "",
      newFormURL: "",
      error: "",
      regionOptions: [],
      programTypeOptions: [],
      selectedRegionValue: [],
      selectedProgramTypeValue: [],
      selectedSiteLocationValue: [],
      errorMessage: "",
      displayNoDataMassage: false,
      isCurrentUserPresentInGroup: false,
      ownersGroup: "",
      membersGroup: "",
      isModalOpen: false,
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      titleErrorMessage: "",
      descriptionErrorMessage: "",
      locationErrorMessage: "",
      categoryErrorMessage: "",
      showMessageBar: false,
      siteLocationOptions: [],
      siteLocationData: [],
      siteLocationColl: [],
      Title: "",
      Location: "",
      Description: "",
      isMarkAsImportant: false,
      categoryOptions: [],
      selectedCategoryValue: [],
      eventStartDate: new Date(),  
      eventEndDate: new Date(),  
      isAllDayEvent: false
    };
    this.eventEmitter.on(CONSTANTS.CONNECTED_WP.SHARE_DATA, this.receiveData.bind(this));
  }

  public dropDownValidation = async (listTitle: string) => {
    if (listTitle == undefined || listTitle == ' ') {
      this.setState({
        error: CONSTANTS.SYS_CONFIG.SELECT_LIST
      });
    }
    else {

      //Check all required fields available in selected list.
      let isValidListColumns = await commonMethod.isValidListColumns(listTitle, CONSTANTS.LIST_VALIDATION_COLUMNS.EVENTS);

      if (isValidListColumns) {
        this.getAllEvents();
      } else {
        this.setState({
          error: CONSTANTS.SYS_CONFIG.EVENTS_LIST_NOT_MATCH
        });
      }

    }
  }

  public componentDidMount(): void {
    this.dropDownValidation(this.props.eventListName);
    this.setState({
      seeAllURL: this.props.seeAllURL == "" ? this.props.siteURL + CONSTANTS.SYS_CONFIG.SITE_LISTS + this.props.eventListName + CONSTANTS.SYS_CONFIG.EVENTS_LIST_PAGE : this.props.seeAllURL,
      newFormURL: this.props.siteURL + CONSTANTS.SYS_CONFIG.SITE_LISTS + this.props.eventListName + CONSTANTS.SYS_CONFIG.EVENTS_LIST_NEWFORM_PAGE
    });

    this.loadOwnerGroup().then((success) => {
      this.loadMemberGroup().then((succeed) => {
        this.LoadCurrentUserGroups();
      });
    });
   
    this.LoadRegion();
    this.LoadCategoryOptions();
  }

  private getAllEvents(): void {
    let eventItems: any = [];
    let listName: string = this.props.eventListName;
    let filterCondition = this.getFilterString();
    commonService.getEventItems(listName, CONSTANTS.SELECTCOLUMNS.EVENTS_LIST, CONSTANTS.SELECTCOLUMNS.EXPAND_EVENTS_COLS, filterCondition, CONSTANTS.ORDERBY.EVENTS, CONSTANTS.SYS_CONFIG.EVENTS_GET_ITEMS_LIMIT).then((listDataIitems: any) => {
      if (listDataIitems.length > 0) {
        //Filter the events and exclude deleted events
        let filterEvents = _.filter(listDataIitems, (p) => {
          return p.Title.startsWith("Deleted:") != true;
        });
        //get top records as configured
        filterEvents.splice((CONSTANTS.SYS_CONFIG.EVENTS_DISPLAY_ITEMS_LIMIT), (filterEvents.length - 1));
        filterEvents.forEach((listDataIitem, index) => {

          let eventItem: any = {};
          //let dateObj: Date = new Date(listDataIitem.EventDate);
          //let endDateObj: Date = new Date(listDataIitem.EndDate);
          let dateObj: Date = null;
          if (listDataIitem.fAllDayEvent) {
            dateObj = new Date(listDataIitem.EventDate);
          } else {
            dateObj = new Date(listDataIitem.FieldValuesAsText.EventDate);
          }

          let endDateObj: Date = new Date(listDataIitem.FieldValuesAsText.EndDate);
          //debugger;
          //alert(listDataIitem.ID);
          let dayOfWeek = dateObj.getDay();
          eventItem.eventId = listDataIitem.ID;
          eventItem.eventTitle = listDataIitem.Title;
          eventItem.fAllDayEvent = listDataIitem.fAllDayEvent;
          eventItem.fRecurrence = listDataIitem.fRecurrence;
          eventItem.EventDate = listDataIitem.fAllDayEvent == true ? listDataIitem.EventDate : listDataIitem.FieldValuesAsText.EventDate;
          eventItem.eventDate = this.getEventDate(dateObj);
          eventItem.eventDay = this.getDayOfWeek(dayOfWeek);
          eventItem.endDate = this.getEndDate(endDateObj);
          eventItem.MarkAsImportant = listDataIitem.Mark_x0020_as_x0020_Important;
          eventItem.DisplayURL = this.props.context.pageContext.site.absoluteUrl + "/Lists/" + listName + "/DispForm.aspx?ID=" + listDataIitem.ID;
          eventItems.push(eventItem);
        });
        this.setState({
          eventData: eventItems,
          displayNoDataMassage: true
        });
      } else {
        this.setState({
          eventData: eventItems,
          errorMessage: CONSTANTS.SYS_CONFIG.NO_DATA_FOUND_ERROR_MESSAGE,
          displayNoDataMassage: true
        });
      }
    });

  }

  private getDayOfWeek(dayOfWeek: any): string {
    return isNaN(dayOfWeek) ? null : CONSTANTS.DAY_NAMES[dayOfWeek];
  }

  private getEventDate(date: any): string {

    let eventDate: string = "";
    let monthNames = CONSTANTS.MONTH_NAMES;
    let day = date.getDate();
    let monthIndex = date.getMonth();
    let monthName = monthNames[monthIndex];
    let year = date.getFullYear();
    eventDate = `${day}-${monthName}-${year}`;
    return eventDate;
  }

  private getEndDate(date: any): string {

    let endDate: string = "";
    let monthNames = CONSTANTS.MONTH_NAMES;
    let day = date.getDate();
    let monthIndex = date.getMonth();
    let monthName = monthNames[monthIndex];
    let year = date.getFullYear();
    endDate = `${day}-${monthName}-${year}`;
    return endDate;
  }

  private receiveData(data: IEventData) {
    g_selectedRegion = data.sharedRegion.text;
    g_selectedProgramType = data.sharedProgramType.text;
    this.getAllEvents();
  }

  private getFilterString = (): string => {

    let filterString: string = "";

    let andCondition: string = " and ";
    if (g_selectedRegion != "" && g_selectedRegion != "Select") {
      filterString += filterString != "" ? andCondition : "";
      filterString += "Region eq '" + g_selectedRegion + "'";
    }
    if (g_selectedProgramType != "" && g_selectedProgramType != "Select") {
      filterString += filterString != "" ? andCondition : "";
      filterString += "Program_x0020_Type eq '" + g_selectedProgramType + "'";
    }
    return filterString;
  }

  private openEventItemView = (url: string): void => {
    window.open(url, "_blank");
  }

  public loadOwnerGroup = ():Promise<any> => {
    return new Promise((resolve,reject) => {
      commonService.getOwnersGroup().then((ownerGroup: any) => {
        this.setState({
          ownersGroup: ownerGroup.Title
        });
        resolve(true);
      });
    });
  }

  public loadMemberGroup = ():Promise<any> => {
    return new Promise((resolve,reject) => {
      commonService.getMembersGroup().then((memberGroup: any) => {
        this.setState({
          membersGroup: memberGroup.Title
        });
        resolve(true);
      });
    });
  }
  public LoadCurrentUserGroups = () => {
    commonService.getCurrentUserGroups().then((userGroups: any) => {
      userGroups.forEach((group: any, index: number) => {
        if((group["Title"] === this.state.ownersGroup) || (group["Title"] ===this.state.membersGroup)) {
          this.setState({
            isCurrentUserPresentInGroup: true
          });
        }
      });
    });
  }

  private openModal = () => {
    this.setState({
      isModalOpen: true
    });
  }

  private closeModalDialog = () => {
    this.setState({
      isModalOpen: false,
      Title: "",
      Location: "",
      eventStartDate: new Date(),  
      eventEndDate: new Date(), 
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      selectedCategoryValue: { key: "", text: "" },
      Description: "",
      isMarkAsImportant: false,
      isAllDayEvent: false,
      titleErrorMessage: "",
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      descriptionErrorMessage: "",
      locationErrorMessage: "",
      categoryErrorMessage: "",
      showMessageBar: false,
    });
  }

  private clearModalDialog = () => {
    this.setState({
      Title: "",
      Location: "",
      eventStartDate: new Date(),  
      eventEndDate: new Date(),
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      selectedCategoryValue: { key: "", text: "" },
      Description: "",
      isMarkAsImportant: false,
      isAllDayEvent: false,
      titleErrorMessage: "",
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      descriptionErrorMessage: "",
      categoryErrorMessage: "",
      locationErrorMessage: "",
    });
  }

  private onTitleChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name] : e.target.value,
      titleErrorMessage: ""
    });     
  }

  private onLocationChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name] : e.target.value,
      locationErrorMessage: ""
    });     
  }

  private onDescriptionChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name] : e.target.value,
      descriptionErrorMessage: ""
    });     
  }

  public onMarkAsImpChange(ev: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) {
    this.setState({
      isMarkAsImportant: isChecked
    });
  }

  public onAllDayEventChange(ev: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) {
    this.setState({
      isAllDayEvent: isChecked
    });
  }

  private __onchangedStartDate = (date: any): void => {  
    this.setState({ eventStartDate: date });  
  } 

  private __onchangedEndDate = (date: any): void => {  
    this.setState({ eventEndDate: date });  
  } 

  private LoadRegion = (): void => {
    let RegionOptions: any[]= [];
    commonService.GetSiteColumnChoices(CONSTANTS.SITE_COLUMN_NAME.REGION_COLS).then((RegionChoices: any) => {
      RegionOptions.push({
        key: "Select",
        text: "Select"
      });

      RegionChoices.Choices.forEach((RegionItem: any, index: number) => {
        RegionOptions.push({
          key: (index + 1),
          text: RegionItem
        });
      });

    
        this.setState({
          regionOptions: RegionOptions,
          selectedRegionValue: { key: "Select", text: "Select" }
        });
      

      this.LoadProgramType();

    });
  }

  private LoadProgramType = (): void => {
    let ProgramTypeOptions: any[] = [];
    commonService.GetSiteColumnChoices(CONSTANTS.SITE_COLUMN_NAME.PROGRAM_TYPE_COLS).then((ProgramTypeChoices: any) => {
      ProgramTypeOptions.push({
        key: "Select",
        text: "Select"
      });
      ProgramTypeChoices.Choices.forEach((ProgramTypeItem: any, index: number) => {
        ProgramTypeOptions.push({
          key: (index + 1),
          text: ProgramTypeItem
        });
      });

     
        this.setState({
          programTypeOptions: ProgramTypeOptions,
          selectedProgramTypeValue: { key: "Select", text: "Select" }
        });
      

      this.LoadSiteLocation();

    });
  }

  private LoadSiteLocation = (): void => {
    let SiteLocationOptions: any[] = [];

    commonService.getSiteLocationConfiguration(CONSTANTS.LIST_NAME.SITE_LOCATION_CONFIGURATION, CONSTANTS.SELECTCOLUMNS.SITE_LOCATION_CONFIGURATION, CONSTANTS.ORDERBY.SITE_LOCATION, CONSTANTS.SYS_CONFIG.GET_ITEMS_LIMIT).then((SiteLocationChoices: any) => {
      SiteLocationOptions.push({
        key: "Select",
        text: "Select"
      });

      let SiteLocationOptionsData: any = [];
      let queryStringParameters = new URLSearchParams(window.location.search);

      //If All region selected then append all choice
      if (queryStringParameters.get("rgn") && queryStringParameters.get("rgn").toLowerCase() == "select") {
        SiteLocationOptionsData = SiteLocationChoices;
      } else {
        SiteLocationOptionsData = _.filter(SiteLocationChoices, (p) => {
          return p.Region == this.state.selectedRegionValue.text;
        });
      }

      let allSiteLocationOption = _.filter(SiteLocationOptionsData, (p) => {
        return p.Region == "Global (Corporate)";
      });

      if (allSiteLocationOption.length > 0) {
        SiteLocationOptions.push({
          key: allSiteLocationOption[0].Id,
          text: allSiteLocationOption[0].Code
        });
        //Tried to splice the all option but it is removed from original array too
        // SiteLocationOptionsData.splice(SiteLocationOptionsData.findIndex(a => a.Id === allSiteLocationOption[0].Id), 1)

      }

      SiteLocationOptionsData.forEach((siteLicationChoice: any, index: number) => {
        //Exclude the all option to add again
        if (siteLicationChoice.Title != "All") {
          SiteLocationOptions.push({
            key: siteLicationChoice.Id,
            text: siteLicationChoice.Title + '-' + siteLicationChoice.Code
          });
        }
      });

      this.setState({
        siteLocationData: SiteLocationChoices,
        siteLocationOptions: SiteLocationOptions,
        selectedSiteLocationValue: { key: "Select", text: "Select" },
        //siteLocationDropDownError: false,
        siteLocationColl: SiteLocationChoices
      });
      //this.getDocumentData(false, this.state.folderServerRelativeUrl);
      //this.setGlobalCssChanges();
    });
  }


  private loadSiteLocationByRegion = (SelectedRegionItem: any): void => {
    let SiteLocationOptionsData: any = [];
    let SiteLocationOptions: any[] = [];
    SiteLocationOptions.push({
      key: "Select",
      text: "Select"
    });

    if (SelectedRegionItem.text != "" && SelectedRegionItem.text != "Select") {
      SiteLocationOptionsData = _.filter(this.state.siteLocationData, (p) => {
        return p.Region.toLowerCase() == SelectedRegionItem.text.toLowerCase();
      });
    } else {
      SiteLocationOptionsData = this.state.siteLocationData;
    }

    let allSiteLocationOption = _.filter(SiteLocationOptionsData, (p) => {
      return p.Region == "Global (Corporate)";
    });

    if (allSiteLocationOption.length > 0) {
      SiteLocationOptions.push({
        key: allSiteLocationOption[0].Id,
        text: allSiteLocationOption[0].Code
      });
      //Tried to splice the all option but it is removed from original array too
      // SiteLocationOptionsData.splice(SiteLocationOptionsData.findIndex(a => a.Id === allSiteLocationOption[0].Id), 1)

    }

    SiteLocationOptionsData.forEach((siteLicationChoice: any, index: number) => {
      //Exclude the all option to add again
      if (siteLicationChoice.Title != "All") {
        SiteLocationOptions.push({
          key: siteLicationChoice.Id,
          text: siteLicationChoice.Title + '-' + siteLicationChoice.Code
        });
      }
    });

    this.setState({
      siteLocationOptions: SiteLocationOptions,
      selectedSiteLocationValue: { key: "Select", text: "Select" }
    });

  }

  public onRegionChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedRegionValue: option,
        regionErrorMessage: ""
      });
      this.loadSiteLocationByRegion(option);
    } else {
      this.setState({
        ...this.state,
        selectedRegionValue: { key: "Select", text: "Select" },
      });
    }
  }

  public onProgramTypeChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedProgramTypeValue: option,
        programTypeErrorMessage: ""
      });
    } else {
      this.setState({
        ...this.state,
        selectedProgramTypeValue: { key: "Select", text: "Select" },
      });
    }
  }

  public onSiteLocationChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    //this.setState({ SingleSelect: option.key });
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedSiteLocationValue: option,
        siteLocationErrorMessage: ""
      });
    } else {
      this.setState({
        ...this.state,
        selectedSiteLocationValue: { key: "Select", text: "Select" },
      });
    }
  }

  private LoadCategoryOptions = (): void => {
    let CategoryOptions: any[]= [];
    commonService.GetCategoryChoices(this.props.eventListName,CONSTANTS.COLULMN_NAME.CATEGORY).then((CategoryChoices: any) => {

      CategoryChoices.Choices.forEach((CategoryItem: any, index: number) => {
        CategoryOptions.push({
          key: (index + 1),
          text: CategoryItem
        });
      });

        this.setState({
          categoryOptions: CategoryOptions,
        });

    });
  }


  public onCategoryOptionChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    //this.setState({ SingleSelect: option.key });
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedCategoryValue: option,
        categoryErrorMessage: ""
      });
    } else {
      this.setState({
        ...this.state,
        selectedCategoryValue: { key: "", text: "" },
      });
    }
  }

  public setAddEventItems = (title, location, startTime, endTime, category, region, programType, siteLocation, description, IsMarkAsImportant, isAllDayEvent): void => {
    if(siteLocation == "Select") {
      commonService.AddItemsInEventList(this.props.eventListName, title, location, startTime, endTime, category, region, programType, description, IsMarkAsImportant, isAllDayEvent).then((result:any) => {
        this.getAllEvents();
        this.clearModalDialog();
        this.setState({  
          message:  "Item: " + title + " - created successfully!", 
          showMessageBar: true,  
          messageType: MessageBarType.success,
        });  
        return result;
      }).catch((error:any) => {  
        this.setState({  
          message: "Item " + title + " creation failed with error: " + error,  
          showMessageBar: true,  
          messageType: MessageBarType.error  
        });  
      });  
    }else {
      commonService.AddItemsInEventListwithSiteLocation(this.props.eventListName, title, location, startTime, endTime, category, region, programType, siteLocation, description, IsMarkAsImportant, isAllDayEvent).then((result:any) => {
        this.getAllEvents();
        this.clearModalDialog();
        this.setState({  
          message:  "Item: " + title + " - created successfully!", 
          showMessageBar: true,  
          messageType: MessageBarType.success,
        });  
        return result;
      }).catch((error:any) => {  
        this.setState({  
          message: "Item " + title + " creation failed with error: " + error,  
          showMessageBar: true,  
          messageType: MessageBarType.error  
        });  
      });  
    }
  }

  public addItemsInEventsList = (): void =>  {
    let title = this.state.Title;
    let location = this.state.Location;
    let startTime = this.state.eventStartDate;
    let endTime = this.state.eventEndDate;
    let selectedCategory = this.state.selectedCategoryValue == undefined ? "Select" : this.state.selectedCategoryValue.text;
    let selectedRegion = this.state.selectedRegionValue == undefined ? "Select" : this.state.selectedRegionValue.text;
    let selectedProgramType = this.state.selectedProgramTypeValue == undefined ? "Select" : this.state.selectedProgramTypeValue.text;
    let selectedSiteLocation = this.state.selectedSiteLocationValue == undefined ? "Select" : this.state.selectedSiteLocationValue.key;
    let description = this.state.Description;
    let IsMarkAsImportant = this.state.isMarkAsImportant;
    let IsAllDayEvent= this.state.isAllDayEvent;

    this.setAddEventItems(title, location, startTime, endTime, selectedCategory, selectedRegion, selectedProgramType, selectedSiteLocation, description, IsMarkAsImportant, IsAllDayEvent);
  }

  public _onSelectEndDate = (date: Date | null | undefined): void => {
    this.setState({ eventStartDate: date});
  }

  public onFormatDate = (date?: Date): string => {
    return !date ? '' : date.toLocaleDateString();
    //return date ? '' : (date.getMonth() + 1) + '/' + date.getDate() + '/' + (date.getFullYear());
  }

  public isEventFormValidate = () => {
    let validation: boolean = true;
    if(this.state.Title == '') {
      validation = false;
      this.setState({
        titleErrorMessage: "Please enter a event title.",
      });
    }
    if(this.state.Location == '') {
      validation = false;
      this.setState({
        locationErrorMessage: "Please enter a event location.",
      });
    }
    if(this.state.selectedCategoryValue.key == "" && this.state.selectedCategoryValue.text == "" || this.state.selectedCategoryValue.length == 0) {
      validation = false;
      this.setState({
        categoryErrorMessage: "Please select a category.",
      });
    }
    if(this.state.selectedRegionValue.key == "Select" && this.state.selectedRegionValue.text == "Select" || this.state.selectedRegionValue.length == 0) {
      validation = false;
      this.setState({
        regionErrorMessage: "Please select a region.",
      });
    }
    if(this.state.selectedProgramTypeValue.key == "Select" && this.state.selectedProgramTypeValue.text == "Select" || this.state.selectedProgramTypeValue.length == 0) {
      validation = false;
      this.setState({
        programTypeErrorMessage: "Please select a program type.",
      });
    }
    if(this.state.Description == '') {
      validation = false;
      this.setState({
        descriptionErrorMessage: "Please enter a announcement description.",
      });
    }
   
    if(validation) {
        this.addItemsInEventsList();
    }
  }


  public render(): React.ReactElement<ISynopsysEventsProps> {
    return (
      //<div className="EventsMainContainer">
      <div className={styles.synopsysEvents} >
        <div>

          <Stack horizontal tokens={outerStackTokens} className={styles.stackclass}>
            <Stack.Item grow={8} >
              <h1 className={styles.webpartHeader}>{this.props.webpartTitle}</h1>
            </Stack.Item>
            { this.state.error == "" ?
              <div className="iconDiv">
                <Stack.Item>
                  <div className={styles.seeAllDiv}>
                    <div className={styles.seeAllEvents}>
                      <a href={this.state.seeAllURL} data-interception="off" target="_blank">{this.props.webpartLabel}</a>
                    </div>
                  </div>
                </Stack.Item>
                { this.state.isCurrentUserPresentInGroup == true ?
                  <Stack.Item>
                    <div className={styles.newItemDiv}>
                      <div className={styles.addNewEvent} onClick={this.openModal}>
                        <a><i className="ms-Icon ms-Icon--CircleAddition" aria-hidden="true"></i></a>
                      </div>
                    </div>
                  </Stack.Item>
                  : ""
                }
              </div>
              : ""
            } 
          </Stack>
        </div>
        {
          this.state.error == "" ?
            <div className={styles.container}>
              <Stack tokens={outerStackTokens} className={styles.stackclass}>
                {/*  <Stack.Item grow >

                  {this.state.error}
                  <div className={styles.seeAllDiv}>
                    <div className={styles.seeAllEvents}>
                      <a href={this.state.seeAllURL} data-interception="off" target="_blank">{this.props.webpartLabel}</a>
                    </div>
                    <a className={styles['next-arrow']} href={this.state.seeAllURL} data-interception="off" target="_blank">
                      <i className="ms-Icon ms-Icon--ChevronRight" aria-hidden="true"></i>
                    </a>
                    <hr className={styles.hrStyle}></hr>
                  </div>

                </Stack.Item>
                */}
                {this.state.eventData.length > 0 ?
                  <React.Fragment>
                    {this.state.eventData.map((dataItem: any, index: number) => (
                      <React.Fragment>
                        <Stack horizontal tokens={innerStackTokens}  className={CONSTANTS.SYS_CONFIG.EVENTS_DISPLAY_ITEMS_LIMIT == (index+1) ? "eventDivLast" : "eventDiv"} >


                          {
                            dataItem.fRecurrence == false ?

                              <Stack.Item grow disableShrink className="eventDayandDate">
                                <div className="eventDate">
                                  {dataItem.eventDate.split('-')[1]}
                                </div>
                                <hr className="newLine" />
                                <div className="eventDate">
                                  {dataItem.eventDate.split('-')[0]}
                                </div>
                              </Stack.Item>
                              :
                              <Stack.Item grow disableShrink className="eventDayandDate">

                                <div className="eventDate">
                                  {dataItem.eventDate.split('-')[1] + " " + dataItem.eventDate.split('-')[0]}
                                </div>
                                <hr className="newLine" />
                                <div className="eventDate">
                                  {dataItem.endDate.split('-')[1] + " " + dataItem.endDate.split('-')[0]}
                                </div>

                              </Stack.Item>
                          }

                          <Stack.Item grow className="eventSection">

                            {/* <div className={styles.eventDate}> */}
                            <div title={dataItem.MarkAsImportant == true ? "Important Event":""} className={dataItem.MarkAsImportant == true ? "eventDateandDayIMP" : "eventDateandDay"}>
                              {dataItem.fRecurrence == false ? (
                                (dataItem.fAllDayEvent == true) ? `${(dataItem.EventDate) ? new Date(dataItem.EventDate).toLocaleString("en-US", {
                                  weekday: 'long',
                                  month: 'long',
                                  day: 'numeric',
                                  year: 'numeric'

                                }) : ''} | All Day`
                                  : (dataItem.EventDate) ? new Date(dataItem.EventDate).toLocaleDateString("en-US", {
                                    weekday: 'long',
                                    month: 'long',
                                    year: 'numeric',
                                    day: 'numeric',
                                    hour: 'numeric',
                                    minute: 'numeric',

                                  }) : ''
                              ) : `${new Date(dataItem.EventDate).toLocaleDateString("en-US", {
                               
                                month: 'long',
                                day: 'numeric',
                                year: 'numeric'

                              }) } - ${new Date(dataItem.endDate).toLocaleDateString("en-US", {
                               
                                month: 'long',
                                day: 'numeric',
                                year: 'numeric'

                              }) }`

                              }

                              {
                                dataItem.fRecurrence == true ?
                                  <a><i className={dataItem.MarkAsImportant == true ? "ms-Icon ms-Icon--SyncOccurence eventRepeatIconIMP":"ms-Icon ms-Icon--SyncOccurence eventRepeatIcon"} aria-hidden="true"></i></a>
                                  : ""
                              }
                            </div>
                            <div><a target='_blank' title={dataItem.eventTitle} onClick={() => this.openEventItemView(dataItem.DisplayURL)} className="eventTitle">
                              {dataItem.eventTitle.length > CONSTANTS.SYS_CONFIG.EVENTS_TITLE_CHARACTER_LENGTH ? dataItem.eventTitle.substring(0, CONSTANTS.SYS_CONFIG.EVENTS_TITLE_CHARACTER_LENGTH) + "..." : dataItem.eventTitle}
                            </a></div>
                          </Stack.Item>



                        </Stack>

                      </React.Fragment>
                    ))}
                  </React.Fragment>
                  :
                  <Stack>
                    <div className={styles.noDataFound}>  {this.state.displayNoDataMassage == true ? this.state.errorMessage : ""}</div>
                  </Stack>
                }
              </Stack>
            </div>
            : <div className={styles.errorDiv}> {this.state.error}</div>
        }


          <Stack>
            <Stack.Item>
              <Modal
                titleAriaId="Modal"
                isOpen={this.state.isModalOpen}
                onDismiss={this.closeModalDialog}
                isBlocking={false}
                containerClassName={styles.container}
              >
                <div className="modalheader">
                  <span id="Modal" className="newItemSpan">Add New Event</span>
                  <IconButton
                    className="modalCloseIcon"
                    iconProps={cancelIcon}
                    ariaLabel="Close popup modal"
                    onClick={this.closeModalDialog}
                  />
                </div>

                {/* <hr></hr> */}
                <div className="eventBodyDiv">
                  <div className="EventModalbody">
                    <Stack className="addItemsRow1">

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <TextField 
                            id="Title"
                            name="Title"
                            label="Title" 
                            required={true}
                            placeholder="Enter value here"
                            value={this.state.Title}
                            onChange={this.onTitleChange}
                            autoComplete="off"
                            errorMessage={this.state.titleErrorMessage}
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <TextField 
                            id="Location"
                            name="Location"
                            label="Location" 
                            required={true}
                            placeholder="Enter value here"
                            value={this.state.Location}
                            onChange={this.onLocationChange}
                            autoComplete="off"
                            errorMessage={this.state.locationErrorMessage}
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <Label required={true}>Start Time</Label>
                          <DateTimePicker 
                            dateConvention={DateConvention.DateTime}
                            timeConvention={TimeConvention.Hours12} 
                            timeDisplayControlType={TimeDisplayControlType.Dropdown}
                            showLabels={false}
                            showGoToToday={true}
                            isMonthPickerVisible={false}
                            formatDate={this.onFormatDate}
                            value={this.state.eventStartDate}
                            onChange={this.__onchangedStartDate} 
                            
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <Label required={true}>End Time</Label>
                          <DateTimePicker
                            dateConvention={DateConvention.DateTime}
                            timeConvention={TimeConvention.Hours12} 
                            timeDisplayControlType={TimeDisplayControlType.Dropdown}
                            showLabels={false}
                            showGoToToday={true}
                            isMonthPickerVisible={false}
                            formatDate={this.onFormatDate}
                            value={this.state.eventEndDate}
                            onChange={this.__onchangedEndDate} 
                            
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <TextField 
                            id="Description"
                            name="Description"
                            label="Description" 
                            multiline={true}
                            required={true}
                            placeholder="Enter value here"
                            value={this.state.Description}
                            onChange={this.onDescriptionChange}
                            autoComplete="off"
                            errorMessage={this.state.descriptionErrorMessage}
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <ComboBox
                            label="Category"
                            allowFreeform={true}
                            //placeholder="Select a category or enter your own choice..."
                            required={true}
                            autoComplete={'off'}
                            options={this.state.categoryOptions}
                            onChange={this.onCategoryOptionChange}
                            selectedKey={this.state.selectedCategoryValue ? this.state.selectedCategoryValue.key : ""}
                            errorMessage={this.state.categoryErrorMessage}
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <ComboBox
                            label="Region"
                            required={true} 
                            allowFreeform={true}
                            autoComplete={'on'}
                            options={this.state.regionOptions}
                            onChange={this.onRegionChange}
                            selectedKey={this.state.selectedRegionValue ? this.state.selectedRegionValue.key : "Select"}
                            errorMessage={this.state.regionErrorMessage}
                          />
                        </div>
                      </Stack.Item>
                      
                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <ComboBox
                            label="Program Type"
                            required={true} 
                            allowFreeform={true}
                            autoComplete={'on'}
                            options={this.state.programTypeOptions}
                            onChange={this.onProgramTypeChange}
                            selectedKey={this.state.selectedProgramTypeValue ? this.state.selectedProgramTypeValue.key : "Select"}
                            errorMessage={this.state.programTypeErrorMessage}
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <ComboBox
                            label="Site Location Code"
                            required={false} 
                            allowFreeform={true}
                            autoComplete={'on'}
                            options={this.state.siteLocationOptions}
                            onChange={this.onSiteLocationChange}
                            selectedKey={this.state.selectedSiteLocationValue ? this.state.selectedSiteLocationValue.key : "Select"}
                            //errorMessage={this.state.siteLocationErrorMessage}
          
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <Label className="allDayEventLabel">All Day Event</Label>
                          <Checkbox 
                            label="Make this an all-day activity that doesn't start or end at a specific hour." 
                            checked={this.state.isAllDayEvent}
                            onChange={this.onAllDayEventChange.bind(this)} 
                            className="checkbox"
                          />
                        </div>
                      </Stack.Item>
                      <Stack.Item>
                        <div className="metadataFieldDiv">
                          <Label className="markAsImpLabel">Mark As Important</Label>
                          <Checkbox 
                            label="Yes" 
                            checked={this.state.isMarkAsImportant}
                            onChange={this.onMarkAsImpChange.bind(this)} 
                            className="checkbox"
                          />
                        </div>
                      </Stack.Item>
                      

                    </Stack>
                  </div>
                </div>
                
                <Stack>
                  <Stack.Item>
                    {  
                      this.state.showMessageBar ?  
                      <div className="form-group">  
                        <Stack {...verticalStackProps}>  
                          <MessageBar className="eventMessageBarDiv" messageBarType={this.state.messageType}>{this.state.message}</MessageBar>  
                        </Stack>  
                      </div>  
                      :  
                      null  
                    }  
                  </Stack.Item>
                </Stack>
                <div className="modalFotter">
                  <PrimaryButton className="saveButton" 
                    onClick={this.isEventFormValidate}
                    text="Save" />
                  <DefaultButton className="closeButton" onClick={this.closeModalDialog} text="Cancel" />
                </div>
              </Modal>
            
            </Stack.Item>
          </Stack>
          
      </div>
    
    );
  }
}
