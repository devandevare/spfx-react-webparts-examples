import * as React from 'react';
import styles from './Announcements.module.scss';
import { IAnnouncementsProps } from './IAnnouncementsProps';
import dataService from "../../../Common/DataService";
import commonMethods from "../../../Common/CommonMethods";
import CONSTANTS from "../../../Common/Constants";
import { Checkbox, ComboBox, DefaultButton, IComboBox, IComboBoxOption, IconButton, IIconProps, IStackProps, IStackTokens, Label, MessageBar, MessageBarType, Modal, PrimaryButton, Stack, TextField } from 'office-ui-fabric-react'; //'@fluentui/react';
import Tooltip from '../../CommonComponents/Tooltip';
import * as _ from 'lodash';
import { RxJsEventEmitter } from '../../RxJsEventEmitter/RxJsEventEmitter';
import Dropzone, { DropEvent, FileRejection } from 'react-dropzone';


let g_selectedRegion: string = "";
let g_selectedProgramType: string = "";

export interface IAnnouncementsState {
  ID: any;
  Title: any;
  Description: any;
  MarkAsImportant: any;
  Attachments: any;
  Created: any;
  Announcements: any;
  error: string;
  seeAllURL: string;
  newFormURL: string;
  DisplayAnnouncement: any;
  selectedRegionValue: any;
  selectedProgramTypeValue: any;
  selectedSiteLocationValue: any;
  errorMessage: string;
  displayNoDataMassage: boolean;
  userName: any;
  userId: any;
  isCurrentUserPresentInGroup: boolean;
  ownersGroup: string;
  membersGroup: string;
  isModalOpen: boolean;
  regionErrorMessage: string;
  programTypeErrorMessage: string;
  siteLocationErrorMessage: string;
  titleErrorMessage: string;
  descriptionErrorMessage: string;
  showMessageBar: boolean;
  messageType?: MessageBarType;    
  message?: string; 
  regionOptions: any[];
  programTypeOptions: any[];
  siteLocationOptions: any;
  siteLocationData: any[];
  siteLocationColl: any[];
  siteLocationDropDownError: boolean;
  programTypeDropDownError: boolean;
  isMarkAsImportant: boolean;
  file: any[];
  fileName: any;
  fileSize: any;
}

export interface IEventData {
  sharedRegion: any;
  sharedProgramType: any;
}

const verticalStackProps: IStackProps = {  
  styles: { root: { overflow: 'hidden', width: '100%' } },  
  tokens: { childrenGap: 20 }  
}; 
const outerStackTokens: IStackTokens = { childrenGap: 5, padding: 10 };
const innerStackTokens: IStackTokens = { childrenGap: 10 };
const innerStackTokens1: IStackTokens = { childrenGap: 0 };
const cancelIcon: IIconProps = { iconName: 'Cancel' };
const commonService = new dataService();
const commonMethod = new commonMethods();
export default class Announcements extends React.Component<IAnnouncementsProps, IAnnouncementsState> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  constructor(props: IAnnouncementsProps, state: IAnnouncementsState) {
    super(props);
    this.state = {
      ID: "",
      Title: "",
      Description: "",
      MarkAsImportant: "",
      Attachments: [],
      Created: "",
      Announcements: [],
      error: '',
      seeAllURL: '',
      newFormURL: '',
      DisplayAnnouncement: [],
      selectedRegionValue: [],
      selectedProgramTypeValue: [],
      selectedSiteLocationValue: [],
      errorMessage: "",
      displayNoDataMassage: false,
      userName: "",
      userId: "",
      isCurrentUserPresentInGroup: false,
      ownersGroup: "",
      membersGroup: "",
      isModalOpen: false,
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      titleErrorMessage: "",
      descriptionErrorMessage: "",
      showMessageBar: false,
      regionOptions: [],
      programTypeOptions: [],
      siteLocationOptions: [],
      siteLocationData: [],
      siteLocationColl: [],
      siteLocationDropDownError: false,
      programTypeDropDownError: false,
      isMarkAsImportant: false,
      file: [],
      fileName: "",
      fileSize: "",
    };
    this.eventEmitter.on(CONSTANTS.CONNECTED_WP.SHARE_DATA, this.receiveData.bind(this));
  }


  private receiveData(data: IEventData) {
    g_selectedRegion = data.sharedRegion.text;
    g_selectedProgramType = data.sharedProgramType.text;
    this.LoadAnnouncements();
  }

  public dropDownValidation = async (listTitle: string) => {
    if (listTitle == undefined || listTitle == ' ') {
      this.setState({
        error: CONSTANTS.SYS_CONFIG.SELECT_LIST
      });
    }
    else {

      let isValidListColumns = await commonMethod.isValidListColumns(listTitle, CONSTANTS.LIST_VALIDATION_COLUMNS.ANNOUNCEMENTS);
      //alert("isValidListColumns:- " + isValidListColumns);
      //Check all required fields available in selected list.
      if (isValidListColumns) {
        this.LoadAnnouncements();
      } else {
        this.setState({
          error: CONSTANTS.SYS_CONFIG.ANNOUNCEMENTS_LIST_NOT_MATCH
        });
      }
      return;

    }
  }

  public async componentDidMount() {

    this.dropDownValidation(this.props.listTitle);
    this.setState({
      seeAllURL: this.props.seeAllURL == "" ? this.props.siteURL + CONSTANTS.SYS_CONFIG.SITE_LISTS + this.props.listTitle + CONSTANTS.SYS_CONFIG.ANNOUNCEMENT_LIST_PAGE : this.props.seeAllURL,
      newFormURL: this.props.siteURL + CONSTANTS.SYS_CONFIG.SITE_LISTS + this.props.listTitle + CONSTANTS.SYS_CONFIG.ANNOUNCEMENT_LIST_NEWFORM_PAGE
    });

    this.loadOwnerGroup().then((success) => {
      this.loadMemberGroup().then((succeed) => {
        this.LoadCurrentUserGroups();
      });
    });
    
    this.LoadRegion();
  }

  private getFilterString = (): string => {

    let filterString: string = "";
    let andCondition: string = " and ";
    if (g_selectedRegion != "" && g_selectedRegion != "Select") {
      filterString = "Region eq '" + g_selectedRegion + "'";
    }
    if (g_selectedProgramType != "" && g_selectedProgramType != "Select") {
      filterString += filterString != "" ? andCondition : "";
      filterString += "Program_x0020_Type eq '" + g_selectedProgramType + "'";
    }
    return filterString;
  }

  public LoadAnnouncements = () => {
    let _announcements = [];
    let filterCondition = this.getFilterString();
    commonService.getAnnouncements(this.props.listTitle, CONSTANTS.ORDERBY.ANNOUNCEMENTS, filterCondition, CONSTANTS.SYS_CONFIG.ANNOUNCEMENT_GET_ITEMS_LIMIT).then((Announcement: any) => {

      if (Announcement.length > 0) {
        Announcement.forEach((AnnouncementItem: any, index: number) => {
          let dateObj: Date = new Date(AnnouncementItem.Created);
          _announcements.push({
            Id: AnnouncementItem.ID,
            Title: AnnouncementItem.Title,
            Created: (AnnouncementItem.Created) ? new Date(AnnouncementItem.Created).toLocaleDateString("en-US", {
              day: 'numeric',
              month: 'short',
              year: 'numeric'
            }) : '',
            MarkAsImportant: AnnouncementItem.Mark_x0020_As_x0020_Important,
            Description: AnnouncementItem.Description,
            Attachments: AnnouncementItem.Attachments,
            announcementtDate: this.getAnnouncemetDate(dateObj),
            DisplayURL: this.props.context.pageContext.site.absoluteUrl + "/Lists/" + this.props.listTitle + "/DispForm.aspx?ID=" + AnnouncementItem.ID //+ "&IsDlg=1"
          });
        });

        this.setState({
          Announcements: _announcements,
          displayNoDataMassage: true
        });
      } else {
        this.setState({
          Announcements: _announcements,
          errorMessage: CONSTANTS.SYS_CONFIG.NO_DATA_FOUND_ERROR_MESSAGE,
          displayNoDataMassage: true
        });

      }
    });
  }


  private getAnnouncemetDate(date: any): string {
    let announcementDate: string = "";
    let monthNames = CONSTANTS.MONTH_NAMES;
    let day = date.getDate();

    let monthIndex = date.getMonth();
    let monthName = monthNames[monthIndex];
    let year = date.getFullYear();
    announcementDate = `${day}-${monthName}-${year}`;
    return announcementDate;
  }

  public getAttachment = (itemId): void => {

    commonService.getAttachment(this.props.listTitle, CONSTANTS.SELECTCOLUMNS.ANNOUNCEMENTS_LIST, CONSTANTS.SELECTCOLUMNS.EXPAND_ANNOUNCEMENT_COLS, itemId).then((result: any) => {
    
      if (result.AttachmentFiles.length > 0) {
        let url: string = `${this.props.siteURL}/_layouts/download.aspx?sourceurl=${result.AttachmentFiles[0].ServerRelativeUrl}`;
        window.open(url, "_blank");
        return false;
      }
    });

  }

  private openAnnouncementItemView = (url: string): void => {
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

        if((group["Title"] === this.state.ownersGroup) || (group["Title"] === this.state.membersGroup)) {
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
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      Description: "",
      isMarkAsImportant: false,
      fileName: "",
      titleErrorMessage: "",
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      descriptionErrorMessage: "",
      showMessageBar: false,
    });
  }

  private clearModalDialog = () => {
    this.setState({
      Title: "",
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      Description: "",
      isMarkAsImportant: false,
      fileName: "",
      titleErrorMessage: "",
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      descriptionErrorMessage: "",
    });
  }

  private onTitleChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name] : e.target.value,
      titleErrorMessage: ""
    });     
  }

  private onDescriptionChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name] : e.target.value,
      descriptionErrorMessage: ""
    });     
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
        siteLocationDropDownError: false,
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

  public setAddAnnouncementItems = (title, region, programType, siteLocation, description, IsMarkAsImportant, filename, file): void => {
    if(siteLocation == "Select") {
      commonService.AddItemsInAnnouncementList(this.props.listTitle, title, region, programType, description, IsMarkAsImportant, filename, file).then((result:any) => {
        this.LoadAnnouncements();
        this.clearModalDialog();
        this.setState({  
          message: "Item: " + title + " - created successfully!",  
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
      commonService.AddItemsInAnnouncementListwithSiteLocation(this.props.listTitle, title, region, programType, siteLocation, description, IsMarkAsImportant, filename, file).then((result:any) => {
        this.LoadAnnouncements();
        this.clearModalDialog();
        this.setState({  
          message: "Item: " + title + " - created successfully!",  
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

  public addItemsInAnnouncementList = (): void =>  {
    let title = this.state.Title;
    let selectedRegion = this.state.selectedRegionValue == undefined ? "Select" : this.state.selectedRegionValue.text;
    let selectedProgramType = this.state.selectedProgramTypeValue == undefined ? "Select" : this.state.selectedProgramTypeValue.text;
    let selectedSiteLocation = this.state.selectedSiteLocationValue == undefined ? "Select" : this.state.selectedSiteLocationValue.key;
    let description = this.state.Description;
    let IsMarkAsImportant = this.state.isMarkAsImportant;
    let filename = this.state.fileName;
    let file = this.state.file[0];

    this.setAddAnnouncementItems(title, selectedRegion, selectedProgramType, selectedSiteLocation, description, IsMarkAsImportant, filename, file);
  }

  public onMarkAsImpChange(ev: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean) {
    this.setState({
      isMarkAsImportant: isChecked
    });
  }

  private _onFileDrop=(files: File[], fileRejection:FileRejection[], event: DropEvent):void=> {​​​​​​
    try{​​​​​​
      //g_attachErrorClass = 'noErrMessage';
      this.setState({​​​​​​file:[], fileName:"", fileSize:"", errorMessage:""}​​​​​​);
      var fileName : string;
      var fileSize : any;
      var file : any = [];
      if(files.length > 0){​​​​​​
        //to check file size in more than 10MB
        if(files[0].size> 9999999){​​​​​​
          files.map(selectedFile => {​​​​​​
            fileName = selectedFile.name,
            fileSize = this._formatFileSize(selectedFile.size);
          }​​​​​​);
          this.setState({​​​​​​
            file: files,
            fileName: fileName,
            fileSize: fileSize,
            //errorMessage: "File size is too large. Please upload a file less than 10MB."
          }​​​​​​);
        }​​​​​​
        else{​​​​​​
          files.map(selectedFile => {​​​​​​
            fileName = selectedFile.name,
            fileSize = this._formatFileSize(selectedFile.size);
          }​​​​​​);
          this.setState({​​​​​​
            file: files,
            fileName: fileName,
            fileSize: fileSize
          }​​​​​​);
        }​​​​​​         
      }​​​​​​
    
    }​​​​​​catch(error){​​​​​​
      console.log(error);
    }​​​​​​
  }​​​​​​

  ​private _formatFileSize = (fileSize: number): any => {
    try{
      if (fileSize == 0) return '0 Bytes';
      var k = 1000;
      var dm = 2;
      var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
      var i = Math.floor(Math.log(fileSize) / Math.log(k));
      return parseFloat((fileSize / Math.pow(k,i)).toFixed(dm)) + ' ' + sizes[i];
    }catch(error) {
      console.log(error);
    }
  }

  public isAnnouncementFormValidate = () => {
    let validation: boolean = true;
    if(this.state.Title == '') {
      validation = false;
      this.setState({
        titleErrorMessage: "Please enter a announcement title.",
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
        this.addItemsInAnnouncementList();
    }
  }


  public render(): React.ReactElement<IAnnouncementsProps> {
    return (
      
      <div className={styles.announcements}>
        <div>
          {/* <h1 className={styles.announcementHeader}>{this.props.webpartTitle}</h1>*/}
          <Stack horizontal tokens={outerStackTokens} className={styles.stackclass}>
            <Stack.Item grow={8} >
              <h1 className={"AnnouncementWebpartHeader"}>{this.props.webpartTitle}</h1>
            </Stack.Item>
            { this.state.error == "" ?
              <div className="iconDiv">
                <Stack.Item>
                  <div className={styles.seeAllDiv}>
                    <div className={styles.seeAllAnnouncements}>
                      <a href={this.state.seeAllURL} data-interception="off" target="_blank">{this.props.webpartLabel}</a>
                    </div>
                  </div>
                </Stack.Item>
                { this.state.isCurrentUserPresentInGroup == true ?
                  <Stack.Item>
                    <div className={styles.newItemDiv}>
                      <div className={styles.addNewAnnouncement} onClick={this.openModal}>
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
              <Stack tokens={outerStackTokens} className={styles.containerParentDiv}>


                {this.state.Announcements && this.state.Announcements.length > 0 ?
                  this.state.Announcements.map((item, key) => {
                    return (
                      <div className={styles.announcementMainDiv} key={key}>

                        <Stack horizontal tokens={innerStackTokens1} className={`${styles.announcementDiv} ${styles.stackclass}`} >
                          <Stack.Item disableShrink className={styles.announcementDateDiv}>

                            <div title={item.MarkAsImportant == true ? "Important Announcement":""} className={`announcementDate ${item.MarkAsImportant == true ? "announcementDayandDateIMP" : "announcementDayandDate"}`}>
                              <div className={item.MarkAsImportant == true ? styles.eventDateIMP : styles.eventDate}>
                                {item.announcementtDate.split('-')[1]}
                              </div>
                              {/*<hr className={styles.newLine} />*/}
                              <div className={item.MarkAsImportant == true ? styles.eventDateIMP : styles.eventDate}>
                                {item.announcementtDate.split('-')[0]}
                              </div>
                            </div>

                          </Stack.Item>
                          <Stack.Item grow className={styles.announcementTitleDescDiv}>



                            <a target='_blank' title={item.Title} className={styles.titleLabel} onClick={() => this.openAnnouncementItemView(item.DisplayURL)}  >
                              {item.Title.length > CONSTANTS.SYS_CONFIG.ANOUNCEMENT_TITLE_CHARACTER_LENGTH ? item.Title.substring(0, CONSTANTS.SYS_CONFIG.ANOUNCEMENT_TITLE_CHARACTER_LENGTH) + "..." : item.Title}
                            </a>
                            {item.Attachments == true ?
                              <a className={styles.attachIcon} onClick={() => this.getAttachment(item.Id)} ><i className="ms-Icon ms-Icon--Attach" aria-hidden="true"></i></a> : ""}
                            {/*<span className={styles.impSpan}>{item.MarkAsImportant == true ? "IMP" : ""}</span>*/}

                            <React.Fragment>
                              {item.Description && item.Description.length > CONSTANTS.SYS_CONFIG.ANNOUNCEMENT_DESCRIPTION_LENGTH ?
                                <Tooltip lenght={CONSTANTS.SYS_CONFIG.ANNOUNCEMENT_DESCRIPTION_LENGTH} value={item.Description} /> :
                                <div className={styles.announcementDescription}>{item.Description}</div>
                              }
                            </React.Fragment>

                          </Stack.Item>

                        </Stack>

                      </div>
                    );
                  })
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
                  <span id="Modal" className="newItemSpan">Add New Announcement</span>
                  <IconButton
                    className="modalCloseIcon"
                    iconProps={cancelIcon}
                    ariaLabel="Close popup modal"
                    onClick={this.closeModalDialog}
                  />
                </div>

                {/* <hr></hr> */}
                <div className="AnnouncementBodyDiv">
                  <div className="AnnouncementModalbody">
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
                          <Label className="markAsImpLabel">Mark As Important</Label>
                          <Checkbox 
                            label="Yes" 
                            checked={this.state.isMarkAsImportant}
                            onChange={this.onMarkAsImpChange.bind(this)} 
                            className="checkbox"
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <Label className={"uploadDocumentLabel"}>Attachments</Label>
                        <div className={"dropzoneDiv"}>
                          <Dropzone onDrop={this._onFileDrop} noDragEventsBubbling={true} multiple={false} >
                            {({getRootProps, getInputProps}) => (
                              <section>
                                <div {...getRootProps()}>
                                    <input {...getInputProps()} />
                                      <div className={"fileUpload cssMarginBottom"} title={this.state.fileName? this.state.fileName : "No file Chosen"}>
                                        {/* <DefaultButton text="Choose File"></DefaultButton> */}
                                        <Label className="addAttachments">{this.state.fileName? this.state.fileName : "Add attachments"}</Label>
                                      </div>
                                </div>
                              </section>
                            )}
                          </Dropzone>
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
                          <MessageBar className="messageBarDiv" messageBarType={this.state.messageType}>{this.state.message}</MessageBar>  
                        </Stack>  
                      </div>  
                      :  
                      null  
                    }  
                  </Stack.Item>
                </Stack>
                <div className="modalFotter">
                  <PrimaryButton className="saveButton" 
                    onClick={this.isAnnouncementFormValidate}
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
