import * as React from 'react';
import styles from './PictureGallery.module.scss';
import { IPictureGalleryProps } from './IPictureGalleryProps';
import dataService from "../../../Common/DataService";
import CONSTANTS from "../../../Common/Constants";
import { escape } from '@microsoft/sp-lodash-subset';
import ImageGallery from 'react-image-gallery';
import "react-image-gallery/styles/css/image-gallery.css";
//import { Label } from '@microsoft/office-ui-fabric-react-bundle';
import { ComboBox, DefaultButton, IComboBox, IComboBoxOption, IconButton, IIconProps, IStackProps, IStackTokens, Label, MessageBar, MessageBarType, Modal, PrimaryButton, Stack, TextField } from 'office-ui-fabric-react'; //'@fluentui/react';
import * as _ from 'lodash';
import { RxJsEventEmitter } from '../../RxJsEventEmitter/RxJsEventEmitter';
import commonMethods from "../../../Common/CommonMethods";
import Dropzone, { DropEvent, FileRejection } from 'react-dropzone';

let g_selectedRegion: string = "";
let g_selectedProgramType: string = "";

export interface IPictureGalleryState {
  galleryImages: string[];
  filteredImages: any[];
  error: string;
  showIndex: boolean;
  showBullets: boolean;
  infinite: boolean;
  showThumbnails: boolean;
  showFullscreenButton: boolean;
  showGalleryFullscreenButton: boolean;
  showPlayButton: boolean;
  showGalleryPlayButton: boolean;
  showNav: boolean;
  isRTL: boolean;
  slideDuration: number;
  slideInterval: number;
  slideOnThumbnailOver: boolean;
  thumbnailPosition: string;
  useWindowKeyDown: boolean;
  selectedRegionValue: any;
  selectedProgramTypeValue: any;
  seeAllURL: string;
  showVideo: any;
  isCurrentUserPresentInGroup: boolean;
  ownersGroup: string;
  membersGroup: string;
  newFormURL: string;
  libraryGUID: string;
  listGUID: string;
  isImageModalOpen: boolean;
  isVideoModalOpen: boolean;
  isWarningModalOpen: boolean;
  file: any[];
  fileName: any;
  fileSize: any;
  errorMessage: string;
  regionOptions: any[];
  programTypeOptions: any[];
  siteLocationOptions: any;
  selectedSiteLocationValue: any;
  siteLocationColl: any[];
  siteLocationDropDownError: boolean;
  programTypeDropDownError: boolean;
  siteLocationData: any[];
  fileNameErrorMessage: any;
  regionErrorMessage: any;
  programTypeErrorMessage: any;
  siteLocationErrorMessage: any;
  showMessageBar: boolean;    
  messageType?: MessageBarType;    
  message?: string; 
  Title: any;
  videoUrl: any;
  ContentTypeOptions: any[];
  titleErrorMessage: any;
  videoUrlErrorMessage: any;
  caption: any;
  captionErrorMessage: any;
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
const cancelIcon: IIconProps = { iconName: 'Cancel' };
const commonService = new dataService();
const commonMethod = new commonMethods();
export default class PictureGallery extends React.Component<IPictureGalleryProps, IPictureGalleryState> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  constructor(props: IPictureGalleryProps, state: IPictureGalleryState) {
    super(props);
    this.state = ({
      galleryImages: [],
      filteredImages: [],
      error: '',
      showIndex: false,
      showBullets: false,
      infinite: true,
      showThumbnails: true,
      showFullscreenButton: true,
      showGalleryFullscreenButton: true,
      showPlayButton: true,
      showGalleryPlayButton: true,
      showNav: true,
      isRTL: false,
      slideDuration: 450,
      slideInterval: 2000,
      slideOnThumbnailOver: false,
      thumbnailPosition: "bottom",
      useWindowKeyDown: true,
      selectedRegionValue: [],
      selectedProgramTypeValue: [],
      seeAllURL: "",
      showVideo: {},
      isCurrentUserPresentInGroup: false,
      ownersGroup: "",
      membersGroup: "",
      newFormURL: "",
      libraryGUID: "",
      listGUID: "",
      isImageModalOpen: false,
      isVideoModalOpen: false,
      isWarningModalOpen: false,
      file: [],
      fileName: "",
      fileSize: "",
      errorMessage: "",
      regionOptions: [],
      programTypeOptions: [],
      siteLocationOptions: [],
      selectedSiteLocationValue: [],
      siteLocationColl: [],
      siteLocationDropDownError: false,
      programTypeDropDownError: false,
      siteLocationData: [],
      fileNameErrorMessage: "",
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      showMessageBar: false,
      Title: "",
      videoUrl: "",
      ContentTypeOptions: [],
      titleErrorMessage: "",
      videoUrlErrorMessage: "",
      caption: "",
      captionErrorMessage: "",
    });
    this._renderVideo = this._renderVideo.bind(this);
    this.eventEmitter.on(CONSTANTS.CONNECTED_WP.SHARE_DATA, this.receiveData.bind(this));
  }

  private receiveData(data: IEventData) {
    debugger;
    g_selectedRegion = data.sharedRegion.text;
    g_selectedProgramType = data.sharedProgramType.text;

    if (this.props.webpartType == "PictureGallery") {
      this.LoadImages();
    } else {
      this.LoadVideos();
    }
  }

  public dropDownValidation = async (listTitle: string) => {
    if (listTitle == undefined || listTitle == ' ') {
      this.setState({
        error: CONSTANTS.SYS_CONFIG.SELECT_LIST
      });
    }
    else {

      //Check all required fields available in selected list.
      let isValidListColumns: boolean;
      if (this.props.webpartType == "PictureGallery") {
        isValidListColumns = await commonMethod.isValidListColumns(listTitle, CONSTANTS.LIST_VALIDATION_COLUMNS.PICTURE);
      } else {
        isValidListColumns = await commonMethod.isValidListColumns(listTitle, CONSTANTS.LIST_VALIDATION_COLUMNS.VIDEO);
      }


      if (isValidListColumns) {
        if (this.props.webpartType == "PictureGallery") {
          this.LoadImages();
        } else {
          this.LoadVideos();
        }
      } else {
        this.setState({
          error: this.props.webpartType == "PictureGallery" ? CONSTANTS.SYS_CONFIG.PICTURE_GALLERY_LIST_NOT_MATCH : CONSTANTS.SYS_CONFIG.VIDEO_GALLERY_LIST_NOT_MATCH
        });
      }

    }
  }
  public async componentDidMount() {
    this.dropDownValidation(this.props.listTitle);
    //If seeAllURL web part properties is empty string then assign default url
    this.getGUID();
    this.setState({
      seeAllURL: this.props.seeAllURL == "" ? this.props.siteURL + "/" + this.props.listTitle + CONSTANTS.SYS_CONFIG.PICTURE_GALLERY_PAGE : this.props.seeAllURL
    });

    this.loadOwnerGroup().then((success) => {
      this.loadMemberGroup().then((succeed) => {
        this.LoadCurrentUserGroups();
      });
    });

    this.LoadRegion();
    //this.loadListContentTypes();
  }

  private getGUID = (): void => {
    if(this.props.webpartType == "PictureGallery") {
      commonService.GetGUID(this.props.listTitle).then((LibraryGUID: any) => {
        this.setState ({
          libraryGUID: LibraryGUID.Id,
          newFormURL: this.props.siteURL + CONSTANTS.SYS_CONFIG.PICTURE_LIST_UPLOAD_PAGE + CONSTANTS.SYS_CONFIG.GUID_START_CODE + LibraryGUID.Id + CONSTANTS.SYS_CONFIG.GUID_END_CODE
        });
      });
    }else {
      this.setState({
        newFormURL: this.props.siteURL + CONSTANTS.SYS_CONFIG.SITE_LISTS + this.props.listTitle + CONSTANTS.SYS_CONFIG.VIDEO_LIST_NEWFORM_PAGE
      });
    }
  }

  private LoadVideos = (): void => {
    let _galleryVideos = [];
    let thumbnailUrl: string = "";
    let videoURL: string = "";
    let videoURLArray: any;
    let caption: string = "";
    let filterCondition = this.getFilterString();

    commonService.GetVideos(this.props.listTitle, filterCondition, CONSTANTS.SELECTCOLUMNS.VIDEO_GALLERY, CONSTANTS.ORDERBY.VIDEO_GALLERY).then((Video: any) => {
      Video.forEach((VideoItem: any, index: number) => {
        thumbnailUrl = "";
        videoURL = "";
        if (VideoItem.Video_x0020_Thumbnail != null) {
          thumbnailUrl = JSON.parse(VideoItem.Video_x0020_Thumbnail).serverUrl + JSON.parse(VideoItem.Video_x0020_Thumbnail).serverRelativeUrl;
        } else {
          thumbnailUrl = this.props.context.pageContext.site.absoluteUrl + "/" + CONSTANTS.ICONS.VIDEO_GALLERY_DEFAULT_IMAGE;
        }

        //Applying the autoplay to video on click
        videoURLArray = VideoItem.Video_x0020_URL.Url.split('?');
        if (videoURLArray.length > 1) {
          let url = new URL(VideoItem.Video_x0020_URL.Url);
          let queryStringParameters = new URLSearchParams(url.search);

          if (queryStringParameters.has("autoplay")) {
            queryStringParameters.set("autoplay", "1");
            videoURL = videoURLArray[0] + "?" + queryStringParameters.toString();
          } else {
            videoURL = VideoItem.Video_x0020_URL.Url + "&autoplay=1";
          }

        } else {
          videoURL = VideoItem.Video_x0020_URL.Url + "?autoplay=1";
        }
        
        if(VideoItem.Caption != null) {
          caption = VideoItem.Caption;
        }else {
          caption = "";
        }
       
        _galleryVideos.push({
          original: thumbnailUrl,
          thumbnail: thumbnailUrl,
          embedUrl: videoURL,
          description: caption,
          renderItem: this._renderVideo.bind(this)
        });
      });
      if (_galleryVideos.length == 0) {
        _galleryVideos.push({
          original: this.props.context.pageContext.site.absoluteUrl + "/" + CONSTANTS.ICONS.VIDEO_GALLERY_IMAGES_NOT_FOUND,
          thumbnail: this.props.context.pageContext.site.absoluteUrl + "/" + CONSTANTS.ICONS.VIDEO_GALLERY_IMAGES_NOT_FOUND
        });
      }
      this.setState({
        galleryImages: _galleryVideos
      });
    });
  }

  private LoadImages = (): void => {
    let _galleryImages = [];
    let filterCondition = this.getFilterString();

    commonService.GetImages(this.props.listTitle, filterCondition, CONSTANTS.SELECTCOLUMNS.PICTURE_GALLERY, CONSTANTS.ORDERBY.PICTURE_GALLERY).then((Img: any) => {
      Img.forEach((ImageItem: any, index: number) => {
        _galleryImages.push({
          original: ImageItem.FileRef,
          thumbnail: ImageItem.FileRef,
          description: ImageItem.Caption

        });
      });
      if (_galleryImages.length == 0) {
        _galleryImages.push({
          original: this.props.context.pageContext.site.absoluteUrl + "/" + CONSTANTS.ICONS.PICTURE_GALLERY_IMAGES_NOT_FOUND,
          thumbnail: this.props.context.pageContext.site.absoluteUrl + "/" + CONSTANTS.ICONS.PICTURE_GALLERY_IMAGES_NOT_FOUND,
        });
      }
      this.setState({
        galleryImages: _galleryImages
      });
    });
  }


  private getFilterString = (): string => {
    // let queryStringParameters = new URLSearchParams(window.location.search);
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

  private _toggleShowVideo(url) {
    this.state.showVideo[url] = !Boolean(this.state.showVideo[url]);
    this.setState({
      showVideo: this.state.showVideo
    });

    if (this.state.showVideo[url]) {
      if (this.state.showPlayButton) {
        this.setState({ showGalleryPlayButton: false });
      }

      if (this.state.showFullscreenButton) {
        this.setState({ showGalleryFullscreenButton: false });
      }
    }
  }

  private _renderVideo(item) {
    return (
      <div>
        {
          this.state.showVideo[item.embedUrl] ?
            <div className='video-wrapper'>
              <a
                className={styles['close-video']}
                onClick={this._toggleShowVideo.bind(this, item.embedUrl)}
              >
              </a>
              <iframe
                //width='350'
                // height='315'
                className="videogallerycontent"
                src={item.embedUrl}
                frameBorder='0'
                allowFullScreen
                allow="accelerometer; autoplay;  clipboard-write; encrypted-media; gyroscope; picture-in-picture"
              >
              </iframe>
            </div>
            :
            <a onClick={this._toggleShowVideo.bind(this, item.embedUrl)}>
              <div className={styles['play-button']}></div>
              <img className='image-gallery-image' src={item.original} />
              { item.description != null ?
               
                item.description &&
                <span
                  className='image-gallery-description'
                  style={{ right: '0', bottom:'0' }}
                >
                  {item.description}
                </span>
                : ""

              }
            </a>
        }
      </div>
    );
  }

  private _onSlide(index) {
    this._resetVideo();
  }

  private _resetVideo() {
    this.setState({ showVideo: {} });

    if (this.state.showPlayButton) {
      this.setState({ showGalleryPlayButton: true });
    }

    if (this.state.showFullscreenButton) {
      this.setState({ showGalleryFullscreenButton: true });
    }
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

  private openWarningModal = () => {
    this.setState({
      isWarningModalOpen: true
    });
  }

  private closeWarningModalDialog = () => {
    this.setState({
      isWarningModalOpen: false
    });
  }

  private openImageModal = () => {
    this.setState({
      isImageModalOpen: true
    });
  }

  private closeImageModalDialog = () => {
    this.setState({
      isImageModalOpen: false,
      fileName: "",
      caption: "",
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      fileNameErrorMessage: "",
      captionErrorMessage: "",
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      showMessageBar: false,
    });
  }

  private openVideoModal = () => {
    this.setState({
      isVideoModalOpen: true
    });
  }

  private closeVideoModalDialog = () => {
    this.setState({
      isVideoModalOpen: false,
      Title: "",
      fileName: "",
      caption: "",
      videoUrl: "",
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      fileNameErrorMessage: "",
      titleErrorMessage: "",
      captionErrorMessage: "",
      videoUrlErrorMessage: "",
      showMessageBar: false,
    });
  }

  private clearImageModalDialog = () => {
    this.setState({
      fileName: "",
      caption: "",
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      fileNameErrorMessage: "",
      captionErrorMessage: "",
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
    });
  }

  private clearVideoModalDialog = () => {
    this.setState({
      Title: "",
      fileName: "",
      caption: "",
      videoUrl: "",
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      fileNameErrorMessage: "",
      captionErrorMessage: "",
      regionErrorMessage: "",
      programTypeErrorMessage: "",
      siteLocationErrorMessage: "",
      titleErrorMessage: "",
      videoUrlErrorMessage: ""
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
        siteLocationDropDownError: true
      });
    }
  }

  public setUploadPictureItemsOLD = async (fileName, file, caption, region, programType, siteLocation): Promise<void> => {
    debugger;
    let fileAddedResult = await commonService.uploadPicture(this.props.listTitle, fileName, file);

    let fileItem = await commonService.getFileItem(fileAddedResult);
    if(fileItem.Id){
      if (siteLocation == "Select") {
        commonService.updatePictureProperties(fileItem, caption, region, programType).then((result: any) => {
          this.LoadImages();
          this.clearImageModalDialog();
          this.setState({  
            message: "Item: " + fileName + " - updated successfully!",
            showMessageBar: true,  
            messageType: MessageBarType.success,
          }); 
          return result;
        }).catch((error:any) => {  
          this.setState({  
            message: "Item " + fileName + " updation failed with error: " + error,  
            showMessageBar: true,  
            messageType: MessageBarType.error  
          });  
        });  
      }else {
        commonService.updatePicturePropertieswithSiteLocation(fileItem, caption, region, programType, siteLocation).then((result: any) => {
          this.LoadImages();
          this.clearImageModalDialog();
          this.setState({  
            message: "Item: " + fileName + " - updated successfully!",
            showMessageBar: true,  
            messageType: MessageBarType.success,
          }); 
          return result;
        }).catch((error:any) => {  
          this.setState({  
            message: "Item " + fileName + " updation failed with error: " + error,  
            showMessageBar: true,  
            messageType: MessageBarType.error  
          });  
        });  
      }
    }
   
  }

  
  public setUploadPictureItems = async (fileName, file, caption, region, programType, siteLocation): Promise<void> => {
    debugger;
    commonService.uploadPicture(this.props.listTitle, fileName, file).then((fileAddedResult) => {
       commonService.getFileItem(fileAddedResult).then((fileItem) => {
        if (siteLocation == "Select") {
          commonService.updatePictureProperties(fileItem, caption, region, programType).then((result: any) => {
            this.LoadImages();
            this.clearImageModalDialog();
            this.setState({  
              message: "Item: " + fileName + " - updated successfully!",
              showMessageBar: true,  
              messageType: MessageBarType.success,
            }); 
            return result;
          }).catch((error:any) => {  
            this.setState({  
              message: "Item " + fileName + " updation failed with error: " + error,  
              showMessageBar: true,  
              messageType: MessageBarType.error  
            });  
          });  
        }else {
          commonService.updatePicturePropertieswithSiteLocation(fileItem, caption, region, programType, siteLocation).then((result: any) => {
            this.LoadImages();
            this.clearImageModalDialog();
            this.setState({  
              message: "Item: " + fileName + " - updated successfully!",
              showMessageBar: true,  
              messageType: MessageBarType.success,
            }); 
            return result;
          }).catch((error:any) => {  
            this.setState({  
              message: "Item " + fileName + " updation failed with error: " + error,  
              showMessageBar: true,  
              messageType: MessageBarType.error  
            });  
          });  
        }
      }).catch((error:any) => {  
        this.setState({  
          message: "Item " + fileName + " updation failed with error: " + error,  
          showMessageBar: true,  
          messageType: MessageBarType.error  
        });  
      });
    }).catch((error:any) => {  
      this.setState({  
        message: "Item " + fileName + " updation failed with error: " + error,  
        showMessageBar: true,  
        messageType: MessageBarType.error  
      });  
    });

    
    
      
    
   
  }

  public UploadItemsInPictureList = (): void =>  {
    let file = this.state.file[0];
    let fileName = this.state.fileName;
    let caption = this.state.caption;
    let selectedRegion = this.state.selectedRegionValue == undefined ? "Select" : this.state.selectedRegionValue.text;
    let selectedProgramType = this.state.selectedProgramTypeValue == undefined ? "Select" : this.state.selectedProgramTypeValue.text;
    let selectedSiteLocation = this.state.selectedSiteLocationValue == undefined ? "Select" : this.state.selectedSiteLocationValue.key;
  
    this.setUploadPictureItems(fileName, file, caption, selectedRegion, selectedProgramType, selectedSiteLocation);
  }

  public isPictureGalleryFormValidate = () => {
    let validation: boolean = true;
    if(this.state.fileName == '') {
      validation = false;
      this.setState({
        fileNameErrorMessage: "Please select a picture.",
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
   
    if(validation) {
        this.UploadItemsInPictureList();
    }
  }

  public isVideoGalleryFormValidate = () => {
    let validation: boolean = true;
    if(this.state.Title == '') {
      validation = false;
      this.setState({
        titleErrorMessage: "Please enter a video title.",
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
    if(this.state.fileName == '') {
      validation = false;
      this.setState({
        fileNameErrorMessage: "Please select a video thumbnail.",
      });
    }
    if(this.state.videoUrl == '') {
      validation = false;
      this.setState({
        videoUrlErrorMessage: "Please enter a video URL.",
      });
    }
    if(validation) {
        this.UploadItemsInVideoList();
    }
  }

  private onVideoTitleChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name] : e.target.value,
      titleErrorMessage: ""
    });     
  }

  private onVideoURLChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name] : e.target.value,
      videoUrlErrorMessage: ""
    });     
  }

  private onImageCaptionChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name] : e.target.value,
      captionErrorMessage: ""
    });     
  }

  public setUploadVideoItems=( title, region, programType, siteLocation, fileName, file, caption, videoURL): void => {
    commonService.GetGUID(this.props.listTitle).then((ListGUID: any) => {
      this.setState ({
        listGUID: ListGUID.Id,
      });
    });
  
    let URL = this.props.siteURL;
    let serverUrl = URL.split('/',4).slice(0, -1).join('/');
    let serverRelativeUrl = URL.split('/',5).slice(3,5).join('/');
    let serverRelativeURL = '/'+ serverRelativeUrl +'/'+ CONSTANTS.SYS_CONFIG.VIDEO_GALLERY_SITEASSETS + this.state.listGUID +'/'+ fileName;
    
    commonService.uploadPicture(CONSTANTS.SYS_CONFIG.VIDEO_GALLERY_SITEASSETS + this.state.listGUID, fileName, file).then((success) => {
      if(siteLocation == "Select") {
        commonService.AddItemsInVideoList(this.props.listTitle, title, region, programType, fileName, file, caption, videoURL, serverUrl, serverRelativeURL).then((result: any) => {
          this.LoadVideos();
          this.clearVideoModalDialog();
          this.setState({  
            message: "Item: " + title + " - updated successfully!",
            showMessageBar: true,  
            messageType: MessageBarType.success,
          });  
          return result;
        }).catch((error:any) => {  
          this.setState({  
            message: "Item " + fileName + " updation failed with error: " + error,  
            showMessageBar: true,  
            messageType: MessageBarType.error  
          });  
        });  
      }else {
        commonService.AddItemsInVideoListwithSiteLocation(this.props.listTitle, title, region, programType, siteLocation, fileName, file, caption, videoURL, serverUrl, serverRelativeURL).then((result: any) => {
          this.LoadVideos();
          this.clearVideoModalDialog();
          this.setState({  
            message: "Item: " + title + " - updated successfully!",
            showMessageBar: true,  
            messageType: MessageBarType.success,
          });  
          return result;
        }).catch((error:any) => {  
          this.setState({  
            message: "Item " + title + " updation failed with error: " + error,  
            showMessageBar: true,  
            messageType: MessageBarType.error  
          });  
        });  
      }
    }).catch((error:any) => {  
      this.setState({  
        message: "Item " + fileName + " updation failed with error: " + error,  
        showMessageBar: true,  
        messageType: MessageBarType.error  
      });  
    });  
  }

  public UploadItemsInVideoList = (): void =>  {
    let file = this.state.file[0];
    let fileName = this.state.fileName;
    let title = this.state.Title;
    let selectedRegion = this.state.selectedRegionValue == undefined ? "Select" : this.state.selectedRegionValue.text;
    let selectedProgramType = this.state.selectedProgramTypeValue == undefined ? "Select" : this.state.selectedProgramTypeValue.text;
    let selectedSiteLocation = this.state.selectedSiteLocationValue == undefined ? "Select" : this.state.selectedSiteLocationValue.key;
    let caption = this.state.caption;
    let videoURL = this.state.videoUrl.replace("watch?v=", "embed/");

    this.setUploadVideoItems(title, selectedRegion, selectedProgramType, selectedSiteLocation, fileName, file, caption, videoURL);
  }

  public render(): React.ReactElement<IPictureGalleryProps> {
    return (
      //<div className={this.props.webpartType == "PictureGallery" ? "PictureMainContainer" : "VideoMainContainer"}>
      <div className={styles.pictureGallery}>
        <div>

          <Stack horizontal tokens={outerStackTokens} className={styles.galleryHeader}>
            <Stack.Item grow={8} >
              {/*<h1 className={"webpartHeader"}>{this.props.webpartTitle}</h1>*/}
              <h1 className={"webpartHeader"}>{this.props.webpartTitle}</h1>
            </Stack.Item>
            { this.state.error == "" ?
              <div className="iconDiv">
                <Stack.Item>
                  <div className={styles.seeAllDiv}>
                    <div className={styles.seeAllItems}>
                      <a href={this.state.seeAllURL} data-interception="off" target="_blank">{this.props.webpartLabel}</a>
                    </div>
                  </div>
                </Stack.Item>
                { this.state.isCurrentUserPresentInGroup == true ?
                  <Stack.Item>
                    <div className={styles.newItemDiv}>
                      <div className={styles.addNewPicture} onClick={this.openWarningModal}>
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
              <Stack>
                {this.state.error}
                {/* <div>
                  <div className={styles.seeAllLabel}>
                    <a href={this.state.seeAllURL} data-interception="off" target="_blank">{this.props.webpartLabel}</a>
                  </div>
                  <a className={styles['next-arrow']} href={this.state.seeAllURL} data-interception="off" target="_blank">
                    <i className="ms-Icon ms-Icon--ChevronRight" aria-hidden="true"></i>
                  </a>
                  <hr className={styles.hrStyle}></hr>
                </div>
               */}
                <div className={styles.imageGalleryDiv}>

                  <ImageGallery
                    items={this.state.galleryImages}
                    infinite={this.props.infinite}
                    showBullets={this.props.showBullets}
                    lazyLoad={false}
                    showFullscreenButton={this.props.webpartType == "PictureGallery" ? this.props.showFullscreenButton && this.state.showGalleryFullscreenButton : this.state.showFullscreenButton}
                    showPlayButton={this.props.showPlayButton && this.state.showGalleryPlayButton}
                    showThumbnails={this.props.showThumbnails}
                    showIndex={this.props.showIndex}
                    showNav={this.props.showNav}
                    isRTL={this.props.isRTL}
                    onSlide={this._onSlide.bind(this)}
                    thumbnailPosition={this.props.thumbnailPosition}
                    slideDuration={(this.props.slideDuration)}
                    slideInterval={(this.props.slideInterval)}
                    slideOnThumbnailOver={this.props.slideOnThumbnailOver}
                    additionalClass="app-image-gallery"
                    useWindowKeyDown={this.props.useWindowKeyDown}
                   
                  />
                </div>
              </Stack>
            </div>
            : <div className={styles.errorDiv}> {this.state.error}</div>
        }
          {this.props.webpartType == "PictureGallery" ?
            <Stack>
              <Stack.Item>
                <Modal
                  titleAriaId="Modal"
                  isOpen={this.state.isWarningModalOpen}
                  onDismiss={this.closeWarningModalDialog}
                  isBlocking={false}
                  containerClassName="warningModalContainer"
                >
                  <div className="modalheader">
                    <span id="Modal" className="warningModalHeader">CONTENT WARNING</span>
                    <IconButton
                      className="modalCloseIcon"
                      iconProps={cancelIcon}
                      ariaLabel="Close popup modal"
                      onClick={this.closeWarningModalDialog}
                    />
                  </div>
  
                  <div className="warningModalContent">
                    <p>
                      Please consider the appropriateness of your images prior to uploading.
                    </p>
                  </div>
                  <div className="modalFotter">
                    <PrimaryButton className="okayButton" 
                      text="Okay" onClick={this.openImageModal}/>
                  </div>
                </Modal>
              </Stack.Item>
            </Stack>
            :
            <Stack>
              <Stack.Item>
                <Modal
                  titleAriaId="Modal"
                  isOpen={this.state.isWarningModalOpen}
                  onDismiss={this.closeWarningModalDialog}
                  isBlocking={false}
                  containerClassName="warningModalContainer"
                >
                  <div className="modalheader">
                    <span id="Modal" className="warningModalHeader">CONTENT WARNING</span>
                    <IconButton
                      className="modalCloseIcon"
                      iconProps={cancelIcon}
                      ariaLabel="Close popup modal"
                      onClick={this.closeWarningModalDialog}
                    />
                  </div>

                  <div className="warningModalContent">
                    <p>
                      Please consider the appropriateness of your videos prior to uploading.
                    </p>
                  </div>
                  <div className="modalFotter">
                    <PrimaryButton className="okayButton" 
                      text="Okay" onClick={this.openVideoModal}/>
                  </div>
                </Modal>
              </Stack.Item>
            </Stack>

          }
         
          <Stack>
            <Stack.Item>
              <Modal
                titleAriaId="Modal"
                isOpen={this.state.isImageModalOpen}
                onDismiss={this.closeImageModalDialog}
                isBlocking={false}
                containerClassName={styles.container}
              >
                <div className="modalheader">
                  <span id="Modal" className="newItemSpan">Add New Picture</span>
                  <IconButton
                    className="modalCloseIcon"
                    iconProps={cancelIcon}
                    ariaLabel="Close popup modal"
                    onClick={()=>{this.closeImageModalDialog(); this.closeWarningModalDialog();}}
                  />
                </div>

                {/* <hr></hr> */}
                <div className="imageBodyDiv">
                  <div className="modalbody">
                    <Stack className="addItemsRow1">
                      <Stack.Item>
                        <Label className={"uploadDocumentLabel"} required={true}>Add a picture</Label>
                        <div className={"dropzoneDiv"}>
                          <Dropzone onDrop={this._onFileDrop} noDragEventsBubbling={true} multiple={false} accept={['image/jpeg','image/png']}>
                            {({getRootProps, getInputProps}) => (
                              <section>
                                <div {...getRootProps()}>
                                  <input {...getInputProps()} />
                                    <div className={"fileUpload cssMarginBottom"} title={this.state.fileName? this.state.fileName : "No file Chosen"}>
                                      <DefaultButton text="Choose File"></DefaultButton>
                                      <Label style={{paddingLeft:'10px'}}>{this.state.fileName? this.state.fileName : "No file Chosen"}</Label>
                                    </div>
                                </div>
                              </section>
                            )}
                          </Dropzone>
                        </div>
                        {this.state.fileName > 0 ? "" : <span className="fileNameErrorSpan">{this.state.fileNameErrorMessage}</span>}
                       
                      </Stack.Item>

                      <Stack.Item>
                        <div className="imageMetadataFieldDiv">
                          <TextField 
                            id="caption"
                            name="caption"
                            label="Caption" 
                            required={false}
                            placeholder="Enter caption"
                            value={this.state.caption}
                            onChange={this.onImageCaptionChange}
                            autoComplete="off"
                            errorMessage={this.state.captionErrorMessage}
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="imageMetadataFieldDiv">
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
                        <div className="imageMetadataFieldDiv">
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
                    onClick={this.isPictureGalleryFormValidate}
                    text="Upload" />
                  <DefaultButton className="closeButton" onClick={()=>{this.closeImageModalDialog(); this.closeWarningModalDialog();}} text="Cancel" />
                </div>
              </Modal>
            
            </Stack.Item>
          </Stack>
          
          <Stack>
            <Stack.Item>
              <Modal
                titleAriaId="Modal"
                isOpen={this.state.isVideoModalOpen}
                onDismiss={this.closeVideoModalDialog}
                isBlocking={false}
                containerClassName={styles.container}
              >
                <div className="modalheader">
                  <span id="Modal" className="newItemSpan">Add New Video</span>
                  <IconButton
                    className="modalCloseIcon"
                    iconProps={cancelIcon}
                    ariaLabel="Close popup modal"
                    onClick={()=>{this.closeVideoModalDialog(); this.closeWarningModalDialog();}}
                  />
                </div>

                {/* <hr></hr> */}
                <div className="videoBodyDiv">
                  <div className="videoFormModalbody">
                    <Stack className="addItemsRow1">

                      <Stack.Item>
                        <div className="videoMetadataFieldDiv">
                          <TextField 
                            id="Title"
                            name="Title"
                            label="Title" 
                            required={true}
                            placeholder="Enter value here"
                            value={this.state.Title}
                            onChange={this.onVideoTitleChange}
                            autoComplete="off"
                            errorMessage={this.state.titleErrorMessage}
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="videoMetadataFieldDiv">
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
                        <div className="videoMetadataFieldDiv">
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
                        <div className="videoMetadataFieldDiv">
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
                        <Label className={"uploadDocumentLabel"} required={true}>Video Thumbnail</Label>
                        <div className={"dropzoneDiv"}>
                          <Dropzone onDrop={this._onFileDrop} noDragEventsBubbling={true} multiple={false} accept={['image/jpeg','image/png']}>
                            {({getRootProps, getInputProps}) => (
                              <section>
                                <div {...getRootProps()}>
                                  <input {...getInputProps()} />
                                    <div className={"fileUpload cssMarginBottom"} title={this.state.fileName? this.state.fileName : "No file Chosen"}>
                                      {/* <DefaultButton text="Choose File"></DefaultButton> */}
                                      <Label className="addThumbnailImage">{this.state.fileName? this.state.fileName : "Add an image"}</Label>
                                    </div>
                                </div>
                              </section>
                            )}
                          </Dropzone>
                        </div>
                        {this.state.fileName > 0 ? "" : <span className="fileNameErrorSpan">{this.state.fileNameErrorMessage}</span>}
                       
                      </Stack.Item>

                      <Stack.Item>
                        <div className="imageMetadataFieldDiv">
                          <TextField 
                            id="caption"
                            name="caption"
                            label="Caption" 
                            required={false}
                            placeholder="Enter caption"
                            value={this.state.caption}
                            onChange={this.onImageCaptionChange}
                            autoComplete="off"
                            errorMessage={this.state.captionErrorMessage}
                          />
                        </div>
                      </Stack.Item>

                      <Stack.Item>
                        <div className="videoMetadataFieldDiv">
                          <TextField 
                            id="videoUrl"
                            name="videoUrl"
                            label="Video URL" 
                            required={true}
                            placeholder="Enter a URL"
                            value={this.state.videoUrl}
                            onChange={this.onVideoURLChange}
                            autoComplete="off"
                            errorMessage={this.state.videoUrlErrorMessage}
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
                          <MessageBar className="videoMessageBarDiv" messageBarType={this.state.messageType}>{this.state.message}</MessageBar>  
                        </Stack>  
                      </div>  
                      :  
                      null  
                    }  
                  </Stack.Item>
                </Stack>
                <div className="modalFotter">
                  <PrimaryButton className="saveButton" 
                    onClick={this.isVideoGalleryFormValidate}
                    text="Upload" />
                  <DefaultButton className="closeButton" onClick={()=>{this.closeVideoModalDialog(); this.closeWarningModalDialog();}} text="Cancel" />
                </div>
              </Modal>
            
            </Stack.Item>
          </Stack>
          
      </div>

    );
  }
}

