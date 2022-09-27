import * as React from 'react';
import styles from './VideoGallery.module.scss';
import { IVideoGalleryProps } from './IVideoGalleryProps';
import dataService from "../../../Common/DataService";
import CONSTANTS from "../../../Common/Constants";
import ImageGallery from 'react-image-gallery';
import "react-image-gallery/styles/css/image-gallery.css";
import { IconButton, IIconProps, Modal, PrimaryButton, Stack } from 'office-ui-fabric-react'; //'@fluentui/react';
import * as _ from 'lodash';
import { RxJsEventEmitter } from '../../RxJsEventEmitter/RxJsEventEmitter';
import commonMethods from "../../../Common/CommonMethods";

let g_selectedRegion: string = "";
let g_selectedProgramType: string = "";

export interface IVideoGalleryState {
  galleryVideos: string[];
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
  showVideo: any;
  selectedRegionValue: any;
  selectedProgramTypeValue: any;
  seeAllURL: string;
  isWarningModalOpen: boolean;
}
export interface IEventData {
  sharedRegion: any;
  sharedProgramType: any;
}

const cancelIcon: IIconProps = { iconName: 'Cancel' };
const commonService = new dataService();
const commonMethod = new commonMethods();
export default class VideoGallery extends React.Component<IVideoGalleryProps, IVideoGalleryState> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  constructor(props: IVideoGalleryProps, state: IVideoGalleryState) {
    super(props);
    this.state = ({
      galleryVideos: [],
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
      thumbnailPosition: 'bottom',
      useWindowKeyDown: true,
      showVideo: {},
      selectedRegionValue: [],
      selectedProgramTypeValue: [],
      seeAllURL: "",
      isWarningModalOpen: false,
    });
    this._renderVideo = this._renderVideo.bind(this);
    this.eventEmitter.on(CONSTANTS.CONNECTED_WP.SHARE_DATA, this.receiveData.bind(this));
  }

  private receiveData(data: IEventData) {
    debugger;
    g_selectedRegion = data.sharedRegion.text;
    g_selectedProgramType = data.sharedProgramType.text;
    this.LoadVideos();
  }
  public dropDownValidation = async (listTitle: string) => {
    if (listTitle == undefined || listTitle == ' ') {
      this.setState({
        error: CONSTANTS.SYS_CONFIG.SELECT_LIST
      });
    }
    else {

      //Check all required fields available in selected list.
      let isValidListColumns = await commonMethod.isValidListColumns(listTitle,CONSTANTS.LIST_VALIDATION_COLUMNS.VIDEO);
      
       if(isValidListColumns){
         this.LoadVideos();        
       }else{
         this.setState({
           error: CONSTANTS.SYS_CONFIG.VIDEO_GALLERY_LIST_NOT_MATCH
         });
       }
      
    }
  }

  public async componentDidMount() {
    this.dropDownValidation(this.props.listTitle);
    this.setState({
      seeAllURL: this.props.seeAllURL == "" ? this.props.siteURL + CONSTANTS.SYS_CONFIG.SITE_LISTS + this.props.listTitle + CONSTANTS.SYS_CONFIG.VIDEO_GALLERY_PAGE : this.props.seeAllURL
    });
  }

  private LoadVideos = (): void => {
    let _galleryVideos = [];
    let thumbnailUrl: string = "";
    let videoURL: string = "";
    let videoURLArray: any;
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
        //debugger;

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

        _galleryVideos.push({
          original: thumbnailUrl,
          thumbnail: thumbnailUrl,
          embedUrl: videoURL,
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
        galleryVideos: _galleryVideos
      });
    });
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
              {
                item.description &&
                <span
                  className='image-gallery-description'
                  style={{ right: '0', left: 'initial' }}
                >
                  {item.description}
                </span>
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

  public render(): React.ReactElement<IVideoGalleryProps> {
    return (
      <div className={styles.videoGallery}>
        <div>
          <h1 className={styles.videoGalleryHeader}>{this.props.webpartTitle}</h1>
        </div>
        {
          this.state.error == "" ?
            <div className={styles.container}>
              <Stack>
                {this.state.error}
                <div>
                  <div className={styles.seeAllLabel}>
                    <a href={this.state.seeAllURL} data-interception="off" target="_blank">{this.props.webpartLabel}</a>
                  </div>
                  <a className={styles['next-arrow']} href={this.state.seeAllURL} data-interception="off" target="_blank">
                    <i className="ms-Icon ms-Icon--ChevronRight" aria-hidden="true"></i>
                  </a>
                  <hr className={styles.hrStyle}></hr>
                </div>

                <div className={styles.videoGalleryDiv}>
                  <ImageGallery
                    items={this.state.galleryVideos}
                    showNav={this.props.showNav}
                    thumbnailPosition={this.props.thumbnailPosition}
                    lazyLoad={false}
                    onSlide={this._onSlide.bind(this)}
                    infinite={this.props.infinite}
                    showBullets={this.props.showBullets}
                    showFullscreenButton={this.props.showFullscreenButton}
                    showPlayButton={this.props.showPlayButton}
                    showThumbnails={this.props.showThumbnails}
                    showIndex={this.props.showIndex}
                    isRTL={this.state.isRTL}
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
                    text="Okay" />
                </div>
              </Modal>
            </Stack.Item>
          </Stack>

      </div>
    );
  }
}
