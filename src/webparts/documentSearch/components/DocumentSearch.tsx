import * as React from 'react';
import styles from './DocumentSearch.module.scss';
import { IDocumentSearchProps } from './IDocumentSearchProps';
import { chunk, escape } from '@microsoft/sp-lodash-subset';
import { Text, Breadcrumb, ComboBox, DatePicker, DayOfWeek, DefaultButton, DetailsList, DetailsListLayoutMode, Selection, Dialog, DialogFooter, DialogType, Dropdown, Fabric, IBreadcrumbItem, IColumn, IComboBox, IComboBoxOption, Icon, IconButton, IDropdownOption, IIconProps, Modal, PrimaryButton, SearchBox, SelectionMode, Stack, Sticky, Link, Label, TextField, ITextFieldStyles, ContextualMenu, IContextualMenuItem, DirectionalHint, getFadedOverflowStyle } from 'office-ui-fabric-react'; //'@fluentui/react';
import { getFileTypeIconProps, FileIconType, initializeFileTypeIcons } from '@uifabric/file-type-icons'; //'@fluentui/react-file-type-icons'; 
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import CONSTANTS from '../../../Common/Constants';
import dataService from "../../../Common/DataService";
import * as _ from 'lodash';
import * as Moment from 'moment';
import Pagination from 'react-js-pagination';
import { Loader } from '../../CommonComponents/Loader';
//import { Loader } from '..//CommonComponents/Loader'
import * as $ from 'jquery';
//import { Item } from '@pnp/sp/items';
import { RxJsEventEmitter } from '../../RxJsEventEmitter/RxJsEventEmitter';
import { stringIsNullOrEmpty } from '@pnp/common';
import Dropzone, { DropEvent, FileRejection } from 'react-dropzone';
//import { RelatedItemManager } from '@pnp/sp/related-items';

let RegionOptions: any[] = [];
let ProgramTypeOptions: any[] = [];
let SiteLocationOptions: any[] = [];
let rootPath: string = "";
//initialize file type icons to use
initializeFileTypeIcons(undefined);
export interface IDocumentSearchState {
  regionOptions: any;
  programTypeOptions: any;
  selectedRegionValue: any;
  selectedProgramTypeValue: any;
  siteLocationOptions: any;
  selectedSiteLocationValue: any;
  modifiedBy: any[];
  modifiedByName: string;
  modifiedDate: any;
  siteLocationMetadatAvailable: boolean;
  leftNavigationItems: any[];
  selectedLeftNavigationItem: any[];
  columns: IColumn[];
  columnsSiteLocation: IColumn[];
  allDocumentitems: IDocument[];
  filtterDocumentitems: IDocument[];
  searchDocumentitems: IDocument[];
  totalItemsCount: number;
  selectedItemPerPage: any;
  counter: Number;
  allFormattedItems: any;
  displayNoDataMassage: boolean;
  rgnptyIsAsAll: boolean;
  showLoader: boolean;
  isOpenDialog: boolean;
  DialogMessage: string;
  isModalOpen: boolean;
  documentUploadDetails: {
    DocumentFailedToUpload: any,
    TotalCount: number,
    FailedCount: number,
    SuccesCount: number
  };
  siteLocationData: any;
  searchTextValue: string;
  isSearchClicked: boolean;
  isMetadatDialogOpen: boolean;
  documentsToBeupload: any;
  errorMessage: string;
  folderServerRelativeUrl: string;
  BreadCrumbItems: IBreadcrumbItem[];
  siteLocationColl: any;
  siteLocationDropDownError: boolean;
  programTypeDropDownError: boolean;
  regionDropDownError: boolean;
  sortColoumn: any;
  isFileUploadModalOpen: boolean;
  file: any[];
  fileName: any;
  fileSize: any;
  documentTitle: any;
  isOpenContextualMenu: boolean;
  isEditMetadataDialogOpen: boolean;
  selectedItemId: any;
  EPSelectedRegionValue: any;
  EPSelectedProgramTypeValue: any;
  EPSelectedSiteLocationValue: any;
  isCurrentUserPresentInGroup: boolean;
  isOpenNoUploadPermissionDialog: boolean;
  noUploadPermissionDialogMessage: string;
  ownersGroup: string;
  membersGroup: string;
}
const cancelIcon: IIconProps = { iconName: 'Cancel' };

export interface IDocument {
  key: string;
  title: string;
  name: string;
  value: string;
  iconPath: string;
  fileType: string;
  modifiedBy: string;
  dateModified: string;
  dateModifiedValue: string;
  region: string;
  programType: string;
  FileRef: string;
  id: any;
  siteLocation: string;
  siteLocationId: number;


}

let ItemPerPageDropdown = [
  { key: 10, text: '10/page' },
  { key: 20, text: '20/page' },
  { key: 30, text: '30/page' },
  { key: 40, text: '40/page' },
  { key: 50, text: '50/page' },
];

export interface IEventData {
  sharedRegion: any;
  sharedProgramType: any;
  sharedLeftNavigation: string;
  sharedSiteLocationMetdataAvailable: boolean;
}

const commonService = new dataService();
export default class DocumentSearch extends React.Component<IDocumentSearchProps, IDocumentSearchState> {
  private readonly eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  private _columns: IColumn[];
  private _columnsSiteLocation: IColumn[];
  private modelProps = {
    isBlocking: false,
    styles: { main: { maxWidth: 450 } },
  };
  private dialogContentProps = {
    type: DialogType.largeHeader,
    title: 'Set Document Details'
    //subText: {this.state.DialogMessage}
  };
  private editDialogContentProps = {
    type: DialogType.largeHeader,
    title: 'Edit Document Details'
    //subText: {this.state.DialogMessage}
  };
  private _selection: Selection;

  constructor(props: IDocumentSearchProps, state: IDocumentSearchState) {
    super(props);

    this._columns = [
      {
        key: 'column1',
        name: 'File Type',
        className: styles.fileIconCell,
        iconClassName: styles.fileIconHeaderIcon,
        ariaLabel: 'Column operations for File type, Press to sort on File type',
        iconName: '',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 30,
        maxWidth: 30,

        onRender: (item: IDocument) => {
          return item.fileType == "Folder" ? <Icon className={styles.DetailsListIcon} {...getFileTypeIconProps({ type: FileIconType.folder, size: 24, imageFileType: 'svg' })} /> :
            <Icon className={styles.DetailsListIcon} {...getFileTypeIconProps({ extension: item.fileType, size: 24, imageFileType: 'svg' })} />;
        },
      },
      {
        key: 'column2',
        name: 'Title',
        fieldName: 'title',
        className: styles.document_Name,
        minWidth: 200,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        isPadded: true,
        onRender: (item: IDocument) => {
          return (
            <>
              <div className={styles.fileTitleDiv}>
                {item.fileType == "Folder" ? <a onClick={() => this.onClickListViewItem(item.FileRef, item.fileType, item.name)}>{item.name}</a> :
                  <a onClick={() => this.onClickListViewItem(item.FileRef, item.fileType, item.name)}>{item.title}</a>}
              </div>
              { this.state.isCurrentUserPresentInGroup == true ?
                <div className={styles.verticalDotDiv}>
                  {item.fileType == "Folder" ? "" :  
                    <IconButton
                      role="menuitem"
                      title="Edit Properties"
                      menuIconProps={{ iconName: 'MoreVertical' }}
                      menuProps={
                        {
                          items:
                          [
                            {
                              key: 'properties',
                              text: 'Edit Properties',
                              onClick: () => this.LoadSelectedDocumentData(this.state.selectedItemId)
                            }
                          ],
                          directionalHint: DirectionalHint.bottomAutoEdge
                        }
                      }
                    />
                  }
                </div> : ""
              }
            </>
          );
        },
      },
      {
        key: 'column3',
        name: 'Region',
        fieldName: 'region',
        className: styles.document_text,
        minWidth: 80,
        maxWidth: 80,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,

        isPadded: true,
      },
      {
        key: 'column4',
        name: 'Program Type',
        fieldName: 'programType',
        className: styles.document_text,
        minWidth: 120,
        maxWidth: 120,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,

        isPadded: true,
      },
      {
        key: 'column6',
        name: 'Modified By',
        fieldName: 'modifiedBy',
        className: styles.document_text,
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,

        isPadded: true,
      },
      {
        key: 'column7',
        name: 'Modified On',
        fieldName: 'dateModifiedValue',
        className: styles.document_text,
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        isPadded: true,
      },

    ];

    this._columnsSiteLocation = [
      {
        key: 'column1',
        name: 'File Type',
        className: styles.fileIconCell,
        iconClassName: styles.fileIconHeaderIcon,
        ariaLabel: 'Column operations for File type, Press to sort on File type',
        iconName: '',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 30,
        maxWidth: 30,
        onRender: (item: IDocument) => {
          return item.fileType == "Folder" ? <Icon className={styles.DetailsListIcon} {...getFileTypeIconProps({ type: FileIconType.folder, size: 24, imageFileType: 'svg' })} /> :
            <Icon className={styles.DetailsListIcon} {...getFileTypeIconProps({ extension: item.fileType, size: 24, imageFileType: 'svg' })} />;
        },
      },
      {
        key: 'column2',
        name: 'Title',
        fieldName: 'title',
        className: styles.document_Name,
        minWidth: 200,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        isPadded: true,
        onRender: (item: IDocument) => {
          return (
            <>
              <div className={styles.fileTitleDiv}>
                {item.fileType == "Folder" ? <a onClick={() => this.onClickListViewItem(item.FileRef, item.fileType, item.name)}>{item.name}</a> :
                  <a onClick={() => this.onClickListViewItem(item.FileRef, item.fileType, item.name)}>{item.title}</a>}
              </div>
              { this.state.isCurrentUserPresentInGroup == true ?
                <div className={styles.verticalDotDiv}>
                  {item.fileType == "Folder" ? "" :  
                    <IconButton
                    role="menuitem"
                    title="Edit Properties"
                    menuIconProps={{ iconName: 'MoreVertical' }}
                    menuProps={
                      {
                        items:
                        [
                          {
                            key: 'properties',
                            text: 'Edit Properties',
                            onClick: () => this.LoadSelectedDocumentData(this.state.selectedItemId)
                          }
                        ],
                        directionalHint: DirectionalHint.bottomAutoEdge
                      }
                    }
                  />
                  }
                </div> : ""
              }
            </>
          );
        },
      },
      {
        key: 'column3',
        name: 'Region',
        fieldName: 'region',
        className: styles.document_text,
        minWidth: 80,
        maxWidth: 80,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,

        isPadded: true,
      },
      {
        key: 'column4',
        name: 'Program Type',
        fieldName: 'programType',
        className: styles.document_text,
        minWidth: 120,
        maxWidth: 120,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,

        isPadded: true,
      },
      {
        key: 'column5',
        name: 'Site Location',
        fieldName: 'siteLocation',
        className: styles.document_text,
        minWidth: 120,
        maxWidth: 120,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,

        isPadded: true,
      },
      {
        key: 'column6',
        name: 'Modified By',
        fieldName: 'modifiedBy',
        className: styles.document_text,
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,

        isPadded: true,
      },
      {
        key: 'column7',
        name: 'Modified On',
        fieldName: 'dateModifiedValue',
        className: styles.document_text,
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        isPadded: true,
      },

    ];


    this.state = {
      regionOptions: [],
      programTypeOptions: [],
      selectedRegionValue: [],
      selectedProgramTypeValue: [],
      siteLocationOptions: [],
      selectedSiteLocationValue: [],
      modifiedBy: [],
      modifiedByName: "",
      modifiedDate: null,
      siteLocationMetadatAvailable: false,
      leftNavigationItems: [],
      selectedLeftNavigationItem: [],
      columns: this._columns,
      allDocumentitems: [],
      filtterDocumentitems: [],
      searchDocumentitems: [],
      totalItemsCount: 0,
      selectedItemPerPage: { key: ItemPerPageDropdown[0].key, text: ItemPerPageDropdown[0].text },
      counter: 0,
      allFormattedItems: [],
      displayNoDataMassage: false,
      rgnptyIsAsAll: false,
      showLoader: false,
      isOpenDialog: false,
      DialogMessage: "",
      isModalOpen: false,
      documentUploadDetails: {
        DocumentFailedToUpload: [],
        FailedCount: 0,
        SuccesCount: 0,
        TotalCount: 0
      },
      columnsSiteLocation: this._columnsSiteLocation,
      siteLocationData: [],
      searchTextValue: "",
      isSearchClicked: false,
      isMetadatDialogOpen: false,
      documentsToBeupload: [],
      errorMessage: "",
      folderServerRelativeUrl: "",
      BreadCrumbItems: [],
      siteLocationColl: [],
      siteLocationDropDownError: false,
      programTypeDropDownError: false,
      regionDropDownError: false,
      sortColoumn: [],
      isFileUploadModalOpen: false,
      file: [],
      fileName: "",
      fileSize: "",
      documentTitle: "",
      isOpenContextualMenu: false,
      isEditMetadataDialogOpen: false,
      selectedItemId: "",
      EPSelectedRegionValue: [],
      EPSelectedProgramTypeValue: [],
      EPSelectedSiteLocationValue: [],
      isCurrentUserPresentInGroup: false,
      isOpenNoUploadPermissionDialog: false,
      noUploadPermissionDialogMessage: "",
      ownersGroup: "",
      membersGroup: "",
    };

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectedItemId: this._getSelectedItem() }),
    });
    this._OnDragLeaveDocument = this._OnDragLeaveDocument.bind(this);
    this._onDragOverDocument = this._onDragOverDocument.bind(this);
    this._onDrop = this._onDrop.bind(this);
    this.eventEmitter.on(CONSTANTS.CONNECTED_WP.SHARE_DATA, this.receiveData.bind(this));
  }

  private _getSelectedItem(): string {
    
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return `${selectionCount}`;
      case 1:
        return (this._selection.getSelection()[0] as IDocument).id;
      default:
       return `${selectionCount}`;
    }
    
  }

  public LoadSelectedDocumentData = (selectedProductId) => {
    //debugger;
    let selectedLibraryName = this.state.selectedLeftNavigationItem[0].LibraryTitle;
    if (this.state.siteLocationMetadatAvailable == false) {
      commonService.getSelectedDocumentData(selectedLibraryName, selectedProductId, CONSTANTS.SELECTCOLUMNS.DOCUMENTS_LIST).then((DocumentItem: any) => {
        let selectedRegion = _.filter(RegionOptions, (p) => {
          return p.text.toLowerCase() == DocumentItem.Region.toLowerCase();
        });
        //alert(DocumentItem.Program_x0020_Type.toLowerCase());
        let selectedProgramType = _.filter(ProgramTypeOptions, (p) => {
          return p.text.toLowerCase() == DocumentItem.Program_x0020_Type.toLowerCase();
        });

          this.setState({
            documentTitle: DocumentItem.Title,
            EPSelectedRegionValue: {key: selectedRegion[0].key, text: DocumentItem.Region},
            EPSelectedProgramTypeValue: {key: selectedProgramType[0].key, text: DocumentItem.Program_x0020_Type},
            isEditMetadataDialogOpen: true,
          });
      });
    } else {
      commonService.getSelectedDocumentDatawithSiteLocation(selectedLibraryName, selectedProductId, CONSTANTS.SELECTCOLUMNS.DOCUMENT_LIST_SITELOCATION,CONSTANTS.EXPAND_COLUMN.DOCUMENT_LIST_SITELOCATION).then((DocumentItem: any) => {
        let selectedRegion = _.filter(RegionOptions, (p) => {
          return p.text.toLowerCase() == DocumentItem.Region.toLowerCase();
        });
        let selectedProgramType = _.filter(ProgramTypeOptions, (p) => {
          return p.text.toLowerCase() == DocumentItem.Program_x0020_Type.toLowerCase();
        });
        let selectedSiteLocation = _.filter(SiteLocationOptions, (p) => {
          return p.text.toLowerCase() == DocumentItem.Site_x0020_Location_x0020_Code.Title.toLowerCase() + "-" + DocumentItem.Site_x0020_Location_x0020_Code.Code.toLowerCase();
        });
          this.setState({
            documentTitle: DocumentItem.Title,
            EPSelectedRegionValue: {key: selectedRegion[0].key, text: DocumentItem.Region},
            EPSelectedProgramTypeValue: {key: selectedProgramType[0].key, text: DocumentItem.Program_x0020_Type},
            EPSelectedSiteLocationValue: {key: selectedSiteLocation[0].key, text: DocumentItem.Site_x0020_Location_x0020_Code.Title + "-" + DocumentItem.Site_x0020_Location_x0020_Code.Code},
            isEditMetadataDialogOpen: true,
          });
      });
    }
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    //debugger;
    let isSortedDescending = column.isSortedDescending;
    // If we've sorted this column, flip it.
    if (column.isSorted) {
      isSortedDescending = !isSortedDescending;
    }
    this.sortDocumentData(column, isSortedDescending, this.state.selectedItemPerPage.key);

  }

  private sortDocumentData = (column: IColumn, isSortedDescending: boolean, itemPerPage: any): void => {

    let { searchDocumentitems } = this.state;

    let folders = _.filter(searchDocumentitems, (p) => {
      return p.fileType == "Folder";
    });
    let files = _.filter(searchDocumentitems, (p) => {
      return p.fileType != "Folder";
    });


    // Sort the items.
    folders = folders!.concat([]).sort((a, b) => {
      const firstValue = a[column.fieldName || ''];
      const secondValue = b[column.fieldName || ''];

      if (isSortedDescending) {
        return firstValue > secondValue ? -1 : 1;
      } else {
        return firstValue > secondValue ? 1 : -1;
      }
    });

    files = files!.concat([]).sort((a, b) => {
      const firstValue = a[column.fieldName || ''];
      const secondValue = b[column.fieldName || ''];

      if (isSortedDescending) {
        return firstValue > secondValue ? -1 : 1;
      } else {
        return firstValue > secondValue ? 1 : -1;
      }
    });

    searchDocumentitems = folders.concat(files);
    column.isSortedDescending = isSortedDescending;
    let gridData: any = this.paging(searchDocumentitems, itemPerPage);
    // Reset the items and columns to match the state.
    if (this.state.siteLocationMetadatAvailable == true) {
      const { columnsSiteLocation } = this.state;
      this.setState({
        columnsSiteLocation: columnsSiteLocation!.map(col => {
          col.isSorted = col.key === column.key;

          if (col.isSorted) {
            col.isSortedDescending = isSortedDescending;

          }
          return col;
        }),
        searchDocumentitems,
        allFormattedItems: gridData,
        filtterDocumentitems: gridData[0],
        totalItemsCount: searchDocumentitems.length,
        counter: 0,
        sortColoumn: column
      });
    } else {
      const { columns } = this.state;
      this.setState({
        columns: columns!.map(col => {
          col.isSorted = col.key === column.key;

          if (col.isSorted) {
            col.isSortedDescending = isSortedDescending;
          }

          return col;
        }),
        searchDocumentitems,
        allFormattedItems: gridData,
        filtterDocumentitems: gridData[0],
        totalItemsCount: searchDocumentitems.length,
        counter: 0,
        sortColoumn: column
      });
    }

  }

  private receiveData(data: IEventData) {
    //debugger;
    let rgn: string = data.sharedRegion.text;
    let pty: string = data.sharedProgramType.text;

    this.setState({
      siteLocationMetadatAvailable: data.sharedSiteLocationMetdataAvailable,
      rgnptyIsAsAll: rgn == "Select" && pty == "Select" ? true : false,
      showLoader: true
    });

    this.GetLeftNavigation(data.sharedLeftNavigation);

    this.LoadRegion();
  }


  private onClickListViewItem = (documentURL: string, type: string, name: string): any => {
    if (type == "Folder") {
      this.setState({ showLoader: true });
      //alert(`documentURL:- ${documentURL} type:-  ${type} Name:- ${name}`);
      let Url: string = `${this.state.folderServerRelativeUrl}/${name}`;
      this.setState({ folderServerRelativeUrl: Url });
      this.getDocumentData(false, Url);
      this.buildBreadCrumb(Url);
    } else {
      //Encode the file name and create document url to download docuement
      documentURL = this.state.folderServerRelativeUrl + "/" + encodeURIComponent(name).replace(/'/g, "%27").replace(/"/g, "%22");
      let url: string = `${this.props.context.pageContext.site.absoluteUrl}/_layouts/download.aspx?sourceurl=${documentURL}`;
      window.open(url, "_blank");
      return false;
    }
  }

  public async componentDidMount() {
    let queryStringParameters = new URLSearchParams(window.location.search);
    //debugger;

    let slmda = queryStringParameters.get("slmda") ? queryStringParameters.get("slmda").toLowerCase() : "";
    let rgn = queryStringParameters.get("rgn") ? queryStringParameters.get("rgn").toLowerCase() : "";
    let pty = queryStringParameters.get("pty") ? queryStringParameters.get("pty").toLowerCase() : "";
    let lnav: string = queryStringParameters.get("lnav") ? queryStringParameters.get("lnav").toLowerCase() : "";

    this.setState({
      siteLocationMetadatAvailable: slmda == "true" ? true : false,
      rgnptyIsAsAll: rgn == "Select" && pty == "Select" ? true : false,
      showLoader: true
    });

    this.GetLeftNavigation(lnav);
    this.LoadRegion();
    // this.setGlobalCssChanges();
    this.loadOwnerGroup().then((success) => {
      this.loadMemberGroup().then((succeed) => {
        this.LoadCurrentUserGroups();
      });
    });

  }

  private GetLeftNavigation = (leftNavigationItemId: string): void => {
    //debugger;
    let leftNavItems: any[] = [];
    let queryStringParameters = new URLSearchParams(window.location.search);
    let queryStringLeftNav: string = "";
    commonService.getLeftNavigationConfigurationData(CONSTANTS.LIST_NAME.LEFT_NAVIGATION_CONFIGURATION_LIST, CONSTANTS.SELECTCOLUMNS.LEFT_NAVIGATION_COLS, CONSTANTS.FILTERCONDITION.LEFT_NAVIGATION_QUERY, CONSTANTS.ORDERBY.LEFT_NAVIGATION).then((data: any) => {

      if (data.length > 0) {
        data.forEach(async (leftNavigationItem: any, index: number) => {

          if (queryStringParameters.get("lnav") == leftNavigationItem.Id) {
            queryStringLeftNav = leftNavigationItemId;
          }

          leftNavItems.push({
            Title: leftNavigationItem.Title,
            Id: leftNavigationItem.Id,
            LibraryName: leftNavigationItem.LibraryName,
            LibraryTitle: leftNavigationItem.LibraryTitle,
            IsSiteLocationMetadataAvailable: leftNavigationItem.IsSiteLocationMetadataAvailable,
            IsActiveLeftNav: queryStringLeftNav == leftNavigationItem.Id ? true : false
          });
        });
      }
      //debugger;
      //get the selected left navigation
      let selectedLeftNavigationItem = _.filter(leftNavItems, (p) => {
        return p.Id == queryStringLeftNav;
      });
      //Set the Page title
      document.title = selectedLeftNavigationItem[0].Title;
      rootPath = `${this.props.context.pageContext.site.serverRelativeUrl}/${selectedLeftNavigationItem[0].LibraryName}`;
      this.setState({
        leftNavigationItems: leftNavItems,
        selectedLeftNavigationItem: selectedLeftNavigationItem,
        folderServerRelativeUrl: rootPath
      });
     
      this.buildBreadCrumb(rootPath);
    });
  }

  private getFilterString = (): string => {
    let filterString: string = "";
    let andCondition: string = " and ";
    if (this.state.selectedRegionValue.text != "Select") {
      filterString = "ListItemAllFields/Region eq '" + this.state.selectedRegionValue.text + "'";
    }
    if (this.state.selectedProgramTypeValue.text != "Select") {
      filterString += filterString != "" ? andCondition : "";
      filterString += "ListItemAllFields/Program_x0020_Type eq '" + this.state.selectedProgramTypeValue.text + "'";
    }
    /*if (this.state.selectedSiteLocationValue.text != "Select") {

      filterString += filterString != "" ? andCondition : "";
      filterString += "ListItemAllFields/Site_x0020_Location_x0020_Code eq '" + this.state.selectedSiteLocationValue.key + "'";
    }
    */

    //Get the selected site location region and added filter, if reion is not selected
    if (this.state.selectedSiteLocationValue.text != "Select" && this.state.selectedRegionValue.text == "Select") {
      let selectedSiteLocation = _.filter(this.state.siteLocationColl, (item) => {
        return item.Id == this.state.selectedSiteLocationValue.key;
      });
      filterString += filterString != "" ? andCondition : "";
      filterString += "ListItemAllFields/Region eq '" + selectedSiteLocation[0].Region + "'";
    }

    //debugger;
    if (this.state.modifiedDate != null) {
      let selectedDate: string = Moment(this.state.modifiedDate).format("YYYY-MM-DD");
      let startDate = selectedDate + 'T00:00:00.000Z';
      let endtDate = selectedDate + 'T23:59:00.000Z';
      filterString += filterString != "" ? andCondition : "";
      filterString += "(ListItemAllFields/Modified ge datetime'" + startDate + "') and (ListItemAllFields/Modified le datetime'" + endtDate + "')";
      //filterString += "(FieldValuesAsText/Modified ge datetime'" + startDate + "') and (FieldValuesAsText/Modified le datetime'" + endtDate + "')";
    }
    /*if (this.state.modifiedBy.length > 0) {
      filterString += filterString != "" ? andCondition : "";      
      filterString += "substringof('" + this.state.modifiedBy[0].ID + "',EditorId)";

    }*/
    return filterString;
  }



  private getDocumentData = async (clearFilter: boolean, folderServerRelativeUrl: string): Promise<void> => {
    const items: IDocument[] = [];
    let FilterCondition: string = "";

    if (!clearFilter) {
      FilterCondition = this.getFilterString();
    }

    //Get the folders data by folder server relative URL
    await commonService.getFolderData(folderServerRelativeUrl, CONSTANTS.SYS_CONFIG.GET_ITEMS_LIMIT, CONSTANTS.ORDERBY.GET_DOCUMENT_ORDERBY, CONSTANTS.EXPAND_COLUMN.GET_DOCUMENTS).then((folderItems: any) => {

      if (folderItems.length > 0) {
        folderItems.forEach((documentItem: any, index: Number) => {
          //Exclude the "root" level Forms folder
          if (documentItem.ServerRelativeUrl != `${decodeURIComponent(rootPath)}/Forms`) {

            var FormateDate = Moment(documentItem.ListItemAllFields.FieldValuesAsText.Modified).format("MM/DD/YYYY");
            //var FormateDate = Moment(documentItem.FieldValuesAsText.Modified).format("MM/DD/YYYY");
            var date: Date = new Date(FormateDate);

            items.push({
              key: index.toString(),
              title: documentItem.ListItemAllFields.FieldValuesAsText.Title,
              name: documentItem.Name,
              value: documentItem.Name,
              iconPath: "",
              fileType: "Folder",
              modifiedBy: documentItem.ListItemAllFields.FieldValuesAsText.Editor != "" ? documentItem.ListItemAllFields.FieldValuesAsText.Editor : "N/A",
              dateModified: date.toLocaleDateString(),
              dateModifiedValue: FormateDate,
              region: "N/A",
              programType: "N/A",
              FileRef: documentItem.ServerRelativeUrl,
              id: documentItem.ListItemAllFields.Id,
              siteLocation: "N/A",
              siteLocationId: 0
            });
          }
        });

      }
    }).catch((error: any) => {
      console.log(error);
      this.setState({
        allDocumentitems: items,
        filtterDocumentitems: items,
        searchDocumentitems: items,
        totalItemsCount: items.length,
        errorMessage: CONSTANTS.SYS_CONFIG.DATA_ERROR,
        showLoader: false,
        displayNoDataMassage: true,
      });

    });

    //Get the files data by folder server relative URL
    await commonService.getDocumentsData(folderServerRelativeUrl, CONSTANTS.SELECTCOLUMNS.GET_DOCUMENTS, CONSTANTS.EXPAND_COLUMN.GET_DOCUMENTS_NEW, FilterCondition, CONSTANTS.SYS_CONFIG.GET_ITEMS_LIMIT, CONSTANTS.ORDERBY.GET_DOCUMENT_ORDERBY).then((dataItems: any) => {

      if (dataItems.length > 0) {

        if (this.state.selectedSiteLocationValue.text != "Select" &&
          this.state.selectedSiteLocationValue.text != undefined) {
          dataItems = _.filter(dataItems, (file) => {
            return file.ListItemAllFields.Site_x0020_Location_x0020_CodeId != null ?
              file.ListItemAllFields.Site_x0020_Location_x0020_CodeId == this.state.selectedSiteLocationValue.key : false;
          });
        }

        if (this.state.modifiedBy.length > 0) {
          dataItems = _.filter(dataItems, (file) => {
            return file.ListItemAllFields.EditorId == this.state.modifiedBy[0].ID;
          });
        }

        dataItems.forEach((documentItem: any, index: number) => {

          var FormateDate = Moment(documentItem.ListItemAllFields.FieldValuesAsText.Modified).format("MM/DD/YYYY");
          var date: Date = new Date(FormateDate);
          let siteLocation_Code: string = documentItem.ListItemAllFields.FieldValuesAsText.Site_x005f_x0020_x005f_Location_x005f_x0020_x005f_Code;
          let siteLocation_Location: string = documentItem.ListItemAllFields.FieldValuesAsText.Site_x005f_x0020_x005f_Location_x005f_x0020_x005f_Code_x005f_x003A_x005f_Location;

          items.push({
            //Append the folders key in file
            key: (items.length + 1).toString(),
            title: documentItem.Title,
            name: documentItem.Name,
            value: documentItem.Name,
            iconPath: "",
            fileType: documentItem.ListItemAllFields.FieldValuesAsText.File_x005f_x0020_x005f_Type,
            modifiedBy: stringIsNullOrEmpty(documentItem.ListItemAllFields.FieldValuesAsText.Editor) ? "N/A" : documentItem.ListItemAllFields.FieldValuesAsText.Editor,
            dateModified: date.toLocaleDateString(),
            dateModifiedValue: FormateDate,
            region: stringIsNullOrEmpty(documentItem.ListItemAllFields.Region) ? "N/A" : documentItem.ListItemAllFields.Region,
            programType: stringIsNullOrEmpty(documentItem.ListItemAllFields.Program_x0020_Type) ? "N/A" : documentItem.ListItemAllFields.Program_x0020_Type,
            FileRef: documentItem.ServerRelativeUrl,
            id: documentItem.ListItemAllFields.Id,
            siteLocation: !stringIsNullOrEmpty(siteLocation_Code) ?
              siteLocation_Location == "All" ? siteLocation_Code :
                siteLocation_Location + " - " + siteLocation_Code : "N/A",
            siteLocationId: documentItem.ListItemAllFields.Site_x0020_Location_x0020_CodeId > 0 ? documentItem.ListItemAllFields.Site_x0020_Location_x0020_CodeId : 0
          });

        });

      }

    }).catch((error: any) => {
      console.log(error);
      this.setState({
        allDocumentitems: items,
        filtterDocumentitems: items,
        searchDocumentitems: items,
        totalItemsCount: items.length,
        errorMessage: CONSTANTS.SYS_CONFIG.DATA_ERROR,
        showLoader: false,
        displayNoDataMassage: true,
      });

    });

    if (items.length > 0) {
      let gridData: any = this.paging(items, this.state.selectedItemPerPage.key);
      this.setState({
        allDocumentitems: items,
        allFormattedItems: gridData,
        filtterDocumentitems: gridData[0],
        searchDocumentitems: items,
        totalItemsCount: items.length,
        counter: 0,
        showLoader: false,
        selectedItemPerPage: { key: ItemPerPageDropdown[0].key, text: ItemPerPageDropdown[0].text },
        displayNoDataMassage: true,
        errorMessage: ""
      });
    } else {
      this.setState({
        allDocumentitems: items,
        filtterDocumentitems: items,
        searchDocumentitems: items,
        totalItemsCount: items.length,
        displayNoDataMassage: true,
        showLoader: false,
        errorMessage: CONSTANTS.SYS_CONFIG.NO_DATA_FOUND_ERROR_MESSAGE

      });
    }
    this.resetIsSortedColumn();
    this.setGlobalCssChanges();
  }

  private resetIsSortedColumn = (): void => {
    if (this.state.siteLocationMetadatAvailable == true) {
      const { columnsSiteLocation } = this.state;
      this.setState({
        columnsSiteLocation: columnsSiteLocation!.map(col => {
          col.isSorted = false;
          return col;
        })
      });
    } else {
      const { columns } = this.state;
      this.setState({
        columns: columns!.map(col => {
          col.isSorted = false;
          return col;
        })
      });
    }

  }


  private buildBreadCrumb = (folderServerRelativepath: string): void => {
    let BreadcrumbItems: IBreadcrumbItem[] = [];
    let FolderStringArr: any[] = folderServerRelativepath.split('/');
    let BreadCrumbKey: string = "";
    //Get the serverRelativeUrl path split index to exclude from the breadcrumb
    let skipIndexToServerRelativeUrl: number = this.props.context.pageContext.site.serverRelativeUrl.split("/").length - 1;
    FolderStringArr.forEach((bItem: any, index: number) => {

      //debugger;
      if (index !== 0) {
        BreadCrumbKey = BreadCrumbKey + "/" + bItem;
        //used index > skipIndexToServerRelativeUrl to skip. e.g. "sites/SafetyHub"
        if (index > skipIndexToServerRelativeUrl) {
          BreadcrumbItems.push({
            text: index == (skipIndexToServerRelativeUrl + 1) ? this.state.selectedLeftNavigationItem[0].LibraryTitle : bItem,
            key: BreadCrumbKey,
            isCurrentItem: FolderStringArr.length === (index + 1) ? true : false,
            onClick: this._onBreadcrumbItemClicked
          });
        }
      }
    });

    this.setState({
      BreadCrumbItems: BreadcrumbItems
    });
  }

  private _onBreadcrumbItemClicked = (ev: React.MouseEvent<HTMLElement>, item: IBreadcrumbItem): void => {

    //debugger;
    this.setState({
      folderServerRelativeUrl: item.key,
      showLoader: true
    });
    //Build the breadcrumb for selected folder
    this.buildBreadCrumb(item.key);
    //Load the folder/file data for selected folder
    this.getDocumentData(false, item.key);
  }


  private LoadRegion = (): void => {
    RegionOptions = [];
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
      let queryStringParameters = new URLSearchParams(window.location.search);

      let selectedRegion: any[] = [];
      if (queryStringParameters.get("rgn")) {
        selectedRegion = _.filter(RegionOptions, (p) => {
          return p.key == queryStringParameters.get("rgn");
        });
      }

      if (selectedRegion.length > 0) {
        this.setState({
          regionOptions: RegionOptions,
          selectedRegionValue: { key: selectedRegion[0].key, text: selectedRegion[0].text }
        });
      } else {
        this.setState({
          regionOptions: RegionOptions,
          selectedRegionValue: { key: "Select", text: "Select" }
        });
      }

      this.LoadProgramType();

    });
  }

  private setGlobalCssChanges = (): void => {

    $(".CanvasZone").css('padding-right', '38px');
    $(".CanvasZone > div").css('max-width', '100%');

    $(".ms-SPLegacyFabricBlock").css('background-color', '#1a458a');

    $(".Canvas--withLayout > div > div > div > div:nth-child(1)").attr('style', 'background-color: #1a458a; padding: 0px; width: 22% !important;');
    $(".Canvas--withLayout > div > div").attr('style', 'padding: 0px;');

    $(".ControlZone").css('margin-top', '0px');
    $(".ControlZone").css('margin-bottom', '0px');

    $(".Canvas--withLayout > div > div > div > div:nth-child(2)").attr('style', 'background-color: #1a458a; padding: 0px; width: 78% !important;');
    $(".webPartContainer").attr('style', 'background: #f0f0f7; display: none;');

  }
  private LoadProgramType = (): void => {

    ProgramTypeOptions = [];
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

      let queryStringParameters = new URLSearchParams(window.location.search);
 
      let selectedProgramType: any[] = [];
    
      if (queryStringParameters.get("pty")) {
        selectedProgramType = _.filter(ProgramTypeOptions, (p) => {
          return p.key == queryStringParameters.get("pty");
        });
      }

      if (selectedProgramType.length > 0) {
        this.setState({
          programTypeOptions: ProgramTypeOptions,
          selectedProgramTypeValue: { key: selectedProgramType[0].key, text: selectedProgramType[0].text }
        });
      } else {
        this.setState({
          programTypeOptions: ProgramTypeOptions,
          selectedProgramTypeValue: { key: "Select", text: "Select" }
        });
      }

      this.LoadSiteLocation();

    });
  }

  private LoadSiteLocation = (): void => {
    SiteLocationOptions = [];

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
      this.getDocumentData(false, this.state.folderServerRelativeUrl);
      //this.setGlobalCssChanges();
    });
  }


  private loadSiteLocationByRegion = (SelectedRegionItem: any): void => {
    let SiteLocationOptionsData: any = [];
    SiteLocationOptions = [];
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
  private getFilterData = (): IDocument[] => {
    let filterDocumentItems: IDocument[] = this.state.allDocumentitems;

    if (this.state.selectedRegionValue.text != "Select") {
      filterDocumentItems = _.filter(filterDocumentItems, (p) => {
        return p.region == this.state.selectedRegionValue.text;
      });
    }
    if (this.state.selectedProgramTypeValue.text != "Select") {
      filterDocumentItems = _.filter(filterDocumentItems, (p) => {
        return p.programType == this.state.selectedProgramTypeValue.text;
      });
    }
    if (this.state.selectedSiteLocationValue.text != "Select") {
      filterDocumentItems = _.filter(filterDocumentItems, (p) => {
        return p.siteLocationId == this.state.selectedSiteLocationValue.key;
      });
    }
    if (this.state.modifiedDate != null) {
      filterDocumentItems = _.filter(filterDocumentItems, (p) => {
        return p.dateModifiedValue == Moment(this.state.modifiedDate).format("MM/DD/YYYY");
      });
    }
    if (this.state.modifiedBy.length > 0) {
      filterDocumentItems = _.filter(filterDocumentItems, (p) => {
        return p.modifiedBy == this.state.modifiedByName;
      });
    }
    return filterDocumentItems;
  }


  private searchData = (): void => {

    let items: IDocument[] = [];

    this.setState({
      showLoader: true,
      isSearchClicked: true,
      searchTextValue: "",
      folderServerRelativeUrl: rootPath
    });

    this.buildBreadCrumb(rootPath);

    if (this.state.rgnptyIsAsAll) {

      items = this.getFilterData();
      let gridData: any = this.paging(items, this.state.selectedItemPerPage.key);
      this.setState({
        allFormattedItems: gridData,
        filtterDocumentitems: gridData[0],
        searchDocumentitems: items,
        totalItemsCount: items.length,
        counter: 0,
        showLoader: false,
        selectedItemPerPage: { key: this.state.selectedItemPerPage.key, text: this.state.selectedItemPerPage.text },
        displayNoDataMassage: true,
        errorMessage: CONSTANTS.SYS_CONFIG.NO_DATA_FOUND_ERROR_MESSAGE
      });

    } else {

      let rooUrl: string = `${this.props.context.pageContext.site.serverRelativeUrl}/${this.state.selectedLeftNavigationItem[0].LibraryName}`;
      this.setState({ folderServerRelativeUrl: rooUrl });
      this.getDocumentData(false, rooUrl);

    }



  }

  private clearFilter = (): void => {

    let emptyUser: any = [];

    this.setState({
      selectedRegionValue: { key: "Select", text: "Select" },
      selectedProgramTypeValue: { key: "Select", text: "Select" },
      selectedSiteLocationValue: { key: "Select", text: "Select" },
      modifiedByName: "",
      modifiedDate: null,
      showLoader: true,
      isSearchClicked: false,
      searchTextValue: "",
      modifiedBy: emptyUser,
      folderServerRelativeUrl: rootPath
    });
    this.buildBreadCrumb(rootPath);
    this.loadSiteLocationByRegion({ key: "Select", text: "Select" });

    if (this.state.rgnptyIsAsAll) {

      let gridData: any = this.paging(this.state.allDocumentitems, this.state.selectedItemPerPage.key);
      this.setState({
        allFormattedItems: gridData,
        filtterDocumentitems: gridData[0],
        searchDocumentitems: this.state.allDocumentitems,
        totalItemsCount: this.state.allDocumentitems.length,
        counter: 0,
        showLoader: false,
        selectedItemPerPage: { key: this.state.selectedItemPerPage.key, text: this.state.selectedItemPerPage.text },
        displayNoDataMassage: true
      });

    } else {
      this.getDocumentData(true, rootPath);
    }
  }

  public _onSelectEndDate = (date: Date | null | undefined): void => {
    this.setState({ modifiedDate: date });
  }

  public onFormatDate = (date?: Date): string => {
    return !date ? '' : date.toLocaleDateString();//.toLocaleString();
  }

  public getModifiedBy = (items: any[]): void => {
    // debugger;
    let getSelectedUsers: any = [];

    items.forEach(User => {
      getSelectedUsers.push({ ID: User.id, LoginName: User.loginName });
    });

    this.setState({
      modifiedBy: getSelectedUsers,
      modifiedByName: items.length > 0 ? items[0].text : ""
    });

  }

  private paging = (allData: any, pItemPerPage: any): any => {
    let data: any = [];
    let itemsPerPage: any;
    if (pItemPerPage != null && pItemPerPage > 0) {
      itemsPerPage = pItemPerPage;
    } else {
      itemsPerPage = this.state.selectedItemPerPage.key;
    }
    if (this.state.counter === 0 && allData.length > 0) {
      let itemLimit: any = allData.length > itemsPerPage ? itemsPerPage : allData.length;
      data = chunk(allData, itemLimit);
    }
    else if (allData.length > 0) {
      {
        let itemLimit: any = allData.length > itemsPerPage ? itemsPerPage : allData.length;
        data = chunk(allData, itemLimit);
      }

    }
    return data;
  }

  private loadPages = (clikedPageNumber: any) => {
    let pageId: number = clikedPageNumber - 1;
    this.setState({
      filtterDocumentitems: this.state.allFormattedItems[pageId],
      counter: pageId
    });
    //this._onColumnClick;
  }

  public onItemPerPageChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {

    if (item.text != "") {

      this.setState({
        selectedItemPerPage: item
      });

      if (this.state.isSearchClicked) {

        let gridData: any = this.paging(this.state.searchDocumentitems, item.key);
        this.setState({
          allFormattedItems: gridData,
          filtterDocumentitems: gridData[0],
          totalItemsCount: this.state.searchDocumentitems.length,
          counter: 0,
          //showLoader: false,
          selectedItemPerPage: { key: item.key, text: item.text },
          displayNoDataMassage: true,
          searchTextValue: ""
        });

      } else {

        let gridData: any = this.paging(this.state.allDocumentitems, item.key);
        this.setState({
          filtterDocumentitems: gridData[0],
          counter: 0,
          totalItemsCount: this.state.allDocumentitems.length,
          allFormattedItems: gridData,
          selectedItemPerPage: { key: item.key, text: item.text },
          displayNoDataMassage: true,
          searchTextValue: ""
        });

      }
      //let column = await this.getIsSortedColumn(this.state.sortColoumn);
      this.sortDocumentData(this.state.sortColoumn, this.state.sortColoumn.isSortedDescending, item.key);
      //this.resetIsSortedColumn();
    }
  }

  private _onDragOverDocument(e) {
    e.preventDefault();
    e.stopPropagation();
    this.hideShowDragDropPanel(true);
    return false;
  }

  private _OnDragLeaveDocument(e) {
    if (e.target.id == "dvDocument" || e.target.id == "dropZone_Document") {
      this.hideShowDragDropPanel(false);
    }
  }

  private hideShowDragDropPanel = (show: boolean): void => {
    if (show) {
      $("#dropZone_Document").removeClass("drop-zone-hide");
      $("#dropZone_Document").addClass("drop-zone-show");
    } else {
      $("#dropZone_Document").removeClass("drop-zone-show");
      $("#dropZone_Document").addClass("drop-zone-hide");
    }
  }

  private _onDrop(e) {
   
    e.preventDefault();
    let files = e.dataTransfer.files;
    var dropTarget = e.target.id;
    if (dropTarget == null && dropTarget == undefined && dropTarget == "") {
      this.hideShowDragDropPanel(false);
      alert("Please try again!");
      return false;
    }
    if (this.state.isCurrentUserPresentInGroup == true) {
      var filename = "";
      for (var i = 0; i < files.length; i++) {
        var file = files[i];
        var dropFile = file.name;
        filename = dropFile.split('.').slice(0, -1).join('.');
        this.setState({
          documentTitle: filename
        });
      }
    
      this.setState({
        documentsToBeupload: files,
      });
    
      let selectedRegion = this.state.selectedRegionValue == undefined ? "Select" : this.state.selectedRegionValue.text;
      let selectedProgramType = this.state.selectedProgramTypeValue == undefined ? "Select" : this.state.selectedProgramTypeValue.text;
      let selectedSiteLocation = this.state.selectedSiteLocationValue == undefined ? "Select" : this.state.selectedSiteLocationValue.key;
      let selectedSearch: string = "";
    
      if (this.state.siteLocationMetadatAvailable) {
    
        if (selectedRegion == "Select") {
          selectedSearch = " region ";
        }
    
        if (selectedProgramType == "Select") {
          selectedSearch = selectedSearch == "" ? " program type " : " region and program type ";
        }
        
        if (selectedSiteLocation == "Select") {
          selectedSearch = selectedSearch == "" ? " site location " : selectedSearch.replace(' and', ',') + " and site location ";
        }
    
        if (selectedSearch != "") {
          this.setState({ isMetadatDialogOpen: true, DialogMessage: CONSTANTS.SYS_CONFIG.DOCUMENT_UPLOAD_VALIDATION.replace('${selector}', selectedSearch) });
          this.hideShowDragDropPanel(false);
          return false;
        }
  
      } else {
        if (selectedRegion == "Select") {
          selectedSearch = " region ";
        }

        if (selectedProgramType == "Select") {
          selectedSearch = selectedSearch == "" ? " program type " : " region and program type ";
        }

        if (selectedSearch != "") {
          this.setState({ isMetadatDialogOpen: true, DialogMessage: CONSTANTS.SYS_CONFIG.DOCUMENT_UPLOAD_VALIDATION.replace('${selector}', selectedSearch) });
          this.hideShowDragDropPanel(false);
          return false;
        }
      }
  
      //Open dialog forcefully to select or update the document upload metadata for user
      this.setState({ isMetadatDialogOpen: true, DialogMessage: "" });
      this.hideShowDragDropPanel(false);
      return false;
  
    }else {
      this.setState({
        isOpenNoUploadPermissionDialog: true,
        noUploadPermissionDialogMessage: CONSTANTS.SYS_CONFIG.DOCUMENT_UPLOAD_NOPERMISSION_MESSAGE,
        rgnptyIsAsAll: false
      });
    }
  }

  public async setUploadParams(docLib, docTitle, region, programType, siteLocation, fileArray) {
      if (fileArray.length > 0 && docLib != "") {
        // debugger;
        let isMultipleFileUploaded: boolean = false;
        let uploadFailedFileColl: any[] = [];
  
        if (fileArray.length > 1) {
          isMultipleFileUploaded = true;
        }
  
        for (var i = 0; i < fileArray.length; i++) {
  
          var file = fileArray[i];
          var iteration = i + 1;
  
          if (isMultipleFileUploaded) {
            docTitle = file.name.split('.').slice(0, -1).join('.');
          }
  
          //console.log(docTitle);
  
          try {
            let fileAddedResult = await commonService.uploadDocument(docLib, file.name, file);
            let fileItem = await commonService.getFileItem(fileAddedResult);
            let fileUpdateResult: any;
  
            if (siteLocation == "Select") {
              fileUpdateResult = await commonService.updateDocumentProperties(fileItem, docTitle, region, programType, this.props.context.pageContext.legacyPageContext["userId"]);
            } else {
              fileUpdateResult = await commonService.updateDocumentPropertiesWithSiteLocation(fileItem, docTitle, region, programType, siteLocation, this.props.context.pageContext.legacyPageContext["userId"]);
            }
          }
          catch (error) {
            uploadFailedFileColl.push({
              errorFileName: file.name,
              error: error.message
            });
            console.log(uploadFailedFileColl);
  
          }
  
          if (iteration == fileArray.length) {
            //Check is document duplicate and get error
            if (uploadFailedFileColl.length > 0) {
  
              //Show message if all uploaded documents get error
              if (uploadFailedFileColl.length === fileArray.length) {
                this.setState({ isOpenDialog: true, showLoader: false, DialogMessage: CONSTANTS.SYS_CONFIG.DOCUMENT_UPLOAD_FAILED_FOR_DUPLICATE_DOCUMENT_MESSAGE });
              } else {
                //Set the failed documents collection details 
                this.setState({
                  documentUploadDetails: {
                    DocumentFailedToUpload: uploadFailedFileColl,
                    FailedCount: uploadFailedFileColl.length,
                    SuccesCount: (fileArray.length - uploadFailedFileColl.length),
                    TotalCount: fileArray.length
                  }
                });
  
                this.getDocumentData(false, this.state.folderServerRelativeUrl);
                //open modal dailog    
                //Show error documents (already exsits documents name) list in modal dailog and message to user                        
                this.setState({ isModalOpen: true, rgnptyIsAsAll: false });
              }
            } else {
              this.getDocumentData(false, this.state.folderServerRelativeUrl);
              //Show message on sucess for all documents
              this.setState({
                isOpenDialog: true,
                DialogMessage: CONSTANTS.SYS_CONFIG.DOCUMENT_UPLOADED_MESSAGE,
                rgnptyIsAsAll: false
              });
            }
          }
        }
      }
      else {
        alert("An error occured. Please try again!");
      }
   
  }

  private closeDialog = (): void => {
    this.setState({ isOpenDialog: false });
  }

  private closeNoUploadPermissionDialog= (): void => {
    this.setState({ isOpenNoUploadPermissionDialog: false });
  }

  private closeMeatadataDialog = (): void => {
      this.setState({
        isMetadatDialogOpen: false,
        documentTitle: "",
        selectedRegionValue: { key: "Select", text: "Select" },
        selectedProgramTypeValue: { key: "Select", text: "Select" },
        selectedSiteLocationValue: { key: "Select", text: "Select" },
      });
  }

  private uploadDocuments = (): void => {
    this.setState({ showLoader: true });

    let documentTitle = this.state.documentTitle;
    let selectedRegion = this.state.selectedRegionValue == undefined ? "Select" : this.state.selectedRegionValue.text;
    let selectedProgramType = this.state.selectedProgramTypeValue == undefined ? "Select" : this.state.selectedProgramTypeValue.text;
    let selectedSiteLocation = this.state.selectedSiteLocationValue == undefined ? "Select" : this.state.selectedSiteLocationValue.key;

    //this.setUploadParams(selectedDocLib, selectedRegion, selectedProgramType, selectedSiteLocation, this.state.documentsToBeupload);
    this.setUploadParams(this.state.folderServerRelativeUrl, documentTitle, selectedRegion, selectedProgramType, selectedSiteLocation, this.state.documentsToBeupload);

    this.closeMeatadataDialog();
  }

  private updateDocuments = (selectedItemId): void => {
    this.setState({ isEditMetadataDialogOpen: false });
    let selectedLibraryName = this.state.selectedLeftNavigationItem[0].LibraryTitle;
    let documentTitle = this.state.documentTitle;
    let selectedRegion = this.state.EPSelectedRegionValue == undefined ? "Select" : this.state.EPSelectedRegionValue.text;
    let selectedProgramType = this.state.EPSelectedProgramTypeValue == undefined ? "Select" : this.state.EPSelectedProgramTypeValue.text;
    let selectedSiteLocation = this.state.EPSelectedSiteLocationValue == undefined ? "Select" : this.state.EPSelectedSiteLocationValue.key;
    
    //this.setUploadParams(selectedDocLib, selectedRegion, selectedProgramType, selectedSiteLocation, this.state.documentsToBeupload);
    if (selectedSiteLocation == "Select") {
      commonService.editDocumentProperties(selectedLibraryName, selectedItemId, documentTitle, selectedRegion, selectedProgramType).then((result: any) => {
        this.getDocumentData(false, this.state.folderServerRelativeUrl);
        return result;
      });
    } else {
      commonService.editDocumentPropertiesWithSiteLocation(selectedLibraryName, selectedItemId, documentTitle, selectedRegion, selectedProgramType, selectedSiteLocation).then((result: any) => {
        this.getDocumentData(false, this.state.folderServerRelativeUrl);
        return result;
      });
    }

    this.setState({
      documentTitle: "",
      EPSelectedRegionValue: { key: "Select", text: "Select" },
      EPSelectedProgramTypeValue: { key: "Select", text: "Select" },
      EPSelectedSiteLocationValue: { key: "Select", text: "Select" },
    });

  }


  private closeModalDialog = (): void => {
    this.setState({ isModalOpen: false });
  }

  private _onClearSearch = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
    console.log("On Clear");
  }

  private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, pText: string): void => {
    //debugger;
    let documentsToFilter: IDocument[] = [];
    let docItems: any[];
    //Check if search clicked
    if (!this.state.isSearchClicked) {
      documentsToFilter = this.state.allDocumentitems;
    } else {
      documentsToFilter = this.state.searchDocumentitems;
    }

    let text = "";
    if (pText != undefined) {
      text = pText.toLowerCase();
      //let selectedSubDomain = this.state.selectedSubDomain.text;


      //Check if site location metadata available or not and add filter depend on that
      if (this.state.siteLocationMetadatAvailable) {
        docItems = text ? _.filter(documentsToFilter, (item) => {
          return (item.name.toString().toLowerCase().indexOf(text) > -1 ||
            item.dateModifiedValue.toLowerCase().indexOf(text) > -1 ||
            item.modifiedBy.toLowerCase().indexOf(text) > -1 ||
            item.programType.toLowerCase().indexOf(text) > -1 ||
            item.siteLocation.toLowerCase().indexOf(text) > -1 ||
            item.region.toLowerCase().indexOf(text) > -1);
        }) : documentsToFilter;
      } else {
        docItems = text ? _.filter(documentsToFilter, (item) => {
          return (item.name.toString().toLowerCase().indexOf(text) > -1 ||
            item.dateModifiedValue.toLowerCase().indexOf(text) > -1 ||
            item.modifiedBy.toLowerCase().indexOf(text) > -1 ||
            item.programType.toLowerCase().indexOf(text) > -1 ||
            item.region.toLowerCase().indexOf(text) > -1);
        }) : documentsToFilter;
      }
    }
    else {
      docItems = documentsToFilter;
    }

    let gridData: any = this.paging(docItems, this.state.selectedItemPerPage.key);
    this.setState({
      filtterDocumentitems: gridData[0],
      totalItemsCount: docItems.length,
      allFormattedItems: gridData,
      searchTextValue: text,
      counter: 0,
      errorMessage: CONSTANTS.SYS_CONFIG.NO_DATA_FOUND_ERROR_MESSAGE
    });
    //this.sortDocumentData(this.state.sortColoumn, this.state.sortColoumn.isSortedDescending);
    this.resetIsSortedColumn();
  }

  public onSiteLocationChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    //this.setState({ SingleSelect: option.key });
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedSiteLocationValue: option,
        siteLocationDropDownError: false
      });
    } else {
      this.setState({
        ...this.state,
        selectedSiteLocationValue: { key: "Select", text: "Select" },
        siteLocationDropDownError: true
      });
    }
  }


  public onProgramTypeChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedProgramTypeValue: option,
        programTypeDropDownError: false
      });
    } else {
      this.setState({
        ...this.state,
        selectedProgramTypeValue: { key: "Select", text: "Select" },
        programTypeDropDownError: true
      });
    }
  }

  public onRegionChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option != undefined) {
      this.setState({
        ...this.state,
        selectedRegionValue: option,
        regionDropDownError: false
      });
      this.loadSiteLocationByRegion(option);
    } else {
      this.setState({
        ...this.state,
        selectedRegionValue: { key: "Select", text: "Select" },
        regionDropDownError: true
      });
    }
  }

  private openFileUploadModalDialog = (): void => {
    this.setState({ isFileUploadModalOpen: true });
  }

  private closeFileUploadModalDialog = (): void => {
    this.setState({ isFileUploadModalOpen: false, file: [], fileName: "", fileSize: "", errorMessage: "" });
  }


  //to validate and set file detail to state on file drop
  private _onFileDrop = (files: File[], fileRejection: FileRejection[], event: DropEvent): void => {
    try {
      //g_attachErrorClass = 'noErrMessage';
      this.setState({ file: [], fileName: "", fileSize: "", errorMessage: "" });
      var fileName: string;
      var fileSize: any;
      var file: any = [];
      if (files.length > 0) {
        //to check file size in more than 10MB
        if (files[0].size > 9999999) {
          files.map(selectedFile => {
            fileName = selectedFile.name,
              fileSize = this._formatFileSize(selectedFile.size);
          });
          this.setState({
            file: files,
            fileName: fileName,
            fileSize: fileSize,
            errorMessage: "File size is too large. Please upload a file less than 10MB."
          });
        }
        else {
          files.map(selectedFile => {
            fileName = selectedFile.name,
              fileSize = this._formatFileSize(selectedFile.size);
          });
          this.setState({
            file: files,
            fileName: fileName,
            fileSize: fileSize
          });
        }
      }
    } catch (error) {
      console.log(error);
    }
  }

  private _formatFileSize = (fileSize: number): any => {
    try {
      if (fileSize == 0) return '0 Bytes';
      var k = 1000;
      var dm = 2;
      var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
      var i = Math.floor(Math.log(fileSize) / Math.log(k));
      return parseFloat((fileSize / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    } catch (error) {
      console.log(error);
    }
  }

  private onDocumentTitleChange = (e): void => {
    this.setState({
      ...this.state,
      [e.target.name]: e.target.value,
    });
  }

  private openEditMetadataDialog = (): void => {
    this.setState({ isEditMetadataDialogOpen: true });
  }
  private closeEditMetadataDialog = (): void => {
    this.setState({ 
      documentTitle: "",
      EPSelectedRegionValue: { key: "Select", text: "Select" },
      EPSelectedProgramTypeValue: { key: "Select", text: "Select" },
      EPSelectedSiteLocationValue: { key: "Select", text: "Select" },
      isEditMetadataDialogOpen: false,
     });
  }

  private loadOnEPSiteLocationByRegion = (SelectedRegionItem: any): void => {
    let SiteLocationOptionsData: any = [];
    SiteLocationOptions = [];
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
      EPSelectedSiteLocationValue: { key: "Select", text: "Select" }
    });

  }

  public onEditSiteLocationChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    //this.setState({ SingleSelect: option.key });
    if (option != undefined) {
      this.setState({
        ...this.state,
        EPSelectedSiteLocationValue: option,
        siteLocationDropDownError: false
      });
    } else {
      this.setState({
        ...this.state,
        EPSelectedSiteLocationValue: { key: "Select", text: "Select" },
        siteLocationDropDownError: true
      });
    }
  }


  public onEditProgramTypeChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option != undefined) {
      this.setState({
        ...this.state,
        EPSelectedProgramTypeValue: option,
        programTypeDropDownError: false
      });
    } else {
      this.setState({
        ...this.state,
        EPSelectedProgramTypeValue: { key: "Select", text: "Select" },
        programTypeDropDownError: true
      });
    }
  }

  public onEditRegionChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    if (option != undefined) {
      this.setState({
        ...this.state,
        EPSelectedRegionValue: option,
        regionDropDownError: false
      });
      this.loadOnEPSiteLocationByRegion(option);
    } else {
      this.setState({
        ...this.state,
        EPSelectedRegionValue: { key: "Select", text: "Select" },
        regionDropDownError: true
      });
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

  public render(): React.ReactElement<IDocumentSearchProps> {
    const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 291 } };
    return (
      <Stack>
        {this.state.showLoader ? <Loader></Loader> : ""}

        <div className={styles.documentSearch}>

          <Stack className={styles.webpartHeaderDiv}>
            <div>
              {this.props.webpartTitle}
            </div>
          </Stack>

          <Stack className={styles.filterParentDiv} >
            <Stack.Item>
              <div>
                <div className={styles.programTypeDiv}>
                  <ComboBox
                    label="PROGRAM TYPE"
                    allowFreeform={true}
                    autoComplete={'on'}
                    options={this.state.programTypeOptions}
                    onChange={this.onProgramTypeChange}
                    selectedKey={this.state.selectedProgramTypeValue ? this.state.selectedProgramTypeValue.key : "Select"}
                    className="siteLocationActive"
                    errorMessage={this.state.programTypeDropDownError == true ? "Please select valid program type." : ""}
                  />
                </div>

                <div className={styles.regionDiv}>

                  <ComboBox
                    label="REGION"
                    allowFreeform={true}
                    autoComplete={'on'}
                    options={this.state.regionOptions}
                    onChange={this.onRegionChange}
                    selectedKey={this.state.selectedRegionValue ? this.state.selectedRegionValue.key : "Select"}
                    errorMessage={this.state.regionDropDownError == true ? "Please select valid region." : ""}

                  />
                </div>
                {this.state.siteLocationMetadatAvailable == true ?
                  <div className={styles.SiteLocationDiv}>
                    <ComboBox
                      label="BY SITE LOCATION"
                      allowFreeform={true}
                      autoComplete={'on'}
                      options={this.state.siteLocationOptions}
                      onChange={this.onSiteLocationChange}
                      selectedKey={this.state.selectedSiteLocationValue ? this.state.selectedSiteLocationValue.key : "Select"}
                      errorMessage={this.state.siteLocationDropDownError == true ? "Please select valid site location." : ""}
                      //disabled={this.state.siteLocationMetadatAvailable == false ? true : false}
                      //className={this.state.siteLocationMetadatAvailable == false ? "siteLocationDisabled" : "siteLocationActive"}
                      className="siteLocationActive"
                    //styles={comboBoxStyles}
                    // Force re-creating the component when the toggles change (for demo purposes)
                    //key={'' + autoComplete + allowFreeform}
                    />
                  </div>
                  : ""
                }

              </div>
            </Stack.Item>
            <Stack.Item>
              <div className={styles.dateParentDiv}>
                <div className={styles.ModifiedDateDiv}>
                  <DatePicker
                    className={styles.modifiedDate}
                    title="Modified Date"
                    firstDayOfWeek={DayOfWeek.Monday}
                    placeholder="Modified Date"
                    ariaLabel="Modified Date"
                    label="MODIFIED DATE"
                    onSelectDate={this._onSelectEndDate}
                    formatDate={this.onFormatDate}
                    isMonthPickerVisible={false}
                    //defaultValue={this.state.modifiedDate}
                    value={this.state.modifiedDate}
                  />
                </div>
                <div className={styles.ModifiedByDiv}>
                  <PeoplePicker
                    peoplePickerCntrlclassName={styles.modifiedBy}
                    context={this.props.context}
                    titleText="MODIFIED BY"
                    personSelectionLimit={1}
                    showtooltip={true}
                    required={false}
                    disabled={false}
                    onChange={this.getModifiedBy}
                    defaultSelectedUsers={[this.state.modifiedByName]}
                    showHiddenInUI={false}
                    ensureUser={true}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}

                  />
                </div>
                <div className={styles.FreeTextSearchDiv}>
                  <div>

                  </div>
                  <div>

                  </div>
                </div>
              </div>

            </Stack.Item>
            <Stack.Item>
              <div className={styles.buttonParendDiv}>
                <div className={styles.searchButtonParentDiv}>
                  <PrimaryButton className={styles.searchButton} onClick={() => this.searchData()} text="Search" />
                </div>
                <div className={styles.clearFilterButtonParentDiv}>
                  <PrimaryButton className={styles.clearFilterButton} onClick={() => this.clearFilter()} text="Clear Filter" />
                </div>

              </div>
            </Stack.Item>
          </Stack>


          <Stack className={styles.navHeaderParentDiv} >

            <Stack.Item grow>
              <div>
                <div className={styles.navHeaderDiv}>
                  {this.state.selectedLeftNavigationItem.length > 0 ? this.state.selectedLeftNavigationItem[0].Title : ""}
                </div>
                <div className={styles.pagingDiv}>
                  <div className={styles.pagingInnerDiv}>
                    <div className="paginationControl">

                      <div className={styles.itemPerPage_div}>

                        {this.state.totalItemsCount > 0 ?
                          <Dropdown
                            id="ItemPerPage"
                            selectedKey={this.state.selectedItemPerPage ? this.state.selectedItemPerPage.key : undefined}
                            onChange={this.onItemPerPageChange}
                            options={ItemPerPageDropdown}
                            className={styles.itemPerPageDropdown}
                          /> : ""
                        }
                      </div>
                      <Pagination
                        activePage={parseInt(this.state.counter.toString()) + 1}
                        itemsCountPerPage={this.state.selectedItemPerPage.key}
                        totalItemsCount={this.state.totalItemsCount}
                        pageRangeDisplayed={CONSTANTS.SYS_CONFIG.PAGE_RANGE_DISPLAYED}
                        onChange={this.loadPages.bind(this)}
                        hideDisabled
                      />
                    </div>


                  </div>
                </div>
              </div>

            </Stack.Item>
            <Stack.Item grow>
              <Text>Drag drop one or more files to upload</Text>
             
              <SearchBox
                placeholder="Search by keyword"
                //onChange={()=> this._onChangeText(this.newValue)}
                onChange={this._onChangeText}
                onClear={this._onClearSearch}
                disableAnimation
                value={this.state.searchTextValue}
                className={styles.searchText}
                ariaLabel="Text Search"
                title="Text Search"

              />

              <Breadcrumb
                items={this.state.BreadCrumbItems}
                maxDisplayedItems={5}
                ariaLabel="Breadcrumb with items rendered as buttons"
                overflowAriaLabel="More links"
                className={this.state.BreadCrumbItems.length > 1 ? styles.ShowBreadcrumb : styles.HideBreadcrumb}
              />

            </Stack.Item>

            <Stack.Item grow>
              <div id="dvDocument" title="Drop your files here to upload" onDragOver={this._onDragOverDocument} onDragLeave={this._OnDragLeaveDocument} onDrop={this._onDrop} className={styles.tableBoxBG}>
                {this.state.isCurrentUserPresentInGroup == true ?
                  <div id="dropZone_Document" className={styles.dropzone + ' ' + 'drop-zone-hide'}><h3>Drop your files here</h3></div> : ''
                }
              
                {
                  this.state.filtterDocumentitems && this.state.filtterDocumentitems.length > 0 ?
                    <DetailsList
                      items={this.state.filtterDocumentitems.length > 0 ? this.state.filtterDocumentitems : []}
                      //columns={this.state.siteLocationMetadatAvailable this.state.columns}
                      columns={this.state.siteLocationMetadatAvailable == true ? this.state.columnsSiteLocation : this.state.columns}
                      selection={this._selection}
                      layoutMode={DetailsListLayoutMode.justified}
                      //selectionMode={SelectionMode.none}
                      isHeaderVisible={true}
                      selectionMode={SelectionMode.single} // controls how/if list manages selection ( non, single, multiple)

                      checkboxVisibility={2} //0 = on hover, 1 = always, 2 = hidden

                    />
                    : <div className={styles.noDataFound}>  {this.state.displayNoDataMassage == true ? this.state.errorMessage : ""}</div>
                }

              </div>
            </Stack.Item>



          </Stack>


        </div>

        <Dialog
          isOpen={this.state.isOpenDialog}
          type={DialogType.close}
          onDismiss={this.closeDialog}
          subText={this.state.DialogMessage}
          isBlocking={false}
          closeButtonAriaLabel='Close'
        >

          <DialogFooter>
            <PrimaryButton onClick={this.closeDialog} text="Ok" />
          </DialogFooter>
        </Dialog>


        <Dialog
          isOpen={this.state.isOpenNoUploadPermissionDialog}
          type={DialogType.close}
          onDismiss={this.closeNoUploadPermissionDialog}
          subText={this.state.noUploadPermissionDialogMessage}
          isBlocking={false}
          closeButtonAriaLabel='Close'
          containerClassName="NoPermissionMsgDialog"
        >

          <DialogFooter>
            <PrimaryButton onClick={this.closeNoUploadPermissionDialog} text="Ok" />
          </DialogFooter>
        </Dialog>


        <Modal
          titleAriaId="Modal"
          isOpen={this.state.isModalOpen}
          onDismiss={this.closeModalDialog}
          isBlocking={false}
          containerClassName={styles.container}
        >
          <div className="modalheader">
            <span id="Modal">Document Upload</span>
            <IconButton
              className="modalCloseIcon"
              iconProps={cancelIcon}
              ariaLabel="Close popup modal"
              onClick={this.closeModalDialog}
            />
          </div>
          <div className="modalbody">

            <div className="modalBodyDiv">
              {CONSTANTS.SYS_CONFIG.DOCUMENT_UPLOAD_FAILED_MODAL_MESSAGE_1.replace('${TotalCount}', this.state.documentUploadDetails.TotalCount.toString()).replace('${SuccesCount}', this.state.documentUploadDetails.SuccesCount.toString())}

            </div>
            <div className="modalBodyDiv">
              {CONSTANTS.SYS_CONFIG.DOCUMENT_UPLOAD_FAILED_MODAL_MESSAGE_2.replace('${FailedCount}', this.state.documentUploadDetails.FailedCount.toString())}

            </div>
            <div className="modalBodyFileMainDiv">
              {this.state.documentUploadDetails.DocumentFailedToUpload.map((file: any, i: number) => {
                return (
                  <div>
                    {/* <p>{file.errorFileName} "  Error:   " {file.error}</p>        */}
                    <div className="modalBodyFileDiv">{file.errorFileName} </div>
                  </div>
                );

              })

              }
            </div>
          </div>
          <div className="modalFotter">
            <DefaultButton className="" onClick={this.closeModalDialog} text="Ok" />
          </div>
        </Modal>

        <Dialog
          isOpen={this.state.isMetadatDialogOpen}
          type={DialogType.close}
          onDismiss={this.closeDialog}
          subText={this.state.DialogMessage}
          isBlocking={false}
          closeButtonAriaLabel='Close'
          dialogContentProps={this.dialogContentProps}
          modalProps={this.modelProps}

        >
          {this.state.documentsToBeupload.length == 1 ?
            <TextField
              name="documentTitle"
              label="TITLE"
              required={false} 
              placeholder="Enter document title "
              value={this.state.documentTitle}
              onChange={this.onDocumentTitleChange}
              autoComplete="off"
              styles={textFieldStyles}
            />
            : ""
          }

          <ComboBox
            label="REGION"
            allowFreeform={true}
            autoComplete={'on'}
            options={this.state.regionOptions}
            onChange={this.onRegionChange}
            selectedKey={this.state.selectedRegionValue ? this.state.selectedRegionValue.key : "Select"}
            errorMessage={this.state.regionDropDownError == true ? "Please select valid region." : ""}
          />

          <ComboBox
            label="PROGRAM TYPE"
            allowFreeform={true}
            autoComplete={'on'}
            options={this.state.programTypeOptions}
            onChange={this.onProgramTypeChange}
            selectedKey={this.state.selectedProgramTypeValue ? this.state.selectedProgramTypeValue.key : "Select"}
            className="siteLocationActive"
            errorMessage={this.state.programTypeDropDownError == true ? "Please select valid program type." : ""}
          />
        
          <ComboBox
            label="BY SITE LOCATION"
            allowFreeform={true}
            autoComplete={'on'}
            options={this.state.siteLocationOptions}
            onChange={this.onSiteLocationChange}
            selectedKey={this.state.selectedSiteLocationValue ? this.state.selectedSiteLocationValue.key : "Select"}
            errorMessage={this.state.siteLocationDropDownError == true ? "Please select valid site location." : ""}
            disabled={this.state.siteLocationMetadatAvailable == false ? true : false}
            className={this.state.siteLocationMetadatAvailable == false ? "siteLocationDisabled" : "siteLocationActive"}
          />

          <DialogFooter>
            <PrimaryButton className={styles.uploadButton} disabled={this.state.documentTitle != "" && this.state.selectedRegionValue.text != "Select" &&
              this.state.selectedProgramTypeValue.text != "Select" ?
              this.state.siteLocationMetadatAvailable == true && this.state.selectedSiteLocationValue.text == "Select" ?
                true : false
              : true}
              onClick={this.uploadDocuments} text="Upload" 
            />
          
           <DefaultButton className={styles.CancelButton} onClick={this.closeMeatadataDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>

        <Stack>
          <Stack.Item>
            <Modal
              titleAriaId="Modal"
              isOpen={this.state.isFileUploadModalOpen}
              onDismiss={this.closeModalDialog}
              isBlocking={false}
              containerClassName={styles.container}
            >
              <div className="fileModalheader">
                <span id="Modal">Document Upload</span>
                <IconButton
                  className="modalCloseIcon"
                  iconProps={cancelIcon}
                  ariaLabel="Close popup modal"
                  onClick={this.closeFileUploadModalDialog}
                />
              </div>
              {/* <hr className="hrStyle"></hr> */}
              <div className="fileModalbody">
                <div>
                  <Dropzone onDrop={this._onFileDrop} noDragEventsBubbling={true} multiple={false}>
                    {({ getRootProps, getInputProps }) => (
                      <section>
                        <div {...getRootProps()}>
                          <input {...getInputProps()} />
                          <div className={"fileUpload cssMarginBottom"} title={this.state.fileName ? this.state.fileName : "No file Chosen"}>
                            <DefaultButton text="Choose File"></DefaultButton>
                            <Label style={{ paddingLeft: '10px' }}>{this.state.fileName ? this.state.fileName : "No file Chosen"}</Label>
                          </div>
                        </div>
                      </section>
                    )}
                  </Dropzone>
                  {this.state.file.length > 0 ?
                    <div id="dvFileNamePicture" className="dvFileName cssMarginBottom" style={{ paddingBottom: '5px' }}>
                      <i className={"ms-Icon ms-Icon--KnowledgeArticle fileIcon"} />
                      <Label style={{ width: '80%' }}>{this.state.fileName ? this.state.fileName : ""}</Label>
                      <Label style={{ width: '20%', textAlign: 'right' }}>{this.state.fileSize ? this.state.fileSize : ""}</Label>
                    </div> : null}
                </div>
              </div>
              <div className="modalFotter">
                <PrimaryButton className="" text="Upload" />
                <DefaultButton className="cancelButton" onClick={this.closeFileUploadModalDialog} text="Cancel" />
              </div>
            </Modal>
          </Stack.Item>
        </Stack>

        <Stack>
          <Dialog
            isOpen={this.state.isEditMetadataDialogOpen}
            type={DialogType.close}
            onDismiss={this.closeEditMetadataDialog}
            subText={this.state.DialogMessage}
            isBlocking={false}
            closeButtonAriaLabel='Close'
            dialogContentProps={this.editDialogContentProps}
            modalProps={this.modelProps}

          >
            <TextField
              name="documentTitle"
              label="TITLE"
              placeholder="Enter document title "
              value={this.state.documentTitle}
              onChange={this.onDocumentTitleChange}
              autoComplete="off"
              styles={textFieldStyles}
            />
          
            <ComboBox
              label="REGION"
              allowFreeform={true}
              autoComplete={'on'}
              options={this.state.regionOptions}
              onChange={this.onEditRegionChange}
              selectedKey={this.state.EPSelectedRegionValue ? this.state.EPSelectedRegionValue.key : "Select"}
              errorMessage={this.state.regionDropDownError == true ? "Please select valid region." : ""}
            />
            
            <ComboBox
              label="PROGRAM TYPE"
              allowFreeform={true}
              autoComplete={'on'}
              options={this.state.programTypeOptions}
              onChange={this.onEditProgramTypeChange}
              selectedKey={this.state.EPSelectedProgramTypeValue ? this.state.EPSelectedProgramTypeValue.key : "Select"}
              className="siteLocationActive"
              errorMessage={this.state.programTypeDropDownError == true ? "Please select valid program type." : ""}
            />
           
            {this.state.siteLocationMetadatAvailable == true ?
              <ComboBox
                label="BY SITE LOCATION"
                allowFreeform={true}
                autoComplete={'on'}
                options={this.state.siteLocationOptions}
                onChange={this.onEditSiteLocationChange}
                selectedKey={this.state.EPSelectedSiteLocationValue ? this.state.EPSelectedSiteLocationValue.key : "Select"}
                errorMessage={this.state.siteLocationDropDownError == true ? "Please select valid site location." : ""}
                //disabled={this.state.siteLocationMetadatAvailable == false ? true : false}
                //className={this.state.siteLocationMetadatAvailable == false ? "siteLocationDisabled" : "siteLocationActive"}
              /> : "" 
            }

            <DialogFooter>
              <PrimaryButton className={styles.uploadButton} onClick={() => this.updateDocuments(this.state.selectedItemId)} text="Update" />
              <DefaultButton className={styles.CancelButton} onClick={this.closeEditMetadataDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </Stack>
      </Stack>
    );
  }

}
