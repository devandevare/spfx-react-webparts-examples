export default class CONSTANTS {
    public static LIST_NAME = {
        LEFT_NAVIGATION_CONFIGURATION_LIST: "LeftNavigation Configuration",
        SITE_LOCATION_CONFIGURATION: "Site Locations Configuration",
        AWARENESS_COMMUNICATIONS: "Awareness & Communications"        
    };

    public static SELECTCOLUMNS = {
        LEFT_NAVIGATION_COLS: "Id,Title,LibraryName,LibraryTitle,IsSiteLocationMetadataAvailable,Icon,ActiveIcon",
        EXPAND_LIST_COLS: "",
        SITE_LOCATION_CONFIGURATION: "Id,Title,Code,Region",
        //GET_DOCUMENTS: "Id,FileLeafRef,FileRef,File_x0020_Type,Region,Program_x0020_Type,Editor/Title,Editor/ID,Modified,Author/Title,Author/ID,Created,Site_x0020_Location_x0020_Code",
        GET_DOCUMENTS: "Id,FileLeafRef,FileRef,File_x0020_Type,Region,Program_x0020_Type,Editor/Title,Editor/ID,Modified,Author/Title,Author/ID,Created,Site_x0020_Location_x0020_Code/ID,Site_x0020_Location_x0020_Code/Code,Site_x0020_Location_x0020_Code/Title,FieldValuesAsText/Modified",
        PICTURE_GALLERY: "FileLeafRef,FileRef,Caption",
        VIDEO_GALLERY: "Video_x0020_URL,Video_x0020_Thumbnail,Region,Program_x0020_Type,Caption",
        EVENTS_LIST:"ID,Title,Attachments,Category,Description,Location,Region,EventDate,EndDate,FieldValuesAsText/EventDate,FieldValuesAsText/EndDate,Program_x0020_Type,fAllDayEvent,fRecurrence,Mark_x0020_as_x0020_Important",
        EXPAND_EVENTS_COLS:"FieldValuesAsText",
        ANNOUNCEMENTS_LIST:"AttachmentFiles/SeverRelativeUrl",
        EXPAND_ANNOUNCEMENT_COLS:"AttachmentFiles",
        DOCUMENTS_LIST: "ID,Title,Region,Program_x0020_Type",
        DOCUMENT_LIST_SITELOCATION: "ID,Title,Region,Program_x0020_Type,Site_x0020_Location_x0020_Code/ID,Site_x0020_Location_x0020_Code/Code,Site_x0020_Location_x0020_Code/Title",
        
    };

    public static ORDERBY = {
        LEFT_NAVIGATION: "Sequence",
        SITE_LOCATION: "Title",
        GET_DOCUMENT_ORDERBY: "ListItemAllFields/Modified",
        //GET_DOCUMENT_ORDERBY: "TimeLastModified",
        PICTURE_GALLERY: "Modified",
        VIDEO_GALLERY: "Modified",
        ANNOUNCEMENTS: "Modified",
        EVENTS: "Modified"
    };

    public static FILTERCONDITION = {
        LEFT_NAVIGATION_QUERY: "ShowInLeftNavigation eq '1'"   
    };

    public static COLULMN_NAME = {
        TITLE: "Title",
        LIBRARY_NAME: "LibraryName",
        LIBRARY_TITLE: "LibraryTitle",
        IS_SITE_LOCATION_METADATA_AVAILVABLE: "IsSiteLocationMetadataAvailable",
        ORDER_BY: "Modified",
        NAME: "Name",
        PIC_SIZE:"Picture Size",
        FILE_SIZE:"File Size",
        REGION: "Region",
        PROGRAM_TYPE: "Program Type",
        VIDEO_THUMBNAIL: "Video Thumbnail",
        VIDEO_URL:"Video URL",
        DESCRIPTION:"Description",
        MARK_AS_IMPORTANT:"Mark As Important",
        BANNER_URL:"Banner URL",
        ATTENDEES:"Attendees",
        GEOLOCATION:"Geolocation",
        CATEGORY: "Category"

    };

    public static LIST_VALIDATION_COLUMNS = {
        ANNOUNCEMENTS: ["Title", "Description", "Mark As Important", "Region", "Program Type"],
        LEFT_NAVIGATION: ["Title", "LibraryName", "LibraryTitle", "IsSiteLocationMetadataAvailable"],
        EVENTS: ["Banner URL", "Attendees", "Geolocation", "Region", "Program Type"],
        VIDEO: ["videolink", "Video Thumbnail", "Region", "Program Type"],
        PICTURE:["Name", "Picture Size", "File Size", "Region", "Program Type"],
    };

    public static ICONS = {
        PDF: "/SiteAssets/SafetyHubAssets/Images/Pdf.svg",
        WORLD: "/SiteAssets/SafetyHubAssets/Images/Word.svg",
        EXCEL: "/SiteAssets/SafetyHubAssets/Images/Excel.svg",
        TEXT: "/SiteAssets/SafetyHubAssets/Images/Text.svg",
        POWERPOINT: "/SiteAssets/SafetyHubAssets/Images/Powerpoint.svg",
        DEFAULT: "/SiteAssets/SafetyHubAssets/Images/Default.svg",
        DELETE: "/SiteAssets/SafetyHubAssets/Images/Delete.svg",    
        FOLDER: "/SiteAssets/SafetyHubAssets/Images/Folder.svg",
        PICTURE_GALLERY_IMAGES_NOT_FOUND: "/SiteAssets/SafetyHubAssets/Images/PictureGalleryImagesNotFound.png",
        VIDEO_GALLERY_IMAGES_NOT_FOUND: "/SiteAssets/SafetyHubAssets/Images/PictureGalleryImagesNotFound.png",
        VIDEO_GALLERY_DEFAULT_IMAGE: "/SiteAssets/SafetyHubAssets/Images/VideoPlayButton.png",
        LEFT_NAVIGATION_DEFAULT_ICON: "/SiteAssets/SafetyHubAssets/Images/LeftNavDefaultIcon.svg",
        LEFT_NAVIGATION_DEFAULT_ACTIVEICON: "/SiteAssets/SafetyHubAssets/Images/LeftNavDefaultActiveIcon.svg",
        LEFT_NAVIGATION_HOME_ICON: "/SiteAssets/SafetyHubAssets/Images/LeftNavHomeIcon.svg",
        LEFT_NAVIGATION_HOME_ACTIVEICON: "/SiteAssets/SafetyHubAssets/Images/LeftNavHomeActiveIcon.svg",
        LEFT_NAVIGATION_HUBWATERMARKLOGO: "/SiteAssets/SafetyHubAssets/Images/Hub logo watermark.SVG"
    };

    public static SYS_CONFIG = {
        NO_DATA_FOUND_ERROR_MESSAGE: "No data found for selected search criteria.",
        ITEMS_COUNT_PER_PAGE: "10",
        PAGE_RANGE_DISPLAYED: "5",
        DATA_ERROR : "Unexpected error occurred while parsing data. Please contact system administrator.",
        DOCUMENT_SEARCH_PAGE: "/sitepages/documentsearch.aspx",
        HOME_PAGE: "/sitepages/homepage.aspx",
        LEFT_NAVIGATION_LIST_NOT_MATCH: "Please select appropriate list with Title, Library Name, Library Title and IsSiteLocationMetadataAvailable fields",
        PICTURE_GALLERY_LIST_NOT_MATCH: "Please select appropriate list with Name, Picture Size, Region, Program Type and File Size fields",
        VIDEO_GALLERY_LIST_NOT_MATCH: "Please select appropriate list with Video URL, Video Thumbnail, Region and Program Type fields",
        ANNOUNCEMENTS_LIST_NOT_MATCH: "Please select appropriate list with Title, Description, Mark As Important, Region and Program Type fields",
        EVENTS_LIST_NOT_MATCH: "Please select appropriate list with Region and Program Type fields",
        SELECT_LIST: "Please edit webpart and select appropriate list",
        GET_ITEMS_LIMIT: 5000,
        DOCUMENT_UPLOAD_VALIDATION: "Please select ${selector} and click upload.",
        DOCUMENT_UPLOAD_VALIDATION_SITE_LOCATION_METADDATA: "Please select region, program type and site location and click upload.",
        DOCUMENT_UPLOAD_VALIDATION_PROGRAM_TYPE: "Please select program type and click upload.",
        DOCUMENT_UPLOAD_VALIDATION_SITE_LOCATION: "Please select site location and click upload.",
        DOCUMENT_UPLOAD_VALIDATION_REGION: "Please select region and click upload.",
        DOCUMENT_UPLOAD_FAILED_FOR_DUPLICATE_DOCUMENT_MESSAGE: "Document with the same already exists in the system.",
        DOCUMENT_UPLOADED_MESSAGE: "The document has been uploaded successfully.",
        DOCUMENT_UPLOAD_NOPERMISSION_MESSAGE: "You do not have permissions to upload documents. Please contact the system administrator",
        DOCUMENT_UPLOAD_FAILED_MODAL_MESSAGE_1: "Out of ${TotalCount}, ${SuccesCount} document/s uploaded successfully.",        
        DOCUMENT_UPLOAD_FAILED_MODAL_MESSAGE_2: "Below ${FailedCount} documents are not uploaded because they already exist in system.",
        SEE_ALL_ITEMS_LINK_TEXT: "see more",
        ANNOUNCEMENT_DESCRIPTION_LENGTH: 80,
        EVENTS_TITLE_CHARACTER_LENGTH: 50,
        ANOUNCEMENT_TITLE_CHARACTER_LENGTH: 30,
        SITE_LISTS: "/Lists/",
        ANNOUNCEMENT_LIST_PAGE: "/AllItems.aspx",
        ANNOUNCEMENT_LIST_NEWFORM_PAGE: "/NewForm.aspx",
        EVENTS_LIST_PAGE: "/calendar.aspx",
        PICTURE_GALLERY_PAGE: "/Forms/Thumbnails.aspx",
        VIDEO_GALLERY_PAGE: "/AllItems.aspx",
        ANNOUNCEMENT_GET_ITEMS_LIMIT: 5,
        EVENTS_GET_ITEMS_LIMIT: 5000,
        EVENTS_DISPLAY_ITEMS_LIMIT: 4,
        EVENTS_LIST_NEWFORM_PAGE: "/NewForm.aspx",
        VIDEO_LIST_NEWFORM_PAGE: "/NewForm.aspx",
        PICTURE_LIST_UPLOAD_PAGE: "/_layouts/15/upload.aspx?List=",
        GUID_START_CODE: "%7B",
        GUID_END_CODE: "%7D",
        LEFT_NAVIGATION_HOME_TITLE: "HOME",
        SAFETYHUB_OWNERS_GP: "SafetyHub Owners",
        SAFETYHUB_MEMBERS_GP: "SafetyHub Members",
        VIDEO_GALLERY_SITEASSETS: "SiteAssets/Lists/"

    };

    public static SITE_COLUMN_NAME = {
        REGION_COLS: "Region",
        PROGRAM_TYPE_COLS: "Program Type",
        SITE_LOCATION: "Site Location"
    };
 
    public static SITE_PAGE_URL = {
        PAGE_URL: "https://sonorasoftware0.sharepoint.com/sites/SafetyHubDev/SitePages/HomePage.aspx"
    };

    public static EXPAND_COLUMN = {
        //GET_DOCUMENTS: "Editor,Author,Site_x0020_Location_x0020_Code,FieldValuesAsText",
        GET_DOCUMENTS: "ListItemAllFields/FieldValuesAsText,ListItemAllFields",
        GET_DOCUMENTS_NEW: "ListItemAllFields/FieldValuesAsText,Editor,Editor,Author,Site_x0020_Location_x0020_Code,Site_x0020_Location_x0020_Code,FieldValuesAsText",    
        DOCUMENT_LIST_SITELOCATION: "Site_x0020_Location_x0020_Code,Site_x0020_Location_x0020_Code,Site_x0020_Location_x0020_Code"  
    };

    public static MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];


    public static DAY_NAMES = ['Sun', 'Mon', 'Tues', 'Wed', 'Thu', 'Fri', 'Sat'];

    public static CONNECTED_WP = {
        SHARE_DATA: "shareData"
    };
}