import { IFileInfo, IFolderInfo, sp } from "@pnp/sp/presets/all";

import { Async } from "office-ui-fabric-react";

/*interface IFilesAndFolders extends IFolderInfo {
    Files: IFileInfo[];
    Folders: IFolderInfo[];
}
*/

export default class dataService {

    public async getData(listName: string, itemId: number) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.getById(itemId).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async getLeftNavigationConfigurationData(listName: string, selectColumn: string, filterCondition: string, orderBy: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.select(selectColumn).filter(filterCondition).orderBy(orderBy).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async GetSiteColumnChoices(itemTitle: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.fields.getByTitle(itemTitle).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });

    }


    public async GetRegion(itemTitle: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.fields.getByTitle(itemTitle).get().then((val) => {
                resolve(val);

            }).catch((error) => {
                reject(error);
            });
        });

    }

    public async GetProgramType(itemTitle: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.fields.getByTitle(itemTitle).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });

    }

    public async getAnnouncements(listName: string, orderBy: string, filterCondition: string, itemLimit: number) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.filter(filterCondition).top(itemLimit).orderBy(orderBy, false).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async GetGUID(listName: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).get().then((val) => {
                resolve(val);

            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async GetImages(listName: string, filterCondition: string, selectColumn: string, orderBy: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.select(selectColumn).filter(filterCondition).orderBy(orderBy, false).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async GetVideos(listName: string, filterCondition: string, selectColumn: string, orderBy: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.select(selectColumn).filter(filterCondition).orderBy(orderBy, false).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }


    public async checkListFields(listName: string, field: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).fields.getByTitle(field).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }


    public async getSiteLocationConfiguration(listName: string, selectColumn: string, orderBy: string, itemLimit: number) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.select(selectColumn).orderBy(orderBy, true).top(itemLimit).get().then((val) => {
            
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async getDocumentsData(folderServerRelativeUrl: string, selectColumn: string, expand: string, filterCondition: string, itemLimit: number, orderBy: string) {

        return new Promise<any>((resolve, reject) => {
          
            sp.web.getFolderByServerRelativeUrl(folderServerRelativeUrl)
              .files.expand("ListItemAllFields/FieldValuesAsText,ListItemAllFields", expand).filter(filterCondition).select("*").top(itemLimit).orderBy(orderBy,false)
              //.files.expand("ListItemAllFields").filter(filterCondition).select("*")
                .get()
                .then((fileData: any) => {
                    resolve(fileData);
                }).catch((error) => {
                    //debugger;
                    reject(error);
                });
           
        });
    }

    public async getFolderData(folderServerRelativeUrl: string, itemLimit: number, orderBy: string, expand: string) {

        return new Promise<any>((resolve, reject) => {

            sp.web.getFolderByServerRelativeUrl(folderServerRelativeUrl)
              .folders.expand(expand).select("*").top(itemLimit).orderBy(orderBy,false)
              //.files.expand("ListItemAllFields").filter(filterCondition).select("*")
                .get()
                .then((folderData: any) => {
                    resolve(folderData);
                }).catch((error) => {
                    reject(error);
                });
             
        });
    }
  
    public async getDocumentsData_Old(listName: string, selectColumn: string, expand: string, filterCondition: string, itemLimit: number, orderBy: string) {

        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.select(selectColumn).filter(filterCondition).expand(expand).top(itemLimit).orderBy(orderBy, false).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public uploadDocument(documentLibraryName: string, fileName: string, fileObject): Promise<any> {
        return new Promise((resolve, reject) => {
            sp.web.getFolderByServerRelativeUrl(documentLibraryName).files.add(fileName, fileObject, false).then((result) => {
                resolve(result);
            }, (error) => {
                reject(error);
            });
        });
    }

    public getFileItem(fileAddResult): Promise<any> {
        return new Promise((resolve, reject) => {
            fileAddResult.file.getItem().then((item) => {
                console.log(item);
                resolve(item);
            }, (error) => {
                reject(error);
            });
        });
    }

    public updateDocumentProperties(fileItem, docTitle, region, programType, documentOwnerId): Promise<any> {
        return new Promise((resolve, reject) => {
            fileItem.update({
                Title: docTitle,
                Region: region,
                Program_x0020_Type: programType,
                EditorId: documentOwnerId
            }).then((result) => {
                resolve(result);
            }, (error) => {
                reject(error);
            });
        });
    }

    public updateDocumentPropertiesWithSiteLocation(fileItem, docTitle, region, programType, siteLocation, documentOwnerId): Promise<any> {
        return new Promise((resolve, reject) => {
            fileItem.update({
                Title: docTitle,
                Region: region,
                Program_x0020_Type: programType,
                EditorId: documentOwnerId,
                Site_x0020_Location_x0020_CodeId: siteLocation
            }).then((result) => {
                resolve(result);
            }, (error) => {
                reject(error);
            });
        });
    }

    public async getEventItems_OLD(listName: string, selectColumn: string, expand: string, filterCondition: string, orderBy: string, itemLimit: number) {
        return new Promise<any>((resolve, reject) => {
            //sp.web.lists.getByTitle(listName).items.select("*").top(5).orderBy("Modified", true).get().then((val) => {
            sp.web.lists.getByTitle(listName).items.select(selectColumn).expand(expand).filter(filterCondition).top(itemLimit).orderBy(orderBy, false).get().then((val) => {
            //sp.web.lists.getByTitle(listName).items.select("*").expand(expand).filter(filterCondition).top(itemLimit).orderBy(orderBy, false).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async getEventItems(listName: string, selectColumn: string, expand: string, filterCondition: string, orderBy: string, itemLimit: number) {
        return new Promise<any>((resolve, reject) => {          
            sp.web.lists.getByTitle(listName).items.select(selectColumn).expand(expand).filter(filterCondition).top(itemLimit).orderBy(orderBy, false).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public getAttachment(listName: string, selectColumn: string, expand: string, itemId: number) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.select(selectColumn).getById(itemId)
                .expand(expand).get().then(result => {
                    resolve(result);
                }).catch((error) => {
                    reject(error);
                });
        });

    } 

    public getCurrentUserGroups() {
        return new Promise<any>((resolve, reject) => {
            sp.web.currentUser.groups.get().then(val => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public getOwnersGroup() {
        return new Promise<any>((resolve, reject) => {
            sp.web.associatedOwnerGroup.get().then(val => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public getMembersGroup() {
        return new Promise<any>((resolve, reject) => {
            sp.web.associatedMemberGroup.get().then(val => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async getSelectedDocumentData(listName: string, selectedProductId: number, fieldName: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.getById(selectedProductId).select(fieldName).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async getSelectedDocumentDatawithSiteLocation(listName: string, selectedProductId: number, fieldName: string,expand: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.getById(selectedProductId).select(fieldName).expand(expand).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });
    }
    
    public editDocumentProperties(listName: string, selectedItemId: number, docTitle: string, region: string, programType: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.getById(selectedItemId).update({
                Title: docTitle,
                Region: region,
                Program_x0020_Type: programType,
            }).then(result => {
                resolve(result);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public editDocumentPropertiesWithSiteLocation(listName: string, selectedItemId: number, docTitle: string, region: string, programType: string, siteLocation:string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.getById(selectedItemId).update({
                Title: docTitle,
                Region: region,
                Program_x0020_Type: programType,  
                Site_x0020_Location_x0020_CodeId: siteLocation
            }).then(result => {
                resolve(result);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public uploadPicture(pictureLibraryName: string, fileName: string, fileObject): Promise<any> {
        return new Promise((resolve, reject) => {
            sp.web.getFolderByServerRelativeUrl(pictureLibraryName).files.add(fileName, fileObject, true).then((result) => {
                //console.log(result);
                resolve(result);
            }, (error) => {
                reject(error);
            });
        });
    }

    public updatePictureProperties(fileItem, caption, region, programType): Promise<any> {
        return new Promise((resolve, reject) => {
            fileItem.update({
                Caption: caption,
                Region: region,
                Program_x0020_Type: programType
            }).then((result) => {
                resolve(result);
            }, (error) => {
                reject(error);
            });
        });
    }

    public updatePicturePropertieswithSiteLocation(fileItem, caption, region, programType, siteLocation): Promise<any> {
        return new Promise((resolve, reject) => {
            fileItem.update({
                Caption: caption,
                Region: region,
                Program_x0020_Type: programType,
                Site_x0020_Location_x0020_CodeId: siteLocation
            }).then((result) => {
                resolve(result);
            }, (error) => {
                reject(error);
            });
        });
    }
    
    public AddItemsInVideoList(listName: string, title: string, region: string, programType: string, filename: string, file, caption: string, videoURL, serverUrl, serverRelativeURL) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.add({
                Title: title,
                Region: region,
                Program_x0020_Type: programType,  
                Caption: caption
            }).then(async result => {
                await sp.web.lists.getByTitle(listName).items.getById(result.data.Id).update({
                    Video_x0020_Thumbnail: JSON.stringify({
                      type: 'thumbnail',
                      fileName: filename,
                      serverUrl: serverUrl,
                      serverRelativeUrl: serverRelativeURL
                    }),
                    Video_x0020_URL:{
                        Description: videoURL,
                        Url: videoURL
                    }
                });
                resolve(result);
            }).catch((error) => {
                reject(error);
            });
        });
    }
    
    public AddItemsInVideoListwithSiteLocation(listName: string, title: string, region: string, programType: string, siteLocation: string, filename: string, file, caption: string, videoURL, serverUrl, serverRelativeURL) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.add({
                Title: title,
                Region: region,
                Program_x0020_Type: programType,  
                Site_x0020_Location_x0020_CodeId: siteLocation,
                Caption: caption
            }).then(async result => {
                await sp.web.lists.getByTitle(listName).items.getById(result.data.Id).update({
                    Video_x0020_Thumbnail: JSON.stringify({
                      type: 'thumbnail',
                      fileName: filename,
                      serverUrl: serverUrl,
                      serverRelativeUrl: serverRelativeURL
                    }),
                    Video_x0020_URL:{
                        Description: videoURL,
                        Url: videoURL
                    }
                });
                resolve(result);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public AddItemsInAnnouncementList(listName: string, title: string, region: string, programType: string, description: string, IsMarkAsImportant: boolean, filename: string, file) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.add({
                Title: title,
                Region: region,
                Program_x0020_Type: programType,  
                Description: description,
                Mark_x0020_As_x0020_Important: IsMarkAsImportant
            }).then(result => {
                result.item.attachmentFiles.add(filename, file).then(r => {
                    resolve(r);
                });
                resolve(result);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public AddItemsInAnnouncementListwithSiteLocation(listName: string, title: string, region: string, programType: string, siteLocation: string, description: string, IsMarkAsImportant: boolean, filename: string, file) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.add({
                Title: title,
                Region: region,
                Program_x0020_Type: programType,  
                Site_x0020_Location_x0020_CodeId: siteLocation,
                Description: description,
                Mark_x0020_As_x0020_Important: IsMarkAsImportant
            }).then(result => {
                result.item.attachmentFiles.add(filename, file).then(r => {
                    resolve(r);
                });
                resolve(result);
            }).catch((error) => {
                reject(error);
            });
        });
    }

    public async GetCategoryChoices(listTitle: string, itemTitle: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listTitle).fields.getByTitle(itemTitle).get().then((val) => {
                resolve(val);
            }).catch((error) => {
                reject(error);
            });
        });

    }

    public AddItemsInEventList(listName: string, title: string, location: string, startTime: Date, endTime: Date , category: string, region: string, programType: string, description: string, IsMarkAsImportant: boolean, IsAllDayEvent: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.add({
                Title: title,
                Location: location,
                EventDate: startTime,
                EndDate: endTime,
                Category: category,
                Region: region,
                Program_x0020_Type: programType,  
                Description: description,
                //RecurrenceData: '<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat><daily dayFrequency="1" /></repeat><repeatForever>FALSE</repeatForever></rule></recurrence>',
                Mark_x0020_as_x0020_Important: IsMarkAsImportant,
                //fRecurrence: true,
                //EventType : 1,
                //RecurrenceData:'<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat><daily dayFrequency="1" /></repeat><windowEnd>2021-11-15T23:59:00Z</windowEnd></rule></recurrence>',
                fAllDayEvent: IsAllDayEvent
            }).then(result => {
                resolve(result);
            }).catch((error) => {
                reject(error);
            });
        });
    }
    
    public AddItemsInEventListwithSiteLocation(listName: string, title: string, location: string, startTime: Date, endTime: Date , category: string, region: string, programType: string, siteLocation: string, description: string, IsMarkAsImportant: boolean, IsAllDayEvent: string) {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getByTitle(listName).items.add({
                Title: title,
                Location: location,
                EventDate: startTime,
                EndDate: endTime,
                Category: category,
                Region: region,
                Program_x0020_Type: programType,  
                Site_x0020_Location_x0020_CodeId: siteLocation,
                Description: description,
                //RecurrenceData: '<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat><daily dayFrequency="1" /></repeat><repeatForever>FALSE</repeatForever></rule></recurrence>',
                Mark_x0020_as_x0020_Important: IsMarkAsImportant,
                //fRecurrence: true,
                //EventType : 1,
                //RecurrenceData:'<recurrence><rule><firstDayOfWeek>su</firstDayOfWeek><repeat><daily dayFrequency="1" /></repeat><windowEnd>2021-11-15T23:59:00Z</windowEnd></rule></recurrence>',
                fAllDayEvent: IsAllDayEvent
            }).then(result => {
                resolve(result);
            }).catch((error) => {
                reject(error);
            });
        });
    }
    
}