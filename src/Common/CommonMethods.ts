import { ColorClassNames } from "@uifabric/styling";
import dataService from "../Common/DataService";

const commonService = new dataService();
export default class commonMethods {

    public isValidListColumns = async (listName: string, columns: string[]): Promise<boolean> => {
        let isValidListColumns: boolean = true;
        for (const columnName of columns) {
            await commonService.checkListFields(listName, columnName).then((fieldData: any) => {
                //debugger;
            }).catch((error: any) => {
                //debugger;
                isValidListColumns = false;
            });
            //If column is not exist in selected list then return false immediately.
            if (!isValidListColumns) {
                break;
            }
        }
        return isValidListColumns;
    }

}