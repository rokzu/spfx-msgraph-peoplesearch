import { ExtendedUser } from "../../models/ExtendedUser";
import { IProfileImage } from "../../models/IProfileImage";
import { PageCollection } from "../../models/PageCollection";
import { ISearchService } from "./ISearchService";
import { SPHttpClient } from "@microsoft/sp-http";
import { isEmpty } from "@microsoft/sp-lodash-subset";
import { ICell, ISharePointPeopleSearchResults } from "./ISharePointPeopleSearchResults";
import { ISharePointPerson } from "./ISharePointPerson";

export class SharePointSearchService implements ISearchService {
    private _spHttpClient: SPHttpClient;
	  private _spWebUrl: string;

    private _selectParameter: string[];
    private _filterParameter: string;
    private _orderByParameter: string;
    private _searchParameter: string;
    private _pageSize: number;
  
    public get selectParameter(): string[] { return this._selectParameter; }
    public set selectParameter(value: string[]) { this._selectParameter = value; }
  
    public get filterParameter(): string { return this._filterParameter; }
    public set filterParameter(value: string) { this._filterParameter = value; }
  
    public get orderByParameter(): string { return this._orderByParameter; }
    public set orderByParameter(value: string) { this._orderByParameter = value; }
  
    public get searchParameter(): string { return this._searchParameter; }
    public set searchParameter(value: string) { this._searchParameter = value; }
  
    public get pageSize(): number { return this._pageSize; }
    public set pageSize(value: number) { this._pageSize = value; }

    constructor(spHttpClient: SPHttpClient, spWebUrl: string) {
        this._spHttpClient = spHttpClient;
        this._spWebUrl = spWebUrl;
    }

    public async searchUsers(): Promise<PageCollection<ExtendedUser>> {
        const blankResult = '{"@odata.count":0, "@odata.nextLink":"", "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users", "value":[]}';
        const resultObj = JSON.parse(blankResult) as PageCollection<ExtendedUser>;

        let query = this.searchParameter;
        if (isEmpty(query)){
          return resultObj;
        }

        query = query === null ? "" : query.replace(/'/g, `''`);

        const headers: HeadersInit = new Headers();
        headers.append("accept", "application/json;odata.metadata=none");

        let selectFields: string  = "FirstName,LastName,UserName,UserProfile_GUID,PreferredName,WorkEmail,PictureURL,WorkPhone,MobilePhone,JobTitle,Department,Skills,PastProjects";
        if (!isEmpty(this.selectParameter)){
            selectFields = this.selectParameter.join(",");
        }

        let orderBy = "LastName:ascending";        
        if (!isEmpty(this.orderByParameter)){
            orderBy = this.orderByParameter;
        }

        const url = `${this._spWebUrl}/_api/search/query?querytext='${query}'&selectproperties='${selectFields}'&sortlist='${orderBy}'&sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'&rowlimit=500`;
        
        let res = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
                headers: headers
        });

        let results = await res.json() as ISharePointPeopleSearchResults;
        
        // convert the SharePoint People Search results to an array of people
        let people: ExtendedUser[] = results.PrimaryQueryResult.RelevantResults.Table.Rows.map(r => {
            return {
              name: this._getValueFromSearchResult('PreferredName', r.Cells),
              givenName: this._getValueFromSearchResult('FirstName', r.Cells),
              surname: this._getValueFromSearchResult('LastName', r.Cells),
              businessPhones: [
                this._getValueFromSearchResult('WorkPhone', r.Cells)
              ],
              mobilePhone: this._getValueFromSearchResult('MobilePhone', r.Cells),
              mail: this._getValueFromSearchResult('WorkEmail', r.Cells),
              photoUrl: `${this._spWebUrl}${"/_layouts/15/userphoto.aspx?size=M&accountname=" + this._getValueFromSearchResult('WorkEmail', r.Cells)}`,
              jobTitle: this._getValueFromSearchResult('JobTitle', r.Cells),
              department: this._getValueFromSearchResult('Department', r.Cells),
              displayName: this._getValueFromSearchResult('PreferredName', r.Cells),
              userPrincipalName: this._getValueFromSearchResult('UserName', r.Cells),
              id: this._getValueFromSearchResult('UserProfile_GUID', r.Cells)
            };
          });

        resultObj["@odata.count"] = people.length +1;
        resultObj.value = people;

        return resultObj;
    }
    public async fetchPage(pageLink: string): Promise<PageCollection<ExtendedUser>> {
      const blankResult = '{"@odata.count":0, "@odata.nextLink":"", "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users", "value":[]}';
      const resultObj = JSON.parse(blankResult) as PageCollection<ExtendedUser>;

      return resultObj;
    }
    public async fetchProfilePictures(users: ExtendedUser[]): Promise<IProfileImage> {
        let returnImages: IProfileImage;        
        return returnImages;
    }

    /**
   * Retrieves the value of the particular managed property for the current search result.
   * If the property is not found, returns an empty string.
   * @param key Name of the managed property to retrieve from the search result
   * @param cells The array of cells for the current search result
   */
  private _getValueFromSearchResult(key: string, cells: ICell[]): string {
    for (let i: number = 0; i < cells.length; i++) {
      if (cells[i].Key === key) {
        return cells[i].Value;
      }
    }

    return '';
  }
}