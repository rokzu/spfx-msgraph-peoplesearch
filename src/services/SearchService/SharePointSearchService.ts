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
        let query = this.searchParameter;

        query = query === null ? "" : query.replace(/'/g, `''`);

        const headers: HeadersInit = new Headers();
        headers.append("accept", "application/json;odata.metadata=none");

        let selectFields: string  = "FirstName,LastName,PreferredName,WorkEmail,PictureURL,WorkPhone,MobilePhone,JobTitle,Department,Skills,PastProjects";
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
        let people: ISharePointPerson[] = results.PrimaryQueryResult.RelevantResults.Table.Rows.map(r => {
            return {
              name: this._getValueFromSearchResult('PreferredName', r.Cells),
              firstName: this._getValueFromSearchResult('FirstName', r.Cells),
              lastName: this._getValueFromSearchResult('LastName', r.Cells),
              phone: this._getValueFromSearchResult('WorkPhone', r.Cells),
              mobile: this._getValueFromSearchResult('MobilePhone', r.Cells),
              email: this._getValueFromSearchResult('WorkEmail', r.Cells),
              photoUrl: `${this._spWebUrl}${"/_layouts/15/userphoto.aspx?size=M&accountname=" + this._getValueFromSearchResult('WorkEmail', r.Cells)}`,
              function: this._getValueFromSearchResult('JobTitle', r.Cells),
              department: this._getValueFromSearchResult('Department', r.Cells),
              skills: this._getValueFromSearchResult('Skills', r.Cells),
              projects: this._getValueFromSearchResult('PastProjects', r.Cells)
            };
          });

        throw new Error("Method not implemented.");
    }
    public async fetchPage(pageLink: string): Promise<PageCollection<ExtendedUser>> {
        throw new Error("Method not implemented.");
    }
    public async fetchProfilePictures(users: ExtendedUser[]): Promise<IProfileImage> {
        throw new Error("Method not implemented.");
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