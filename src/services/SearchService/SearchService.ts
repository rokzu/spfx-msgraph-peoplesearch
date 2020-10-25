import { ISearchService } from "./ISearchService";
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { isEmpty } from "@microsoft/sp-lodash-subset";
import { PageCollection } from "../../models/PageCollection";
import { ExtendedUser } from "../../models/ExtendedUser";
import { IGraphBatchResponseBody } from "./IGraphBatchResponseBody";
import { IGraphBatchRequestBody } from "./IGraphBatchRequestBody";
import { IProfileImage } from "../../models/IProfileImage";

export class SearchService implements ISearchService {
  private _msGraphClientFactory: MSGraphClientFactory;
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

  constructor(msGraphClientFactory: MSGraphClientFactory, spWebUrl: string) {
    this._msGraphClientFactory = msGraphClientFactory;
    this._spWebUrl = spWebUrl;
  }

  public async searchUsers(): Promise<PageCollection<ExtendedUser>> {
    const graphClient = await this._msGraphClientFactory.getClient();

    let resultQuery = graphClient
      .api('/users')
      .version("v1.0")
      .header("ConsistencyLevel", "eventual")
      .count(true)
      .top(this.pageSize);

    if (!isEmpty(this.selectParameter)) {
      resultQuery = resultQuery.select(this.selectParameter);
    }

    if (!isEmpty(this.filterParameter)) {
      resultQuery = resultQuery.filter(this.filterParameter);
    }

    if (!isEmpty(this.orderByParameter)) {
      resultQuery = resultQuery.orderby(this.orderByParameter);
    }

    if (!isEmpty(this.searchParameter)) {
      resultQuery = resultQuery.query({ $search: `"displayName:${this.searchParameter}"` });
    }

    if (isEmpty(this.searchParameter)){
      //do not search
      const blankResult = '{"@odata.count":0, "@odata.nextLink":"", "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users", "value":[]}';
      const blankResultObj = JSON.parse(blankResult) as PageCollection<ExtendedUser>;
      return blankResultObj;
    }

    var searchResults = await resultQuery.get();
    //sort the results
    try{
      let resUsers = searchResults.value;
      let sortedUsers = resUsers.sort((n1,n2) => {
          if (n1.surname > n2.surname) {
              return 1;
          }
      
          if (n1.surname < n2.surname) {
              return -1;
          }
      
          return 0;
      });
      
      searchResults.value = [];
      sortedUsers.forEach(userElement => {
        let userUpn = userElement.userPrincipalName;
        if (!isEmpty(userUpn)){
          userElement["photoUrl"] =  `${this._spWebUrl}${"/_layouts/15/userphoto.aspx?size=M&accountname=" + userUpn}`;
        }
        searchResults.value.push(userElement);  
      });

      searchResults.value = sortedUsers;
    }
    catch(e){
      console.error(e);
    }
    return searchResults;
  }

  public async fetchPage(pageLink: string): Promise<PageCollection<ExtendedUser>>  {
    const graphClient = await this._msGraphClientFactory.getClient();

    let resultQuery = graphClient.api(pageLink).header("ConsistencyLevel", "eventual");

    return await resultQuery.get();
  }

  public async fetchProfilePictures(users: ExtendedUser[]): Promise<IProfileImage> {
    const graphClient = await this._msGraphClientFactory.getClient();

    let body: IGraphBatchRequestBody = { requests: [] };
        
    users.forEach((user) => {
      let requestUrl: string = `/users/${user.id}/photo/$value`;
      body.requests.push({ id: user.id.toString(), method: 'GET', url: requestUrl });
    });

    var response: IGraphBatchResponseBody = await graphClient.api('$batch').version('v1.0').post(body);

    var results: IProfileImage = {};
      
      for (let i=0;i<response.responses.length;i++){
        const r = response.responses[i];
        if (r.status === 200) {
          results[r.id] = `data:${r.headers["Content-Type"]};base64,${r.body}`;
        }
        else{
          const userUpn = this._getUserUpn(users, r.id);
          if (!isEmpty(userUpn)){
            const photoUrl = `${this._spWebUrl}${"/_layouts/15/userphoto.aspx?size=M&accountname=" + userUpn}`;
            results[r.id] = photoUrl;
          }
        }
      }

    return results;
  }

  // get user upn account name
  private _getUserUpn(users: ExtendedUser[], id: string): string {
    if (users == null || !Array.isArray(users)){
      return null;
    }

    for (let i=0;i<users.length;i++){
      if(users[i].id == id){
        return users[i].userPrincipalName;
      }
    }
  }
}
