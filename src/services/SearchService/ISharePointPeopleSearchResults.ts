/**
 * Interface describing people search results coming from SharePoint Search
 */
export interface ISharePointPeopleSearchResults {
  PrimaryQueryResult: {
    RelevantResults: {
      RowCount: number;
      Table: {
        Rows: {
          Cells: ICell[];
        }[];
      };
      TotalRows: number;
    };
  };
  error?: {
    code: string;
    message: string;
  };
}

export interface ICell {
  Key: string;
  Value: string;
  ValueType: string;
}