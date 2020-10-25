import ResultsLayoutOption from "../../models/ResultsLayoutOption";
import { DynamicProperty } from '@microsoft/sp-component-base';
import SearchParameterOption from "../../models/SearchParameterOption";
import SearchServiceToUse from "../../models/SearchServiceToUse";

export interface IPeopleSearchWebPartProps {
  selectParameter: string;
  filterParameter: string;
  orderByParameter: string;
  searchParameter: DynamicProperty<string>;
  searchParameterOption: SearchParameterOption;
  searchEngineUse: SearchServiceToUse;
  pageSize: string;
  showPagination: boolean;
  showLPC: boolean;
  showResultsCount: boolean;
  showBlank: boolean;
  selectedLayout: ResultsLayoutOption;
  webPartTitle: string;
  templateParameters: { [key:string]: any };
}