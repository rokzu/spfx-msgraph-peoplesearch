import { DisplayMode, ServiceScope } from "@microsoft/sp-core-library";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ISearchService } from "../../../../services/SearchService";
import ResultsLayoutOption from "../../../../models/ResultsLayoutOption";
import { TemplateService } from "../../../../services/TemplateService/TemplateService";
import SearchParameterOption from "../../../../models/SearchParameterOption";
import SearchServiceToUse from "../../../../models/SearchServiceToUse";

export interface IPeopleSearchContainerProps {
      /**
     * The web part title
     */
    webPartTitle: string;

    /**
     * The search data provider instance
     */
    searchService: ISearchService;

    searchParameterOption: SearchParameterOption;

    searchServiceToUse: SearchServiceToUse;

    /**
     * Show the result count and entered keywords
     */
    showResultsCount: boolean;

    /**
     * Show nothing if no result
     */
    showBlank: boolean;

    showPagination: boolean;

    showLPC: boolean;

    /**
     * The current display mode of Web Part
     */
    displayMode: DisplayMode;

    /**
     * The current selected layout
     */
    selectedLayout: ResultsLayoutOption;

    /**
     * The current theme variant
     */
    themeVariant: IReadonlyTheme | undefined;

        /**
     * The template helper instance
     */
    templateService: TemplateService;

    /**
     * Template parameters from Web Part property pane
     */
    templateParameters: { [key:string]: any };

    serviceScope: ServiceScope;
    
    updateWebPartTitle: (value: string) => void;

    updateSearchParameter: (value: string) => void;
}