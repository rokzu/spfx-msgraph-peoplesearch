import * as React from "react";
import * as ReactDom from "react-dom";
import { Version, Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { ThemeProvider, IReadonlyTheme, ThemeChangedEventArgs } from '@microsoft/sp-component-base';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from "@microsoft/sp-webpart-base";
import { DisplayMode } from "@microsoft/sp-core-library";
import { isEqual } from '@microsoft/sp-lodash-subset';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  IPropertyPaneField,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField,
  IPropertyPaneGroup,
  IPropertyPaneConditionalGroup,
  DynamicDataSharedDepth,
  PropertyPaneDynamicField,
  PropertyPaneDynamicFieldSet
} from "@microsoft/sp-property-pane";
import * as update from 'immutability-helper';
import * as strings from "PeopleSearchWebPartStrings";
import { IPeopleSearchWebPartProps } from "./IPeopleSearchWebPartProps";
import { ISearchService, MockSearchService, SearchService } from "../../services/SearchService";
import { IPeopleSearchContainerProps, PeopleSearchContainer } from "./components/PeopleSearchContainer";
import ResultsLayoutOption from "../../models/ResultsLayoutOption";
import { TemplateService } from "../../services/TemplateService/TemplateService";
import SearchParameterOption from "../../models/SearchParameterOption";
import SearchServiceToUse from "../../models/SearchServiceToUse";
import { SharePointSearchService } from "../../services/SearchService/SharePointSearchService";

export default class PeopleSearchWebPart extends BaseClientSideWebPart<IPeopleSearchWebPartProps> {
  private _searchService: ISearchService;
  private _templateService: TemplateService;
  private _placeholder = null;
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme;
  private _initComplete = false;
  private _templatePropertyPaneOptions: IPropertyPaneField<any>[] = [];

  public async render(): Promise<void> {

    if (!this._initComplete) {
        return;
    }

    await this._initTemplate();

    if (this.displayMode === DisplayMode.Edit) {
        const { Placeholder } = await import(
            /* webpackChunkName: 'search-property-pane' */
            '@pnp/spfx-controls-react/lib/Placeholder'
        );
        this._placeholder = Placeholder;
    }

    this.renderCompleted();
  }

  protected get isRenderAsync(): boolean {
    return true;
  }

  protected renderCompleted(): void {
    super.renderCompleted();
    let renderElement = null;

    if (this._isWebPartConfigured()) {

      const searchParameter: string | undefined = this.properties.searchParameter.tryGetValue();

      this._searchService = update(this._searchService, {
        selectParameter: { $set: this.properties.selectParameter ? this.properties.selectParameter.split(',') : [] },
        filterParameter: { $set: this.properties.filterParameter },
        orderByParameter: { $set: this.properties.orderByParameter },
        searchParameter: { $set: searchParameter },
        pageSize: { $set: parseInt(this.properties.pageSize) }
      });

      renderElement = React.createElement(
        PeopleSearchContainer,
        {
          webPartTitle: this.properties.webPartTitle,
          displayMode: this.displayMode,
          showBlank: this.properties.showBlank,
          showResultsCount: this.properties.showResultsCount,
          showPagination: this.properties.showPagination,
          showLPC: this.properties.showLPC,
          searchParameterOption: this.properties.searchParameterOption,
          searchService: this._searchService,
          searchServiceToUse: this.properties.searchEngineUse,
          templateService: this._templateService,
          templateParameters: this.properties.templateParameters,
          selectedLayout: this.properties.selectedLayout,
          themeVariant: this._themeVariant,
          serviceScope: this.context.serviceScope,
          updateWebPartTitle: (value: string) => {
            this.properties.webPartTitle = value;
          },
          updateSearchParameter: (value: string) => {
            this.properties.searchParameter.setValue(value);
            this.render();
          }
        } as IPeopleSearchContainerProps
      );
    } else {
      if (this.displayMode === DisplayMode.Edit) {
          const placeholder: React.ReactElement<any> = React.createElement(
              this._placeholder,
              {
                  iconName: strings.PlaceHolderEditLabel,
                  iconText: strings.PlaceHolderIconText,
                  description: strings.PlaceHolderDescription,
                  buttonLabel: strings.PlaceHolderConfigureBtnLabel,
                  onConfigure: this._setupWebPart.bind(this)
              }
          );
          renderElement = placeholder;
      } else {
          renderElement = React.createElement('div', null);
      }
    }

    ReactDom.render(renderElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._initializeRequiredProperties();

    await this._initSearchService();

    this._initThemeVariant();

    this._templateService = new TemplateService();

    this._initComplete = true;

    return super.onInit();
  }



  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    const templateParametersGroup = this._getTemplateFieldsGroup();

    let propertyPaneGroups: (IPropertyPaneGroup | IPropertyPaneConditionalGroup)[] = [
      {
        groupName: strings.UseSearchEngineSettingsGroup,
        groupFields: this._getSearchServiceToUse()
      },
      {
        groupName: strings.QuerySettingsGroupName,
        groupFields: this._getQueryFields()
      },
      {
        groupName: strings.SearchQuerySettingsGroupName,
        groupFields: this._getSearchQueryFields()
      },
      {
        groupName: strings.StylingSettingsGroupName,
        groupFields: this._getStylingFields(),
      }
    ];

    if (templateParametersGroup) {
      propertyPaneGroups.push(templateParametersGroup);
    }

    return {
      pages: [
        {
          groups: propertyPaneGroups,
          displayGroupsAsAccordion: true
        }
      ]
    };
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string) {

    if (propertyPath.localeCompare('searchEngineUse') === 0) {
      await this._initSearchService();
      this.context.propertyPane.refresh();
    }

    if (propertyPath.localeCompare('selectedLayout') === 0) {
      await this._initTemplate();
      this.context.propertyPane.refresh();
    }

    if (propertyPath.localeCompare('searchParameterOption') === 0) {
      if (this.properties.searchParameterOption === SearchParameterOption.None) {
        this.properties.searchParameter.setValue("");
      }
    }
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'searchParameter': {
        dynamicPropertyType: 'string'
      }
    } as any as IWebPartPropertiesMetadata;
  }

  private _getSearchServiceToUse(): IPropertyPaneField<any>[] {

   const searchEngineOption = [
      {
          text: strings.MSGraphSearchEngine,
          key: SearchServiceToUse.MSGraph
      },
      {
        text: strings.SharePointSearch,
        key: SearchServiceToUse.SharePoint,
      }
    ] as IPropertyPaneChoiceGroupOption[];

    let searchEngineUseFields: IPropertyPaneField<any>[] = [
      PropertyPaneChoiceGroup('searchEngineUse', {
        label: strings.UseSearchEngine,
        options: searchEngineOption
      }),
    ];

    return searchEngineUseFields;
  }


  /**
   * Determines the group fields for search query options inside the property pane
   */
  private _getSearchQueryFields(): IPropertyPaneField<any>[] {
    const searchParameterOptions = [
      {
          text: strings.NoneSearchParameterOption,
          key: SearchParameterOption.None
      },
      {
        text: strings.BoxSearchParameterOption,
        key: SearchParameterOption.SearchBox,
      },
      {
          text: strings.StaticSearchParameterOption,
          key: SearchParameterOption.Static
      },
      {
          text: strings.DynamicSearchParameterOption,
          key: SearchParameterOption.Dynamic
      }
    ] as IPropertyPaneChoiceGroupOption[];

    let searchQueryFields: IPropertyPaneField<any>[] = [
      PropertyPaneChoiceGroup('searchParameterOption', {
        label: strings.SearchParameterOption,
        options: searchParameterOptions
      }),
    ];

    if (this.properties.searchParameterOption == SearchParameterOption.Static) {
      searchQueryFields.push(
        PropertyPaneTextField('searchParameter', {
          label: strings.SearchParameter
        })
      );
    }

    if (this.properties.searchParameterOption == SearchParameterOption.Dynamic) {
      searchQueryFields.push(
        PropertyPaneDynamicFieldSet({
          label: strings.SearchParameter,
          fields: [
            PropertyPaneDynamicField('searchParameter', {
              label: strings.SearchParameter
            })
          ],
          sharedConfiguration: {
            depth: DynamicDataSharedDepth.Property
          }
        })
      );
    }

    return searchQueryFields;
  }

  /**
   * Determines the group fields for query options inside the property pane
   */
  private _getQueryFields(): IPropertyPaneField<any>[] {
    let queryFields: IPropertyPaneField<any>[] = [      
      PropertyPaneTextField('selectParameter', {
          label: strings.SelectParameter,
          multiline: true
      }),
      PropertyPaneTextField('filterParameter', {
        label: strings.FilterParameter,
        multiline: true
      }),
      PropertyPaneTextField('orderByParameter', {
        label: strings.OrderByParameter,
        multiline: true
      }),
      PropertyPaneTextField('pageSize', {
        label: strings.PageSizeParameter,
        value: this.properties.pageSize.toString(),
        maxLength: 3,
        deferredValidationTime: 300,
        onGetErrorMessage: (value: string) => {
          return this._validateNumber(value);
        } 
      }),
    ];

    return queryFields;
  }

    /**
   * Init the template according to the property pane current configuration
   * @returns the template content as a string
   */
  private async _initTemplate(): Promise<void> {
    this._templatePropertyPaneOptions = this._templateService.getTemplateParameters(this.properties.selectedLayout, this.properties);
  }

   /**
   * Init the search service according to the property pane current configuration
   * @returns none
   */
  private async _initSearchService(): Promise<void> {    
    if (Environment.type === EnvironmentType.Local) {
      this._searchService = new MockSearchService();
    } else {
      if (this.properties.searchEngineUse == SearchServiceToUse.SharePoint){
        this._searchService = new SharePointSearchService(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl);
      }
      else{
        this._searchService = new SearchService(this.context.msGraphClientFactory, this.context.pageContext.web.absoluteUrl);
      }
    }    
  }

  /**
   * Determines the group fields for styling options inside the property pane
   */
  private _getStylingFields(): IPropertyPaneField<any>[] {
    const layoutOptions = [
        {
            iconProps: {
                officeFabricIconFontName: 'People'
            },
            text: strings.PeopleLayoutOption,
            key: ResultsLayoutOption.People
        },
        {
            iconProps: {
                officeFabricIconFontName: 'Code'
            },
            text: strings.DebugLayoutOption,
            key: ResultsLayoutOption.Debug
        }
    ] as IPropertyPaneChoiceGroupOption[];

    let stylingFields: IPropertyPaneField<any>[] = [
      PropertyPaneToggle('showPagination', {
        label: strings.ShowPaginationControl,
      }),
      PropertyPaneToggle('showBlank', {
          label: strings.ShowBlankLabel,
          checked: this.properties.showBlank,
      }),
      PropertyPaneToggle('showResultsCount', {
          label: strings.ShowResultsCountLabel,
          checked: this.properties.showResultsCount,
      }),
      PropertyPaneToggle('showLPC', {
        label: strings.ShowLivePersonaCard,
      }),
      PropertyPaneChoiceGroup('selectedLayout', {
          label: strings.ResultsLayoutLabel,
          options: layoutOptions
      }),
    ];

    return stylingFields;
  }

  /**
   * Gets template parameters fields
   */
  private _getTemplateFieldsGroup(): IPropertyPaneGroup {

    let templateFieldsGroup: IPropertyPaneGroup = null;

    if (this._templatePropertyPaneOptions.length > 0) {
        templateFieldsGroup = {
            groupFields: this._templatePropertyPaneOptions,
            isCollapsed: false,
            groupName: strings.TemplateParameters.TemplateParametersGroupName
        };
    }

    return templateFieldsGroup;
  }

  /**
  * Checks if all webpart properties have been configured
  */ 
  private _isWebPartConfigured(): boolean {
    return true;
  }

  /**
  * Initializes the Web Part required properties if there are not present in the manifest (i.e. during an update scenario)
  */
  private _initializeRequiredProperties() {
    this.properties.selectedLayout = (this.properties.selectedLayout !== undefined && this.properties.selectedLayout !== null) ? this.properties.selectedLayout : ResultsLayoutOption.People;
    this.properties.searchParameterOption = (this.properties.searchParameterOption !== undefined && this.properties.searchParameterOption !== null) ? this.properties.searchParameterOption : SearchParameterOption.None;
    this.properties.templateParameters = this.properties.templateParameters ? this.properties.templateParameters : {};
    this.properties.searchEngineUse = (this.properties.searchEngineUse !== undefined && this.properties.searchEngineUse !== null) ? this.properties.searchEngineUse : SearchServiceToUse.MSGraph;
  }

  /**
   * Initializes theme variant properties
   */
  private _initThemeVariant(): void {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent.bind(this));
  }

  /**
   * Update the current theme variant reference and re-render.
   * @param args The new theme
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
      if (!isEqual(this._themeVariant, args.theme)) {
          this._themeVariant = args.theme;
          this.render();
      }
  }

    /**
   * Opens the Web Part property pane
   */
  private _setupWebPart() {
      this.context.propertyPane.open();
  }

  private _validateNumber(value: string): string {
    let number = parseInt(value);
    if (isNaN(number)) {
        return strings.InvalidNumberIntervalMessage;
    }
    if (number < 1 || number > 999) {
        return strings.InvalidNumberIntervalMessage;
    }
    return '';
  }
}
