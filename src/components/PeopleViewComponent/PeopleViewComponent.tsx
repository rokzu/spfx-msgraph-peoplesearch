import * as React from 'react';
import ITemplateContext from '../../models/ITemplateContext';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { PersonaCard } from '../PersonaCard/PersonaCard';
import styles from './PeopleViewComponent.module.scss';
import { Text } from '@microsoft/sp-core-library';
import * as strings from "PeopleSearchWebPartStrings";
import {
    Log, Environment, EnvironmentType,
  } from '@microsoft/sp-core-library';
  import { SPComponentLoader } from '@microsoft/sp-loader';
  import { initializeIcons } from '@uifabric/icons';
import { Icon } from 'office-ui-fabric-react/lib/Icon';

const LIVE_PERSONA_COMPONENT_ID: string = "914330ee-2df2-4f6e-a858-30c23a812408";

initializeIcons();

export interface IPeopleViewProps {
    templateContext: ITemplateContext;
}

export interface IPeopleViewState {
    isComponentLoaded: boolean;
}

export class PeopleViewComponent extends React.Component<IPeopleViewProps, IPeopleViewState> {
    private sharedLibrary: any;

    constructor(props: IPeopleViewProps) {
      super(props);
  
      this.state = {
        isComponentLoaded: false,
      };
  
      this.sharedLibrary = null;

      if (Environment.type !== EnvironmentType.Local && props.templateContext.showLPC) {
        this._loadSpfxSharedLibrary();
      }
    }
  
    private async _loadSpfxSharedLibrary() {
      if (!this.state.isComponentLoaded) {
          try {
              this.sharedLibrary = await SPComponentLoader.loadComponentById(LIVE_PERSONA_COMPONENT_ID);   
  
              this.setState({
                  isComponentLoaded: true
              });
  
          } catch (error) {
             Log.error(`[LivePersona_Component]`, error, this.props.templateContext.serviceScope);
          }
      }        
    }

    private _formatPhoneNumber(phoneNumber: string, isMobile: boolean){

        if (isEmpty(phoneNumber)){
            return phoneNumber;
        }

        try{
            let cleanNumber = phoneNumber.replace("+","").trim();
        
                //@ts-ignore
            if (phoneNumber.startsWith("386")){
                cleanNumber = cleanNumber.substr(2).trim();
            }        
            cleanNumber = cleanNumber.replace(" ","");

            if (isMobile){
                cleanNumber = cleanNumber.replace(/(\d{1})(\d{2})(\d{3})(\d{3})/,"+386 $2 $3 $4");
            }
            else{
                cleanNumber = cleanNumber.replace(/(\d{1})(\d{1})(\d{3})(\d{2})(\d{2})/,"+386 $2 $3 $4 $5");
            }

            return cleanNumber;
        }
        catch(e)
        {
            return phoneNumber;
        }
    }

    public render() {
        const ctx = this.props.templateContext;
        let mainElement: JSX.Element = null;
        let resultCountElement: JSX.Element = null;
        let paginationElement: JSX.Element = null;

        if (!isEmpty(ctx.items) && !isEmpty(ctx.items.value)) {
            if (ctx.showResultsCount) {
                resultCountElement = <div className={styles.resultCount}>
                        <label className="ms-fontWeight-semibold">{Text.format(strings.ResultsCount, ctx.resultCount)}</label>
                    </div>;
            }

            if (ctx.showPagination) {
                paginationElement = null;
            }

            const phoneNumberStyle = {
                marginLeft:72 + 16,
                marginTop:0
            };
            
            switch (ctx.personaSize){
                case 11:
                    {
                        phoneNumberStyle.marginLeft = 32 + 16;                        
                        break;
                    }
                    case 12:
                    {
                        phoneNumberStyle.marginLeft = 40 + 16;                        
                        break;
                    }
                    case 13:
                    {
                        phoneNumberStyle.marginLeft = 48 + 16;                        
                        break;
                    }
                    case 14:
                    {
                        phoneNumberStyle.marginLeft = 72 + 25;                        
                        break;
                    }
                    case 15:
                    {
                        phoneNumberStyle.marginLeft = 100 + 16;
                        break;
                    }


            }

            const personaCards = [];
            for (let i = 0; i < ctx.items.value.length; i++) {

                //add mobile number to a separate div
                let bussinessPhoneDiv = [];
                let mobilePhoneDiv;
                let phonesDiv;

                const currentItem = ctx.items.value[i];
                if (currentItem != null){
                    if (currentItem.businessPhones != null && currentItem.businessPhones.length>0){
                        for (let iNumber: number = 0;iNumber<currentItem.businessPhones.length;iNumber++){
                            bussinessPhoneDiv.push(<div><Icon iconName="Phone" /> {this._formatPhoneNumber(currentItem.businessPhones[iNumber], false)}</div>);
                        }
                    }
                    if (!isEmpty(currentItem.mobilePhone)){
                        mobilePhoneDiv = (<div><Icon iconName="CellPhone" /> {this._formatPhoneNumber(currentItem.mobilePhone, true)}</div>);
                    }
                    phonesDiv = (
                        <div style={phoneNumberStyle}>
                            {bussinessPhoneDiv}
                            {mobilePhoneDiv}
                        </div>
                    );
                }

                // businessPhones
                // mobilePhone
                personaCards.push(
                <div className={styles.documentCardItem} key={i}>
                    <div className={styles.personaCard}>
                        <PersonaCard serviceScope={ctx.serviceScope} fieldsConfiguration={ctx.peopleFields} item={ctx.items.value[i]} themeVariant={ctx.themeVariant} personaSize={ctx.personaSize} showLPC={ctx.showLPC} lpcLibrary={this.sharedLibrary} />                        
                    </div>
                                        
                    {phonesDiv}
                </div>);
            }

            mainElement = <React.Fragment>
                <div className={styles.defaultCard}>
                    {resultCountElement}
                    <div className={styles.documentCardContainer}>
                        {personaCards}
                    </div>
                </div>
                {paginationElement}
            </React.Fragment>;
        }
        else if (!ctx.showBlank) {
            mainElement = <div className={styles.noResults}>{strings.NoResultMessage}</div>;
        }

        return <div className={styles.peopleView}>{mainElement}</div>;
    }
}