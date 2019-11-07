import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'MultipleSearchQueriesWebPartStrings';
import MultipleSearchQueries from './components/MultipleSearchQueries';
import { IMultipleSearchQueriesProps } from './components/IMultipleSearchQueriesProps';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import MockClient from './MockClient';
import JSOMClient from './JSOMClient';
import { IDynamicDataPropertyDefinition, IDynamicDataCallables } from '@microsoft/sp-dynamic-data';

export interface IMultipleSearchQueriesWebPartProps {
  TermGroupName: string;
  TermSetName: string;
  EventName:string;
}
export interface ISPTermGroup {
  Name: string;
  Id: string;
}

export interface ISPTermSet {
  GroupId: string;
  Name: string;
  Id: string;
}

export interface ISPTermGroups {
  value: ISPTermGroup[];
}



export interface ISPTermSets {
  value: ISPTermSet[];
}

export interface ISPTerm {
  Name: string;
  Id: string;
}
export default class MultipleSearchQueriesWebPart extends BaseClientSideWebPart<IMultipleSearchQueriesWebPartProps> implements IDynamicDataCallables {
  private termGroupNames: IPropertyPaneDropdownOption[];
  private termGroupDropdownDisabled: boolean = true;
  private termSetNames: IPropertyPaneDropdownOption[];
  private terms: string[]=[];
  private termSetDropdownDisabled: boolean = true;
  private client:JSOMClient;
  private EventName:string;

  protected onInit():Promise<void>
  {
    this.EventName = this.properties.EventName?this.properties.EventName:"TermsEventData";
    // register this web part as dynamic data source
    this.context.dynamicDataSourceManager.initializeSource(this);

    if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) 
        {   
          if(!this.client)
          {
            this.client = new JSOMClient();
            this.client.siteUrl = this.context.pageContext.web.absoluteUrl;
            return this.client.loadJSOMLib().then(
              ()=>
              {
                return this.loadTerms().then(
                  (data:string[])=>
                  {

                    this.terms = data?data:[];
                    if(data)
                    {
                    // notify subscribers that the selected event has changed
                      this.context.dynamicDataSourceManager.notifyPropertyChanged(this.EventName);
                    }
                  }
                );
              }
            );
          }
        }
    
  }
  /**
   * Return list of dynamic data properties that this dynamic data source
   * returns
   */
  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: this.EventName,
        title: this.EventName
      }
    ];
  }

  /**
   * Return the current value of the specified dynamic data set
   * @param propertyId ID of the dynamic data set to retrieve the value for
   */
  public getPropertyValue(propertyId: string): string[] {
    switch (propertyId) {
      case this.EventName:
        return this.terms;
      
    }

    throw new Error('Bad property id');
  }

  public render(): void {

        const element: React.ReactElement<IMultipleSearchQueriesProps > = React.createElement(
          MultipleSearchQueries,
          {
            TermGroupName: this.properties.TermGroupName,
            TermSetName: this.properties.TermSetName,
            Terms: this.terms
          }
        );
        
        ReactDom.render(element, this.domElement);
        
      
    
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private loadTermGroups(): Promise<IPropertyPaneDropdownOption[]> 
  {
    //If it's a local workbench & SharePoint is not available.
    if (Environment.type === EnvironmentType.Local) 
    {
      return new Promise<IPropertyPaneDropdownOption[]>(
        (resolve: (options: IPropertyPaneDropdownOption[]) => void, 
        reject: (error: any) => void) => 
      {
        // setTimeout to simulate some delay whih we normally associate with http requests.
        setTimeout(() => 
        {
          //load some mock data.
          MockClient.getTermGroups().then((data: ISPTermGroup[]) => {
            var ddOptions: IPropertyPaneDropdownOption[] = data.map(g => <IPropertyPaneDropdownOption>{ key: g.Id, text: g.Name });
            resolve(ddOptions);
          }
          );

        }, 2000);
      });


    }
    else if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) 
    {   // if SharPoint is available.
      return new Promise<IPropertyPaneDropdownOption[]>(
        (resolve: (options: IPropertyPaneDropdownOption[]) => void, 
        reject: (error: any) => void) => 
        {
          // load real data
        this.client.getTermGroups().then((data: ISPTermGroup[]) => 
        {
          // Transform ISPTermGroup[] into IPropertyPaneDropdownOption[]
          var ddOptions: IPropertyPaneDropdownOption[] = data.map(g => <IPropertyPaneDropdownOption>{ key: g.Id, text: g.Name });
          resolve(ddOptions);
        }
        );
      });

    }
  }

  private loadTermSets(): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.TermGroupName) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }

    const wp: MultipleSearchQueriesWebPart = this;
    if (Environment.type === EnvironmentType.Local) {
      return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
        setTimeout(() => {
          MockClient.getTermSets(wp.properties.TermGroupName).then((data: ISPTermSet[]) => {
            var ddOptions: IPropertyPaneDropdownOption[] = data.map(g => <IPropertyPaneDropdownOption>{ key: g.Id, text: g.Name });
            resolve(ddOptions);
          }
          );

        }, 2000);


      });
    }
    else if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
        this.client.getTermSets(wp.properties.TermGroupName).then((data: ISPTermSet[]) => {
          var ddOptions: IPropertyPaneDropdownOption[] = data.map(g => <IPropertyPaneDropdownOption>{ key: g.Id, text: g.Name });
          resolve(ddOptions);
        }
        );


      });
    }
  }

  private loadTerms(): Promise<string[]> {
    if (!this.properties.TermGroupName || !this.properties.TermSetName) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }

    const wp: MultipleSearchQueriesWebPart = this;
    if (Environment.type === EnvironmentType.Local) {
      return new Promise<string[]>((resolve: (options: string[]) => void, reject: (error: any) => void) => {
        setTimeout(() => {
          MockClient.getTerms(wp.properties.TermGroupName, wp.properties.TermSetName).then((data: ISPTerm[]) => {
            var terms: string[] = data.map(t => t.Name);
            
            resolve(terms);
          }
          );

        }, 2000);


      });
    }
    else if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
      return new Promise<string[]>((resolve: (options: string[]) => void, reject: (error: any) => void) => {
        this.client.getTerms(wp.properties.TermGroupName, wp.properties.TermSetName).then((data: ISPTerm[]) => {
          var terms: string[] = data.map(t => t.Name);
            resolve(terms);
        }
        );


      });
    }
  }
  
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('TermGroupName', {
                  label: strings.TermGroupNameFieldLabel,
                  options: this.termGroupNames,
                  disabled: this.termGroupDropdownDisabled
                }),
                PropertyPaneDropdown('TermSetName', {
                  label: strings.TermSetNameFieldLabel,
                  options: this.termSetNames,
                  disabled: this.termSetDropdownDisabled,
                  selectedKey: this.properties.TermSetName || ''
                }),
                PropertyPaneTextField('EventName',{label: strings.EventNameFieldLabel})
              ]
            }
          ]
        }
      ]
    };
  }
  protected onPropertyPaneConfigurationStart(): void 
  {
    // If Group names have not been fetched yet, keep group drop down disabled.
    this.termGroupDropdownDisabled = !this.termGroupNames; 

    // If Group or Term Set names have not been fetched yet, keep Term Set drop down disabled.
    this.termSetDropdownDisabled = !this.properties.TermGroupName || !this.termSetNames;

    // If Group names have already been fetched, return; don't proceed any further in this method.
    if (this.termGroupNames) {
      return;
    }

    ReactDom.unmountComponentAtNode(this.domElement);
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Options');

    // call the asynchronus method to load Term Groups defined in the webpart's main class
    this.loadTermGroups()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => 
      {
        //Once Groups have been loaded, set the instance variable. 
        this.termGroupNames = listOptions;
        this.termGroupDropdownDisabled = false;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);

        // In case a term group name was previously selected, call the asynchronus method to load Term Sets 
        this.loadTermSets()
        .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
          this.termSetNames = itemOptions;
          this.termSetDropdownDisabled = !this.properties.TermGroupName;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        });
      })
      ;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'TermGroupName' && newValue) 
    {
      // push new term group value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // get previously selected item
      const previousTermSet: string = this.properties.TermSetName;
      // reset selected item
      this.properties.TermSetName = undefined;
      
      this.terms = [];
      // push new item value
      this.onPropertyPaneFieldChanged('TermSetName', previousTermSet, this.properties.TermSetName);
      this.context.propertyPane.refresh();
      // disable item selector until new items are loaded
      this.termSetDropdownDisabled = true;
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      ReactDom.unmountComponentAtNode(this.domElement);
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Term Sets');

      this.loadTermSets()
        .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
          // store items
          this.termSetNames = itemOptions;
          // enable item selector
          this.termSetDropdownDisabled = false;
          // clear status indicator
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          // re-render the web part as clearing the loading indicator removes the web part body
          this.render();
          
          // refresh the item selector control by repainting the property pane
          this.context.propertyPane.refresh();
        });
    }
    else if(propertyPath === 'TermSetName' && newValue ) {
      // push new term group value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      
      ReactDom.unmountComponentAtNode(this.domElement);
      // communicate loading items
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'Terms');
      // Since Term set has been selected now, load Terms now.
        this.loadTerms()
          .then((data: string[]): void => {
            this.terms = data;
            // notify subscribers that the selected event has changed
            this.context.dynamicDataSourceManager.notifyPropertyChanged(this.EventName);
            // clear status indicator
            this.context.statusRenderer.clearLoadingIndicator(this.domElement);
            // re-render the web part as clearing the loading indicator removes the web part body
            this.render();
            this.context.propertyPane.refresh();
          });
        
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.context.propertyPane.refresh();
    }
  }
}
