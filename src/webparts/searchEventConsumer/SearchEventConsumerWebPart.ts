import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { DynamicProperty } from '@microsoft/sp-component-base';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IWebPartPropertiesMetadata,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SearchEventConsumerWebPartStrings';
import SearchEventConsumer from './components/SearchEventConsumer';
import { ISearchEventConsumerProps } from './components/ISearchEventConsumerProps';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http'; 
export interface ISearchEventConsumerWebPartProps {
  description: string;
  Query: string;
  SelectProperties:string;
  EventData: DynamicProperty<string>;
}

export default class SearchEventConsumerWebPart extends BaseClientSideWebPart<ISearchEventConsumerWebPartProps> {

  public render(): void {

    const eventData: string[] | undefined = this.properties.EventData.tryGetValues();
    //let searchResult: any = [];
    if (eventData.length > 0 && this.properties.Query) {
      this.getResults(this.properties.Query, this.properties.SelectProperties, eventData).then(
        data => {
          const element: React.ReactElement<ISearchEventConsumerProps> = React.createElement(
            SearchEventConsumer,
            {
              description: this.properties.description,
              Query: this.properties.Query,
              EventData: this.properties.EventData,
              SearchResults: data
            }
          );

          ReactDom.render(element, this.domElement);
        }
      );
    }
    else {

      const element: React.ReactElement<ISearchEventConsumerProps> = React.createElement(
        SearchEventConsumer,
        {
          description: this.properties.description,
          Query: this.properties.Query,
          EventData: this.properties.EventData,
          SearchResults: []
        }
      );

      ReactDom.render(element, this.domElement);
    }
  }
  private getResults (query:string,selectProperties:string, eventData:string[]):Promise<any>
  {
    let searchResults:any = [];
    let allPromises: any= [];
    eventData.forEach(datum => {
      //let query_copy = (' ' + query).slice(1);
      let q:string=`${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${query.replace("#_EVENT_DATUM_#",datum)}'&selectproperties='${selectProperties}'`;
      allPromises.push(this.context.spHttpClient.get(q, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {  
        response.json().then((responseJSON: any) => 
        {  

          if (responseJSON.PrimaryQueryResult.RelevantResults.Table.Rows.length > 0) {
            let results:any = [];
            responseJSON.PrimaryQueryResult.RelevantResults.Table.Rows.forEach(row => 
            {
              let sProps: string[] = selectProperties.split(",");
              let resultItem:any = [];
              sProps.forEach(
                p => {
                  let item  : any = {};
                  item[p] = row.Cells.filter(c => { return c.Key.toLowerCase() == p.toLowerCase() })[0].Value;
                  resultItem.push(item);

                }
              );
              results.push(resultItem);
            });
            let searchResult:any = {};
            searchResult[datum] = results;
            searchResults.push(
              searchResult
            );
            
          } 
        });  
      })); 
    });
    return Promise.all(allPromises).then( () => {
      return searchResults;
    });
    
  }
  /*protected onInit():Promise<void>
  {
    const eventData: string[] | undefined = this.properties.EventData.tryGetValues();
    this.getResults(this.properties.Query,this.properties.SelectProperties,eventData);
    return Promise.resolve();
  }*/
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      // Specify the web part properties data type to allow the address
      // information to be serialized by the SharePoint Framework.
      'EventData': {
        dynamicPropertyType: 'object'
      }
    };
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
                PropertyPaneTextField('Query', {
                  label: strings.QueryFieldLabel
                }),
                PropertyPaneTextField('SelectProperties', {
                  label: strings.SelectPropertiesFieldLabel
                }),
                PropertyPaneDynamicFieldSet({
                  label: 'Select event source',
                  fields: [
                    PropertyPaneDynamicField('EventData', {
                      label: 'Event source'
                    })
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
