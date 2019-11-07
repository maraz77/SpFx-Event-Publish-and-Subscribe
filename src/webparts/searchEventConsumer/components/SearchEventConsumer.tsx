import * as React from 'react';
import styles from './SearchEventConsumer.module.scss';
import { ISearchEventConsumerProps } from './ISearchEventConsumerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { element } from 'prop-types';

export default class SearchEventConsumer extends React.Component<ISearchEventConsumerProps, {}> {
  public render(): React.ReactElement<ISearchEventConsumerProps> {
    const eventData: string[] | undefined = this.props.EventData.tryGetValues();
    let results:any =[];
    if(this.props.SearchResults && this.props.SearchResults.length>0)
    {
      this.props.SearchResults.forEach(element => {
        for(var key in element)
        {
          results.push(<h3>{key}</h3>);
          element[key].forEach(e1 => {
            e1.forEach(e2 => {

              for(var itemKey in e2)
              {
                results.push(<div className={styles.halfcolumn}>{itemKey}</div>);
                results.push(<div className={styles.halfcolumn}>{e2[itemKey]}</div>);
                
              }
              
            });
            results.push(<hr/>);
          });
          
        }
        //results += element[0];
      });
      
    }
    //let termsArray:string[] = eventData?eventData[0]:[];
    return (
      <div className={ styles.searchEventConsumer }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              
            </div>
          </div>
          <div className={ styles.row }>{results}</div>
        </div>
      </div>
    );
  }
}
