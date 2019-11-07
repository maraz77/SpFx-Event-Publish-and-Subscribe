import { DynamicProperty } from "@microsoft/sp-component-base";
export interface ISearchEventConsumerProps {
  description: string;
  Query: string;
  EventData: DynamicProperty<string>;
  SearchResults: any;
}
