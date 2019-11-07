import { ISPTermGroup } from './MultipleSearchQueriesWebPart';
import { ISPTermSet } from './MultipleSearchQueriesWebPart';
import { ISPTerm } from './MultipleSearchQueriesWebPart';
export default class MockClient  {

    private static _termGroups: ISPTermGroup[] = [{
        Id: 'TermGroup1',
        Name: 'Term Group 1'
      },
      {
        Id: 'TermGroup2',
        Name: 'Term Group 2'
      }];
      private static _terms: ISPTermGroup[] = [{
        Id: 'Term1',
        Name: 'Term 1'
      },
      {
        Id: 'Term2',
        Name: 'Term 2'
      }];

    private static _termSets: ISPTermSet[] = [
        {
            GroupId: 'TermGroup1',
          Id: 'Group1Set1',
          Name: 'Group 1 Set 1'
        },
        {
            GroupId: 'TermGroup1',
          Id: 'Group1Set2',
          Name: 'Group 1 Set 2'
        },
        
            {
                GroupId: 'TermGroup2',
              Id: 'Group2Set1',
              Name: 'Group 2 Set 1'
            },
            {
                GroupId: 'TermGroup2',
              Id: 'Group2Set2',
              Name: 'Group 2 Set 2'
            }
          
      ];

    public static getTermGroups(): Promise<ISPTermGroup[]> {
    return new Promise<ISPTermGroup[]>((resolve) => {
            resolve(MockClient._termGroups);
        });
    }
    public static getTermSets(groupId:string): Promise<ISPTermSet[]> {
        return new Promise<ISPTermSet[]>((resolve) => {
                resolve(MockClient._termSets.filter(s=>s.GroupId==groupId));
            });
        }
        public static getTerms(groupId: string,termSetId:string): Promise<ISPTerm[]> {
          return new Promise<ISPTerm[]>((resolve) => {
                  resolve(MockClient._terms);
              });
          }
}