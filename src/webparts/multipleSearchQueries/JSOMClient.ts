import { ISPTermGroup } from './MultipleSearchQueriesWebPart';
import { ISPTermSet } from './MultipleSearchQueriesWebPart';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ISPTerm } from './MultipleSearchQueriesWebPart';
export default class JSOMClient {
    private _siteUrl: string;
    get siteUrl(): string {
        return this._siteUrl;
    }
    set siteUrl(value: string) {
        this._siteUrl = value;
    }
    // This method loads JSOM libraries from _layouts directory
    public loadJSOMLib(): Promise<{}> {
        return SPComponentLoader.loadScript('/_layouts/15/init.js', {
            globalExportsName: '$_global_init'
        })
            .then((): Promise<{}> => {
                return SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', {
                    globalExportsName: 'Sys'
                });
            })
            .then((): Promise<{}> => {
                return SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', {
                    globalExportsName: 'SP'
                });
            })
            .then((): Promise<{}> => {
                return SPComponentLoader.loadScript('/_layouts/15/SP.js', {
                    globalExportsName: 'SP'
                });
            })
            .then((): Promise<{}> => {
                return SPComponentLoader.loadScript('/_layouts/15/SP.taxonomy.js', {
                    globalExportsName: 'SP'
                });
            })
            ;
    }

    public getTermGroups(): Promise<ISPTermGroup[]> {
        const context: SP.ClientContext = new SP.ClientContext(this._siteUrl);
        let taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        let termStore = taxSession.getDefaultSiteCollectionTermStore();
        let termGroups = termStore.get_groups();
        context.load(termGroups);


        let groups: ISPTermGroup[] = [];
        return new Promise<ISPTermGroup[]>((resolve, reject) => {
            context.executeQueryAsync(() => {

                let e = termGroups.getEnumerator();
                while (e.moveNext()) {
                    let g = e.get_current();
                    groups.push(<ISPTermGroup>{ Id: g.get_id().toString(), Name: g.get_name() });

                }
                resolve(groups);

            },
                (sender: any, args: SP.ClientRequestFailedEventArgs) => {
                    reject(groups);
                });

        });
    }

    public getTermSets(groupId: string): Promise<ISPTermSet[]> {
        const context: SP.ClientContext = new SP.ClientContext(this._siteUrl);
        let taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        let termStore = taxSession.getDefaultSiteCollectionTermStore();
        let termGroups = termStore.get_groups();
        let termGroup = termGroups.getById(new SP.Guid(groupId));
        let termSets = termGroup.get_termSets();
        let sets: ISPTermSet[] = [];
        context.load(termSets);


        return new Promise<ISPTermSet[]>((resolve, reject) => {
            context.executeQueryAsync(() => {
                let e = termSets.getEnumerator();
                while (e.moveNext()) {
                    let g = e.get_current();
                    sets.push(<ISPTermSet>{ GroupId: groupId, Id: g.get_id().toString(), Name: g.get_name() });

                }
                resolve(sets.filter(s => s.GroupId == groupId));
            }, (sender: any, args: SP.ClientRequestFailedEventArgs) => { reject(sets); });

        });
    }
    public getTerms(groupId: string,termSetId: string): Promise<ISPTerm[]> {
        const context: SP.ClientContext = new SP.ClientContext(this._siteUrl);
        let taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
        let termStore = taxSession.getDefaultSiteCollectionTermStore();
        let termGroups = termStore.get_groups();
        let termGroup = termGroups.getById(new SP.Guid(groupId));
        let termSets = termGroup.get_termSets();
        let termSet = termSets.getById(new SP.Guid(termSetId));
        let terms = termSet.get_terms();
        let spTerm: ISPTerm[] = [];
        context.load(terms);


        return new Promise<ISPTerm[]>((resolve, reject) => {
            context.executeQueryAsync(() => {
                let e = terms.getEnumerator();
                while (e.moveNext()) {
                    let s = e.get_current();
                    spTerm.push(<ISPTerm>{ Id: s.get_id().toString(), Name: s.get_name() });

                }
                resolve(spTerm);
            }, (sender: any, args: SP.ClientRequestFailedEventArgs) => { reject(spTerm); });

        });
    }
}