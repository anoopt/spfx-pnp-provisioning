import { ISubSiteService } from "./ISubSiteService";
import { ISubSite } from "../interfaces/ISubSite";
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import pnp, { List, ItemAddResult } from "sp-pnp-js";

const SITE_REQUEST_LIST_NAME: string = "Site Request";

export class SubSiteService implements ISubSiteService {
    public static readonly serviceKey: ServiceKey<ISubSiteService> = ServiceKey.create<ISubSiteService>('cc:ISubSiteService', SubSiteService);
    private _pageContext: PageContext;
    private _currentWebUrl: string;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
            this._currentWebUrl = this._pageContext.web.absoluteUrl;

            //Setup pnp-js to work with the current web url
            pnp.setup({
                sp: {
                    baseUrl: this._currentWebUrl
                }
            });
        });
    }

    public async addSubSite(subSite: ISubSite): Promise<boolean> {
        return pnp.sp.web.lists.getByTitle(SITE_REQUEST_LIST_NAME).items.add({
            'Title': subSite.Title,
            'Description': subSite.Description,
            'Url': subSite.Url,
            'GroupName': subSite.GroupName,
            'GroupOwnerId': subSite.GroupOwnerId,
            'GroupMemberId': subSite.GroupMemberId
        }).then(async (result: ItemAddResult): Promise<boolean> => {
            let addedItem: ISubSite = result.data;
            console.log(addedItem);
            return pnp.sp.web.lists.getByTitle(SITE_REQUEST_LIST_NAME).items.getById(addedItem.Id).update({
                Region: {
                    __metadata: { "type": "SP.Taxonomy.TaxonomyFieldValue" },
                    Label: subSite.Region.Label,
                    TermGuid: subSite.Region.TermGuid,
                    WssId: -1
                }
            }).then(i => {
                console.log(i);
                return true;
            }, (error: any): boolean => {
                return false;
            });
        }, (error: any): boolean => {
            return false;
        });
    }
}