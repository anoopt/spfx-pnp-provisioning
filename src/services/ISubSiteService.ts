import { ISubSite } from "../interfaces/ISubSite";

export interface ISubSiteService {
    addSubSite(subSite: ISubSite): Promise<boolean>;
}