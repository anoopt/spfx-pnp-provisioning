import { ISubSite } from "../../../interfaces/ISubSite";

export interface ICreateSiteState {
    loadingScripts: boolean;
    errors?: string[];
    status: JSX.Element;
    subSiteInContext: ISubSite;
}