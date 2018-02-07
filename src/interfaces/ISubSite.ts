export interface ISubSite {
    Id?: number;
    Title: string;
    Description: string;
    Url: string;
    GroupName: string;
    GroupOwnerId: number;
    GroupMemberId: number;
    Region: {
        Label: string,
        TermGuid: string
    };
}