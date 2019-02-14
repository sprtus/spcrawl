/**
 * SharePoint content type
 */
export interface IContentType {
    Description: string;
    Group: string;
    Id: string;
    Name: string;
    [index: string]: any;
}
