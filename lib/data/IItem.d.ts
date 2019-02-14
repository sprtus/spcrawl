/**
 * SharePoint list item
 */
export interface IItem {
    AuthorId: number;
    ContentTypeId: string;
    Created: string;
    EditorId: number;
    FielSystemObjectType: number;
    GUID: string;
    Id: number;
    Modified: string;
    Title?: string;
    [index: string]: any;
}
