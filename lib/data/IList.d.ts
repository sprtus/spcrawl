import { IContentType } from './IContentType';
import { IField } from './IField';
import { IItem } from './IItem';
/**
 * SharePoint list
 */
export interface IList {
    BaseTemplate: number;
    BaseType: number;
    ContentTypes: IContentType[];
    ContentTypesEnabled: boolean;
    Created: string;
    Description: string;
    DocumentTemplateUrl: string;
    DraftVersionVisibility: number;
    EnableAttachments: boolean;
    EnableFolderCreation: boolean;
    EnableMinorVersions: boolean;
    EnableModeration: boolean;
    EnableRequestSignOff: boolean;
    EnableVersioning: boolean;
    EntityTypeName: string;
    Fields: IField[];
    ForceCheckout: boolean;
    HasExternalDataSource: boolean;
    Hidden: boolean;
    Id: string;
    ImageUrl: string;
    IsApplicationList: boolean;
    IsCatalog: boolean;
    IsPrivate: boolean;
    ItemCount: number;
    Items: IItem[];
    LastItemDeletedDate: string;
    LastItemModifiedDate: string;
    LastItemUserModifiedDate: string;
    ListExperienceOptions: number;
    ListItemEntityTypeFullName: string;
    MajorVersionLimit: number;
    MajorWithMinorVersionsLimit: number;
    MultipleDataList: boolean;
    NoCrawl: boolean;
    ParentWebUrl: string;
    TemplateFeatureId: string;
    Title: string;
    [index: string]: any;
}
