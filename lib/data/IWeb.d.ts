import { IContentType } from './IContentType';
import { IFeature } from './IFeature';
import { IField } from './IField';
import { IList } from './IList';
/**
 * SharePoint web or subsite
 */
export interface IWeb {
    ContentTypes?: IContentType[];
    Created: string;
    CustomMasterUrl: string;
    Description: string;
    Features?: IFeature[];
    Fields?: IField[];
    Id: string;
    Language: number;
    LastItemModifiedDate: string;
    Lists?: IList[];
    MasterUrl: string;
    ServerRelativeUrl: string;
    Title: string;
    Url: string;
    WebTemplate: string;
    [index: string]: any;
}
/**
 * Default web
 */
export declare const NewWeb: IWeb;
