import { IWeb } from './IWeb';
/**
 * SharePoint site collection
 */
export interface ISite {
    LastScan: string;
    Url: string;
    Webs: IWeb[];
    [index: string]: any;
}
