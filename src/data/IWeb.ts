import { IContentType } from './IContentType';
import { IField } from './IField';
import { IList } from './IList';

/**
 * SharePoint web or subsite
 */
export interface IWeb {
  ContentTypes: IContentType[];
  Created: string;
  CustomMasterUrl: string;
  Description: string;
  Fields: IField[];
  Id: string;
  Language: number;
  LastItemModifiedDate: string;
  Lists: IList[];
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
export const NewWeb: IWeb = {
  ContentTypes: [],
  Created: '',
  CustomMasterUrl: '',
  Description: '',
  Fields: [],
  Id: '',
  Language: 0,
  LastItemModifiedDate: '',
  Lists: [],
  MasterUrl: '',
  ServerRelativeUrl: '',
  Title: '',
  Url: '',
  WebTemplate: '',
};
