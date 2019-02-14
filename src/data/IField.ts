/**
 * SharePoint field
 */
export interface IField {
  DefaultFormula?: string;
  DefaultValue?: any;
  Description: string;
  EntityPropertyName: string;
  FieldTypeKind: number;
  Group: string;
  Hidden: boolean;
  Indexed: boolean;
  InternalName: string;
  MaxLength?: number;
  ReadOnlyField: boolean;
  Required: boolean;
  SchemaXml: string;
  Scope: string;
  Sortable: boolean;
  StaticName: string;
  Title: string;
  TypeAsString: string;
  TypeDisplayName: string;
  TypeShortDescription: string;
  [index: string]: any;
}

/**
 * Default field
 */
export const NewField: IField = {
  DefaultFormula: undefined,
  DefaultValue: undefined,
  Description: '',
  EntityPropertyName: '',
  FieldTypeKind: 0,
  Group: '',
  Hidden: false,
  Indexed: true,
  InternalName: '',
  MaxLength: undefined,
  ReadOnlyField: false,
  Required: false,
  SchemaXml: '',
  Scope: '',
  Sortable: true,
  StaticName: '',
  Title: '',
  TypeAsString: '',
  TypeDisplayName: '',
  TypeShortDescription: '',
};
