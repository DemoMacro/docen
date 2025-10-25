import {
  OutputType,
  ISectionOptions,
  IImageOptions,
  IPropertiesOptions,
  ITableCellOptions,
  IParagraphOptions,
  ITableOptions,
  ITableRowOptions,
} from "docx";

/**
 * Options for generating DOCX documents
 */
export interface PropertiesOptions<T extends OutputType = OutputType> {
  // === IPropertiesOptions fields (in order) ===
  sections?: ISectionOptions[];
  title?: string;
  subject?: string;
  creator?: string;
  keywords?: string;
  description?: string;
  lastModifiedBy?: string;
  revision?: number;
  externalStyles?: IPropertiesOptions["externalStyles"];
  styles?: IPropertiesOptions["styles"];
  numbering?: IPropertiesOptions["numbering"];
  comments?: IPropertiesOptions["comments"];
  footnotes?: IPropertiesOptions["footnotes"];
  background?: IPropertiesOptions["background"];
  features?: IPropertiesOptions["features"];
  compatabilityModeVersion?: IPropertiesOptions["compatabilityModeVersion"];
  compatibility?: IPropertiesOptions["compatibility"];
  customProperties?: IPropertiesOptions["customProperties"];
  evenAndOddHeaderAndFooters?: IPropertiesOptions["evenAndOddHeaderAndFooters"];
  defaultTabStop?: IPropertiesOptions["defaultTabStop"];
  fonts?: IPropertiesOptions["fonts"];
  hyphenation?: IPropertiesOptions["hyphenation"];

  // === Specific options ===
  image?: {
    paragraph?: Partial<IParagraphOptions>;
    run?: Partial<IImageOptions>;
  };

  table?: {
    properties?: Partial<ITableOptions>;
    row?: Partial<ITableRowOptions>;
    cell?: Partial<ITableCellOptions>;
    header?: Partial<ITableCellOptions>;
  };

  // Export options
  outputType: T;
}
