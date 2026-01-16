import {
  OutputType,
  ISectionOptions,
  IImageOptions,
  IPropertiesOptions,
  ITableCellOptions,
  IParagraphOptions,
  ITableOptions,
  ITableRowOptions,
  ITableOfContentsOptions,
  IParagraphStyleOptions,
} from "docx";

/**
 * Options for exporting TipTap content to DOCX
 */
export interface DocxExportOptions<T extends OutputType = OutputType> {
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
  tableOfContents?: {
    title?: string;
    run?: Partial<ITableOfContentsOptions>;
  };

  image?: {
    // Style definition for image paragraphs
    style?: IParagraphStyleOptions;

    // Image-specific run options (global defaults, can be overridden by node.attrs)
    run?: Partial<IImageOptions>;
  };

  table?: {
    // Style definition for table paragraphs
    style?: IParagraphStyleOptions;

    // Table-level options
    run?: Partial<ITableOptions>;

    row?: {
      paragraph?: Partial<IParagraphOptions>;
      run?: Partial<ITableRowOptions>;
    };
    cell?: {
      paragraph?: Partial<IParagraphOptions>;
      run?: Partial<ITableCellOptions>;
    };
    header?: {
      paragraph?: Partial<IParagraphOptions>;
      run?: Partial<ITableCellOptions>;
    };
  };

  details?: {
    summary?: {
      paragraph?: Partial<IParagraphOptions>;
    };
    content?: {
      paragraph?: Partial<IParagraphOptions>;
    };
  };

  horizontalRule?: {
    paragraph?: Partial<IParagraphOptions>;
  };

  // Export options
  outputType: T;
}
