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
import type { DocxImageExportHandler } from "./utils/image";

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
    // Custom image handler to replace default fetch behavior
    handler?: DocxImageExportHandler;

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

  code?: {
    // Style definition for code block paragraphs
    style?: IParagraphStyleOptions;
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
