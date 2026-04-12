import {
  OutputType,
  PatchType,
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
import type { JSONContent } from "@tiptap/core";
import type { DocxImageExportHandler } from "./utils/image";

export interface DocxExportOptions<T extends OutputType = OutputType> {
  // Document properties
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

  tableOfContents?: {
    title?: string;
    run?: Partial<ITableOfContentsOptions>;
  };

  image?: {
    handler?: DocxImageExportHandler;
    style?: IParagraphStyleOptions;
    run?: Partial<IImageOptions>;
  };

  table?: {
    style?: IParagraphStyleOptions;
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

  outputType: T;
}

export interface DocxPatchContent {
  /** @default PatchType.DOCUMENT */
  type?: typeof PatchType.DOCUMENT | typeof PatchType.PARAGRAPH;
  content: JSONContent;
}

export interface DocxPatchOptions<T extends OutputType = OutputType> {
  template: ArrayBuffer | Buffer | Uint8Array | Blob | string;
  patches: Record<string, DocxPatchContent>;
  placeholderDelimiters?: { start?: string; end?: string };
  /** @default false */
  keepOriginalStyles?: boolean;
  exportOptions: Omit<DocxExportOptions<T>, "outputType">;
  outputType: T;
}
