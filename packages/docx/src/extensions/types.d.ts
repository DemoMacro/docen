import type { RunOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

declare module "@tiptap/core" {
  interface NodeConfig<Options, Storage> {
    /**
     * DOCX serialization: Tiptap JSON node → DOCX opts.
     * Each node extension defines this to convert its attrs to DOCX properties.
     */
    renderDocx?: (node: JSONContent) => Record<string, unknown>;
    /**
     * DOCX deserialization: DOCX opts → Tiptap JSON attrs.
     * Each node extension defines this to convert DOCX properties back to attrs.
     */
    parseDocx?: (opts: Record<string, unknown>) => Record<string, unknown>;
  }

  interface MarkConfig<Options, Storage> {
    /**
     * DOCX serialization: mark attrs → RunOptions properties.
     * Each mark extension defines this to contribute run-level properties.
     */
    renderDocx?: (attrs: Record<string, unknown>) => Partial<RunOptions>;
    /**
     * DOCX deserialization: RunOptions → mark attrs.
     * Each mark extension defines this to extract its attrs from run properties.
     */
    parseDocx?: (opts: RunOptions) => Record<string, unknown>;
  }
}
