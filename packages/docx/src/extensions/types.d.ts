import type { RunOptions } from "@office-open/docx";
import type { JSONContent } from "@tiptap/core";

declare module "@tiptap/core" {
  interface NodeConfig<Options = any, Storage = any> {
    /**
     * DOCX serialization: Tiptap JSON node → DOCX opts, or null when the node
     * cannot be serialized (e.g. an image with no embedded data — DocxManager
     * then drops it). Each node extension defines this to convert its attrs to
     * DOCX properties.
     */
    renderDocx?: (node: JSONContent) => Record<string, unknown> | null;
    /**
     * DOCX deserialization: DOCX opts → Tiptap JSON attrs.
     * Each node extension defines this to convert DOCX properties back to attrs.
     */
    parseDocx?: (opts: Record<string, unknown>) => Record<string, unknown>;
  }

  interface MarkConfig<Options = any, Storage = any> {
    /**
     * DOCX serialization: mark attrs → run-level properties (merged into the
     * run's options). Each mark extension defines this to contribute rPr fields.
     */
    renderDocx?: (attrs: Record<string, unknown>) => Record<string, unknown>;
    /**
     * DOCX deserialization: RunOptions → mark attrs, or null when the run does
     * not carry this mark (DocxManager then skips emitting it). Each mark
     * extension defines this to extract its attrs from run properties.
     */
    parseDocx?: (opts: RunOptions) => Record<string, unknown> | null;
  }
}
