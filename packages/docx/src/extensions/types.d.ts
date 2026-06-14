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
