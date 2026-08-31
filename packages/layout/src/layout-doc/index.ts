// LayoutDoc — the engine's own input projection. Adapters (docx from
// Tiptap/ProseMirror, pptx from its shape tree, xlsx from its grid) build
// these plain blocks; the engine never sees a Tiptap node or an OOXML part.
//
// Projection contract:
// - All geometry is px. Unit conversions happen once, in the adapter.
// - Style cascades are resolved. `spacing`/`indent`/text styles arrive as the
//   effective values (direct attrs → style chain → document defaults merged);
//   the engine has no styles table and no `styleId` — re-resolving a cascade
//   here would duplicate each format's semantics it must not know about.
// - Only layout semantics keep their OOXML shape (a line rule, a grid pitch,
//   a snap flag) — those are the rules this engine exists to implement.

export * from "./inline";
export * from "./drawing";
export * from "./table";
export * from "./block";
export * from "./page";
