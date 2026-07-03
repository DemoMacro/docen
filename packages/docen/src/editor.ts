/**
 * Full-featured editor entry — a facade over @docen/editor.
 *
 * Importing this module registers the <docen-document> / <docen-workbook> /
 * <docen-presentation> custom elements as a side effect (and pulls in the
 * Fluent UI shell), so it lives on the `docen/editor` subpath rather than the
 * package root. The root entry stays side-effect-free so pure-converter usage
 * (`import { parseDOCX } from "docen"`) remains tree-shakable and does not
 * drag in the UI bundle.
 *
 * @module
 */
export * from "@docen/editor";
