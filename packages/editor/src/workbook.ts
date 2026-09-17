/**
 * `<docen-workbook>` — XLSX editor element (stub).
 *
 * Placeholder for the future RevoGrid-based workbook. Signature mirrors
 * `<docen-document>` so the public API is stable when the engine lands.
 * The export is wired in three places that move together — src/index.ts
 * (`DocenWorkbook`), package.json `exports` (`./workbook`), vite.config.ts
 * `pack.entry` — update them in lockstep, don't drop any one.
 */
class DocenWorkbook extends HTMLElement {
  connectedCallback(): void {
    console.warn("[docen-workbook] not yet implemented");
  }
}

customElements.define("docen-workbook", DocenWorkbook);

export default DocenWorkbook;
