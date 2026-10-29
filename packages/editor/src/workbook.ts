/**
 * `<docen-workbook>` — XLSX editor element (stub).
 *
 * Placeholder for the future RevoGrid-based workbook. This is an explicit
 * stub; its final host API lands with the workbook engine.
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
