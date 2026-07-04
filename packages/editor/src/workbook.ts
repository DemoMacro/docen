/**
 * `<docen-workbook>` — XLSX editor element (stub).
 *
 * Placeholder for the future RevoGrid-based workbook. Signature mirrors
 * `<docen-document>` so the public API is stable when the engine lands.
 */
class DocenWorkbook extends HTMLElement {
  connectedCallback(): void {
    console.warn("[docen-workbook] not yet implemented");
  }
}

customElements.define("docen-workbook", DocenWorkbook);

export default DocenWorkbook;
