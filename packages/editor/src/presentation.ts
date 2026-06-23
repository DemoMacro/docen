/**
 * `<docen-presentation>` — PPTX editor super-component (stub).
 *
 * Placeholder for the future LeaferJS-based presentation. Signature mirrors
 * `<docen-document>` so the public API is stable when the engine lands.
 */
class DocenPresentation extends HTMLElement {
  connectedCallback(): void {
    console.warn("[docen-presentation] not yet implemented");
  }
}

customElements.define("docen-presentation", DocenPresentation);

export default DocenPresentation;
