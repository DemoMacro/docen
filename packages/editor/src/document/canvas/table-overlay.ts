/**
 * Canvas table overlay — Word's interactive table controls on the canvas:
 * 1. Guideline: visual line indicating the column or row boundary being dragged.
 * 2. Corner handle: bottom-right square handle to resize the table proportionally.
 * 3. Quick-Insert buttons: floating (+) buttons on column and row boundaries.
 */

import type { TableZone } from "./caret-map";

export interface TableOverlayCallbacks {
  scale(): number;
  insertColumnAt(colIndex: number, tablePos: number): void;
  insertRowAt(rowIndex: number, tablePos: number): void;
  onCornerDown?(e: MouseEvent): void;
}

export class TableCanvasOverlay {
  readonly el: HTMLDivElement;
  readonly guideline: HTMLDivElement;
  readonly guidelineBox: HTMLDivElement;
  readonly cornerHandle: HTMLDivElement;
  readonly quickColBtn: HTMLDivElement;
  readonly quickRowBtn: HTMLDivElement;
  readonly quickGuide: HTMLDivElement;

  #callbacks: TableOverlayCallbacks;
  #currentZone: TableZone | null = null;
  #quickColIndex = -1;
  #quickRowIndex = -1;

  constructor(callbacks: TableOverlayCallbacks) {
    this.#callbacks = callbacks;

    this.el = document.createElement("div");
    this.el.setAttribute("data-docen-overlay", "");
    Object.assign(this.el.style, {
      position: "absolute",
      inset: "0",
      pointerEvents: "none",
      zIndex: "6",
    } satisfies Partial<CSSStyleDeclaration>);

    // 1. Guideline for single column or row border dragging
    this.guideline = document.createElement("div");
    this.guideline.setAttribute("data-docen-overlay", "");
    Object.assign(this.guideline.style, {
      position: "absolute",
      display: "none",
      pointerEvents: "none",
      backgroundColor: "#2b7cd3",
      boxShadow: "0 0 3px rgba(43, 124, 211, 0.8)",
      zIndex: "10",
    } satisfies Partial<CSSStyleDeclaration>);
    this.el.append(this.guideline);

    // 2. Guideline bounding box for whole-table corner resizing
    this.guidelineBox = document.createElement("div");
    this.guidelineBox.setAttribute("data-docen-overlay", "");
    Object.assign(this.guidelineBox.style, {
      position: "absolute",
      display: "none",
      pointerEvents: "none",
      border: "1.5px dashed #2b7cd3",
      boxSizing: "border-box",
      zIndex: "10",
    } satisfies Partial<CSSStyleDeclaration>);
    this.el.append(this.guidelineBox);

    // 3. Bottom-right corner resize handle
    this.cornerHandle = document.createElement("div");
    this.cornerHandle.setAttribute("data-docen-overlay", "");
    this.cornerHandle.title = "Resize table";
    Object.assign(this.cornerHandle.style, {
      position: "absolute",
      width: "8px",
      height: "8px",
      backgroundColor: "#ffffff",
      border: "1.5px solid #2b7cd3",
      borderRadius: "1px",
      boxSizing: "border-box",
      cursor: "nwse-resize",
      pointerEvents: "auto",
      display: "none",
      zIndex: "8",
    } satisfies Partial<CSSStyleDeclaration>);
    this.cornerHandle.addEventListener("mouseenter", () => {
      this.cornerHandle.style.backgroundColor = "#e8f1fa";
    });
    this.cornerHandle.addEventListener("mouseleave", () => {
      this.cornerHandle.style.backgroundColor = "#ffffff";
    });
    this.cornerHandle.addEventListener("mousedown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.#callbacks.onCornerDown?.(e);
    });
    this.el.append(this.cornerHandle);

    // 4. Quick insert preview guide (dashed line)
    this.quickGuide = document.createElement("div");
    this.quickGuide.setAttribute("data-docen-overlay", "");
    Object.assign(this.quickGuide.style, {
      position: "absolute",
      display: "none",
      pointerEvents: "none",
      zIndex: "7",
    } satisfies Partial<CSSStyleDeclaration>);
    this.el.append(this.quickGuide);

    // 5. Quick-insert column button (+)
    this.quickColBtn = document.createElement("div");
    this.quickColBtn.setAttribute("data-docen-overlay", "");
    this.quickColBtn.title = "Insert column";
    Object.assign(this.quickColBtn.style, {
      position: "absolute",
      width: "16px",
      height: "16px",
      borderRadius: "50%",
      backgroundColor: "#ffffff",
      border: "1.5px solid #2b7cd3",
      color: "#2b7cd3",
      display: "none",
      alignItems: "center",
      justifyContent: "center",
      cursor: "pointer",
      pointerEvents: "auto",
      boxShadow: "0 1px 4px rgba(0, 0, 0, 0.2)",
      zIndex: "9",
      fontSize: "12px",
      lineHeight: "12px",
      userSelect: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    this.quickColBtn.innerHTML =
      '<svg width="10" height="10" viewBox="0 0 10 10" style="display:block;"><path d="M5 1v8M1 5h8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
    this.quickColBtn.addEventListener("mouseenter", () => {
      this.quickColBtn.style.backgroundColor = "#2b7cd3";
      this.quickColBtn.style.color = "#ffffff";
      this.#showColGuide();
    });
    this.quickColBtn.addEventListener("mouseleave", () => {
      this.quickColBtn.style.backgroundColor = "#ffffff";
      this.quickColBtn.style.color = "#2b7cd3";
      this.quickGuide.style.display = "none";
    });
    this.quickColBtn.addEventListener("mousedown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (this.#quickColIndex > 0 && this.#currentZone?.tablePos != null) {
        this.#callbacks.insertColumnAt(this.#quickColIndex, this.#currentZone.tablePos);
      }
    });
    this.el.append(this.quickColBtn);

    // 6. Quick-insert row button (+)
    this.quickRowBtn = document.createElement("div");
    this.quickRowBtn.setAttribute("data-docen-overlay", "");
    this.quickRowBtn.title = "Insert row";
    Object.assign(this.quickRowBtn.style, {
      position: "absolute",
      width: "16px",
      height: "16px",
      borderRadius: "50%",
      backgroundColor: "#ffffff",
      border: "1.5px solid #2b7cd3",
      color: "#2b7cd3",
      display: "none",
      alignItems: "center",
      justifyContent: "center",
      cursor: "pointer",
      pointerEvents: "auto",
      boxShadow: "0 1px 4px rgba(0, 0, 0, 0.2)",
      zIndex: "9",
      fontSize: "12px",
      lineHeight: "12px",
      userSelect: "none",
    } satisfies Partial<CSSStyleDeclaration>);
    this.quickRowBtn.innerHTML =
      '<svg width="10" height="10" viewBox="0 0 10 10" style="display:block;"><path d="M5 1v8M1 5h8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
    this.quickRowBtn.addEventListener("mouseenter", () => {
      this.quickRowBtn.style.backgroundColor = "#2b7cd3";
      this.quickRowBtn.style.color = "#ffffff";
      this.#showRowGuide();
    });
    this.quickRowBtn.addEventListener("mouseleave", () => {
      this.quickRowBtn.style.backgroundColor = "#ffffff";
      this.quickRowBtn.style.color = "#2b7cd3";
      this.quickGuide.style.display = "none";
    });
    this.quickRowBtn.addEventListener("mousedown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (this.#quickRowIndex > 0 && this.#currentZone?.tablePos != null) {
        this.#callbacks.insertRowAt(this.#quickRowIndex, this.#currentZone.tablePos);
      }
    });
    this.el.append(this.quickRowBtn);
  }

  /** Ensure the overlay element is attached to the target page frame container. */
  attachTo(frame: HTMLElement): void {
    if (this.el.parentElement !== frame) {
      frame.append(this.el);
    }
  }

  /** Show and place the corner resize handle at the table's bottom-right corner. */
  showCorner(zone: TableZone, scale: number): void {
    this.#currentZone = zone;
    const left = (zone.xPx + zone.widthPx) * scale - 4;
    const top = (zone.yPx + zone.heightPx) * scale - 4;
    this.cornerHandle.style.left = `${left}px`;
    this.cornerHandle.style.top = `${top}px`;
    this.cornerHandle.style.display = "block";
  }

  /** Hide the corner handle. */
  hideCorner(): void {
    this.cornerHandle.style.display = "none";
  }

  /** Show the quick-insert column (+) button at the top boundary. */
  showQuickCol(zone: TableZone, colIndex: number, scale: number): void {
    this.#currentZone = zone;
    this.#quickColIndex = colIndex;
    const borderX = zone.xPx + zone.colEdges[colIndex]!;
    this.quickColBtn.style.left = `${borderX * scale - 8}px`;
    this.quickColBtn.style.top = `${zone.yPx * scale - 18}px`;
    this.quickColBtn.style.display = "flex";
  }

  /** Hide the quick-insert column button and guide. */
  hideQuickCol(): void {
    this.quickColBtn.style.display = "none";
    if (this.#quickColIndex > 0) {
      this.quickGuide.style.display = "none";
      this.#quickColIndex = -1;
    }
  }

  /** Show the quick-insert row (+) button at the left boundary. */
  showQuickRow(zone: TableZone, rowIndex: number, scale: number): void {
    this.#currentZone = zone;
    this.#quickRowIndex = rowIndex;
    const borderY = zone.yPx + zone.rowEdges[rowIndex]!;
    this.quickRowBtn.style.left = `${zone.xPx * scale - 18}px`;
    this.quickRowBtn.style.top = `${borderY * scale - 8}px`;
    this.quickRowBtn.style.display = "flex";
  }

  /** Hide the quick-insert row button and guide. */
  hideQuickRow(): void {
    this.quickRowBtn.style.display = "none";
    if (this.#quickRowIndex > 0) {
      this.quickGuide.style.display = "none";
      this.#quickRowIndex = -1;
    }
  }

  #showColGuide(): void {
    if (!this.#currentZone || this.#quickColIndex <= 0) return;
    const scale = this.#callbacks.scale();
    const borderX = this.#currentZone.xPx + this.#currentZone.colEdges[this.#quickColIndex]!;
    Object.assign(this.quickGuide.style, {
      left: `${borderX * scale}px`,
      top: `${this.#currentZone.yPx * scale}px`,
      width: "1px",
      height: `${this.#currentZone.heightPx * scale}px`,
      borderLeft: "1.5px dashed #2b7cd3",
      borderTop: "none",
      display: "block",
    } satisfies Partial<CSSStyleDeclaration>);
  }

  #showRowGuide(): void {
    if (!this.#currentZone || this.#quickRowIndex <= 0) return;
    const scale = this.#callbacks.scale();
    const borderY = this.#currentZone.yPx + this.#currentZone.rowEdges[this.#quickRowIndex]!;
    Object.assign(this.quickGuide.style, {
      left: `${this.#currentZone.xPx * scale}px`,
      top: `${borderY * scale}px`,
      width: `${this.#currentZone.widthPx * scale}px`,
      height: "1px",
      borderTop: "1.5px dashed #2b7cd3",
      borderLeft: "none",
      display: "block",
    } satisfies Partial<CSSStyleDeclaration>);
  }

  /** Show the column drag guideline at a specific page-local X. */
  showColGuideline(xPx: number, topPx: number, heightPx: number, scale: number): void {
    Object.assign(this.guideline.style, {
      left: `${xPx * scale - 1}px`,
      top: `${topPx * scale}px`,
      width: "2px",
      height: `${heightPx * scale}px`,
      display: "block",
    } satisfies Partial<CSSStyleDeclaration>);
  }

  /** Show the row drag guideline at a specific page-local Y. */
  showRowGuideline(leftPx: number, yPx: number, widthPx: number, scale: number): void {
    Object.assign(this.guideline.style, {
      left: `${leftPx * scale}px`,
      top: `${yPx * scale - 1}px`,
      width: `${widthPx * scale}px`,
      height: "2px",
      display: "block",
    } satisfies Partial<CSSStyleDeclaration>);
  }

  /** Hide the drag guideline. */
  hideGuideline(): void {
    this.guideline.style.display = "none";
  }

  /** Show the bounding box guideline for whole-table corner resize. */
  showBoxGuideline(
    xPx: number,
    yPx: number,
    widthPx: number,
    heightPx: number,
    scale: number,
  ): void {
    Object.assign(this.guidelineBox.style, {
      left: `${xPx * scale}px`,
      top: `${yPx * scale}px`,
      width: `${widthPx * scale}px`,
      height: `${heightPx * scale}px`,
      display: "block",
    } satisfies Partial<CSSStyleDeclaration>);
  }

  /** Hide the bounding box guideline. */
  hideBoxGuideline(): void {
    this.guidelineBox.style.display = "none";
  }

  /** Hide all table overlays (e.g. when leaving the table or during a non-table drag). */
  hideAll(): void {
    this.hideCorner();
    this.hideQuickCol();
    this.hideQuickRow();
    this.hideGuideline();
    this.hideBoxGuideline();
    this.quickGuide.style.display = "none";
    this.#currentZone = null;
  }

  destroy(): void {
    this.hideAll();
    this.el.remove();
  }
}
