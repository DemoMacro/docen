import type { Box } from "../drawing/geometry";
import {
  tableEdges,
  tableSelectionRects,
  type TableSelectionRange,
  type TableSelectionView,
} from "./table-selection";

export interface TableOverlayCallbacks {
  scale(): number;
  selectGrip(kind: "row" | "col" | "table", index?: number): void;
}

/** Word's table selection layer: one translucent box per selected grid cell
 *  and edge grips for row/column/table selection. It deliberately has no
 *  resize or rotate handles — a table is selected and edited as cells. */
export class TableSelectionOverlay {
  readonly el: HTMLDivElement;
  #callbacks: TableOverlayCallbacks;
  #member: TableSelectionView | null = null;
  #stripY = 0;
  #selection: TableSelectionRange | null = null;
  #highlights: HTMLDivElement[] = [];
  #grips: HTMLDivElement[] = [];

  constructor(callbacks: TableOverlayCallbacks) {
    this.#callbacks = callbacks;
    this.el = document.createElement("div");
    this.el.setAttribute("data-docen-overlay", "table-selection");
    Object.assign(this.el.style, {
      position: "absolute",
      zIndex: "7",
      pointerEvents: "none",
      display: "none",
    } satisfies Partial<CSSStyleDeclaration>);
  }

  show(member: TableSelectionView, stripY: number, selection: TableSelectionRange | null): void {
    this.#member = member;
    this.#stripY = stripY;
    this.#selection = selection;
    this.el.style.display = "block";
    this.#place();
  }

  setSelection(selection: TableSelectionRange | null): void {
    this.#selection = selection;
    this.#placeHighlights();
  }

  hide(): void {
    this.#member = null;
    this.#selection = null;
    this.el.style.display = "none";
    this.#placeHighlights();
  }

  #place(): void {
    const member = this.#member;
    if (!member) return;
    const scale = this.#callbacks.scale();
    const width = member.table.columnWidthsPx.reduce((sum, width) => sum + width, 0);
    const height = member.table.rows.reduce((sum, row) => sum + row.heightPx, 0);
    Object.assign(this.el.style, {
      left: `${member.x * scale}px`,
      top: `${(member.y + this.#stripY) * scale}px`,
      width: `${width * scale}px`,
      height: `${height * scale}px`,
    });
    this.#placeHighlights();
    this.#placeGrips();
  }

  #placeHighlights(): void {
    const member = this.#member;
    if (!member || !this.#selection) {
      this.#highlights.forEach((el) => {
        el.style.display = "none";
      });
      return;
    }
    const origin = member;
    const rects = tableSelectionRects(member, this.#selection);
    while (this.#highlights.length < rects.length) {
      const rect = document.createElement("div");
      Object.assign(rect.style, {
        position: "absolute",
        pointerEvents: "none",
        background: "rgba(0, 120, 215, 0.25)",
      } satisfies Partial<CSSStyleDeclaration>);
      this.el.append(rect);
      this.#highlights.push(rect);
    }
    const scale = this.#callbacks.scale();
    this.#highlights.forEach((el, index) => {
      const rect: Box | undefined = rects[index];
      if (!rect) {
        el.style.display = "none";
        return;
      }
      Object.assign(el.style, {
        display: "block",
        left: `${(rect.x - origin.x) * scale}px`,
        top: `${(rect.y - origin.y) * scale}px`,
        width: `${rect.width * scale}px`,
        height: `${rect.height * scale}px`,
      });
    });
  }

  #placeGrips(): void {
    const member = this.#member;
    if (!member) return;
    const { colEdges, rowEdges } = tableEdges(member);
    const scale = this.#callbacks.scale();
    const wanted = colEdges.length - 1 + rowEdges.length - 1 + 1;
    while (this.#grips.length < wanted) {
      const grip = document.createElement("div");
      grip.className = "table-grip";
      grip.setAttribute("data-docen-overlay", "table-grip");
      Object.assign(grip.style, {
        position: "absolute",
        pointerEvents: "auto",
        background: "transparent",
        cursor: "pointer",
      } satisfies Partial<CSSStyleDeclaration>);
      grip.addEventListener("pointerdown", (event) => {
        event.preventDefault();
        event.stopPropagation();
        this.#callbacks.selectGrip(
          grip.dataset.gripKind as "row" | "col" | "table",
          grip.dataset.gripIndex ? Number(grip.dataset.gripIndex) : undefined,
        );
      });
      this.el.append(grip);
      this.#grips.push(grip);
    }
    this.#grips.forEach((grip, index) => {
      if (index < colEdges.length - 1) {
        grip.dataset.gripKind = "col";
        grip.dataset.gripIndex = String(index);
        Object.assign(grip.style, {
          left: `${colEdges[index]! * scale}px`,
          top: `-${8 * scale}px`,
          width: `${(colEdges[index + 1]! - colEdges[index]!) * scale}px`,
          height: `${8 * scale}px`,
        });
        return;
      }
      const rowIndex = index - (colEdges.length - 1);
      if (rowIndex < rowEdges.length - 1) {
        grip.dataset.gripKind = "row";
        grip.dataset.gripIndex = String(rowIndex);
        Object.assign(grip.style, {
          left: `-${8 * scale}px`,
          top: `${rowEdges[rowIndex]! * scale}px`,
          width: `${8 * scale}px`,
          height: `${(rowEdges[rowIndex + 1]! - rowEdges[rowIndex]!) * scale}px`,
        });
        return;
      }
      grip.dataset.gripKind = "table";
      delete grip.dataset.gripIndex;
      Object.assign(grip.style, {
        left: `-${10 * scale}px`,
        top: `-${10 * scale}px`,
        width: `${10 * scale}px`,
        height: `${10 * scale}px`,
        cursor: "crosshair",
      });
    });
  }
}
