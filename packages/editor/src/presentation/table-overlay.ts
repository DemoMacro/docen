import type { Box } from "../drawing/geometry";
import {
  tableGripAt,
  tableEdges,
  tableSelectionRects,
  type TableGripHit,
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
  #grip: HTMLDivElement | null = null;
  #gripHit: TableGripHit | null = null;

  constructor(callbacks: TableOverlayCallbacks) {
    this.#callbacks = callbacks;
    this.el = document.createElement("div");
    this.el.setAttribute("data-docen-overlay", "table-selection");
    Object.assign(this.el.style, {
      position: "absolute",
      zIndex: "8",
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
    this.#gripHit = null;
    this.el.style.display = "none";
    this.#placeHighlights();
    if (this.#grip) this.#grip.style.display = "none";
  }

  /** DOCX shows one grip at a time: the hover point resolves a strip, corner,
   *  or table-wide square and the overlay paints just that target. */
  hover(localX: number, localY: number, editing = false): void {
    const member = this.#member;
    if (!member) return;
    this.#gripHit = tableGripAt(member, member.x + localX, member.y + localY, true);
    if (editing && this.#gripHit && !this.#gripHit.clickable) this.#gripHit = null;
    this.#placeGrip();
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
    this.#placeGrip();
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
        background: "rgba(0, 120, 215, 0.38)",
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

  #placeGrip(): void {
    const member = this.#member;
    if (!member || !this.#gripHit) {
      if (this.#grip) this.#grip.style.display = "none";
      return;
    }
    const { colEdges, rowEdges } = tableEdges(member);
    const scale = this.#callbacks.scale();
    if (!this.#grip) {
      this.#grip = document.createElement("div");
      this.#grip.setAttribute("data-docen-overlay", "table-grip");
      this.#grip.innerHTML =
        '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M1 6h7M5 2l4 4-4 4z" fill="#454545"/></svg>' +
        '<svg width="12" height="12" viewBox="0 0 12 12"><rect x="0.5" y="0.5" width="11" height="11" fill="#ffffff" stroke="#7f7f7f"/><path d="M6 1v10M1 6h10" stroke="#7f7f7f"/></svg>';
      this.#grip.addEventListener("pointerdown", (event) => {
        if (!this.#gripHit?.clickable) return;
        event.preventDefault();
        event.stopPropagation();
        this.#callbacks.selectGrip(this.#gripHit.kind, this.#gripHit.index);
      });
      this.el.append(this.#grip);
    }
    const grip = this.#grip;
    const hit = this.#gripHit;
    const box =
      hit.kind === "col"
        ? {
            left: `${colEdges[hit.index]! * scale}px`,
            top: `-${14 * scale}px`,
            width: `${(colEdges[hit.index + 1]! - colEdges[hit.index]!) * scale}px`,
            height: `${12 * scale}px`,
          }
        : hit.kind === "row"
          ? {
              left: `-${14 * scale}px`,
              top: `${rowEdges[hit.index]! * scale}px`,
              width: `${12 * scale}px`,
              height: `${(rowEdges[hit.index + 1]! - rowEdges[hit.index]!) * scale}px`,
            }
          : {
              left: `-${13 * scale}px`,
              top: `-${13 * scale}px`,
              width: `${12 * scale}px`,
              height: `${12 * scale}px`,
            };
    Object.assign(grip.style, {
      ...box,
      position: "absolute",
      pointerEvents: hit.clickable ? "auto" : "none",
      cursor: hit.clickable ? "pointer" : "default",
      display: "block",
      background: "transparent",
    } satisfies Partial<CSSStyleDeclaration>);
    const [arrow, grid] = Array.from(grip.children) as SVGElement[];
    const rotate = hit.kind === "col" ? "rotate(90deg)" : "";
    arrow?.setAttribute(
      "style",
      hit.kind === "table" ? "display:none" : `display:block;margin:auto;transform:${rotate}`,
    );
    grid?.setAttribute("style", hit.kind === "table" ? "display:block" : "display:none");
  }
}
