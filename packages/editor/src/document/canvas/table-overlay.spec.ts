// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest";

import type { TableZone } from "./caret-map";
import { TableCanvasOverlay } from "./table-overlay";

describe("TableCanvasOverlay", () => {
  const mockZone: TableZone = {
    page: 0,
    xPx: 100,
    yPx: 50,
    widthPx: 300,
    heightPx: 150,
    colEdges: [0, 100, 200, 300],
    rowEdges: [0, 50, 100, 150],
    tablePos: 12,
  };

  it("creates overlay elements with correct default styles and attributes", () => {
    const callbacks = {
      scale: () => 1,
      insertColumnAt: vi.fn(),
      insertRowAt: vi.fn(),
    };
    const overlay = new TableCanvasOverlay(callbacks);

    expect(overlay.el.hasAttribute("data-docen-overlay")).toBe(true);
    expect(overlay.guideline.hasAttribute("data-docen-overlay")).toBe(true);
    expect(overlay.cornerHandle.hasAttribute("data-docen-overlay")).toBe(true);
    expect(overlay.quickColBtn.hasAttribute("data-docen-overlay")).toBe(true);
    expect(overlay.quickRowBtn.hasAttribute("data-docen-overlay")).toBe(true);

    expect(overlay.guideline.style.display).toBe("none");
    expect(overlay.cornerHandle.style.display).toBe("none");
    expect(overlay.quickColBtn.style.display).toBe("none");
    expect(overlay.quickRowBtn.style.display).toBe("none");
  });

  it("attaches to container frame", () => {
    const frame = document.createElement("div");
    const overlay = new TableCanvasOverlay({
      scale: () => 1,
      insertColumnAt: vi.fn(),
      insertRowAt: vi.fn(),
    });

    overlay.attachTo(frame);
    expect(overlay.el.parentElement).toBe(frame);
  });

  it("positions and displays corner handle correctly", () => {
    const overlay = new TableCanvasOverlay({
      scale: () => 2,
      insertColumnAt: vi.fn(),
      insertRowAt: vi.fn(),
    });

    overlay.showCorner(mockZone, 2);
    // (xPx + widthPx) * scale - 4 = (100 + 300) * 2 - 4 = 796px
    // (yPx + heightPx) * scale - 4 = (50 + 150) * 2 - 4 = 396px
    expect(overlay.cornerHandle.style.left).toBe("796px");
    expect(overlay.cornerHandle.style.top).toBe("396px");
    expect(overlay.cornerHandle.style.display).toBe("block");

    overlay.hideCorner();
    expect(overlay.cornerHandle.style.display).toBe("none");
  });

  it("dispatches onCornerDown callback when corner handle is clicked", () => {
    const onCornerDown = vi.fn();
    const overlay = new TableCanvasOverlay({
      scale: () => 1,
      insertColumnAt: vi.fn(),
      insertRowAt: vi.fn(),
      onCornerDown,
    });

    const event = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
    overlay.cornerHandle.dispatchEvent(event);

    expect(onCornerDown).toHaveBeenCalledTimes(1);
    expect(event.defaultPrevented).toBe(true);
  });

  it("positions quick column insert button and triggers insertColumnAt on mousedown", () => {
    const insertColumnAt = vi.fn();
    const overlay = new TableCanvasOverlay({
      scale: () => 1,
      insertColumnAt,
      insertRowAt: vi.fn(),
    });

    overlay.showQuickCol(mockZone, 1, 1);
    // borderX = 100 + 100 = 200. left = 200 * 1 - 8 = 192px. top = 50 * 1 - 18 = 32px.
    expect(overlay.quickColBtn.style.left).toBe("192px");
    expect(overlay.quickColBtn.style.top).toBe("32px");
    expect(overlay.quickColBtn.style.display).toBe("flex");

    const event = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
    overlay.quickColBtn.dispatchEvent(event);

    expect(insertColumnAt).toHaveBeenCalledWith(1, 12);
  });

  it("positions quick row insert button and triggers insertRowAt on mousedown", () => {
    const insertRowAt = vi.fn();
    const overlay = new TableCanvasOverlay({
      scale: () => 1,
      insertColumnAt: vi.fn(),
      insertRowAt,
    });

    overlay.showQuickRow(mockZone, 2, 1);
    // left = 100 * 1 - 18 = 82px. borderY = 50 + 100 = 150. top = 150 * 1 - 8 = 142px.
    expect(overlay.quickRowBtn.style.left).toBe("82px");
    expect(overlay.quickRowBtn.style.top).toBe("142px");
    expect(overlay.quickRowBtn.style.display).toBe("flex");

    const event = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
    overlay.quickRowBtn.dispatchEvent(event);

    expect(insertRowAt).toHaveBeenCalledWith(2, 12);
  });

  it("shows column and row guidelines accurately", () => {
    const overlay = new TableCanvasOverlay({
      scale: () => 1,
      insertColumnAt: vi.fn(),
      insertRowAt: vi.fn(),
    });

    overlay.showColGuideline(250, 50, 150, 1);
    expect(overlay.guideline.style.left).toBe("249px");
    expect(overlay.guideline.style.top).toBe("50px");
    expect(overlay.guideline.style.width).toBe("2px");
    expect(overlay.guideline.style.height).toBe("150px");
    expect(overlay.guideline.style.display).toBe("block");

    overlay.hideGuideline();
    expect(overlay.guideline.style.display).toBe("none");

    overlay.showRowGuideline(100, 120, 300, 1);
    expect(overlay.guideline.style.left).toBe("100px");
    expect(overlay.guideline.style.top).toBe("119px");
    expect(overlay.guideline.style.width).toBe("300px");
    expect(overlay.guideline.style.height).toBe("2px");
    expect(overlay.guideline.style.display).toBe("block");
  });

  it("shows table bounding box guideline for corner scaling", () => {
    const overlay = new TableCanvasOverlay({
      scale: () => 1,
      insertColumnAt: vi.fn(),
      insertRowAt: vi.fn(),
    });

    overlay.showBoxGuideline(100, 50, 320, 160, 1);
    expect(overlay.guidelineBox.style.left).toBe("100px");
    expect(overlay.guidelineBox.style.top).toBe("50px");
    expect(overlay.guidelineBox.style.width).toBe("320px");
    expect(overlay.guidelineBox.style.height).toBe("160px");
    expect(overlay.guidelineBox.style.display).toBe("block");

    overlay.hideBoxGuideline();
    expect(overlay.guidelineBox.style.display).toBe("none");
  });

  it("hides all overlays and removes element on destroy", () => {
    const frame = document.createElement("div");
    const overlay = new TableCanvasOverlay({
      scale: () => 1,
      insertColumnAt: vi.fn(),
      insertRowAt: vi.fn(),
    });
    overlay.attachTo(frame);
    overlay.showCorner(mockZone, 1);
    overlay.showQuickCol(mockZone, 1, 1);

    expect(frame.contains(overlay.el)).toBe(true);

    overlay.destroy();
    expect(frame.contains(overlay.el)).toBe(false);
    expect(overlay.cornerHandle.style.display).toBe("none");
    expect(overlay.quickColBtn.style.display).toBe("none");
  });
});
