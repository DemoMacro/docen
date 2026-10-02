// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

import { ribbonIcon, type RibbonControlOrLayout } from "../ui";
import { presentationRibbonTabs, tableDesignTab, tableLayoutTab } from "./ribbon";

function controlsOf(controls: readonly RibbonControlOrLayout[]): RibbonControlOrLayout[] {
  return controls.flatMap((control) =>
    "controls" in control ? controlsOf(control.controls) : [control],
  );
}

function tabs(): ReturnType<typeof presentationRibbonTabs> {
  return [...presentationRibbonTabs(), tableDesignTab(), tableLayoutTab()];
}

function iconNames(controls: readonly RibbonControlOrLayout[]): string[] {
  return controls.flatMap((control) => {
    const itemIcons =
      control.type === "split"
        ? control.items.flatMap((item) => (item.icon ? [item.icon] : []))
        : [];
    const controlIcon = "icon" in control && typeof control.icon === "string" ? [control.icon] : [];
    return [...controlIcon, ...itemIcons];
  });
}

function layoutOf(control: RibbonControlOrLayout | undefined) {
  return control?.type === "layout" ? control : undefined;
}

describe("pptx ribbon icons", () => {
  it("registers transition and animation gallery glyphs", () => {
    const splits = tabs().flatMap((tab) =>
      tab.groups.flatMap((group) =>
        controlsOf(group.controls).filter((control) => control.type === "split"),
      ),
    ) as Extract<RibbonControlOrLayout, { type: "split" }>[];

    for (const event of ["transition", "effect-options", "animate", "add-animation"]) {
      const split = splits.filter((split) => split.event === event);
      expect(split).not.toHaveLength(0);
      expect(split.every((split) => ribbonIcon(split.icon ?? "")?.includes("<svg"))).toBe(true);
    }

    const items = splits
      .filter((split) => ["transition", "animate", "add-animation"].includes(split.event ?? ""))
      .flatMap((split) => split.items);
    expect(items.length).toBeGreaterThan(0);
    for (const item of items) {
      expect(item.icon).toMatch(/^(transition|animation)-/);
      expect(ribbonIcon(item.icon ?? "")?.includes("<svg")).toBe(true);
    }
  });

  it("previews every built-in pptx table style", () => {
    const tab = tableDesignTab();
    const controls = tab.groups.flatMap((group) => controlsOf(group.controls));
    const split = controls.find(
      (control): control is Extract<RibbonControlOrLayout, { type: "split" }> =>
        control.type === "split" && control.event === "table-style",
    );
    expect(split).toBeDefined();
    expect(split!.size).toBe("large");
    expect(split!.items).toHaveLength(5);
    expect(ribbonIcon(split!.icon ?? "")?.includes("<svg")).toBe(true);
    for (const item of split!.items) {
      expect(ribbonIcon(item.icon ?? "")?.includes("<svg")).toBe(true);
    }

    const shading = controls.find(
      (control): control is Extract<RibbonControlOrLayout, { type: "color-picker" }> =>
        control.type === "color-picker" && control.event === "cell-shading",
    );
    expect(shading?.size).toBe("large");
    const borders = controls.find(
      (control): control is Extract<RibbonControlOrLayout, { type: "split" }> =>
        control.type === "split" && control.event === "table-borders",
    );
    expect(borders).toBeDefined();
    expect(borders!.size).toBe("large");
  });

  it("resolves every ribbon control and menu icon", () => {
    for (const tab of tabs()) {
      for (const icon of iconNames(tab.groups.flatMap((group) => group.controls))) {
        expect(ribbonIcon(icon)?.includes("<svg"), icon).toBe(true);
      }
    }
  });

  it("lays out table layout groups like PowerPoint", () => {
    const groups = Object.fromEntries(tableLayoutTab().groups.map((group) => [group.id, group]));

    expect(groups["table"]!.controls).toHaveLength(3);
    expect(groups["table"]!.controls[2]).toMatchObject({
      type: "split",
      size: "large",
      items: [{ value: "columns" }, { value: "rows" }, { value: "table" }],
    });

    expect(groups["rows-columns"]!.controls[0]).toMatchObject({
      type: "button",
      size: "large",
    });
    const otherInserts = layoutOf(groups["rows-columns"]!.controls[1]);
    expect(otherInserts?.layout).toBe("column");
    expect(otherInserts?.controls).toHaveLength(3);
    expect(otherInserts?.controls.every((control) => control.type === "button")).toBe(true);

    expect(groups["merge"]!.controls).toHaveLength(2);
    expect(
      groups["merge"]!.controls.every((control) => "size" in control && control.size === "large"),
    ).toBe(true);

    const size = layoutOf(groups["cell-size"]!.controls[0]);
    expect(
      size?.controls.map((control) => (control.type === "layout" ? control.layout : "")),
    ).toEqual(["column", "column"]);

    expect(groups["cell-alignment"]!.controls).toHaveLength(3);
    expect(
      layoutOf(groups["cell-alignment"]!.controls[0])?.controls.map((control) =>
        control.type === "layout" ? control.controls.length : 0,
      ),
    ).toEqual([3, 3]);
    expect(
      controlsOf(groups["table-arrange"]!.controls).every(
        (control) => "size" in control && control.size === "large",
      ),
    ).toBe(true);
  });
});
