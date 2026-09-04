// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";

import { LOCAL_HANDLED } from "./file-formats";
import "./i18n";
import { headerFooterContextTab } from "./ribbon";

describe("headerFooterContextTab", () => {
  it("defines the contextual tab for Header & Footer tools", () => {
    const tab = headerFooterContextTab();
    expect(tab.id).toBe("header-footer-tab");
    expect(tab.contextual).toBe(true);

    const groupIds = tab.groups.map((g) => g.id);
    expect(groupIds).toEqual(["navigation", "options", "close"]);
  });

  it("wires header and footer contextual commands into LOCAL_HANDLED", () => {
    expect(LOCAL_HANDLED.has("close-header-footer")).toBe(true);
    expect(LOCAL_HANDLED.has("goto-header")).toBe(true);
    expect(LOCAL_HANDLED.has("goto-footer")).toBe(true);
    expect(LOCAL_HANDLED.has("header-option")).toBe(true);
  });
});
