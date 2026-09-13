// The presentation editor's built-in ribbon schema: the PowerPoint tab set
// (Home/Insert/Draw/Design/Transitions/Animations/Slide Show/Review/View).
// Commands mirror the PowerPoint surface; the host greys every control whose
// event has no handler yet (#applyRibbonGreying), so the skeleton stays honest
// as wiring lands batch by batch. External add-ins layer more tabs/groups on
// top via mergeRibbonSchema, exactly like the document editor.

import type {
  RibbonButton,
  RibbonControlOrLayout,
  RibbonControlSize,
  RibbonGroup,
  RibbonTab,
} from "../ui";

const cmd = (name: string): string => `ppt.ribbon.cmd.${name}`;

const btn = (
  event: string,
  label: string,
  opts: { icon?: string; size?: RibbonControlSize; iconOnly?: boolean } = {},
): RibbonButton => ({ type: "button", event, label, ...opts });

const group = (id: string, controls: readonly RibbonControlOrLayout[]): RibbonGroup => ({
  id,
  label: `ppt.ribbon.group.${id}`,
  controls,
});

const columnOf = (...controls: readonly RibbonControlOrLayout[]): RibbonControlOrLayout => ({
  type: "layout",
  layout: "column",
  controls,
});

const rowOf = (...controls: readonly RibbonControlOrLayout[]): RibbonControlOrLayout => ({
  type: "layout",
  layout: "row",
  controls,
});

export function presentationRibbonTabs(): RibbonTab[] {
  return [
    {
      id: "home",
      label: "ppt.ribbon.tab.home",
      groups: [
        group("clipboard", [btn("paste", cmd("paste"), { icon: "paste", size: "large" })]),
        group("slides", [
          columnOf(
            btn("new-slide", cmd("new-slide"), { icon: "new", size: "large" }),
            rowOf(
              btn("layout", cmd("layout"), { icon: "page-size", iconOnly: true }),
              btn("reset", cmd("reset"), { icon: "sync", iconOnly: true }),
            ),
            rowOf(btn("section", cmd("section"), { icon: "columns", iconOnly: true })),
          ),
        ]),
        group("font", [
          columnOf(
            rowOf(
              {
                type: "combobox",
                event: "font-face",
                label: cmd("font-face"),
                icon: "text-font",
                source: "local-fonts",
              },
              {
                type: "combobox",
                event: "font-size",
                label: cmd("font-size"),
                comboboxSize: "short",
              },
            ),
            rowOf(
              btn("bold", cmd("bold"), { icon: "bold", iconOnly: true }),
              btn("italic", cmd("italic"), { icon: "italic", iconOnly: true }),
              btn("underline", cmd("underline"), { icon: "underline", iconOnly: true }),
              btn("strike", cmd("strike"), { icon: "strike", iconOnly: true }),
              btn("line-spacing", cmd("line-spacing"), { icon: "line-spacing", iconOnly: true }),
            ),
          ),
        ]),
        group("paragraph", [
          columnOf(
            rowOf(
              btn("align-left", cmd("align-left"), { icon: "align-left", iconOnly: true }),
              btn("align-center", cmd("align-center"), { icon: "align-center", iconOnly: true }),
              btn("align-right", cmd("align-right"), { icon: "align-right", iconOnly: true }),
              btn("justify", cmd("justify"), { icon: "justify", iconOnly: true }),
            ),
            rowOf(
              btn("list", cmd("list"), { icon: "list", iconOnly: true }),
              btn("numbering", cmd("numbering"), { icon: "numbering", iconOnly: true }),
            ),
          ),
        ]),
        group("drawing", [
          columnOf(
            btn("shapes", cmd("shapes"), { size: "large" }),
            rowOf(btn("arrange", cmd("arrange"), { icon: "group-objects", iconOnly: true })),
          ),
        ]),
        group("editing", [
          columnOf(
            rowOf(
              btn("find", cmd("find"), { icon: "replace", iconOnly: true }),
              btn("replace", cmd("replace"), { icon: "replace", iconOnly: true }),
              btn("select", cmd("select"), { icon: "selection-pane", iconOnly: true }),
            ),
          ),
        ]),
      ],
    },
    {
      id: "insert",
      label: "ppt.ribbon.tab.insert",
      groups: [
        group("tables", [
          btn("insert-table", cmd("insert-table"), { icon: "table-add", size: "large" }),
        ]),
        group("images", [
          columnOf(
            btn("insert-picture", cmd("insert-picture"), { icon: "insert-picture", size: "large" }),
            rowOf(
              btn("online-picture", cmd("online-picture"), {
                icon: "online-picture",
                iconOnly: true,
              }),
            ),
          ),
        ]),
        group("illustrations", [
          columnOf(
            btn("shapes", cmd("shapes"), { size: "large" }),
            rowOf(
              btn("smartart", cmd("smartart"), { icon: "smartart", iconOnly: true }),
              btn("chart", cmd("chart"), { icon: "chart", iconOnly: true }),
              btn("3d-model", cmd("3d-model"), { icon: "3d-model", iconOnly: true }),
            ),
          ),
        ]),
        group("links", [btn("hyperlink", cmd("hyperlink"), { icon: "hyperlink", size: "large" })]),
        group("text", [
          columnOf(
            btn("text-box", cmd("text-box"), { icon: "text-box", size: "large" }),
            rowOf(
              btn("header-footer", cmd("header-footer"), { icon: "header", iconOnly: true }),
              btn("wordart", cmd("wordart"), { icon: "wordart", iconOnly: true }),
              btn("date-time", cmd("date-time"), { icon: "date-time", iconOnly: true }),
            ),
            rowOf(
              btn("slide-number", cmd("slide-number"), { icon: "page-number", iconOnly: true }),
              btn("object", cmd("object"), { icon: "object", iconOnly: true }),
              btn("symbol", cmd("symbol"), { icon: "symbol", iconOnly: true }),
            ),
          ),
        ]),
        group("media", [
          btn("video", cmd("video"), { size: "large" }),
          btn("audio", cmd("audio"), { size: "large" }),
        ]),
      ],
    },
    {
      id: "draw",
      label: "ppt.ribbon.tab.draw",
      groups: [
        group("tools", [
          btn("draw-select", cmd("draw-select")),
          btn("draw-pen", cmd("draw-pen"), { icon: "action-pen" }),
          btn("draw-eraser", cmd("draw-eraser")),
        ]),
      ],
    },
    {
      id: "design",
      label: "ppt.ribbon.tab.design",
      groups: [
        group("themes", [
          columnOf(
            btn("themes", cmd("themes"), { icon: "theme", size: "large" }),
            rowOf(btn("variants", cmd("variants"), { icon: "page-color", iconOnly: true })),
          ),
        ]),
        group("customize", [
          columnOf(
            btn("format-background", cmd("format-background"), {
              icon: "page-color",
              size: "large",
            }),
            rowOf(btn("slide-size", cmd("slide-size"), { icon: "page-size", iconOnly: true })),
          ),
        ]),
      ],
    },
    {
      id: "transitions",
      label: "ppt.ribbon.tab.transitions",
      groups: [
        group("transition-to-slide", [
          columnOf(
            btn("transition", cmd("transition"), { size: "large" }),
            rowOf(
              btn("effect-options", cmd("effect-options")),
              btn("apply-to-all", cmd("apply-to-all"), { icon: "repeat", iconOnly: true }),
            ),
          ),
        ]),
      ],
    },
    {
      id: "animations",
      label: "ppt.ribbon.tab.animations",
      groups: [
        group("animation", [
          columnOf(
            btn("animate", cmd("animate"), { size: "large" }),
            rowOf(
              btn("animation-pane", cmd("animation-pane"), {
                icon: "selection-pane",
                iconOnly: true,
              }),
            ),
          ),
        ]),
        group("advanced-animation", [
          btn("add-animation", cmd("add-animation"), { size: "large" }),
        ]),
      ],
    },
    {
      id: "slide-show",
      label: "ppt.ribbon.tab.slide-show",
      groups: [
        group("start-slide-show", [
          columnOf(
            btn("from-beginning", cmd("from-beginning"), { size: "large" }),
            rowOf(btn("from-current", cmd("from-current"))),
          ),
        ]),
        group("set-up", [btn("set-up-show", cmd("set-up-show"), { size: "large" })]),
      ],
    },
    {
      id: "review",
      label: "ppt.ribbon.tab.review",
      groups: [
        group("proofing", [
          btn("spell-check", cmd("spell-check"), { icon: "spell-check", size: "large" }),
        ]),
        group("comments", [
          columnOf(
            btn("new-comment", cmd("new-comment"), { icon: "comment-add", size: "large" }),
            rowOf(
              btn("show-comments", cmd("show-comments"), { icon: "comment-add", iconOnly: true }),
            ),
          ),
        ]),
      ],
    },
    {
      id: "view",
      label: "ppt.ribbon.tab.view",
      groups: [
        group("presentation-views", [
          columnOf(
            btn("normal", cmd("normal"), { size: "large" }),
            rowOf(btn("slide-sorter", cmd("slide-sorter")), btn("notes", cmd("notes"))),
          ),
        ]),
        group("show", [btn("gridlines", cmd("gridlines"), { icon: "gridlines", iconOnly: true })]),
        // Zoom is the one wired group: the stage scales with it.
        group("zoom", [
          columnOf(
            rowOf(
              btn("zoom-out", cmd("zoom-out"), { icon: "zoom-out", iconOnly: true }),
              btn("zoom-in", cmd("zoom-in"), { icon: "zoom-in", iconOnly: true }),
            ),
            rowOf(btn("zoom-100", cmd("zoom-100"))),
          ),
        ]),
      ],
    },
  ];
}
