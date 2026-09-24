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
  RibbonMenuItem,
  RibbonSplit,
  RibbonTab,
} from "../ui";

const cmd = (name: string): string => `ppt.ribbon.cmd.${name}`;

const btn = (
  event: string,
  label: string,
  opts: { icon?: string; size?: RibbonControlSize; iconOnly?: boolean } = {},
): RibbonButton => ({ type: "button", event, label, ...opts });

const splitBtn = (
  event: string,
  label: string,
  items: RibbonMenuItem[],
  opts: { icon?: string; size?: RibbonControlSize; iconOnly?: boolean } = {},
): RibbonSplit => ({ type: "split", event, label, items, ...opts });

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

/** The Shapes gallery, grouped the way PowerPoint's drop-down is. Items carry
 *  the prstGeom token as their value; the host inserts a fresh 2" shape of
 *  that geometry. */
const SHAPE_GROUPS: readonly { header: string; shapes: readonly string[] }[] = [
  {
    header: "Rectangles",
    shapes: ["rect", "roundRect", "snip1Rect", "snip2SameRect", "snip2DiagRect", "round1Rect"],
  },
  {
    header: "Basic Shapes",
    shapes: [
      "triangle",
      "rtTriangle",
      "ellipse",
      "diamond",
      "parallelogram",
      "trapezoid",
      "pentagon",
      "hexagon",
      "heptagon",
      "octagon",
      "decagon",
      "dodecagon",
      "pie",
      "chord",
      "teardrop",
      "smileyFace",
      "sun",
      "moon",
      "cloud",
      "heart",
      "lightningBolt",
      "arc",
      "donut",
      "noSmoking",
      "blockArc",
      "plus",
      "plaque",
    ],
  },
  {
    header: "Block Arrows",
    shapes: [
      "rightArrow",
      "leftArrow",
      "upArrow",
      "downArrow",
      "leftRightArrow",
      "upDownArrow",
      "quadArrow",
      "bentArrow",
      "uturnArrow",
      "circularArrow",
    ],
  },
  {
    header: "Stars and Banners",
    shapes: [
      "star4",
      "star5",
      "star6",
      "star8",
      "star10",
      "star12",
      "star16",
      "star24",
      "star32",
      "irregularSeal1",
      "irregularSeal2",
      "ribbon",
      "ribbon2",
    ],
  },
  {
    header: "Callouts",
    shapes: [
      "wedgeRectCallout",
      "wedgeRoundRectCallout",
      "wedgeEllipseCallout",
      "cloudCallout",
      "borderCallout1",
      "borderCallout2",
    ],
  },
  {
    header: "Flowchart",
    shapes: ["flowChartProcess", "flowChartDecision", "flowChartTerminator", "flowChartConnector"],
  },
];

const shapesGallery = (): RibbonControlOrLayout => ({
  type: "gallery",
  event: "shapes",
  label: cmd("shapes"),
  icon: "shapes",
  size: "large",
  visibleCount: 8,
  items: SHAPE_GROUPS.flatMap(({ header, shapes }) => [
    { text: header, header: true },
    ...shapes.map((token) => ({
      text: token.replace(/([a-z])([A-Z])/g, "$1 $2").replace(/^./, (c) => c.toUpperCase()),
      value: token,
    })),
  ]),
});

/** The transitions gallery: the common effects plus none (clears the
 *  slide's transition). Values are the TransitionType tokens. */
const TRANSITIONS = [
  "none",
  "fade",
  "push",
  "wipe",
  "split",
  "blinds",
  "checker",
  "dissolve",
  "circle",
  "wheel",
  "zoom",
  "random",
] as const;

export function presentationRibbonTabs(): RibbonTab[] {
  return [
    {
      id: "home",
      label: "ppt.ribbon.tab.home",
      groups: [
        group("clipboard", [btn("paste", cmd("paste"), { icon: "paste", size: "large" })]),
        group("slides", [
          btn("new-slide", cmd("new-slide"), { icon: "new", size: "large" }),
          columnOf(
            btn("delete-slide", cmd("delete-slide"), { icon: "delete-slide", iconOnly: true }),
            btn("duplicate-slide", cmd("duplicate-slide"), {
              icon: "duplicate-slide",
              iconOnly: true,
            }),
            btn("layout", cmd("layout"), { icon: "page-size", iconOnly: true }),
            btn("reset", cmd("reset"), { icon: "sync", iconOnly: true }),
            btn("section", cmd("section"), { icon: "columns", iconOnly: true }),
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
              splitBtn(
                "line-spacing",
                cmd("line-spacing"),
                ["1.0", "1.5", "2.0", "2.5", "3.0"].map((value) => ({
                  text: `ppt.ribbon.lineSpacing.${value}`,
                  value,
                })),
                { icon: "line-spacing", iconOnly: true },
              ),
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
          shapesGallery(),
          columnOf(
            btn("bring-front", cmd("bring-front"), { icon: "bring-front", iconOnly: true }),
            btn("send-back", cmd("send-back"), { icon: "send-back", iconOnly: true }),
          ),
        ]),
        group("editing", [
          columnOf(
            rowOf(
              btn("find", cmd("find"), { icon: "search", iconOnly: true }),
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
          btn("insert-picture", cmd("insert-picture"), { icon: "insert-picture", size: "large" }),
          btn("online-picture", cmd("online-picture"), { icon: "online-picture", size: "large" }),
        ]),
        group("illustrations", [
          shapesGallery(),
          btn("smartart", cmd("smartart"), { icon: "smartart", size: "large" }),
          btn("chart", cmd("chart"), { icon: "chart", size: "large" }),
          btn("3d-model", cmd("3d-model"), { icon: "3d-model", size: "large" }),
        ]),
        group("links", [btn("hyperlink", cmd("hyperlink"), { icon: "hyperlink", size: "large" })]),
        group("text", [
          btn("text-box", cmd("text-box"), { icon: "text-box", size: "large" }),
          btn("wordart", cmd("wordart"), { icon: "wordart", size: "large" }),
          columnOf(
            btn("header-footer", cmd("header-footer"), { icon: "header", iconOnly: true }),
            btn("date-time", cmd("date-time"), { icon: "date-time", iconOnly: true }),
            btn("slide-number", cmd("slide-number"), { icon: "page-number", iconOnly: true }),
            btn("object", cmd("object"), { icon: "object", iconOnly: true }),
            btn("symbol", cmd("symbol"), { icon: "symbol", iconOnly: true }),
          ),
        ]),
        group("media", [
          btn("video", cmd("video"), { icon: "video", size: "large" }),
          btn("audio", cmd("audio"), { icon: "audio", size: "large" }),
        ]),
      ],
    },
    {
      id: "draw",
      label: "ppt.ribbon.tab.draw",
      groups: [
        group("tools", [
          btn("draw-select", cmd("draw-select"), { icon: "cursor" }),
          btn("draw-pen", cmd("draw-pen"), { icon: "action-pen" }),
          btn("draw-eraser", cmd("draw-eraser"), { icon: "eraser" }),
        ]),
      ],
    },
    {
      id: "design",
      label: "ppt.ribbon.tab.design",
      groups: [
        group("themes", [
          btn("themes", cmd("themes"), { icon: "theme", size: "large" }),
          columnOf(btn("variants", cmd("variants"), { icon: "page-color", iconOnly: true })),
        ]),
        group("customize", [
          {
            type: "color-picker",
            event: "format-background",
            label: cmd("format-background"),
            icon: "page-color",
            size: "large",
          },
          columnOf(btn("slide-size", cmd("slide-size"), { icon: "page-size", iconOnly: true })),
        ]),
      ],
    },
    {
      id: "transitions",
      label: "ppt.ribbon.tab.transitions",
      groups: [
        group("transition-to-slide", [
          splitBtn(
            "transition",
            cmd("transition"),
            TRANSITIONS.map((value) => ({
              text: `ppt.ribbon.transition.${value}`,
              value,
            })),
            { icon: "transition", size: "large" },
          ),
          columnOf(
            splitBtn(
              "effect-options",
              cmd("effect-options"),
              ["slow", "medium", "fast"].map((value) => ({
                text: `ppt.ribbon.effectOptions.${value}`,
                value,
              })),
            ),
            btn("apply-to-all", cmd("apply-to-all"), { icon: "repeat", iconOnly: true }),
          ),
        ]),
      ],
    },
    {
      id: "animations",
      label: "ppt.ribbon.tab.animations",
      groups: [
        group("animation", [
          btn("animate", cmd("animate"), { icon: "animate", size: "large" }),
          columnOf(
            btn("animation-pane", cmd("animation-pane"), {
              icon: "selection-pane",
              iconOnly: true,
            }),
          ),
        ]),
        group("advanced-animation", [
          btn("add-animation", cmd("add-animation"), { icon: "add-animation", size: "large" }),
        ]),
      ],
    },
    {
      id: "slide-show",
      label: "ppt.ribbon.tab.slide-show",
      groups: [
        group("start-slide-show", [
          btn("from-beginning", cmd("from-beginning"), { icon: "from-beginning", size: "large" }),
          columnOf(btn("from-current", cmd("from-current"), { icon: "from-current" })),
        ]),
        group("set-up", [
          btn("set-up-show", cmd("set-up-show"), { icon: "set-up-show", size: "large" }),
        ]),
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
          btn("new-comment", cmd("new-comment"), { icon: "comment-add", size: "large" }),
          columnOf(
            btn("show-comments", cmd("show-comments"), { icon: "comment-add", iconOnly: true }),
          ),
        ]),
      ],
    },
    {
      id: "view",
      label: "ppt.ribbon.tab.view",
      groups: [
        group("presentation-views", [
          btn("normal", cmd("normal"), { icon: "normal", size: "large" }),
          columnOf(
            btn("slide-sorter", cmd("slide-sorter"), { icon: "slide-sorter" }),
            btn("notes", cmd("notes"), { icon: "notes" }),
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
