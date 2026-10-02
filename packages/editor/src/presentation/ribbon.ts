// The presentation editor's built-in ribbon schema: the PowerPoint tab set
// (Home/Insert/Draw/Design/Transitions/Animations/Slide Show/Review/View).
// Commands mirror the PowerPoint surface; the host greys every control whose
// event has no handler yet (#applyRibbonGreying), so the skeleton stays honest
// as wiring lands batch by batch. External add-ins layer more tabs/groups on
// top via mergeRibbonSchema, exactly like the document editor.

import type { SlideAnimation, TransitionType } from "@docen/pptx";

import { ensureShapePreviewIcons } from "../document/ribbon";
import { registerIcon } from "../ui";
import type {
  RibbonButton,
  RibbonControlOrLayout,
  RibbonControlSize,
  RibbonGroup,
  RibbonInput,
  RibbonMenuItem,
  RibbonSplit,
  RibbonTab,
} from "../ui";

const cmd = (name: string): string => `ppt.ribbon.cmd.${name}`;

const btn = (
  event: string,
  label: string,
  opts: { value?: string; icon?: string; size?: RibbonControlSize; iconOnly?: boolean } = {},
): RibbonButton => ({ type: "button", event, label, ...opts });

const splitBtn = (
  event: string,
  label: string,
  items: RibbonMenuItem[],
  opts: { icon?: string; size?: RibbonControlSize; iconOnly?: boolean } = {},
): RibbonSplit => ({ type: "split", event, label, items, ...opts });

const input = (event: string, value: string, opts: { label?: string } = {}): RibbonInput => ({
  type: "input",
  event,
  value,
  ...opts,
});

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

const shapesGallery = (): RibbonControlOrLayout => {
  ensureShapePreviewIcons(SHAPE_GROUPS.flatMap(({ shapes }) => shapes));
  return {
    type: "gallery",
    event: "shapes",
    label: cmd("shapes"),
    icon: "shapes",
    size: "large",
    visibleCount: 6,
    items: SHAPE_GROUPS.flatMap(({ header, shapes }) => [
      { text: header, header: true },
      ...shapes.map((token) => ({
        text: token.replace(/([a-z])([A-Z])/g, "$1 $2").replace(/^./, (c) => c.toUpperCase()),
        icon: `shape-${token}`,
        value: token,
      })),
    ]),
  };
};

/** Insert's Shapes picker: a large split whose drop-down lists every
 *  category with previewed shape cards (PowerPoint keeps the inline gallery
 *  on Home and a picker here — an inline strip this size overflows the
 *  Illustrations group and pushes its neighbors off the tab). */
const shapesPicker = (): RibbonControlOrLayout => {
  const items: RibbonMenuItem[] = SHAPE_GROUPS.flatMap(({ header, shapes }) => [
    { text: header, header: true },
    ...shapes.map((token) => ({
      text: token.replace(/([a-z])([A-Z])/g, "$1 $2").replace(/^./, (c) => c.toUpperCase()),
      icon: `shape-${token}`,
      value: token,
    })),
  ]);
  ensureShapePreviewIcons(SHAPE_GROUPS.flatMap(({ shapes }) => shapes));
  return splitBtn("shapes", cmd("shapes"), items, { icon: "shapes", size: "large" });
};

/** The transition tokens the gallery offers — compiler-checked against the
 *  office-open TransitionType union so a renamed token breaks the build. */
const TRANSITION_TYPES = [
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
] as const satisfies readonly TransitionType[];

/** The transitions gallery: the typed tokens plus none (clears the slide's
 *  transition). */
const TRANSITIONS = ["none", ...TRANSITION_TYPES] as const;

/** The gallery's transition tokens, as the runtime set #setTransition
 *  validates the ribbon value against before it touches the model. */
export const TRANSITION_PRESETS: ReadonlySet<TransitionType> = new Set(TRANSITION_TYPES);

/** The Animate / Add Animation splits' presets: the entrance effects the
 *  show playback tweens, plus none (clears the shape's entry). Values are
 *  the AnimationType tokens. */
const ANIMATION_TYPES = [
  "appear",
  "fade",
  "fly",
  "zoom",
] as const satisfies readonly SlideAnimation["type"][];

const ANIMATIONS = ["none", ...ANIMATION_TYPES] as const;

/** The entrance presets, as the runtime set #applyAnimation validates
 *  against — the same tokens the show playback tweens. */
export const ANIMATION_PRESETS: ReadonlySet<SlideAnimation["type"]> = new Set(ANIMATION_TYPES);

const animationItems = (): RibbonMenuItem[] =>
  ANIMATIONS.map((value) => ({
    text: `ppt.ribbon.animate.${value}`,
    icon: `animation-${value}`,
    value,
  }));

/** The style flags PowerPoint's Table Design checkbox grid exposes. */
const TABLE_LOOK_FLAGS = [
  "firstRow",
  "lastRow",
  "firstCol",
  "lastCol",
  "bandRow",
  "bandCol",
] as const;

/** Built-in DrawingML style GUIDs the projector already resolves; values are
 *  passed straight to the official `tableStyleId` field. */
const tableStyle = (
  key: string,
  value: string,
): RibbonMenuItem & { value: string; icon: string } => ({
  value,
  text: `ppt.ribbon.tableStyle.${key}`,
  icon: value,
});

const TABLE_STYLES = [
  tableStyle("none", "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"),
  tableStyle("grid", "{5940675A-B579-460E-94D1-54222C63F5DA}"),
  tableStyle("light1", "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"),
  tableStyle("medium2", "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"),
  tableStyle("dark2", "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}"),
];

/** PowerPoint-style table style thumbnails: the menu previews the fill/grid
 *  contract behind each official GUID instead of leaving text-only rows. */
const tableStylePreview = (background: string, body: string): string =>
  `<svg width="24" height="24" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><rect x="2.5" y="4.5" width="19" height="15" rx="1.5" fill="${background}" stroke="#595959"/>${body}</svg>`;

const TABLE_STYLE_PREVIEWS: Record<string, string> = {
  none: tableStylePreview(
    "#ffffff",
    '<path d="M8.5 5v14M15.5 5v14M3 10h18M3 15h18" stroke="#bfbfbf" stroke-dasharray="2 2"/>',
  ),
  grid: tableStylePreview(
    "#ffffff",
    '<path d="M8.5 5v14M15.5 5v14M3 10h18M3 15h18" stroke="#4472c4"/>',
  ),
  light1: tableStylePreview(
    "#ffffff",
    '<rect x="3" y="5" width="18" height="5" fill="#d9e2f3"/><path d="M8.5 5v14M15.5 5v14M3 10h18M3 15h18" stroke="#8faadc"/>',
  ),
  medium2: tableStylePreview(
    "#ffffff",
    '<rect x="3" y="5" width="18" height="5" fill="#4472c4"/><rect x="3" y="10" width="18" height="5" fill="#d9e2f3"/><path d="M8.5 5v14M15.5 5v14M3 10h18M3 15h18" stroke="#ffffff"/>',
  ),
  dark2: tableStylePreview(
    "#1f4e79",
    '<path d="M8.5 5v14M15.5 5v14M3 10h18M3 15h18" stroke="#ffffff"/>',
  ),
};

for (const style of TABLE_STYLES)
  registerIcon(style.value, TABLE_STYLE_PREVIEWS[style.text.split(".").at(-1)!]!);

const TEXT_DIRECTIONS = [
  { text: "ppt.ribbon.textDirection.horizontal", value: "horizontal" },
  { text: "ppt.ribbon.textDirection.rotate90", value: "vertical" },
  { text: "ppt.ribbon.textDirection.rotate270", value: "vertical270" },
  { text: "ppt.ribbon.textDirection.stacked", value: "wordArtVertical" },
] satisfies RibbonMenuItem[];

const CELL_MARGINS = [
  { text: "ppt.ribbon.cellMargin.normal", value: "normal" },
  { text: "ppt.ribbon.cellMargin.none", value: "none" },
  { text: "ppt.ribbon.cellMargin.narrow", value: "narrow" },
  { text: "ppt.ribbon.cellMargin.wide", value: "wide" },
] satisfies RibbonMenuItem[];

export const TABLE_STYLE_IDS: ReadonlySet<string> = new Set(
  TABLE_STYLES.map((style) => style.value),
);

const tableBorderItems = (): RibbonMenuItem[] =>
  ["all", "outside", "none", "top", "bottom", "left", "right"].map((value) => ({
    text: `ppt.ribbon.tableBorder.${value}`,
    value,
  }));

/** PowerPoint's contextual Table Design tab. The host appends/removes it as
 *  the selection enters/leaves a table, exactly like the product tab set. */
export function tableDesignTab(): RibbonTab {
  return {
    id: "ppt-table-design",
    label: "ppt.ribbon.tab.table-design",
    contextual: true,
    groups: [
      group("table-style-options", [
        {
          type: "layout",
          layout: "grid",
          controls: TABLE_LOOK_FLAGS.map((flag) => ({
            type: "checkbox",
            event: "toggle-table-look",
            value: flag,
            label: `ppt.ribbon.tableLook.${flag}`,
          })),
        },
      ]),
      group("table-styles", [
        splitBtn("table-style", cmd("table-style"), [...TABLE_STYLES], {
          icon: TABLE_STYLES[3]!.value,
          size: "large",
        }),
      ]),
      group("table-shading", [
        {
          type: "color-picker",
          event: "cell-shading",
          label: cmd("cell-shading"),
          icon: "shading",
          defaultColor: "FFFF00",
          size: "large",
        },
      ]),
      group("table-borders", [
        splitBtn("table-borders", cmd("table-borders"), tableBorderItems(), {
          icon: "border",
          size: "large",
        }),
      ]),
    ],
  };
}

/** PowerPoint's contextual Table Layout tab: the structural commands that
 *  operate on the live cell selection. */
export function tableLayoutTab(): RibbonTab {
  return {
    id: "ppt-table-layout",
    label: "ppt.ribbon.tab.table-layout",
    contextual: true,
    groups: [
      group("table", [
        btn("select", cmd("select"), { icon: "selection-pane", size: "large" }),
        btn("gridlines", cmd("gridlines"), { icon: "gridlines", size: "large" }),
        splitBtn(
          "delete-table",
          cmd("delete-table"),
          [
            { text: "ppt.ribbon.tableDelete.columns", value: "columns" },
            { text: "ppt.ribbon.tableDelete.rows", value: "rows" },
            { text: "ppt.ribbon.tableDelete.table", value: "table" },
          ],
          { icon: "table-delete", size: "large" },
        ),
      ]),
      group("rows-columns", [
        btn("insert-above", cmd("insert-above"), { icon: "table-stack-above", size: "large" }),
        columnOf(
          btn("insert-below", cmd("insert-below"), { icon: "table-stack-below" }),
          btn("insert-left", cmd("insert-left"), { icon: "table-stack-left" }),
          btn("insert-right", cmd("insert-right"), { icon: "table-stack-right" }),
        ),
      ]),
      group("merge", [
        btn("merge-cells", cmd("merge-cells"), { icon: "merge-cells", size: "large" }),
        btn("split-cells", cmd("split-cells"), { icon: "split-cells", size: "large" }),
      ]),
      group("cell-size", [
        rowOf(
          columnOf(
            input("cell-height", "", { label: cmd("cell-height") }),
            input("cell-width", "", { label: cmd("cell-width") }),
          ),
          columnOf(
            btn("distribute-rows", cmd("distribute-rows"), { icon: "distribute-rows" }),
            btn("distribute-columns", cmd("distribute-columns"), { icon: "distribute-columns" }),
          ),
        ),
      ]),
      group("cell-alignment", [
        rowOf(
          columnOf(
            btn("table-align", cmd("table-align"), {
              value: "left",
              icon: "align-left",
              iconOnly: true,
            }),
            btn("table-align", cmd("table-align"), {
              value: "center",
              icon: "align-center",
              iconOnly: true,
            }),
            btn("table-align", cmd("table-align"), {
              value: "right",
              icon: "align-right",
              iconOnly: true,
            }),
          ),
          columnOf(
            btn("table-align", cmd("table-align"), {
              value: "top",
              icon: "align-top",
              iconOnly: true,
            }),
            btn("table-align", cmd("table-align"), {
              value: "middle",
              icon: "align-middle",
              iconOnly: true,
            }),
            btn("table-align", cmd("table-align"), {
              value: "bottom",
              icon: "align-bottom",
              iconOnly: true,
            }),
          ),
        ),
        splitBtn("text-direction", cmd("text-direction"), TEXT_DIRECTIONS, {
          icon: "text-direction",
          size: "large",
        }),
        splitBtn("cell-margins", cmd("cell-margins"), CELL_MARGINS, {
          icon: "cell-margin",
          size: "large",
        }),
      ]),
      group("table-arrange", [
        {
          type: "layout",
          layout: "grid",
          columns: 2,
          controls: [
            splitBtn(
              "bring-forward",
              cmd("bring-forward"),
              [
                { text: "ppt.ribbon.cmd.bring-front", value: "front" },
                { text: "ppt.ribbon.cmd.bring-forward", value: "forward" },
              ],
              { icon: "bring-forward", size: "large" },
            ),
            {
              type: "menu",
              event: "align-objects",
              label: cmd("align-objects"),
              icon: "align-left",
              size: "large",
              items: [
                { text: "ppt.ribbon.cmd.align-left", value: "left" },
                { text: "ppt.ribbon.cmd.align-center", value: "center" },
                { text: "ppt.ribbon.cmd.align-right", value: "right" },
                { text: "ppt.ribbon.cmd.align-top", value: "top" },
                { text: "ppt.ribbon.cmd.align-middle", value: "middle" },
                { text: "ppt.ribbon.cmd.align-bottom", value: "bottom" },
                { text: "ppt.ribbon.cmd.distribute-horizontal", value: "horizontal" },
                { text: "ppt.ribbon.cmd.distribute-vertical", value: "vertical" },
              ],
            },
            splitBtn(
              "send-backward",
              cmd("send-backward"),
              [
                { text: "ppt.ribbon.cmd.send-back", value: "back" },
                { text: "ppt.ribbon.cmd.send-backward", value: "backward" },
              ],
              { icon: "send-backward", size: "large" },
            ),
            {
              type: "menu",
              event: "drawing-group",
              label: cmd("drawing-group"),
              icon: "group-objects",
              size: "large",
              items: [{ text: "ppt.ribbon.cmd.group", value: "group" }],
            },
            btn("selection-pane", cmd("selection-pane"), {
              icon: "selection-pane",
              size: "large",
            }),
            {
              type: "menu",
              event: "rotate",
              label: cmd("rotate"),
              icon: "rotate",
              size: "large",
              items: [
                { text: "ppt.ribbon.rotate.right", value: "90" },
                { text: "ppt.ribbon.rotate.left", value: "-90" },
                { text: "ppt.ribbon.rotate.flip-horizontal", value: "flip-horizontal" },
                { text: "ppt.ribbon.rotate.flip-vertical", value: "flip-vertical" },
              ],
            },
          ],
        },
      ]),
    ],
  };
}

export function presentationRibbonTabs(): RibbonTab[] {
  return [
    {
      id: "home",
      label: "ppt.ribbon.tab.home",
      groups: [
        group("clipboard", [btn("paste", cmd("paste"), { icon: "paste", size: "large" })]),
        group("slides", [
          btn("new-slide", cmd("new-slide"), { icon: "new", size: "large" }),
          btn("delete-slide", cmd("delete-slide"), { icon: "delete-slide", size: "large" }),
          btn("duplicate-slide", cmd("duplicate-slide"), {
            icon: "duplicate-slide",
            size: "large",
          }),
          btn("layout", cmd("layout"), { icon: "page-size", size: "large" }),
          btn("reset", cmd("reset"), { icon: "sync", size: "large" }),
          btn("section", cmd("section"), { icon: "columns", size: "large" }),
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
          btn("bring-front", cmd("bring-front"), { icon: "bring-front", size: "large" }),
          btn("send-back", cmd("send-back"), { icon: "send-back", size: "large" }),
        ]),
        group("editing", [
          btn("find", cmd("find"), { icon: "search", size: "large" }),
          btn("replace", cmd("replace"), { icon: "replace", size: "large" }),
          btn("select", cmd("select"), { icon: "selection-pane", size: "large" }),
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
          shapesPicker(),
          btn("smartart", cmd("smartart"), { icon: "smartart", size: "large" }),
          btn("chart", cmd("chart"), { icon: "chart", size: "large" }),
          btn("3d-model", cmd("3d-model"), { icon: "3d-model", size: "large" }),
        ]),
        group("links", [btn("hyperlink", cmd("hyperlink"), { icon: "hyperlink", size: "large" })]),
        group("text", [
          btn("text-box", cmd("text-box"), { icon: "text-box", size: "large" }),
          btn("header-footer", cmd("header-footer"), { icon: "header", size: "large" }),
          btn("wordart", cmd("wordart"), { icon: "wordart", size: "large" }),
          btn("date-time", cmd("date-time"), { icon: "date-time", size: "large" }),
          btn("slide-number", cmd("slide-number"), { icon: "page-number", size: "large" }),
          btn("object", cmd("object"), { icon: "object", size: "large" }),
          btn("symbol", cmd("symbol"), { icon: "symbol", size: "large" }),
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
          btn("draw-select", cmd("draw-select"), { icon: "cursor", size: "large" }),
          btn("draw-pen", cmd("draw-pen"), { icon: "action-pen", size: "large" }),
          btn("draw-eraser", cmd("draw-eraser"), { icon: "eraser", size: "large" }),
        ]),
      ],
    },
    {
      id: "design",
      label: "ppt.ribbon.tab.design",
      groups: [
        group("themes", [
          btn("themes", cmd("themes"), { icon: "theme", size: "large" }),
          btn("variants", cmd("variants"), { icon: "page-color", size: "large" }),
        ]),
        group("customize", [
          {
            type: "color-picker",
            event: "format-background",
            label: cmd("format-background"),
            icon: "page-color",
            size: "large",
          },
          btn("slide-size", cmd("slide-size"), { icon: "page-size", size: "large" }),
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
              icon: `transition-${value}`,
              value,
            })),
            { icon: "transition", size: "large" },
          ),
          splitBtn(
            "effect-options",
            cmd("effect-options"),
            ["slow", "medium", "fast"].map((value) => ({
              text: `ppt.ribbon.effectOptions.${value}`,
              value,
            })),
            { icon: "effect-options", size: "large" },
          ),
          btn("apply-to-all", cmd("apply-to-all"), { icon: "repeat", size: "large" }),
        ]),
      ],
    },
    {
      id: "animations",
      label: "ppt.ribbon.tab.animations",
      groups: [
        group("animation", [
          splitBtn("animate", cmd("animate"), animationItems(), {
            icon: "animate",
            size: "large",
          }),
          btn("animation-pane", cmd("animation-pane"), {
            icon: "selection-pane",
            size: "large",
          }),
        ]),
        group("advanced-animation", [
          splitBtn("add-animation", cmd("add-animation"), animationItems(), {
            icon: "add-animation",
            size: "large",
          }),
        ]),
      ],
    },
    {
      id: "slide-show",
      label: "ppt.ribbon.tab.slide-show",
      groups: [
        group("start-slide-show", [
          btn("from-beginning", cmd("from-beginning"), { icon: "from-beginning", size: "large" }),
          btn("from-current", cmd("from-current"), { icon: "from-current", size: "large" }),
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
          btn("show-comments", cmd("show-comments"), { icon: "comment-add", size: "large" }),
        ]),
      ],
    },
    {
      id: "view",
      label: "ppt.ribbon.tab.view",
      groups: [
        group("presentation-views", [
          btn("normal", cmd("normal"), { icon: "normal", size: "large" }),
          btn("slide-sorter", cmd("slide-sorter"), { icon: "slide-sorter", size: "large" }),
          btn("notes", cmd("notes"), { icon: "notes", size: "large" }),
        ]),
        group("show", [btn("gridlines", cmd("gridlines"), { icon: "gridlines", size: "large" })]),
        // Zoom is the one wired group: the stage scales with it.
        group("zoom", [
          btn("zoom-out", cmd("zoom-out"), { icon: "zoom-out", size: "large" }),
          btn("zoom-in", cmd("zoom-in"), { icon: "zoom-in", size: "large" }),
          btn("zoom-100", cmd("zoom-100"), { icon: "zoom-in", size: "large" }),
        ]),
      ],
    },
  ];
}
