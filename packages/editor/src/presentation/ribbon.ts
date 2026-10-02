// The presentation editor's built-in ribbon schema: the PowerPoint tab set
// (Home/Insert/Draw/Design/Transitions/Animations/Slide Show/Review/View).
// Commands mirror the PowerPoint surface; the host greys every control whose
// event has no handler yet (#applyRibbonGreying), so the skeleton stays honest
// as wiring lands batch by batch. External add-ins layer more tabs/groups on
// top via mergeRibbonSchema, exactly like the document editor.

import type { SlideAnimation, TransitionType } from "@docen/pptx";

import { ensureShapePreviewIcons } from "../document/ribbon";
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

const input = (event: string, value: string): RibbonInput => ({
  type: "input",
  event,
  value,
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
  ANIMATIONS.map((value) => ({ text: `ppt.ribbon.animate.${value}`, value }));

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
const TABLE_STYLES = [
  { value: "{2D5ABB26-0587-4C30-8999-92F81FD0307C}", text: "ppt.ribbon.tableStyle.none" },
  { value: "{5940675A-B579-460E-94D1-54222C63F5DA}", text: "ppt.ribbon.tableStyle.grid" },
  { value: "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}", text: "ppt.ribbon.tableStyle.light1" },
  { value: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", text: "ppt.ribbon.tableStyle.medium2" },
  { value: "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}", text: "ppt.ribbon.tableStyle.dark2" },
] satisfies RibbonMenuItem[];

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
      group("table-styles", [splitBtn("table-style", cmd("table-style"), [...TABLE_STYLES])]),
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
        splitBtn("table-borders", cmd("table-borders"), tableBorderItems(), { icon: "border" }),
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
      group("rows-columns", [
        columnOf(
          rowOf(
            btn("insert-above", cmd("insert-above"), { icon: "row-insert-above", iconOnly: true }),
            btn("insert-below", cmd("insert-below"), { icon: "row-insert-below", iconOnly: true }),
          ),
          rowOf(
            btn("insert-left", cmd("insert-left"), { icon: "column-insert-left", iconOnly: true }),
            btn("insert-right", cmd("insert-right"), {
              icon: "column-insert-right",
              iconOnly: true,
            }),
          ),
        ),
      ]),
      group("delete", [
        splitBtn(
          "delete-table",
          cmd("delete-table"),
          [
            { text: "ppt.ribbon.tableDelete.rows", value: "rows" },
            { text: "ppt.ribbon.tableDelete.columns", value: "columns" },
            { text: "ppt.ribbon.tableDelete.table", value: "table" },
          ],
          { icon: "delete-slide" },
        ),
      ]),
      group("merge", [
        btn("merge-cells", cmd("merge-cells"), { icon: "merge-cells" }),
        btn("split-cells", cmd("split-cells"), { icon: "split-cells" }),
      ]),
      group("cell-size", [
        input("cell-height", ""),
        input("cell-width", ""),
        rowOf(
          btn("distribute-rows", cmd("distribute-rows"), {
            icon: "distribute-rows",
            iconOnly: true,
          }),
          btn("distribute-columns", cmd("distribute-columns"), {
            icon: "distribute-columns",
            iconOnly: true,
          }),
        ),
      ]),
      group("cell-alignment", [
        rowOf(
          btn("table-align", cmd("align-left"), {
            value: "left",
            icon: "align-left",
            iconOnly: true,
          }),
          btn("table-align", cmd("align-center"), {
            value: "center",
            icon: "align-center",
            iconOnly: true,
          }),
          btn("table-align", cmd("align-right"), {
            value: "right",
            icon: "align-right",
            iconOnly: true,
          }),
        ),
        rowOf(
          btn("table-align", cmd("align-top"), { value: "top", icon: "align-top", iconOnly: true }),
          btn("table-align", cmd("align-middle"), {
            value: "middle",
            icon: "align-middle",
            iconOnly: true,
          }),
          btn("table-align", cmd("align-bottom"), {
            value: "bottom",
            icon: "align-bottom",
            iconOnly: true,
          }),
        ),
        rowOf(
          splitBtn("text-direction", cmd("text-direction"), TEXT_DIRECTIONS, {
            icon: "text-direction",
            iconOnly: true,
          }),
          splitBtn("cell-margins", cmd("cell-margins"), CELL_MARGINS, {
            icon: "cell-margin",
            iconOnly: true,
          }),
        ),
      ]),
      group("table-arrange", [
        rowOf(
          btn("bring-front", cmd("bring-front"), { icon: "bring-front", iconOnly: true }),
          btn("send-back", cmd("send-back"), { icon: "send-back", iconOnly: true }),
        ),
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
          columnOf(
            btn("delete-slide", cmd("delete-slide"), { icon: "delete-slide", iconOnly: true }),
            btn("duplicate-slide", cmd("duplicate-slide"), {
              icon: "duplicate-slide",
              iconOnly: true,
            }),
          ),
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
          columnOf(
            btn("date-time", cmd("date-time"), { icon: "date-time", iconOnly: true }),
            btn("slide-number", cmd("slide-number"), { icon: "page-number", iconOnly: true }),
          ),
          columnOf(
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
            { size: "large" },
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
