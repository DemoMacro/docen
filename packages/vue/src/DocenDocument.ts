import type { SectionPropertiesOptions, StylesOptions } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import type { PropType } from "vue";
import { computed, defineComponent, h, ref, shallowRef, watch } from "vue";
// Side-effect: registers the <docen-document> custom element on first import.
import "@docen/editor";

type TaskPaneId = "navigation" | "properties";

type DocenEl = HTMLElement & {
  editor?: Editor;
  // Office.context.displayLanguage equivalent — read-only current locale.
  displayLanguage?: string;
  showTaskpane?: (id: TaskPaneId) => void;
  hideTaskpane?: (id: TaskPaneId) => void;
  getTaskpaneState?: (id: TaskPaneId) => boolean;
  setZoom?: (pct: number) => void;
  getZoom?: () => number;
  setShowMarks?: (on: boolean) => void;
  getShowMarks?: () => boolean;
};

/**
 * Vue 3 wrapper around the <docen-document> web component, shaped after Nuxt
 * UI's UEditor adapter:
 *   - `v-model` for content (HTML) — two-way: modelValue → editor.setContent,
 *     docen:change → editor.getHTML → update:modelValue (with an echo guard).
 *   - `v-slot="{ editor }"` exposes the underlying Tiptap editor (undefined
 *     until docen:ready) so a parent can render ad-hoc UI alongside the editor.
 *   - a template ref exposes `{ editor, getElement() }`.
 *
 * Only @docen/editor is a runtime dependency; @docen/docx types are imported
 * for prop typing only.
 */
export const DocenDocument = defineComponent({
  name: "DocenDocument",
  props: {
    /** Content (HTML) — two-way via v-model. Also seeds the editor on connect. */
    modelValue: { type: String, default: undefined },
    filename: { type: String, default: undefined },
    editable: { type: Boolean, default: undefined },
    spellcheck: { type: Boolean, default: undefined },
    user: { type: String, default: undefined },
    avatar: { type: String, default: undefined },
    sectionProperties: { type: Object as PropType<SectionPropertiesOptions>, default: undefined },
    styles: { type: Object as PropType<StylesOptions>, default: undefined },
    /** External add-ins (ribbon/task-pane data contributions). Functions can't
     *  cross the attribute boundary — register command/pane-render add-ins via
     *  the template ref's getElement().addAddin() instead. */
    addins: { type: Array as PropType<unknown[]>, default: undefined },
    /** Color scheme: `"light"` | `"dark"` | `""` (system). Reflected to the
     *  reactive `theme` attribute; the host applies it globally. */
    theme: { type: String, default: undefined },
    /** Initial visibility for the navigation (left) task pane. Reflected as
     *  `navigation-pane` on mount (once-attribute); runtime changes route through
     *  `showTaskpane`/`hideTaskpane`. Track the actual state via
     *  `taskpane-visibility-change`. */
    navigationPane: { type: Boolean, default: undefined },
    /** Initial visibility for the properties (right) task pane. Same lifecycle
     *  as `navigationPane`. */
    propertiesPane: { type: Boolean, default: undefined },
    /** Initial zoom percent. Reflected as `zoom` on mount; runtime changes route
     *  through `setZoom`. Track via `zoom-change`. */
    zoom: { type: Number, default: undefined },
    /** Initial editing-marks visibility. Reflected as `show-marks` on mount;
     *  runtime changes route through `setShowMarks`. Track via `marks-change`. */
    showMarks: { type: Boolean, default: undefined },
    /** UI locale (BCP-47 tag, e.g. "zh-CN" / "en" / "fr"). Reflected to the
     *  reactive `lang` attribute; the host forwards it to the workspace so
     *  every label re-resolves. Internal changes (status-bar cycle, Options
     *  OK) surface via `lang-change`. */
    lang: { type: String, default: undefined },
  },
  emits: [
    "update:modelValue",
    "change",
    "save",
    "save-as",
    "open",
    "new",
    "print",
    "zoom-change",
    "taskpane-visibility-change",
    "marks-change",
    "lang-change",
  ],
  setup(props, { emit, expose, slots }) {
    const el = ref<DocenEl | null>(null);
    /** The underlying Tiptap editor (undefined until docen:ready). Exposed on
     *  the template ref and the default slot scope. */
    const editor = shallowRef<Editor | undefined>(undefined);

    // Reflect props onto the web component's kebab-case attributes. Vue diffs
    // them, so initial mount sets them before connectedCallback reads them,
    // and runtime changes reach attributeChangedCallback (the observed ones).
    // Props left undefined emit no attribute, preserving the web component's
    // own default.
    const attrs = computed<Record<string, string>>(() => {
      const a: Record<string, string> = {};
      if (props.modelValue != null) a.content = props.modelValue;
      if (props.filename != null) a.filename = props.filename;
      if (props.editable != null) a.editable = props.editable ? "true" : "false";
      if (props.spellcheck != null) a.spellcheck = props.spellcheck ? "true" : "false";
      if (props.user != null) a.user = props.user;
      if (props.avatar != null) a.avatar = props.avatar;
      if (props.sectionProperties != null)
        a["section-properties"] = JSON.stringify(props.sectionProperties);
      if (props.styles != null) a.styles = JSON.stringify(props.styles);
      if (props.addins != null) a.addins = JSON.stringify(props.addins);
      if (props.theme != null) a.theme = props.theme;
      if (props.navigationPane != null)
        a["navigation-pane"] = props.navigationPane ? "true" : "false";
      if (props.propertiesPane != null)
        a["properties-pane"] = props.propertiesPane ? "true" : "false";
      if (props.zoom != null) a.zoom = String(props.zoom);
      if (props.showMarks != null) a["show-marks"] = props.showMarks ? "true" : "false";
      if (props.lang != null) a.lang = props.lang;
      return a;
    });

    // v-model: external modelValue → editor.setContent. The getHTML() equality
    // check breaks the feedback loop — onChange emits the editor's own HTML back
    // as modelValue, so the watch sees html === getHTML() and skips setContent.
    watch(
      () => props.modelValue,
      (html) => {
        if (html == null) return;
        const ed = editor.value;
        if (ed && ed.getHTML() !== html) ed.commands.setContent(html);
      },
    );

    // Once-attribute states (navigation-pane/properties-pane/zoom/show-marks)
    // are only seeded on connect — the host ignores runtime attribute writes —
    // so a prop change after mount must route through the host methods. `theme`
    // is a true reactive attribute and is handled by the `attrs` reflection alone.
    watch(
      () => props.navigationPane,
      (on) => {
        if (on == null) return;
        const host = el.value;
        if (!host) return;
        if (on) host.showTaskpane?.("navigation");
        else host.hideTaskpane?.("navigation");
      },
    );
    watch(
      () => props.propertiesPane,
      (on) => {
        if (on == null) return;
        const host = el.value;
        if (!host) return;
        if (on) host.showTaskpane?.("properties");
        else host.hideTaskpane?.("properties");
      },
    );
    watch(
      () => props.zoom,
      (pct) => {
        if (pct == null) return;
        el.value?.setZoom?.(pct);
      },
    );
    watch(
      () => props.showMarks,
      (on) => {
        if (on == null) return;
        el.value?.setShowMarks?.(on);
      },
    );

    function onReady(): void {
      editor.value = el.value?.editor;
    }

    function onChange(e: Event): void {
      emit("change", (e as CustomEvent).detail);
      const ed = editor.value;
      if (ed) emit("update:modelValue", ed.getHTML());
    }

    const onSave = (e: Event): void => emit("save", (e as CustomEvent).detail);
    const onSaveAs = (e: Event): void => emit("save-as", (e as CustomEvent).detail);
    const onOpen = (e: Event): void => emit("open", (e as CustomEvent).detail);
    const onNew = (e: Event): void => emit("new", (e as CustomEvent).detail);
    const onPrint = (e: Event): void => emit("print", (e as CustomEvent).detail);
    const onZoomChange = (e: Event): void => emit("zoom-change", (e as CustomEvent).detail);
    const onTaskpaneVisibilityChange = (e: Event): void =>
      emit("taskpane-visibility-change", (e as CustomEvent).detail);
    const onMarksChange = (e: Event): void => emit("marks-change", (e as CustomEvent).detail);
    const onLangChange = (e: Event): void => emit("lang-change", (e as CustomEvent).detail);

    expose({
      editor,
      getElement: (): HTMLElement | null => el.value,
      /** Current UI locale (Office.context.displayLanguage equivalent). */
      getDisplayLanguage: (): string | undefined => el.value?.displayLanguage,
    });

    return () => [
      // Default slot exposes the editor (undefined until docen:ready) for ad-hoc
      // external UI rendered alongside the editor; leave it empty to render the
      // editor alone.
      slots.default?.({ editor: editor.value }),
      h("docen-document", {
        ref: el,
        ...attrs.value,
        onDocenChange: onChange,
        onDocenNew: onNew,
        onDocenOpen: onOpen,
        onDocenPrint: onPrint,
        onDocenReady: onReady,
        onDocenSave: onSave,
        onDocenSaveAs: onSaveAs,
        onDocenZoomChange: onZoomChange,
        onDocenTaskpaneVisibilityChange: onTaskpaneVisibilityChange,
        onDocenMarksChange: onMarksChange,
        onDocenLangChange: onLangChange,
      }),
    ];
  },
});
