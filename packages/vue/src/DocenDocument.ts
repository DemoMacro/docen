import type { JSONContent, SectionPropertiesOptions, StylesOptions } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import type { DocenAddin, VisibilityMode } from "@docen/editor";
import type { PropType } from "vue";
import {
  computed,
  defineComponent,
  h,
  markRaw,
  onBeforeUnmount,
  ref,
  shallowRef,
  watch,
} from "vue";
// Side-effect: registers the <docen-document> custom element on first import.
import "@docen/editor";

type TaskPaneId = "navigation" | "properties";

type DocenEl = HTMLElement & {
  editor?: Editor;
  // Office.context.displayLanguage equivalent — read-only current locale.
  displayLanguage?: string;
  // DocenHost content surface — getJSON unwraps page nodes (external consumers
  // see flat doc > block+); setJSON wraps pages and preserves doc-level attrs
  // (styles/core/sectionProperties) that editor.commands.setContent drops.
  setJSON?: (json: JSONContent) => void;
  getJSON?: () => JSONContent;
  showTaskpane?: (id: TaskPaneId) => void;
  hideTaskpane?: (id: TaskPaneId) => void;
  getTaskpaneState?: (id: TaskPaneId) => boolean;
  setZoom?: (pct: number) => void;
  getZoom?: () => number;
  setShowMarks?: (on: boolean) => void;
  getShowMarks?: () => boolean;
  // Runtime add-in registration (commands/task-pane/ribbon/mini-toolbar). Add-ins
  // carrying functions can't cross the attribute boundary — register them here,
  // not through the `addins` prop (which only accepts JSON-serializable data).
  addAddin?: (addin: DocenAddin) => void;
};

/**
 * Host action events the adapter forwards as the **native `CustomEvent`** (not
 * just `e.detail`) so a consumer can call `preventDefault()` to take over the
 * host's built-in action. Without `preventDefault` the host proceeds with its
 * default — save/save-as → native file dialog, open → file picker, print →
 * browser print. The payload stays on `e.detail` (undefined for save/open/new/
 * print; `{ format }` for save-as).
 */
export type DocenSaveEvent = CustomEvent;
export type DocenSaveAsEvent = CustomEvent<{ format?: "docx" | "markdown" | "html" }>;
export type DocenOpenEvent = CustomEvent;
export type DocenNewEvent = CustomEvent;
export type DocenPrintEvent = CustomEvent;

/**
 * Vue 3 wrapper around the <docen-document> web component:
 *   - `v-model` for content (Tiptap JSON) — two-way: modelValue → host.setJSON
 *     (preserves doc.attrs.styles that editor.commands.setContent drops),
 *     docen:change → host.getJSON → update:modelValue (debounced; echo-broken
 *     by reference equality so the round-trip doesn't re-inject).
 *   - `v-slot="{ editor }"` exposes the underlying Tiptap editor (undefined
 *     until docen:ready) so a parent can render ad-hoc UI alongside the editor.
 *   - a template ref exposes `{ editor, getElement(), getJSON(), setJSON() }`.
 *
 * Why JSON not HTML: HTML round-trips (getHTML/setContent) drop DOCX-rich attrs
 * (styles/sectionProperties) and serialize O(n) per change; JSON carries the
 * full runtime model and injects via host.setJSON which keeps doc-level attrs.
 * The editor change → emit path is debounced (300 ms) so a large DOCX import
 * (many pagination-reflow change events) produces one getJSON, not one per
 * transaction.
 *
 * Only @docen/editor is a runtime dependency; @docen/docx types are imported
 * for prop typing only.
 */
export const DocenDocument = defineComponent({
  name: "DocenDocument",
  props: {
    /** Content (Tiptap JSON, page nodes unwrapped) — two-way via v-model.
     *  Seeded through host.setJSON on ready; emitted as host.getJSON()
     *  (debounced) on editor change. */
    modelValue: { type: Object as PropType<JSONContent>, default: undefined },
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
    /** Debounce (ms) for the modelValue emit on editor change. 0 = emit
     *  synchronously on each change (suitable for tests). @default 300 */
    debounce: { type: Number, default: 300 },
  },
  emits: {
    "update:modelValue": (_json: JSONContent) => true,
    // Host docen:* events forward the native CustomEvent — read `e.detail` for
    // the payload; call `e.preventDefault()` on the cancelable actions (save,
    // save-as, open, new, print) to suppress the host's built-in dialog.
    change: (_e: CustomEvent<{ dirty: true }>) => true,
    save: (_e: DocenSaveEvent) => true,
    "save-as": (_e: DocenSaveAsEvent) => true,
    open: (_e: DocenOpenEvent) => true,
    new: (_e: DocenNewEvent) => true,
    print: (_e: DocenPrintEvent) => true,
    "zoom-change": (_e: CustomEvent<{ zoom: number }>) => true,
    "taskpane-visibility-change": (
      _e: CustomEvent<{ id: TaskPaneId; visibilityMode: VisibilityMode }>,
    ) => true,
    "marks-change": (_e: CustomEvent<{ showMarks: boolean }>) => true,
    "lang-change": (_e: CustomEvent<{ lang: string }>) => true,
    "theme-change": (_e: CustomEvent<{ theme: string }>) => true,
  },
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
      // modelValue is NOT reflected to the `content` attribute — it is injected
      // through host.setJSON (see onReady + the modelValue watch below), which
      // preserves doc.attrs.styles and avoids re-serializing a large string
      // attribute on every change.
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

    // v-model JSON: external modelValue → host.setJSON. host.setJSON routes
    // through #loadDoc (fresh EditorState) so doc-level attrs (styles/core/
    // sectionProperties) survive — editor.commands.setContent would drop them.
    // lastEmitted breaks the round-trip echo: onChange emits getJSON() back as
    // modelValue, the watch sees the same reference and skips re-injecting.
    const lastEmitted = shallowRef<JSONContent | undefined>(undefined);
    let emitTimer: ReturnType<typeof setTimeout> | undefined;
    // True while applying a modelValue ourselves (onReady seed + watch setJSON)
    // — suppresses the change→emit round-trip so the parent's initial reference
    // isn't replaced and our own injection doesn't echo back. Cleared after the
    // pagination reflow that follows a load settles (a couple of frames).
    let suppressEmit = false;

    watch(
      () => props.modelValue,
      (json) => {
        if (json == null || json === lastEmitted.value) return;
        const host = el.value;
        if (host?.editor) host.setJSON?.(json);
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
      const host = el.value;
      editor.value = host?.editor;
      // Seed initial content via setJSON (not the content attribute) so a
      // modelValue carrying doc.attrs.styles is applied properly. Suppress the
      // follow-on change emit — the parent passed this content, no need to emit
      // it back and replace the parent's initial reference. rAF×2 covers the
      // pagination reflow that follows a load (dispatched on subsequent frames).
      if (props.modelValue != null) {
        suppressEmit = true;
        host?.setJSON?.(props.modelValue);
        requestAnimationFrame(() =>
          requestAnimationFrame(() => {
            suppressEmit = false;
          }),
        );
      }
    }

    function onChange(e: Event): void {
      emit("change", e as CustomEvent<{ dirty: true }>);
      if (suppressEmit) return; // our own injection — don't echo back to parent
      // Debounce getJSON: a DOCX import triggers many docen:change events as
      // pagination reflows; one getJSON per quiet window instead of one per
      // transaction (getJSON is O(n) on large docs).
      const emitJSON = (): void => {
        const host = el.value;
        const json = host?.getJSON?.();
        if (!json) return;
        const raw = markRaw(json) as JSONContent;
        lastEmitted.value = raw;
        emit("update:modelValue", raw);
      };
      clearTimeout(emitTimer);
      if (props.debounce <= 0)
        emitJSON(); // synchronous (tests)
      else emitTimer = setTimeout(emitJSON, props.debounce);
    }

    onBeforeUnmount(() => clearTimeout(emitTimer));

    // Forward the native CustomEvent (not e.detail) so a consumer can call
    // preventDefault() on the cancelable actions to take over the host default.
    const onSave = (e: Event): void => emit("save", e as DocenSaveEvent);
    const onSaveAs = (e: Event): void => emit("save-as", e as DocenSaveAsEvent);
    const onOpen = (e: Event): void => emit("open", e as DocenOpenEvent);
    const onNew = (e: Event): void => emit("new", e as DocenNewEvent);
    const onPrint = (e: Event): void => emit("print", e as DocenPrintEvent);
    const onZoomChange = (e: Event): void =>
      emit("zoom-change", e as CustomEvent<{ zoom: number }>);
    const onTaskpaneVisibilityChange = (e: Event): void =>
      emit(
        "taskpane-visibility-change",
        e as CustomEvent<{ id: TaskPaneId; visibilityMode: VisibilityMode }>,
      );
    const onMarksChange = (e: Event): void =>
      emit("marks-change", e as CustomEvent<{ showMarks: boolean }>);
    const onLangChange = (e: Event): void =>
      emit("lang-change", e as CustomEvent<{ lang: string }>);
    const onThemeChange = (e: Event): void =>
      emit("theme-change", e as CustomEvent<{ theme: string }>);

    expose({
      editor,
      getElement: (): HTMLElement | null => el.value,
      /** Current UI locale (Office.context.displayLanguage equivalent). */
      getDisplayLanguage: (): string | undefined => el.value?.displayLanguage,
      /** Read the document as Tiptap JSON (page nodes unwrapped). */
      getJSON: (): JSONContent | undefined => el.value?.getJSON?.(),
      /** Replace the document from Tiptap JSON (routes through #loadDoc, so
       *  doc.attrs.styles/core are preserved). */
      setJSON: (json: JSONContent): void => {
        el.value?.setJSON?.(json);
      },
      /** Register an add-in (commands/task-pane/ribbon/mini-toolbar). Add-ins
       *  carrying functions must register here — functions can't survive the
       *  `addins` attribute serialization. */
      addAddin: (addin: DocenAddin): void => {
        el.value?.addAddin?.(addin);
      },
    });

    return () => [
      // Default slot exposes the editor (undefined until docen:ready) for ad-hoc
      // external UI rendered alongside the editor; leave it empty to render the
      // editor alone.
      slots.default?.({ editor: editor.value }),
      h("docen-document", {
        ref: el,
        ...attrs.value,
        // Host events are colon-namespaced (docen:change — Office.js style).
        // Vue's onXxx identifier hyphenates to "docen-change" (no colon) and
        // silently misses the host's "docen:change", so the keys must be string
        // literals preserving the colon: parseName("onDocen:change") runs
        // hyphenate("Docen:change") → "docen:change" (colon kept, no uppercase
        // left to hyphenate). See Vue runtime-dom events.ts parseName.
        "onDocen:change": onChange,
        "onDocen:new": onNew,
        "onDocen:open": onOpen,
        "onDocen:print": onPrint,
        "onDocen:ready": onReady,
        "onDocen:save": onSave,
        "onDocen:save-as": onSaveAs,
        "onDocen:zoom-change": onZoomChange,
        "onDocen:taskpane-visibility-change": onTaskpaneVisibilityChange,
        "onDocen:marks-change": onMarksChange,
        "onDocen:lang-change": onLangChange,
        "onDocen:theme-change": onThemeChange,
      }),
    ];
  },
});
