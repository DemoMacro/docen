import type { SectionPropertiesOptions, StylesOptions } from "@docen/docx";
import type { Editor } from "@docen/docx/core";
import type { PropType } from "vue";
import { computed, defineComponent, h, ref, shallowRef, watch } from "vue";
// Side-effect: registers the <docen-document> custom element on first import.
import "@docen/editor";

type DocenEl = HTMLElement & { editor?: Editor };

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
    toolbar: { type: Boolean, default: undefined },
    header: { type: Boolean, default: undefined },
    statusBar: { type: Boolean, default: undefined },
    navigationPane: { type: String, default: undefined },
    propertiesPane: { type: String, default: undefined },
    tabs: { type: String, default: undefined },
    closable: { type: Boolean, default: undefined },
    user: { type: String, default: undefined },
    avatar: { type: String, default: undefined },
    sectionProperties: { type: Object as PropType<SectionPropertiesOptions>, default: undefined },
    styles: { type: Object as PropType<StylesOptions>, default: undefined },
  },
  emits: [
    "update:modelValue",
    "change",
    "save",
    "save-as",
    "open",
    "new",
    "print",
    "request-close",
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
      if (props.toolbar != null) a.toolbar = props.toolbar ? "true" : "false";
      if (props.header != null) a.header = props.header ? "true" : "false";
      if (props.statusBar != null) a["status-bar"] = props.statusBar ? "true" : "false";
      if (props.navigationPane != null) a["navigation-pane"] = props.navigationPane;
      if (props.propertiesPane != null) a["properties-pane"] = props.propertiesPane;
      if (props.tabs != null) a.tabs = props.tabs;
      if (props.closable != null) a.closable = props.closable ? "true" : "false";
      if (props.user != null) a.user = props.user;
      if (props.avatar != null) a.avatar = props.avatar;
      if (props.sectionProperties != null)
        a["section-properties"] = JSON.stringify(props.sectionProperties);
      if (props.styles != null) a.styles = JSON.stringify(props.styles);
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
    const onRequestClose = (e: Event): void => emit("request-close", (e as CustomEvent).detail);

    expose({
      editor,
      getElement: (): HTMLElement | null => el.value,
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
        onDocenRequestClose: onRequestClose,
        onDocenSave: onSave,
        onDocenSaveAs: onSaveAs,
      }),
    ];
  },
});
