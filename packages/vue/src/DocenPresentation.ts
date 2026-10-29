import type { DataType, PresentationOptions } from "@docen/pptx";
import type { PropType } from "vue";
import { defineComponent, h, markRaw, shallowRef, watch } from "vue";
// Side-effect: registers the presentation custom element on first import.
import "@docen/editor";

type PresentationEl = HTMLElement & {
  getContent?: () => PresentationOptions | null;
  setContent?: (content: PresentationOptions) => void;
  openPresentation?: (data: DataType) => Promise<void>;
  savePresentation?: () => Promise<Uint8Array>;
  closePresentation?: () => void;
};

/** Cancelable host-action events for presentation persistence owners. */
export type DocenPresentationOpenEvent = CustomEvent<void>;
export type DocenPresentationSaveEvent = CustomEvent<{ format: "pptx" }>;
export type DocenPresentationSaveAsEvent = CustomEvent<{ format: "pptx" }>;
export type DocenPresentationCloseEvent = CustomEvent<void>;

export const DocenPresentation = defineComponent({
  name: "DocenPresentation",
  props: {
    /** Presentation shown in the title bar and used as the download name. */
    filename: { type: String, default: undefined },
    /** Color scheme: `"light"` | `"dark"` | `""` (system). */
    theme: { type: String, default: undefined },
    /** UI locale (BCP-47 tag, e.g. `"zh-CN"`). */
    lang: { type: String, default: undefined },
    /** Canonical presentation JSON. Absent until the host opens a deck. */
    modelValue: { type: Object as PropType<PresentationOptions | undefined>, default: undefined },
    /** Delay before a host change is emitted to `modelValue` in milliseconds. */
    debounce: { type: Number, default: 300 },
  },
  emits: {
    "update:modelValue": (_content: PresentationOptions) => true,
    change: (_event: CustomEvent<{ dirty: true }>) => true,
    ready: () => true,
    open: (_event: DocenPresentationOpenEvent) => true,
    save: (_event: DocenPresentationSaveEvent) => true,
    "save-as": (_event: DocenPresentationSaveAsEvent) => true,
    close: (_event: DocenPresentationCloseEvent) => true,
  },
  setup(props, { emit, expose }) {
    const el = shallowRef<PresentationEl | null>(null);
    let emitTimer = 0;
    let suppressEmit = false;
    let seeded = false;
    let lastEmitted: PresentationOptions | undefined;

    const emitContent = (): void => {
      const content = el.value?.getContent?.();
      if (!content) return;
      lastEmitted = markRaw(content);
      emit("update:modelValue", lastEmitted);
    };

    const onReady = (): void => {
      if (seeded) {
        emit("ready");
        return;
      }
      seeded = true;
      suppressEmit = true;
      if (props.modelValue) {
        lastEmitted = props.modelValue;
        el.value?.setContent?.(structuredClone(props.modelValue));
      }
      requestAnimationFrame(() => {
        suppressEmit = false;
      });
      emit("ready");
    };

    const onChange = (event: CustomEvent<{ dirty: true }>): void => {
      if (suppressEmit) return;
      emit("change", event);
      clearTimeout(emitTimer);
      if (props.debounce <= 0) emitContent();
      else emitTimer = window.setTimeout(emitContent, props.debounce);
    };

    watch(
      () => props.modelValue,
      (content) => {
        if (!content || content === lastEmitted) return;
        lastEmitted = content;
        suppressEmit = true;
        el.value?.setContent?.(structuredClone(content));
        requestAnimationFrame(() => {
          suppressEmit = false;
        });
      },
    );

    expose({
      getElement: (): HTMLElement | null => el.value,
      getContent: (): PresentationOptions | undefined => el.value?.getContent?.() ?? undefined,
      setContent: (content: PresentationOptions): void => {
        el.value?.setContent?.(structuredClone(content));
      },
      openPresentation: async (data: DataType): Promise<void> => {
        await el.value?.openPresentation?.(data);
      },
      savePresentation: async (): Promise<Uint8Array> => {
        if (!el.value?.savePresentation) throw new Error("Presentation is not mounted");
        return el.value.savePresentation();
      },
      closePresentation: (): void => {
        el.value?.closePresentation?.();
      },
    });

    return () => {
      const attrs: Record<string, string> = {};
      if (props.filename != null) attrs.filename = props.filename;
      if (props.theme != null) attrs.theme = props.theme;
      if (props.lang != null) attrs.lang = props.lang;

      return h("docen-presentation", {
        ref: el,
        ...attrs,
        "onDocen:ready": onReady,
        "onDocen:change": onChange,
        "onDocen:open": (event: DocenPresentationOpenEvent): void => emit("open", event),
        "onDocen:save": (event: DocenPresentationSaveEvent): void => emit("save", event),
        "onDocen:save-as": (event: DocenPresentationSaveAsEvent): void => emit("save-as", event),
        "onDocen:close": (event: DocenPresentationCloseEvent): void => emit("close", event),
      });
    };
  },
});
