// https://nuxt.com/docs/api/configuration/nuxt-config
import { defineNuxtConfig } from "nuxt/config";

export default defineNuxtConfig({
  compatibilityDate: "2025-07-15",

  vite: {
    optimizeDeps: {
      include: [
        "@nuxt/ui > prosemirror-state",
        "@nuxt/ui > prosemirror-transform",
        "@nuxt/ui > prosemirror-model",
        "@nuxt/ui > prosemirror-view",
        "@nuxt/ui > prosemirror-gapcursor",
        "@vue/devtools-core",
        "@vue/devtools-kit",
        "docx",
        "image-meta",
        "xast-util-from-xml",
        "fflate",
        "undio",
        "@tiptap/extension-text-align",
        "@tiptap/extension-document",
        "@tiptap/extension-text",
        "@tiptap/extension-paragraph",
        "@tiptap/extension-heading",
        "@tiptap/extension-blockquote",
        "@tiptap/extension-horizontal-rule",
        "@tiptap/extension-code-block-lowlight",
        "@tiptap/extension-bullet-list",
        "@tiptap/extension-ordered-list",
        "@tiptap/extension-list-item",
        "@tiptap/extension-task-list",
        "@tiptap/extension-task-item",
        "@tiptap/extension-table",
        "@tiptap/extension-image",
        "@tiptap/extension-hard-break",
        "@tiptap/extension-details",
        "@tiptap/extension-emoji",
        "@tiptap/extension-mention",
        "@tiptap/extension-mathematics",
        "@tiptap/extension-bold",
        "@tiptap/extension-italic",
        "@tiptap/extension-underline",
        "@tiptap/extension-strike",
        "@tiptap/extension-code",
        "@tiptap/extension-link",
        "@tiptap/extension-highlight",
        "@tiptap/extension-subscript",
        "@tiptap/extension-superscript",
        "@tiptap/extension-text-style",
      ],
    },
  },

  css: ["~/assets/css/main.css"],

  devtools: { enabled: true },

  modules: ["@nuxt/ui"],

  ui: {
    mdc: true,
    content: true,
    experimental: {
      componentDetection: true,
    },
  },
});
