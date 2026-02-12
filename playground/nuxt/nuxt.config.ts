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
