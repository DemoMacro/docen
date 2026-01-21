// https://nuxt.com/docs/api/configuration/nuxt-config
import { defineNuxtConfig } from "nuxt/config";

export default defineNuxtConfig({
  compatibilityDate: "2025-07-15",

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
