// Component-internal i18n entry. The locale tables seed inline in ./localize
// (en/zh-CN) at module load — no registration call here, just the re-export of
// the public API. Business strings (ribbon/pane/status) register separately
// from the editor package.
export * from "./localize";
