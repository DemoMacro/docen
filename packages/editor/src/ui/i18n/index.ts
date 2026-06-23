import { en } from "./locales/en";
import { zhCN } from "./locales/zh-CN";
import { registerTranslation } from "./localize";

// Register the built-in locales on import. English is the fallback.
registerTranslation(en);
registerTranslation(zhCN);

export * from "./localize";
export { en, zhCN };
