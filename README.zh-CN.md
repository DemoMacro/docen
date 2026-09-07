> **⚠️ 警告：** 本项目尚未稳定，在达到 1.0.0 版本之前可能发生重大变化。我们强烈建议不要在生产环境中使用。

# Docen

[English](./README.md) | 简体中文

[![npm downloads](https://img.shields.io/npm/dm/docen)](https://www.npmjs.com/package/docen)
[![GitHub Stars](https://img.shields.io/github/stars/DemoMacro/docen)](https://github.com/DemoMacro/docen/stargazers)
![GitHub License](https://img.shields.io/github/license/DemoMacro/docen)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)

> 画布 DOCX 编辑器——在浏览器中以 MS Office 的排版保真度渲染并编辑文档，基于 TipTap/ProseMirror 与 LeaferJS 构建——并内置经统一 Tiptap JSON 模型的 headless Markdown ⇄ DOCX 转换。全量类型，无需服务器。

[在线演示](https://docen.office-open.com/) · [讨论区](https://github.com/DemoMacro/docen/discussions) · [报告问题](https://github.com/DemoMacro/docen/issues)

⭐ **如果 Docen 对你有用，一个 star 能帮更多开发者发现它。**

![Docen 编辑器](./assets/editor-demo.png)

## 包

| 包                                                          | 版本                                                    | 说明                                                                    |
| ----------------------------------------------------------- | ------------------------------------------------------- | ----------------------------------------------------------------------- |
| [docen](./packages/docen/README.md)                         | ![npm](https://img.shields.io/npm/v/docen)              | 全家桶——headless Markdown/DOCX 转换 + 完整 `<docen-document>` 编辑器    |
| [@docen/vue](./packages/vue/README.md)                      | ![npm](https://img.shields.io/npm/v/@docen/vue)         | Vue 3 适配器——`<DocenDocument>` 组件（v-model + v-slot 编辑器）         |
| [@docen/editor](./packages/editor/README.md)                | ![npm](https://img.shields.io/npm/v/@docen/editor)      | 组装层——Fluent UI 宿主 + docx 引擎，产出 `<docen-document>`             |
| [@docen/docx](./packages/docx/README.md)                    | ![npm](https://img.shields.io/npm/v/@docen/docx)        | DOCX 引擎——Tiptap schema + 转换器 + 排版投影，由 @office-open/docx 驱动 |
| [@docen/layout](./packages/layout/README.md)                | ![npm](https://img.shields.io/npm/v/@docen/layout)      | 分页引擎——测量 → 分页 LayoutDoc，Word 的堆叠规则                        |
| [@docen/pretext](./packages/pretext/README.md)              | ![npm](https://img.shields.io/npm/v/@docen/pretext)     | @chenglou/pretext 的 vendored fork——CJK 测量修正 + Word/CJK 排版修复    |
| [@docen/core](./packages/core/README.md)                    | ![npm](https://img.shields.io/npm/v/@docen/core)        | 场景绘制器——LayoutDoc → LeaferJS 树，供画布编辑器使用                   |
| [leafer-x-metafile](./packages/leafer-x-metafile/README.md) | ![npm](https://img.shields.io/npm/v/leafer-x-metafile)  | 零依赖 WMF/EMF+ 图元文件回放 → 中立 drawing 成员                        |
| [@docen/deduplicate](./packages/deduplicate/README.md)      | ![npm](https://img.shields.io/npm/v/@docen/deduplicate) | 文档比对（SimHash + Winnowing），供 compare 功能使用                    |

## 快速开始

### headless 转换（`docen`）

通过单一统一 API 在 Markdown、纯文本与 DOCX 之间无缝转换：

```bash
# 使用 pnpm 安装
$ pnpm add docen
```

```typescript
import { parseMarkdown, generateDOCX, parseDOCX, generateMarkdown } from "docen";

// Markdown → DOCX
const doc = parseMarkdown("# 标题\n\n你好，世界");
const docx = await generateDOCX(doc);

// DOCX → Markdown
const json = await parseDOCX(buffer);
const markdown = generateMarkdown(json);
```

编辑器支持把剪贴板中的带样式 HTML 作为**粘贴输入**——扩展的 `parseHTML` 规则将其转为文档 JSON。项目任何地方都不存在 HTML 生成。

> 💡 `docen` 包同时捆绑了完整引擎与编辑器——`import { createDocxEditor } from "docen/docx"` 或 `import { DocenDocument } from "docen/editor"`——一个依赖覆盖 headless 转换、引擎与 Web Component。

### DOCX 引擎（`@docen/docx`）

DOCX 引擎——Tiptap schema、转换器与排版投影——近无损往返转换：

```bash
$ pnpm add @docen/docx
```

```typescript
import { docxExtensions, parseDOCX, generateDOCX } from "@docen/docx";
import { Editor } from "@docen/docx/core";

// 无视图：编辑器就是编辑模型——渲染属于宿主。
const editor = new Editor({
  element: null,
  extensions: docxExtensions,
  content: await parseDOCX(buffer),
});
const output = await generateDOCX(editor.getJSON());
```

### 可视化编辑器（`@docen/editor`）

开箱即用的 Web Component 编辑器（`<docen-document>`），捆绑 Fluent UI 宿主、`@docen/docx` 引擎与 LeaferJS 画布舞台：

```bash
$ pnpm add @docen/editor
```

```html
<docen-document id="doc" filename="Welcome.docx"></docen-document>

<script type="module">
  import { registerComponents, applyTheme } from "@docen/editor";
  registerComponents();
  applyTheme("light");
</script>
```

### Vue（`@docen/vue`）

面向 Vue 3 的类型化 `<DocenDocument>` 组件——`v-model` 绑定内容、`v-slot="{ editor }"` 作用域插槽、模板 ref 暴露：

```bash
$ pnpm add @docen/vue
```

```vue
<script setup lang="ts">
import { ref } from "vue";
import type { JSONContent } from "@docen/docx";
import { DocenDocument } from "@docen/vue";
import { parseDOCX } from "@docen/docx";

// v-model 携带 Tiptap JSON；模板 ref 暴露 Tiptap editor
// 以及 getJSON/setJSON 方法对。
const content = ref<JSONContent>({ type: "doc", content: [{ type: "paragraph" }] });
const editorRef = ref();

async function open(file: File) {
  const json = await parseDOCX(await file.arrayBuffer());
  editorRef.value?.setJSON(json); // 保留 doc.attrs.styles
}
</script>

<template>
  <DocenDocument ref="editorRef" v-model="content" filename="Welcome.docx" editable />
</template>
```

## 开发

### 前置要求

- **Node.js** 18.x 或更高
- **pnpm** 9.x 或更高（推荐包管理器）
- **Git** 版本控制

### 开始

1. **克隆仓库**：

   ```bash
   git clone https://github.com/DemoMacro/docen.git
   cd docen
   ```

2. **安装依赖**：

   ```bash
   pnpm install
   ```

3. **构建全部包**：

   ```bash
   pnpm build
   ```

### 开发命令

```bash
pnpm build                       # 构建全部包
cd packages/<pkg> && pnpm build  # 构建单个包
vp check                         # Lint 与格式化
```

## 版本策略

本项目遵循[语义化版本](https://semver.org/)。主版本号为 `0`（pre-1.0）期间，破坏性 API 变更以 **minor** 版本（`0.x.0`）而非 patch 发布——公开 API 预计持续演进至 `1.0.0` 稳定版。若需要在 minor 更新之间保持稳定，下游项目请锁定精确版本。

## 参与贡献

欢迎贡献！完整贡献工作流、编码规范与 PR 清单见 [CONTRIBUTING.md](./CONTRIBUTING.md)。

## 支持与社区

- 📫 [报告问题](https://github.com/DemoMacro/docen/issues)
- 💬 [讨论区](https://github.com/DemoMacro/docen/discussions) — 提问、想法与作品展示

如果 Docen 对你有用，一个 [⭐ star](https://github.com/DemoMacro/docen/stargazers) 能帮更多开发者发现它。

## 许可证

本项目基于 MIT 许可证开源——详见 [LICENSE](./LICENSE)。

---

Built with ❤️ by [Demo Macro](https://www.demomacro.com/)
