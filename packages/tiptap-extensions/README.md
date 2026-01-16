# @docen/tiptap-extensions

![npm version](https://img.shields.io/npm/v/@docen/tiptap-extensions)
![npm downloads](https://img.shields.io/npm/dw/@docen/tiptap-extensions)
![npm license](https://img.shields.io/npm/l/@docen/tiptap-extensions)

> Curated collection of TipTap extensions with comprehensive TypeScript type definitions for Docen.

## Features

- 📦 **All-in-One Package** - All Docen-required extensions in a single dependency
- 🔒 **Full Type Safety** - Comprehensive TypeScript definitions for all content nodes and marks
- 🎯 **Curated Selection** - Only includes extensions actively used in Docen, no bloat
- 📤 **Type Exports** - Direct access to all node types (DocumentNode, ParagraphNode, etc.)
- 🚀 **Ready to Use** - Pre-configured extension arrays for blocks and marks

## Installation

```bash
# Install with npm
$ npm install @docen/tiptap-extensions

# Install with yarn
$ yarn add @docen/tiptap-extensions

# Install with pnpm
$ pnpm add @docen/tiptap-extensions
```

## Quick Start

```typescript
import { tiptapExtensions, tiptapMarkExtensions } from "@docen/tiptap-extensions";

const editor = new Editor({
  extensions: [
    ...tiptapExtensions, // All block extensions
    ...tiptapMarkExtensions, // All mark extensions
  ],
  content: "<p>Hello, world!</p>",
});
```

## Exports

### Extension Arrays

**`tiptapExtensions`** - Block-level extensions:

- Document - Root document node
- Paragraph - Standard paragraphs
- Heading - H1-H6 headings
- Blockquote - Blockquote sections
- CodeBlock - Code blocks with Lowlight syntax highlighting
- HorizontalRule - Horizontal dividers
- Image - Image embedding
- Details - Collapsible details/summary sections
- Table - Table containers
- TableRow - Table rows
- TableCell - Table body cells
- TableHeader - Table header cells
- BulletList - Unordered lists
- OrderedList - Ordered lists with start support
- ListItem - List item containers
- TaskList - Task list containers
- TaskItem - Task items with checkboxes

**`tiptapMarkExtensions`** - Text formatting marks:

- Bold - **Bold text**
- Italic - _Italic text_
- Underline - <u>Underlined text</u>
- Strike - ~~Strikethrough text~~
- Code - `Inline code`
- Highlight - Text highlighting
- Subscript - Sub~script~
- Superscript - Super^script^
- TextStyle - Text styling (colors, fonts, sizes)
- Link - Hyperlinks with href, target, rel attributes

### TypeScript Types

This package exports comprehensive TypeScript types for type-safe development:

```typescript
import type {
  // Root and document types
  JSONContent,

  // Block nodes
  DocumentNode,
  ParagraphNode,
  HeadingNode,
  BlockquoteNode,
  CodeBlockNode,
  HorizontalRuleNode,
  ImageNode,
  DetailsNode,

  // List nodes
  BulletListNode,
  OrderedListNode,
  ListItemNode,
  TaskListNode,
  TaskItemNode,

  // Table nodes
  TableNode,
  TableRowNode,
  TableCellNode,
  TableHeaderNode,

  // Text nodes
  TextNode,
  HardBreakNode,

  // Type unions
  BlockNode,
  TextContent,
} from "@docen/tiptap-extensions";
```

### Type Definitions

**Content Nodes:**

```typescript
// Paragraph with text alignment and spacing
interface ParagraphNode {
  type: "paragraph";
  attrs?: {
    textAlign?: "left" | "right" | "center" | "justify";
    indentLeft?: number;
    indentRight?: number;
    indentFirstLine?: number;
    spacingBefore?: number;
    spacingAfter?: number;
  };
  content?: Array<TextContent>;
}

// Heading with level and spacing
interface HeadingNode {
  type: "heading";
  attrs: {
    level: 1 | 2 | 3 | 4 | 5 | 6;
    indentLeft?: number;
    indentRight?: number;
    indentFirstLine?: number;
    spacingBefore?: number;
    spacingAfter?: number;
  };
  content?: Array<TextContent>;
}

// Table with colspan/rowspan
interface TableCellNode {
  type: "tableCell" | "tableHeader";
  attrs?: {
    colspan?: number;
    rowspan?: number;
    colwidth?: number[] | null;
  };
  content?: Array<ParagraphNode>;
}

// Image with attributes (extended from TipTap)
interface ImageNode {
  type: "image";
  attrs?: {
    src: string;
    alt?: string | null;
    title?: string | null;
    width?: number | null;
    height?: number | null;
    rotation?: number; // Additional attribute: rotation in degrees (not in TipTap core)
  };
}
```

**Text and Marks:**

```typescript
// Text node with marks
interface TextNode {
  type: "text";
  text: string;
  marks?: Array<Mark>;
}

// Mark with attributes
interface Mark {
  type:
    | "bold"
    | "italic"
    | "underline"
    | "strike"
    | "code"
    | "textStyle"
    | "link"
    | "highlight"
    | "subscript"
    | "superscript";
  attrs?: {
    // TextStyle attributes
    color?: string;
    backgroundColor?: string;
    fontSize?: string;
    fontFamily?: string;
    lineHeight?: string;

    // Link attributes
    href?: string;
    target?: string;
    rel?: string;
    class?: string | null;

    // Other attributes
    [key: string]: unknown;
  };
}
```

## Usage Examples

### Type-Safe Content Creation

```typescript
import type { JSONContent, ParagraphNode } from "@docen/tiptap-extensions";

const doc: JSONContent = {
  type: "doc",
  content: [
    {
      type: "paragraph",
      content: [
        {
          type: "text",
          marks: [{ type: "bold" }],
          text: "Hello, world!",
        },
      ],
    },
  ],
};

// Type narrowing with type guards
function isParagraph(node: JSONContent): node is ParagraphNode {
  return node.type === "paragraph";
}
```

### Working with Tables

```typescript
import type { TableNode, TableCellNode } from "@docen/tiptap-extensions";

const table: TableNode = {
  type: "table",
  content: [
    {
      type: "tableRow",
      content: [
        {
          type: "tableHeader",
          attrs: { colspan: 2, rowspan: 1 },
          content: [
            {
              type: "paragraph",
              content: [{ type: "text", text: "Header" }],
            },
          ],
        },
      ],
    },
  ],
};
```

### Custom Editor Setup

```typescript
import { Editor } from "@tiptap/core";
import { tiptapExtensions } from "@docen/tiptap-extensions";

const editor = new Editor({
  extensions: [
    ...tiptapExtensions,
    // Add your custom extensions here
  ],
});
```

## Import Paths

This package provides two import paths for flexibility:

```typescript
// Main entry point - extensions and types
import { tiptapExtensions, tiptapMarkExtensions } from "@docen/tiptap-extensions";
import type { JSONContent, ParagraphNode } from "@docen/tiptap-extensions";

// Types-only path for type definitions
import type { JSONContent } from "@docen/tiptap-extensions/types";
```

## Contributing

Contributions are welcome! Please read our [Contributor Covenant](https://www.contributor-covenant.org/version/2/1/code_of_conduct/) and submit pull requests to the [main repository](https://github.com/DemoMacro/docen).

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
