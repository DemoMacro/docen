<script setup lang="ts">
import { generateDOCX } from "@docen/export-docx";
import { parseDOCX } from "@docen/import-docx";
import type { DocxExportOptions } from "@docen/export-docx";
import type { JSONContent } from "@docen/extensions/types";
import { tiptapExtensions } from "@docen/extensions";
import type { EditorToolbarItem, EditorSuggestionMenuItem } from "@nuxt/ui";

// Editor content - MUST be defined first!
const content = ref<JSONContent>({ type: "doc" });

// UEditor already includes these built-in extensions, so we exclude them
const builtInExtensions = [
  "link",
  "mention",
  "doc",
  "text",
  "hardBreak",
  "blockquote",
  "orderedList",
  "bulletList",
  "listItem",
  "codeBlock",
  "horizontalRule",
  "bold",
  "code",
  "italic",
  "strike",
  "underline",
];

const extensions = computed(() => {
  const exts = tiptapExtensions.filter((ext) => {
    const name = ext.name;
    return !builtInExtensions.includes(name);
  });
  return exts as any;
});

// Import DOCX
const isImporting = ref(false);
const importFileInput = ref<HTMLInputElement>();

async function importDOCX(event: Event) {
  const target = event.target as HTMLInputElement;
  const file = target.files?.[0];
  if (!file) return;

  try {
    isImporting.value = true;
    const arrayBuffer = await file.arrayBuffer();
    const jsonContent = await parseDOCX(arrayBuffer);
    content.value = jsonContent;

    useToast().add({
      title: "Import successful",
      description: `File imported: ${file.name}`,
      color: "success",
      icon: "i-lucide-check-circle",
    });
  } catch (error) {
    console.error("Import failed:", error);
    useToast().add({
      title: "Import failed",
      description: error instanceof Error ? error.message : "Unknown error",
      color: "error",
      icon: "i-lucide-x-circle",
    });
  } finally {
    isImporting.value = false;
    if (target) target.value = "";
  }
}

// Export DOCX
const isExporting = ref(false);

async function exportDOCX() {
  try {
    isExporting.value = true;

    const options: DocxExportOptions = {
      outputType: "blob",
      title: "Docen Document",
      creator: "Docen Editor",
      description: "Document created with Docen",
    };

    const blob = (await generateDOCX(content.value, options)) as Blob;

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `document-${Date.now()}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    useToast().add({
      title: "Export successful",
      description: "Document exported as DOCX",
      color: "success",
      icon: "i-lucide-check-circle",
    });
  } catch (error) {
    console.error("Export failed:", error);
    useToast().add({
      title: "Export failed",
      description: error instanceof Error ? error.message : "Unknown error",
      color: "error",
      icon: "i-lucide-x-circle",
    });
  } finally {
    isExporting.value = false;
  }
}

// Toolbar items
const toolbarItems: EditorToolbarItem[][] = [
  // Undo/Redo
  [
    { kind: "undo", icon: "i-lucide-undo", tooltip: { text: "Undo" } },
    { kind: "redo", icon: "i-lucide-redo", tooltip: { text: "Redo" } },
  ],
  // Headings
  [
    {
      icon: "i-lucide-heading",
      tooltip: { text: "Headings" },
      content: {
        align: "start",
      },
      items: [
        {
          kind: "heading",
          level: 1,
          icon: "i-lucide-heading-1",
          label: "Heading 1",
        },
        {
          kind: "heading",
          level: 2,
          icon: "i-lucide-heading-2",
          label: "Heading 2",
        },
        {
          kind: "heading",
          level: 3,
          icon: "i-lucide-heading-3",
          label: "Heading 3",
        },
        {
          kind: "heading",
          level: 4,
          icon: "i-lucide-heading-4",
          label: "Heading 4",
        },
        {
          kind: "heading",
          level: 5,
          icon: "i-lucide-heading-5",
          label: "Heading 5",
        },
        {
          kind: "heading",
          level: 6,
          icon: "i-lucide-heading-6",
          label: "Heading 6",
        },
      ],
    },
  ],
  // Lists & Block
  [
    { kind: "bulletList", icon: "i-lucide-list", label: "Bullet List" },
    {
      kind: "orderedList",
      icon: "i-lucide-list-ordered",
      label: "Ordered List",
    },
    {
      kind: "blockquote",
      icon: "i-lucide-text-quote",
      tooltip: { text: "Blockquote" },
    },
    {
      kind: "codeBlock",
      icon: "i-lucide-square-code",
      tooltip: { text: "Code Block" },
    },
  ],
  // Text formatting
  [
    {
      kind: "mark",
      mark: "bold",
      icon: "i-lucide-bold",
      tooltip: { text: "Bold" },
    },
    {
      kind: "mark",
      mark: "italic",
      icon: "i-lucide-italic",
      tooltip: { text: "Italic" },
    },
    {
      kind: "mark",
      mark: "underline",
      icon: "i-lucide-underline",
      tooltip: { text: "Underline" },
    },
    {
      kind: "mark",
      mark: "strike",
      icon: "i-lucide-strikethrough",
      tooltip: { text: "Strikethrough" },
    },
    {
      kind: "mark",
      mark: "code",
      icon: "i-lucide-code",
      tooltip: { text: "Code" },
    },
  ],
  // Alignment
  [
    {
      icon: "i-lucide-align-left",
      tooltip: { text: "Align Left" },
      kind: "textAlign",
      align: "left",
    },
    {
      icon: "i-lucide-align-center",
      tooltip: { text: "Align Center" },
      kind: "textAlign",
      align: "center",
    },
    {
      icon: "i-lucide-align-right",
      tooltip: { text: "Align Right" },
      kind: "textAlign",
      align: "right",
    },
    {
      icon: "i-lucide-align-justify",
      tooltip: { text: "Align Justify" },
      kind: "textAlign",
      align: "justify",
    },
  ],
  // Import/Export
  [
    {
      icon: "i-lucide-upload",
      tooltip: { text: "Import DOCX" },
      loading: isImporting.value,
      disabled: isImporting.value,
      onClick: () => importFileInput.value?.click(),
    },
    {
      icon: "i-lucide-download",
      tooltip: { text: "Export DOCX" },
      loading: isExporting.value,
      disabled: isExporting.value,
      onClick: exportDOCX,
    },
  ],
];

// Suggestion menu items
const suggestionItems: EditorSuggestionMenuItem[][] = [
  [
    { type: "label", label: "Basic Blocks" },
    { kind: "paragraph", label: "Paragraph", icon: "i-lucide-type" },
    {
      kind: "heading",
      level: 1,
      label: "Heading 1",
      icon: "i-lucide-heading-1",
    },
    {
      kind: "heading",
      level: 2,
      label: "Heading 2",
      icon: "i-lucide-heading-2",
    },
    {
      kind: "heading",
      level: 3,
      label: "Heading 3",
      icon: "i-lucide-heading-3",
    },
    { kind: "bulletList", label: "Bullet List", icon: "i-lucide-list" },
    {
      kind: "orderedList",
      label: "Ordered List",
      icon: "i-lucide-list-ordered",
    },
    { kind: "blockquote", label: "Blockquote", icon: "i-lucide-text-quote" },
    { kind: "codeBlock", label: "Code Block", icon: "i-lucide-square-code" },
  ],
  [
    { type: "label", label: "Formatting" },
    { kind: "mark", mark: "bold", label: "Bold", icon: "i-lucide-bold" },
    { kind: "mark", mark: "italic", label: "Italic", icon: "i-lucide-italic" },
    {
      kind: "mark",
      mark: "underline",
      label: "Underline",
      icon: "i-lucide-underline",
    },
    {
      kind: "mark",
      mark: "strike",
      label: "Strikethrough",
      icon: "i-lucide-strikethrough",
    },
    { kind: "mark", mark: "code", label: "Code", icon: "i-lucide-code" },
  ],
  [
    { type: "label", label: "Insert" },
    {
      kind: "horizontalRule",
      label: "Horizontal Rule",
      icon: "i-lucide-separator-horizontal",
    },
  ],
];
</script>

<template>
  <div class="flex flex-col h-screen">
    <!-- Header with Import/Export buttons -->
    <UContainer>
      <UHeader title="Docen">
        <template #right>
          <div class="flex items-center gap-2">
            <!-- Hidden file input -->
            <input
              ref="importFileInput"
              type="file"
              accept=".docx"
              class="hidden"
              @change="importDOCX"
            />
          </div>
        </template>
      </UHeader>
    </UContainer>

    <!-- Main Editor with all slots -->
    <UContainer class="flex-1 py-6">
      <UEditor
        v-slot="{ editor }"
        v-model="content"
        :extensions="extensions"
        :starter-kit="{
          heading: false,
          paragraph: false,
        }"
        :image="false"
        autofocus
        placeholder="Type / for commands..."
        class="w-full min-h-37 flex flex-col gap-4"
      >
        <!-- Fixed Toolbar -->
        <UEditorToolbar :editor="editor" :items="toolbarItems" class="sm:px-8 overflow-x-auto" />

        <!-- Drag Handle -->
        <UEditorDragHandle :editor="editor" />

        <!-- Suggestion Menu -->
        <UEditorSuggestionMenu :editor="editor" :items="suggestionItems" />
      </UEditor>
    </UContainer>
  </div>
</template>
