/**
 * Document demo — mounts the full `<docen-document>` editor with theme +
 * add-in registration. Mirrors the inline setup that used to live in
 * packages/editor/index.html (now replaced by the tabbed demo shell).
 */
import {
  applyTheme,
  createLightTheme,
  registerTheme,
  registerTranslation,
  type BrandVariants,
} from "@docen/editor";

export const mountDocumentDemo = (stage: HTMLElement): void => {
  applyTheme("light");

  const contosoGreen: BrandVariants = {
    10: "#052506",
    20: "#0a3909",
    30: "#0f4d0d",
    40: "#14610c",
    50: "#1a750b",
    60: "#1f8a0a",
    70: "#259e08",
    80: "#2ab207",
    90: "#43c322",
    100: "#62d340",
    110: "#75d450",
    120: "#93de6b",
    130: "#a8e380",
    140: "#b8e896",
    150: "#c9eda9",
    160: "#dff3c5",
  };
  registerTheme("contoso", createLightTheme(contosoGreen));
  registerTranslation({
    languageTag: "en",
    translations: { "theme.contoso": "Contoso Green" },
  });
  registerTranslation({
    languageTag: "zh-CN",
    translations: { "theme.contoso": "Contoso 绿" },
  });

  registerTranslation({
    languageTag: "fr",
    $name: "Français",
    translations: {
      "ribbon.tab.home": "Accueil",
      "ribbon.tab.insert": "Insertion",
      "header.autosave": "Enregistrement automatique",
      "header.save": "Enregistrer",
      "header.undo": "Annuler",
      "header.redo": "Rétablir",
      "header.open": "Ouvrir…",
      "header.save-as": "Enregistrer sous…",
      "header.print": "Imprimer",
      "header.options": "Options",
      "options.title": "Options",
      "options.ok": "OK",
      "options.cancel": "Annuler",
      "options.language": "Langue",
      "status.page-of": "Page {page} sur {total}",
      "status.words": "{n} mots",
      "theme.contoso": "Contoso Vert",
    },
  });

  const doc = document.createElement("docen-document") as HTMLElement & {
    addAddin: (addin: unknown) => void;
    openDOCX: (input: File | ArrayBuffer | Uint8Array) => Promise<void>;
  };
  doc.setAttribute("user", "Demo Macro");
  doc.setAttribute("filename", "Welcome.docx");
  doc.setAttribute("navigation-pane", "");
  doc.className = "demo-doc";
  stage.append(doc);

  doc.addAddin({
    id: "about",
    ribbon: [
      {
        tab: "about",
        label: "about.tab",
        groups: [
          {
            id: "about-help",
            label: "about.group.help",
            controls: [{ type: "button", id: "help", label: "about.cmd.help", event: "open-help" }],
          },
        ],
      },
    ],
    localizationInfo: {
      defaultLanguageTag: "en",
      additionalLanguages: [
        {
          languageTag: "en",
          translations: {
            "about.tab": "About",
            "about.group.help": "Help",
            "about.cmd.help": "Help",
          },
        },
        {
          languageTag: "zh-CN",
          translations: {
            "about.tab": "关于",
            "about.group.help": "帮助",
            "about.cmd.help": "帮助",
          },
        },
      ],
    },
    commands: {
      "open-help": () => window.open("https://github.com/DemoMacro/docen", "_blank"),
    },
  });

  doc.addAddin({
    id: "mini-toolbar-demo",
    miniToolbar: [
      {
        icon: "highlight",
        event: "highlight",
        label: "ribbon.cmd.highlight",
        activeMark: "highlight",
      },
    ],
  });
};
