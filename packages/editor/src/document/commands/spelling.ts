import type { Editor } from "@docen/docx/core";

import {
  addSpellWord,
  checkSpelling,
  ignoreSpellWord,
  spellSuggestions,
  type SpellingIssue,
} from "../spelling";

/** The spelling commands' view of the host — resolved per call so the
 *  controller can be built before a document opens. */
export interface SpellingHost {
  /** The headless editor — undefined before a document opens. */
  editor(): Editor | null | undefined;
  /** The story bridge — squiggle overlay feeds and jump scrolling. */
  bridge():
    | { setSpellingIssues(issues: SpellingIssue[]): void; scrollIntoView(pos: number): void }
    | undefined;
  /** The host element — the shadow-DOM root for the proofing pane and the
   *  status-bar book. */
  element(): HTMLElement;
}

/**
 * The spelling domain, split out of the host element: the debounced check
 * after every render, the issue list that feeds the squiggle overlay, the
 * proofing pane, and the status-bar book, plus the replace/ignore/navigation
 * interactions.
 */
export class SpellingCommands {
  constructor(private readonly host: SpellingHost) {}

  #issues: SpellingIssue[] = [];
  /** The pane's active issue (document order); -1 = nothing selected. */
  #active = -1;
  #timer: ReturnType<typeof setTimeout> | null = null;

  /** Tear down pending timers (the host's disconnectedCallback). */
  dispose(): void {
    if (this.#timer != null) clearTimeout(this.#timer);
  }

  /** Re-check after the user pauses (debounced; driven by every render). */
  schedule(): void {
    if (this.#timer != null) clearTimeout(this.#timer);
    this.#timer = setTimeout(() => {
      this.#timer = null;
      this.run();
    }, 400);
  }

  /** The live issue list (the ribbon's spell-check command reads it). */
  issues(): SpellingIssue[] {
    return this.#issues;
  }

  run(): void {
    const editor = this.host.editor();
    if (!editor) return;
    this.#issues = checkSpelling(editor.state.doc);
    this.host.bridge()?.setSpellingIssues(this.#issues);
    this.#active = this.#issues.length ? 0 : -1;
    const bar = this.host.element().shadowRoot?.querySelector("docen-status-bar");
    bar?.setAttribute("proofing", this.#issues.length ? "issues" : "ok");
    this.#syncPane();
  }

  /** Push the active issue (with its suggestions, computed here — the pane
   *  stays data-only) into the proofing pane when it's open. */
  #syncPane(): void {
    const pane = this.host.element().shadowRoot?.querySelector("docen-spelling-pane") as
      | (HTMLElement & {
          entries: Array<{ word: string; suggestions: string[] }>;
          active: number;
          total: number;
        })
      | null;
    if (!pane) return;
    const issues = this.#issues;
    pane.total = issues.length;
    pane.entries = issues.map((i) => ({ word: i.word, suggestions: [] }));
    if (this.#active >= 0 && this.#active < issues.length) {
      pane.active = this.#active;
      pane.entries[this.#active].suggestions = spellSuggestions(issues[this.#active].word);
      pane.entries = [...pane.entries];
    } else {
      pane.active = -1;
    }
  }

  /** Select and scroll to a spelling issue (the pane / command navigation). */
  goto(index: number): void {
    const issues = this.#issues;
    if (!issues.length) return;
    this.#active = ((index % issues.length) + issues.length) % issues.length;
    const issue = issues[this.#active];
    this.host.editor()?.commands.setTextSelection({ from: issue.from, to: issue.to });
    this.host.bridge()?.scrollIntoView(issue.from);
    this.#syncPane();
  }

  /** The pane's active issue index (the nav stepper offsets from it). */
  activeIndex(): number {
    return this.#active;
  }

  /** Replace the active issue's text with a suggestion — one transaction, so
   *  undo steps the whole replacement; the re-check rides the render. */
  replace(replacement: string): void {
    const issue = this.#issues[this.#active];
    const editor = this.host.editor();
    if (!issue || !editor) return;
    editor.commands.insertContentAt({ from: issue.from, to: issue.to }, replacement);
    editor.commands.setTextSelection({ from: issue.from, to: issue.from + replacement.length });
  }

  /** Add the active word to the session dictionary, or skip it for this
   *  session — then re-check, which drops the flagged occurrences. */
  ignore(mode: "ignore" | "add"): void {
    const issue = this.#issues[this.#active];
    if (!issue) return;
    if (mode === "add") addSpellWord(issue.word);
    else ignoreSpellWord(issue.word);
    this.run();
  }
}
