/** One reversible deck mutation. The closures capture the minimum source
 *  slices they need — the whole deck is never snapshotted. */
export interface DeckEdit {
  undo(): void;
  redo(): void;
}

/** The deck's linear undo/redo stack: a new edit truncates the redo branch,
 *  and the oldest edit falls off after `limit` entries. */
export class DeckHistory {
  readonly #edits: DeckEdit[] = [];
  #index = -1;

  constructor(private readonly limit = 50) {}

  get canUndo(): boolean {
    return this.#index >= 0;
  }

  get canRedo(): boolean {
    return this.#index < this.#edits.length - 1;
  }

  push(edit: DeckEdit): void {
    this.#edits.length = this.#index + 1;
    this.#edits.push(edit);
    if (this.#edits.length > this.limit) this.#edits.shift();
    this.#index = this.#edits.length - 1;
  }

  undo(): void {
    this.#edits[this.#index--]?.undo();
  }

  redo(): void {
    this.#edits[++this.#index]?.redo();
  }

  clear(): void {
    this.#edits.length = 0;
    this.#index = -1;
  }
}
