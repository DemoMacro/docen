// Hyperlink interaction for the canvas route — Word's desktop behavior:
// hovering a link shows its URL with the Ctrl+Click hint, and Ctrl+Click
// follows it (a plain click keeps editing, dropping the caret).

import { t } from "../../ui";

/** A resolved link mark at a caret position. */
export interface LinkHit {
  href: string;
  target: string | null;
}

export interface LinkHoverDeps {
  /** The tooltip's positioning host — the same box the input surface lives
   *  in (its rect turns client coordinates into host-local ones). */
  host: HTMLElement;
  /** Viewport point → caret position (null off-text). */
  posAtClient: (x: number, y: number) => number | null;
  /** The link mark at a caret position, null off-link. */
  linkAt: (pos: number) => LinkHit | null;
}

/** Word's link tooltip: the URL on top, the follow hint under it. */
export function installLinkHover(deps: LinkHoverDeps): {
  onMove: (event: MouseEvent) => void;
  hide: () => void;
} {
  const tip = document.createElement("div");
  Object.assign(tip.style, {
    position: "absolute",
    display: "none",
    zIndex: "9",
    pointerEvents: "none",
    maxWidth: "420px",
    padding: "6px 10px",
    background: "#ffffff",
    border: "1px solid #d4d4d4",
    borderRadius: "4px",
    boxShadow: "0 2px 6px rgba(0,0,0,.13)",
    fontFamily: "inherit",
    fontSize: "12px",
    lineHeight: "1.5",
  } satisfies Partial<CSSStyleDeclaration>);
  const url = document.createElement("div");
  Object.assign(url.style, {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  } satisfies Partial<CSSStyleDeclaration>);
  const hint = document.createElement("div");
  Object.assign(hint.style, { color: "#616161" } satisfies Partial<CSSStyleDeclaration>);
  tip.append(url, hint);
  deps.host.append(tip);

  let shownHref: string | null = null;

  const hide = (): void => {
    if (shownHref == null) return;
    shownHref = null;
    tip.style.display = "none";
  };

  const onMove = (event: MouseEvent): void => {
    const pos = deps.posAtClient(event.clientX, event.clientY);
    const link = pos != null ? deps.linkAt(pos) : null;
    if (!link || !link.href) {
      hide();
      return;
    }
    const hostRect = deps.host.getBoundingClientRect();
    const hrefChanged = link.href !== shownHref;
    if (hrefChanged) {
      shownHref = link.href;
      url.textContent = link.href;
      hint.textContent = t("docen.link.hint");
      tip.style.display = "block";
    }
    // Park the tooltip under the pointer, clamped into the host's right edge.
    const left = Math.min(event.clientX - hostRect.left + 12, hostRect.width - tip.offsetWidth - 4);
    tip.style.left = `${Math.max(0, left)}px`;
    tip.style.top = `${event.clientY - hostRect.top + 20}px`;
  };

  return { onMove, hide };
}

/** Follow the link Word-style: the mark's target (external links parse as
 *  _blank) in a new tab, never leaking the opener. Returns false when there
 *  is nothing to open (the click keeps its editing meaning). */
export function followLink(link: LinkHit | null): boolean {
  if (!link?.href || link.href.startsWith("#")) return false;
  window.open(link.href, link.target ?? "_blank", "noopener,noreferrer");
  return true;
}
