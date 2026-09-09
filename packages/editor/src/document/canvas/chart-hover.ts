// Word's chart ScreenTip: hovering a plot element names its series and, on
// a data point, its category and value — read off the chart node the hit
// pairs with, not the painter's payload (a stale chart must not mislabel).

/** What the tip shows, resolved by the bridge from the hovered hit: the
 *  series name plus, on a data point, its category and value. */
export interface ChartTip {
  name: string;
  category?: string;
  value?: number;
}

export function installChartHover(deps: { host: HTMLElement }): {
  onMove: (tip: ChartTip | null, clientX: number, clientY: number) => void;
  hide: () => void;
} {
  const el = document.createElement("div");
  Object.assign(el.style, {
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
  const name = document.createElement("div");
  const detail = document.createElement("div");
  Object.assign(detail.style, { color: "#616161" } satisfies Partial<CSSStyleDeclaration>);
  el.append(name, detail);
  deps.host.append(el);

  let shown: string | null = null;
  const hide = (): void => {
    if (shown == null) return;
    shown = null;
    el.style.display = "none";
  };

  const onMove = (tip: ChartTip | null, clientX: number, clientY: number): void => {
    if (!tip) {
      hide();
      return;
    }
    const key = `${tip.name}/${tip.category ?? ""}/${tip.value ?? ""}`;
    if (key !== shown) {
      shown = key;
      name.textContent = tip.name;
      detail.textContent = tip.category != null ? `${tip.category}: ${tip.value ?? ""}` : "";
      detail.style.display = detail.textContent ? "block" : "none";
      el.style.display = "block";
    }
    // Park the tip under the pointer, clamped into the host's right edge.
    const rect = deps.host.getBoundingClientRect();
    const left = Math.min(clientX - rect.left + 12, rect.width - el.offsetWidth - 4);
    el.style.left = `${Math.max(0, left)}px`;
    el.style.top = `${clientY - rect.top + 20}px`;
  };

  return { onMove, hide };
}
