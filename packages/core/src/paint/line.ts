/** OOXML prstDash tokens → dash patterns in px (line-width units, the host's
 *  preset line styles); unlisted tokens render solid. */
const PRSTDASH_PATTERN: Record<string, number[]> = {
  dot: [1, 3],
  // A 1px-on/1px-off antialiased hairline blends to a faint tint — Word's
  // sysDot boxes read as clear 2px dots at hairline widths (user-verified).
  sysDot: [2, 2],
  dash: [4, 3],
  sysDash: [4, 2],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [4, 2, 1, 2],
  dashDotDot: [4, 3, 1, 3, 1, 3],
  sysDashDotDot: [4, 2, 1, 2, 1, 2],
  lgDash: [12, 3],
  lgDashDot: [12, 3, 1, 3],
  lgDashDotDot: [12, 3, 1, 3, 1, 3],
};

/** One line's outline as Leafer stroke props: hex color (ink when absent),
 *  the 1.5 px floor that keeps dashed hairlines from vanishing in their gaps,
 *  and the dressing Leafer spells differently (cap/join — a flat cap has no
 *  Leafer counterpart and stays unset). Shared by path members, shape boxes
 *  and picture borders; strokeAlign stays at the call sites (closed shapes
 *  need the explicit center, open paths already default to it). */
export function strokePropsOf(
  line:
    | {
        px: number;
        color?: string;
        cap?: "round" | "square" | "flat";
        join?: "round" | "bevel" | "miter";
        dash?: string;
      }
    | undefined,
) {
  return {
    stroke: line ? (line.color ? `#${line.color}` : "#000000") : undefined,
    strokeWidth: line?.px != null && line.dash ? Math.max(line.px, 1.5) : line?.px,
    strokeCap: line?.cap === "round" || line?.cap === "square" ? line.cap : undefined,
    strokeJoin: line?.join === "round" || line?.join === "bevel" ? line.join : undefined,
    dashPattern: line?.dash ? PRSTDASH_PATTERN[line.dash] : undefined,
  };
}
