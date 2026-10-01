import type { CustomGeometryOptions, PathCommand } from "@office-open/core/drawing";

export interface CustomGeometryOutline {
  d: string;
  fill: boolean;
  stroke: boolean;
}

/** A numeric custGeom coordinate in path coordinate space. Guide references
 *  need the geometry guide evaluator; a pen stroke only writes literal points.
 *  Null keeps an unsupported command from producing a corrupt visual path. */
const coordinate = (value: string): number | null => {
  if (!/^[+-]?\d+(?:\.\d+)?$/.test(value)) return null;
  return Number(value);
};

const pointOf = (
  point: { x: string; y: string },
  scaleX: number,
  scaleY: number,
): [number, number] | null => {
  const x = coordinate(point.x);
  const y = coordinate(point.y);
  return x == null || y == null ? null : [x * scaleX, y * scaleY];
};

const commandData = (command: PathCommand, scaleX: number, scaleY: number): string | null => {
  const pt = (point: { x: string; y: string }): string => {
    const value = pointOf(point, scaleX, scaleY);
    if (!value) throw new Error("unsupported custom geometry guide");
    return `${value[0].toFixed(3)} ${value[1].toFixed(3)}`;
  };
  switch (command.command) {
    case "moveTo":
      return `M ${pt(command.point)}`;
    case "lineTo":
      return `L ${pt(command.point)}`;
    case "quadBezTo":
      return `Q ${command.points.map(pt).join(" ")}`;
    case "cubicBezTo":
      return `C ${command.points.map(pt).join(" ")}`;
    case "close":
      return "Z";
    case "arcTo":
      throw new Error("unsupported custom geometry arc");
  }
};

/** CT_CustomGeometry2D paths → the renderer-native outline contract. Literal
 *  coordinates scale from each path's declared space into the shape box. */
export function customGeometryOutlines(
  geometry: CustomGeometryOptions | undefined,
  width: number,
  height: number,
): CustomGeometryOutline[] {
  if (!geometry) return [];
  const outlines: CustomGeometryOutline[] = [];
  for (const path of geometry.pathList ?? []) {
    const pathWidth = path.w && path.w > 0 ? path.w : width;
    const pathHeight = path.h && path.h > 0 ? path.h : height;
    if (!pathWidth || !pathHeight) continue;
    const scaleX = width / pathWidth;
    const scaleY = height / pathHeight;
    const commands = (path.commands ?? []).map((command) => commandData(command, scaleX, scaleY));
    if (commands.some((command) => command == null) || commands.length === 0) continue;
    outlines.push({
      d: (commands as string[]).join(" "),
      fill: path.fill !== "none",
      stroke: path.stroke !== false,
    });
  }
  return outlines;
}
