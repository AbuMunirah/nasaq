/*
 * Pure geometry helpers for the Nasaq commands. All values are in points,
 * matching PowerPoint.Shape left/top/width/height.
 */

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface Position {
  left: number;
  top: number;
}

/** "center" lines up horizontal centers (shapes move sideways); "middle" lines up vertical centers. */
export type Edge = "left" | "right" | "top" | "bottom" | "center" | "middle";

/**
 * Returns the new position of each box so that the given edge lines up with the
 * same edge of the target box (e.g. "left" -> every left edge moves to target.left).
 */
export function alignToEdge(boxes: Box[], edge: Edge, target: Box): Position[] {
  switch (edge) {
    case "left":
      return boxes.map((b) => ({ left: target.left, top: b.top }));
    case "right":
      return boxes.map((b) => ({ left: target.left + target.width - b.width, top: b.top }));
    case "top":
      return boxes.map((b) => ({ left: b.left, top: target.top }));
    case "bottom":
      return boxes.map((b) => ({ left: b.left, top: target.top + target.height - b.height }));
    case "center":
      return boxes.map((b) => ({ left: target.left + (target.width - b.width) / 2, top: b.top }));
    case "middle":
      return boxes.map((b) => ({ left: b.left, top: target.top + (target.height - b.height) / 2 }));
  }
}

/**
 * Returns the new positions of a and b after swapping them center to center:
 * each box ends up centered where the other one was centered.
 */
export function swapCenters(a: Box, b: Box): [Position, Position] {
  const centerA = { x: a.left + a.width / 2, y: a.top + a.height / 2 };
  const centerB = { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  return [
    { left: centerB.x - a.width / 2, top: centerB.y - a.height / 2 },
    { left: centerA.x - b.width / 2, top: centerA.y - b.height / 2 },
  ];
}
