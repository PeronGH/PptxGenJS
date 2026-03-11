import {
  SlideTransitionCornerDirection,
  SlideTransitionEightDirection,
  SlideTransitionEightDirectionProps,
  SlideTransitionInOutDirection,
  SlideTransitionOptionalBlackProps,
  SlideTransitionOrientation,
  SlideTransitionOrientationProps,
  SlideTransitionProps,
  SlideTransitionSideDirection,
  SlideTransitionSideDirectionProps,
  SlideTransitionSplitProps,
  SlideTransitionStripsProps,
  SlideTransitionWheelProps,
  SlideTransitionZoomProps,
} from "./core-interfaces.ts";

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";

const EMPTY_TRANSITIONS = new Set([
  "circle",
  "diamond",
  "dissolve",
  "newsflash",
  "plus",
  "random",
  "wedge",
]);
const ORIENTATION_TRANSITIONS = new Set([
  "blinds",
  "checker",
  "comb",
  "randomBar",
]);
const EIGHT_DIRECTION_TRANSITIONS = new Set(["cover", "pull"]);
const OPTIONAL_BLACK_TRANSITIONS = new Set(["cut", "fade"]);
const SIDE_DIRECTION_TRANSITIONS = new Set(["push", "wipe"]);
const STRIPS_TRANSITIONS = new Set(["strips"]);
const ZOOM_TRANSITIONS = new Set(["zoom"]);
const ALL_TRANSITIONS = [
  ...EMPTY_TRANSITIONS,
  ...ORIENTATION_TRANSITIONS,
  ...EIGHT_DIRECTION_TRANSITIONS,
  ...OPTIONAL_BLACK_TRANSITIONS,
  ...SIDE_DIRECTION_TRANSITIONS,
  "split",
  ...STRIPS_TRANSITIONS,
  "wheel",
  ...ZOOM_TRANSITIONS,
];
const COMMON_KEYS = new Set([
  "advanceOnClick",
  "advanceTime",
  "duration",
  "speed",
  "type",
]);

const ORIENTATION_TO_XML: Record<SlideTransitionOrientation, string> = {
  horizontal: "horz",
  vertical: "vert",
};
const SIDE_DIRECTION_TO_XML: Record<SlideTransitionSideDirection, string> = {
  down: "d",
  left: "l",
  right: "r",
  up: "u",
};
const CORNER_DIRECTION_TO_XML: Record<SlideTransitionCornerDirection, string> =
  {
    leftDown: "ld",
    leftUp: "lu",
    rightDown: "rd",
    rightUp: "ru",
  };
const IN_OUT_DIRECTION_TO_XML: Record<SlideTransitionInOutDirection, string> = {
  in: "in",
  out: "out",
};

export function normalizeSlideTransition(
  transition?: SlideTransitionProps | null,
): SlideTransitionProps | undefined {
  if (transition === null || typeof transition === "undefined") {
    return undefined;
  }
  if (typeof transition !== "object" || Array.isArray(transition)) {
    throw new Error("slide.transition must be an object");
  }

  const candidate = { ...transition } as SlideTransitionProps;
  const type = candidate.type;
  if (!type || !ALL_TRANSITIONS.includes(type)) {
    throw new Error(
      `slide.transition.type must be one of: ${ALL_TRANSITIONS.join(", ")}`,
    );
  }

  validateCommonProps(candidate);

  if (EMPTY_TRANSITIONS.has(type)) {
    validateKnownKeys(candidate, COMMON_KEYS, type);
    return candidate;
  }
  if (ORIENTATION_TRANSITIONS.has(type)) {
    const normalized = candidate as SlideTransitionOrientationProps;
    validateKnownKeys(
      normalized,
      new Set([...COMMON_KEYS, "orientation"]),
      type,
    );
    if (
      typeof normalized.orientation !== "undefined" &&
      !hasOwn(ORIENTATION_TO_XML, normalized.orientation)
    ) {
      throw new Error(
        'slide.transition.orientation must be "horizontal" or "vertical"',
      );
    }
    return normalized;
  }
  if (EIGHT_DIRECTION_TRANSITIONS.has(type)) {
    const normalized = candidate as SlideTransitionEightDirectionProps;
    validateKnownKeys(normalized, new Set([...COMMON_KEYS, "direction"]), type);
    if (typeof normalized.direction !== "undefined") {
      mapEightDirection(normalized.direction);
    }
    return normalized;
  }
  if (OPTIONAL_BLACK_TRANSITIONS.has(type)) {
    const normalized = candidate as SlideTransitionOptionalBlackProps;
    validateKnownKeys(
      normalized,
      new Set([...COMMON_KEYS, "throughBlack"]),
      type,
    );
    if (
      typeof normalized.throughBlack !== "undefined" &&
      typeof normalized.throughBlack !== "boolean"
    ) {
      throw new Error("slide.transition.throughBlack must be a boolean");
    }
    return normalized;
  }
  if (SIDE_DIRECTION_TRANSITIONS.has(type)) {
    const normalized = candidate as SlideTransitionSideDirectionProps;
    validateKnownKeys(normalized, new Set([...COMMON_KEYS, "direction"]), type);
    if (typeof normalized.direction !== "undefined") {
      mapSideDirection(normalized.direction);
    }
    return normalized;
  }
  if (type === "split") {
    const normalized = candidate as SlideTransitionSplitProps;
    validateKnownKeys(
      normalized,
      new Set([...COMMON_KEYS, "direction", "orientation"]),
      type,
    );
    if (typeof normalized.direction !== "undefined") {
      mapInOutDirection(normalized.direction);
    }
    if (
      typeof normalized.orientation !== "undefined" &&
      !hasOwn(ORIENTATION_TO_XML, normalized.orientation)
    ) {
      throw new Error(
        'slide.transition.orientation must be "horizontal" or "vertical"',
      );
    }
    return normalized;
  }
  if (STRIPS_TRANSITIONS.has(type)) {
    const normalized = candidate as SlideTransitionStripsProps;
    validateKnownKeys(normalized, new Set([...COMMON_KEYS, "direction"]), type);
    if (
      typeof normalized.direction !== "undefined" &&
      !hasOwn(CORNER_DIRECTION_TO_XML, normalized.direction)
    ) {
      throw new Error(
        "slide.transition.direction must be one of: leftUp, rightUp, leftDown, rightDown",
      );
    }
    return normalized;
  }
  if (type === "wheel") {
    const normalized = candidate as SlideTransitionWheelProps;
    validateKnownKeys(normalized, new Set([...COMMON_KEYS, "spokes"]), type);
    if (
      typeof normalized.spokes !== "undefined" &&
      ![1, 2, 3, 4, 8].includes(normalized.spokes)
    ) {
      throw new Error("slide.transition.spokes must be one of: 1, 2, 3, 4, 8");
    }
    return normalized;
  }
  if (ZOOM_TRANSITIONS.has(type)) {
    const normalized = candidate as SlideTransitionZoomProps;
    validateKnownKeys(normalized, new Set([...COMMON_KEYS, "direction"]), type);
    if (typeof normalized.direction !== "undefined") {
      mapInOutDirection(normalized.direction);
    }
    return normalized;
  }

  return candidate;
}

export function genXmlSlideTransition(
  transition?: SlideTransitionProps | null,
): string {
  const normalized = normalizeSlideTransition(transition);
  if (!normalized) return "";

  const childXml = genXmlTransitionChild(normalized);
  if (typeof normalized.duration === "undefined") {
    return genXmlTransitionElement(normalized, childXml);
  }

  return (
    `<mc:AlternateContent xmlns:mc="${MC_NS}" xmlns:p14="${P14_NS}">` +
    `<mc:Choice Requires="p14">${
      genXmlTransitionElement(normalized, childXml, true)
    }</mc:Choice>` +
    `<mc:Fallback>${
      genXmlTransitionElement(normalized, childXml)
    }</mc:Fallback>` +
    "</mc:AlternateContent>"
  );
}

function validateCommonProps(transition: SlideTransitionProps): void {
  if (
    typeof transition.speed !== "undefined" &&
    !["slow", "med", "fast"].includes(transition.speed)
  ) {
    throw new Error("slide.transition.speed must be one of: slow, med, fast");
  }
  if (
    typeof transition.advanceOnClick !== "undefined" &&
    typeof transition.advanceOnClick !== "boolean"
  ) {
    throw new Error("slide.transition.advanceOnClick must be a boolean");
  }
  if (typeof transition.advanceTime !== "undefined") {
    validateUnsignedInt(
      "slide.transition.advanceTime",
      transition.advanceTime,
      2147483647,
    );
  }
  if (typeof transition.duration !== "undefined") {
    validateUnsignedInt("slide.transition.duration", transition.duration);
  }
}

function validateKnownKeys(
  transition: SlideTransitionProps,
  allowedKeys: Set<string>,
  type: SlideTransitionProps["type"],
): void {
  Object.keys(transition).forEach((key) => {
    if (!allowedKeys.has(key)) {
      throw new Error(`slide.transition.${type} does not support "${key}"`);
    }
  });
}

function validateUnsignedInt(
  name: string,
  value: number,
  maxValue = Number.MAX_SAFE_INTEGER,
): void {
  if (!Number.isInteger(value) || value < 0 || value > maxValue) {
    throw new Error(`${name} must be a non-negative integer`);
  }
}

function genXmlTransitionChild(transition: SlideTransitionProps): string {
  switch (transition.type) {
    case "circle":
    case "diamond":
    case "dissolve":
    case "newsflash":
    case "plus":
    case "random":
    case "wedge":
      return genXmlLeafTransition(transition.type);
    case "blinds":
    case "checker":
    case "comb":
    case "randomBar":
      return genXmlLeafTransition(transition.type, [[
        "dir",
        ORIENTATION_TO_XML[transition.orientation],
      ]]);
    case "cover":
    case "pull":
      return genXmlLeafTransition(transition.type, [[
        "dir",
        mapEightDirection(transition.direction),
      ]]);
    case "cut":
    case "fade":
      return genXmlLeafTransition(transition.type, [[
        "thruBlk",
        transition.throughBlack,
      ]]);
    case "push":
    case "wipe":
      return genXmlLeafTransition(transition.type, [[
        "dir",
        mapSideDirection(transition.direction),
      ]]);
    case "split":
      return genXmlLeafTransition("split", [
        ["orient", ORIENTATION_TO_XML[transition.orientation]],
        ["dir", mapInOutDirection(transition.direction)],
      ]);
    case "strips":
      return genXmlLeafTransition("strips", [[
        "dir",
        CORNER_DIRECTION_TO_XML[transition.direction],
      ]]);
    case "wheel":
      return genXmlLeafTransition("wheel", [["spokes", transition.spokes]]);
    case "zoom":
      return genXmlLeafTransition("zoom", [[
        "dir",
        mapInOutDirection(transition.direction),
      ]]);
    default:
      throw new Error(
        `slide.transition.type must be one of: ${ALL_TRANSITIONS.join(", ")}`,
      );
  }
}

function genXmlTransitionElement(
  transition: SlideTransitionProps,
  childXml: string,
  includeDuration = false,
): string {
  const attrs: Array<[string, boolean | number | string]> = [];
  if (transition.speed) attrs.push(["spd", transition.speed]);
  if (includeDuration && typeof transition.duration !== "undefined") {
    attrs.push(["p14:dur", transition.duration]);
  }
  if (typeof transition.advanceOnClick !== "undefined") {
    attrs.push(["advClick", transition.advanceOnClick]);
  }
  if (typeof transition.advanceTime !== "undefined") {
    attrs.push(["advTm", transition.advanceTime]);
  }

  return `<p:transition${genXmlAttributes(attrs)}>${childXml}</p:transition>`;
}

function genXmlLeafTransition(
  name: string,
  attrs: Array<[string, boolean | number | string]> = [],
): string {
  return `<p:${name}${genXmlAttributes(attrs)}/>`;
}

function genXmlAttributes(
  attrs: Array<[string, boolean | number | string]>,
): string {
  const filtered = attrs.filter(([, value]) => typeof value !== "undefined");
  if (filtered.length === 0) return "";

  return " " +
    filtered.map(([name, value]) =>
      `${name}="${formatXmlAttributeValue(value)}"`
    ).join(" ");
}

function formatXmlAttributeValue(value: boolean | number | string): string {
  if (typeof value === "boolean") return value ? "1" : "0";
  return String(value);
}

function mapEightDirection(
  direction?: SlideTransitionEightDirection,
): string | undefined {
  if (typeof direction === "undefined") return undefined;
  if (hasOwn(SIDE_DIRECTION_TO_XML, direction)) {
    return SIDE_DIRECTION_TO_XML[direction as SlideTransitionSideDirection];
  }
  if (hasOwn(CORNER_DIRECTION_TO_XML, direction)) {
    return CORNER_DIRECTION_TO_XML[direction as SlideTransitionCornerDirection];
  }
  throw new Error(
    "slide.transition.direction must be one of: left, right, up, down, leftUp, rightUp, leftDown, rightDown",
  );
}

function mapSideDirection(
  direction?: SlideTransitionSideDirection,
): string | undefined {
  if (typeof direction === "undefined") return undefined;
  if (!hasOwn(SIDE_DIRECTION_TO_XML, direction)) {
    throw new Error(
      "slide.transition.direction must be one of: left, right, up, down",
    );
  }
  return SIDE_DIRECTION_TO_XML[direction];
}

function mapInOutDirection(
  direction?: SlideTransitionInOutDirection,
): string | undefined {
  if (typeof direction === "undefined") return undefined;
  if (!hasOwn(IN_OUT_DIRECTION_TO_XML, direction)) {
    throw new Error('slide.transition.direction must be "in" or "out"');
  }
  return IN_OUT_DIRECTION_TO_XML[direction];
}

function hasOwn<T extends object>(value: T, key: PropertyKey): key is keyof T {
  return Object.prototype.hasOwnProperty.call(value, key);
}
