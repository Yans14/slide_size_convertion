import { describe, it, expect } from "vitest";
const {
  transformElement,
  resolveOverlaps,
  splitByHeight,
  scoreSlideQuality,
  buildReviewDecision,
  buildDefaultPolicy,
  isFlowElement,
  isLowPriority,
  clamp,
  round,
} = require("../src/transform");

// ── Helpers ─────────────────────────────────────────────────────────────

describe("clamp", () => {
  it("clamps below minimum", () => expect(clamp(-5, 0, 10)).toBe(0));
  it("clamps above maximum", () => expect(clamp(15, 0, 10)).toBe(10));
  it("passes through in range", () => expect(clamp(5, 0, 10)).toBe(5));
});

describe("round", () => {
  it("rounds to 4 decimal places by default", () => {
    expect(round(3.14159265)).toBe(3.1416);
  });
  it("supports custom precision", () => {
    expect(round(3.14159265, 2)).toBe(3.14);
  });
});

// ── isFlowElement ───────────────────────────────────────────────────────

describe("isFlowElement", () => {
  it("returns true for text", () => expect(isFlowElement({ type: "text" })).toBe(true));
  it("returns true for table-text", () => expect(isFlowElement({ type: "table-text" })).toBe(true));
  it("returns true for placeholder", () => expect(isFlowElement({ type: "placeholder" })).toBe(true));
  it("returns true for table", () => expect(isFlowElement({ type: "table" })).toBe(true));
  it("returns true for chart", () => expect(isFlowElement({ type: "chart" })).toBe(true));
  it("returns false for image", () => expect(isFlowElement({ type: "image" })).toBe(false));
  it("returns false for shape", () => expect(isFlowElement({ type: "shape" })).toBe(false));
});

// ── isLowPriority ───────────────────────────────────────────────────────

describe("isLowPriority", () => {
  it("considers tiny text low priority", () => {
    expect(isLowPriority({ type: "text", text: "Hi", w: 0.01, h: 0.01 }, 100)).toBe(true);
  });

  it("does not consider substantial text low priority", () => {
    expect(isLowPriority({ type: "text", text: "Hello World This is substantial", w: 5, h: 2 }, 100)).toBe(false);
  });

  it("considers tiny images low priority", () => {
    expect(isLowPriority({ type: "image", w: 0.1, h: 0.1 }, 100)).toBe(true);
  });

  it("does not consider large images low priority", () => {
    expect(isLowPriority({ type: "image", w: 5, h: 5 }, 100)).toBe(false);
  });
});

// ── transformElement ────────────────────────────────────────────────────

describe("transformElement", () => {
  const policy = buildDefaultPolicy({ readabilityMinFontPt: 10 });

  it("scales position and dimensions", () => {
    const item = { type: "image", x: 1, y: 2, w: 3, h: 4 };
    const result = transformElement(item, 0.5, 0.1, 0.2, policy);
    expect(result.x).toBeCloseTo(0.6);
    expect(result.y).toBeCloseTo(1.2);
    expect(result.w).toBeCloseTo(1.5);
    expect(result.h).toBeCloseTo(2);
  });

  it("scales font size for text elements", () => {
    const item = { type: "text", x: 0, y: 0, w: 5, h: 2, fontSizePt: 20 };
    const result = transformElement(item, 0.5, 0, 0, policy);
    expect(result.fontSizePt).toBe(10);
  });

  it("clamps font size to minimum", () => {
    const item = { type: "text", x: 0, y: 0, w: 5, h: 2, fontSizePt: 8 };
    const result = transformElement(item, 0.5, 0, 0, policy);
    expect(result.fontSizePt).toBe(10); // readabilityMinFontPt
  });

  it("scales table cell font sizes", () => {
    const item = {
      type: "table", x: 0, y: 0, w: 10, h: 5,
      rows: [{ cells: [{ text: "A", fontSizePt: 20 }] }],
      colWidths: [10],
    };
    const result = transformElement(item, 0.5, 0, 0, policy);
    expect(result.rows[0].cells[0].fontSizePt).toBe(10);
    expect(result.colWidths[0]).toBeCloseTo(5);
  });

  it("transforms group children recursively", () => {
    const item = {
      type: "group", x: 0, y: 0, w: 10, h: 10,
      children: [
        { type: "text", x: 1, y: 1, w: 3, h: 2, fontSizePt: 20 },
      ],
    };
    const result = transformElement(item, 0.5, 0, 0, policy);
    expect(result.children[0].x).toBeCloseTo(0.5);
    expect(result.children[0].w).toBeCloseTo(1.5);
    expect(result.children[0].fontSizePt).toBe(10);
  });
});

// ── resolveOverlaps ─────────────────────────────────────────────────────

describe("resolveOverlaps", () => {
  it("repositions overlapping flow elements", () => {
    const items = [
      { type: "text", name: "A", x: 0, y: 0, w: 5, h: 2 },
      { type: "text", name: "B", x: 0, y: 1, w: 5, h: 2 },
    ];
    const actions = resolveOverlaps(items, 10);
    expect(actions.length).toBeGreaterThan(0);
    expect(actions[0].type).toBe("reposition");
    expect(actions[0].elementName).toBe("B");
    expect(items[1].y).toBeGreaterThan(1); // B was moved down
  });

  it("doesn't touch non-overlapping items", () => {
    const items = [
      { type: "text", name: "A", x: 0, y: 0, w: 5, h: 1 },
      { type: "text", name: "B", x: 0, y: 2, w: 5, h: 1 },
    ];
    const actions = resolveOverlaps(items, 10);
    expect(actions).toEqual([]);
  });

  it("ignores non-flow elements", () => {
    const items = [
      { type: "text", name: "A", x: 0, y: 0, w: 5, h: 2 },
      { type: "image", name: "B", x: 0, y: 1, w: 5, h: 2 },
    ];
    const actions = resolveOverlaps(items, 10);
    expect(actions).toEqual([]);
  });
});

// ── splitByHeight ───────────────────────────────────────────────────────

describe("splitByHeight", () => {
  const dstLayout = { width: 10, height: 7.5 };

  it("keeps items that fit on one slide", () => {
    const items = [
      { type: "text", name: "A", x: 0, y: 0.5, w: 5, h: 2 },
    ];
    const policy = buildDefaultPolicy({ allowSlideSplit: true });
    const { slides, actions } = splitByHeight(items, dstLayout, policy);
    expect(slides).toHaveLength(1);
    expect(slides[0]).toHaveLength(1);
    expect(actions).toHaveLength(0);
  });

  it("splits overflowing items to new slides", () => {
    const items = [
      { type: "text", name: "A", x: 0, y: 0.5, w: 5, h: 2 },
      { type: "text", name: "B", x: 0, y: 8, w: 5, h: 2 },
    ];
    const policy = buildDefaultPolicy({ allowSlideSplit: true });
    const { slides, actions } = splitByHeight(items, dstLayout, policy);
    expect(slides.length).toBeGreaterThanOrEqual(2);
    expect(actions.some((a) => a.type === "split")).toBe(true);
  });

  it("clamps items when splitting is disabled", () => {
    const items = [
      { type: "text", name: "A", x: 0, y: 8, w: 5, h: 2 },
    ];
    const policy = buildDefaultPolicy({ allowSlideSplit: false });
    const { slides, actions } = splitByHeight(items, dstLayout, policy);
    expect(slides).toHaveLength(1);
    expect(actions.some((a) => a.reason === "overflow-clamped")).toBe(true);
  });

  it("drops low-priority oversized items when deletion allowed", () => {
    const items = [
      { type: "text", name: "A", x: 0, y: 0, w: 0.1, h: 8, text: "x" },
    ];
    const policy = buildDefaultPolicy({ allowElementDeletion: true });
    const { slides, dropped } = splitByHeight(items, dstLayout, policy);
    expect(dropped).toHaveLength(1);
  });
});

// ── scoreSlideQuality ───────────────────────────────────────────────────

describe("scoreSlideQuality", () => {
  it("returns 1 for a perfect slide", () => {
    const score = scoreSlideQuality({
      sourceSlide: { unsupportedCount: 0, placeholderCount: 0 },
      actions: [],
      dropped: [],
      outputSlideCount: 1,
    });
    expect(score).toBe(1);
  });

  it("penalizes unsupported elements", () => {
    const score = scoreSlideQuality({
      sourceSlide: { unsupportedCount: 2, placeholderCount: 0 },
      actions: [],
      dropped: [],
      outputSlideCount: 1,
    });
    expect(score).toBeLessThan(1);
  });

  it("gives bonus for extracted tables and charts", () => {
    const score = scoreSlideQuality({
      sourceSlide: { unsupportedCount: 0, placeholderCount: 0, tableCount: 2, chartCount: 1 },
      actions: [],
      dropped: [],
      outputSlideCount: 1,
    });
    expect(score).toBe(1); // clamped to max 1 even with bonus
  });
});

// ── buildDefaultPolicy ──────────────────────────────────────────────────

describe("buildDefaultPolicy", () => {
  it("has sensible defaults", () => {
    const policy = buildDefaultPolicy();
    expect(policy.allowSlideSplit).toBe(true);
    expect(policy.allowElementDeletion).toBe(false);
    expect(policy.readabilityMinFontPt).toBe(12);
    expect(policy.reviewThreshold).toBe(0.78);
    expect(policy.strictReview).toBe(true);
  });

  it("overrides defaults", () => {
    const policy = buildDefaultPolicy({ readabilityMinFontPt: 8, allowSlideSplit: false });
    expect(policy.readabilityMinFontPt).toBe(8);
    expect(policy.allowSlideSplit).toBe(false);
  });
});
