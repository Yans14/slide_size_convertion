import { describe, it, expect, vi } from "vitest";
const {
  buildTextProps,
  normalizeColor,
} = require("../src/generate");

// ── normalizeColor ──────────────────────────────────────────────────────

describe("normalizeColor", () => {
  it("returns valid hex", () => expect(normalizeColor("FF3399")).toBe("FF3399"));
  it("strips hash prefix", () => expect(normalizeColor("#aabbcc")).toBe("AABBCC"));
  it("returns fallback for null", () => expect(normalizeColor(null, "000000")).toBe("000000"));
  it("returns fallback for invalid", () => expect(normalizeColor("xyz", "111111")).toBe("111111"));
});

// ── buildTextProps ──────────────────────────────────────────────────────

describe("buildTextProps", () => {
  it("builds TextProps from rich paragraphs", () => {
    const paragraphs = [
      {
        runs: [
          { text: "Hello ", bold: true, fontSizePt: 24, color: "FF0000" },
          { text: "World", italic: true, fontSizePt: 18 },
        ],
        align: "center",
      },
    ];
    const result = buildTextProps(paragraphs);
    expect(result).toHaveLength(2);

    expect(result[0].text).toBe("Hello ");
    expect(result[0].options.bold).toBe(true);
    expect(result[0].options.fontSize).toBe(24);
    expect(result[0].options.color).toBe("FF0000");
    expect(result[0].options.align).toBe("center");

    expect(result[1].text).toBe("World");
    expect(result[1].options.italic).toBe(true);
    expect(result[1].options.fontSize).toBe(18);
  });

  it("adds breakType between paragraphs", () => {
    const paragraphs = [
      { runs: [{ text: "Line 1" }] },
      { runs: [{ text: "Line 2" }] },
    ];
    const result = buildTextProps(paragraphs);
    expect(result).toHaveLength(2);
    expect(result[0].options.breakType).toBe("break");
    expect(result[1].options.breakType).toBeUndefined();
  });

  it("handles empty paragraphs as newlines", () => {
    const paragraphs = [
      { runs: [] },
    ];
    const result = buildTextProps(paragraphs);
    expect(result).toHaveLength(1);
    expect(result[0].text).toBe("\n");
  });

  it("handles underline", () => {
    const paragraphs = [
      { runs: [{ text: "Underlined", underline: true }] },
    ];
    const result = buildTextProps(paragraphs);
    expect(result[0].options.underline).toEqual({ style: "sng" });
  });

  it("handles font face", () => {
    const paragraphs = [
      { runs: [{ text: "Custom", fontFace: "Arial" }] },
    ];
    const result = buildTextProps(paragraphs);
    expect(result[0].options.fontFace).toBe("Arial");
  });

  it("handles bullet levels", () => {
    const paragraphs = [
      { runs: [{ text: "Bullet item" }], bulletLevel: 1 },
    ];
    const result = buildTextProps(paragraphs);
    expect(result[0].options.indentLevel).toBe(1);
  });
});

// ── Integration-style tests using mock slide ────────────────────────────

describe("renderElement dispatching", () => {
  // We test that the generator module exports renderElement
  const { renderElement, applySlideBackground } = require("../src/generate");

  function createMockSlide() {
    const calls = [];
    return {
      calls,
      addText: vi.fn((...args) => calls.push({ method: "addText", args })),
      addShape: vi.fn((...args) => calls.push({ method: "addShape", args })),
      addImage: vi.fn((...args) => calls.push({ method: "addImage", args })),
      addTable: vi.fn((...args) => calls.push({ method: "addTable", args })),
      addChart: vi.fn((...args) => calls.push({ method: "addChart", args })),
      background: undefined,
    };
  }

  const mockPptx = { ChartType: { bar: "bar", line: "line", pie: "pie" } };

  it("renders text element with rich paragraphs", () => {
    const slide = createMockSlide();
    const element = {
      type: "text", x: 1, y: 2, w: 5, h: 3,
      text: "Hello World",
      paragraphs: [{ runs: [{ text: "Hello ", bold: true }, { text: "World" }] }],
    };
    renderElement(slide, element, {}, mockPptx);
    expect(slide.addText).toHaveBeenCalledTimes(1);
    const args = slide.addText.mock.calls[0];
    expect(Array.isArray(args[0])).toBe(true); // TextProps[]
    expect(args[0][0].text).toBe("Hello ");
    expect(args[0][0].options.bold).toBe(true);
  });

  it("renders simple text fallback when no paragraphs", () => {
    const slide = createMockSlide();
    const element = {
      type: "text", x: 1, y: 2, w: 5, h: 3,
      text: "Simple", fontSizePt: 14, bold: false, color: "000000",
    };
    renderElement(slide, element, {}, mockPptx);
    expect(slide.addText).toHaveBeenCalledTimes(1);
    const args = slide.addText.mock.calls[0];
    expect(args[0]).toBe("Simple"); // plain string
  });

  it("renders image element", () => {
    const slide = createMockSlide();
    const element = {
      type: "image", x: 1, y: 2, w: 5, h: 3,
      dataUri: "data:image/png;base64,abc123",
    };
    renderElement(slide, element, {}, mockPptx);
    expect(slide.addImage).toHaveBeenCalledTimes(1);
  });

  it("renders shape element", () => {
    const slide = createMockSlide();
    const element = {
      type: "shape", x: 1, y: 2, w: 5, h: 3,
      shapeType: "roundRect", fillColor: "3366CC", lineColor: "000000", lineWidthPt: 1,
    };
    renderElement(slide, element, {}, mockPptx);
    expect(slide.addShape).toHaveBeenCalledTimes(1);
    expect(slide.addShape.mock.calls[0][0]).toBe("roundRect");
  });

  it("renders shape with text inside", () => {
    const slide = createMockSlide();
    const element = {
      type: "shape", x: 1, y: 2, w: 5, h: 3,
      shapeType: "rect", fillColor: null, lineColor: null, lineWidthPt: 0.75,
      shapeParagraphs: [{ runs: [{ text: "Label" }] }],
      shapeText: "Label",
    };
    renderElement(slide, element, {}, mockPptx);
    expect(slide.addShape).toHaveBeenCalledTimes(1);
    expect(slide.addText).toHaveBeenCalledTimes(1); // text inside shape
  });

  it("renders table element", () => {
    const slide = createMockSlide();
    const element = {
      type: "table", x: 1, y: 2, w: 8, h: 4,
      rows: [
        { cells: [{ text: "A", fontSizePt: 12 }, { text: "B", fontSizePt: 12 }] },
        { cells: [{ text: "C", fontSizePt: 12 }, { text: "D", fontSizePt: 12 }] },
      ],
      colWidths: [4, 4],
    };
    renderElement(slide, element, {}, mockPptx);
    expect(slide.addTable).toHaveBeenCalledTimes(1);
    const rows = slide.addTable.mock.calls[0][0];
    expect(rows).toHaveLength(2);
    expect(rows[0]).toHaveLength(2);
  });

  it("renders chart element", () => {
    const slide = createMockSlide();
    const element = {
      type: "chart", x: 1, y: 2, w: 8, h: 5,
      chartType: "column_clustered", name: "Revenue Chart",
      categories: ["Q1", "Q2"],
      series: [{ name: "Revenue", values: [100, 120] }],
      title: "Sales",
    };
    renderElement(slide, element, {}, mockPptx);
    expect(slide.addChart).toHaveBeenCalledTimes(1);
  });

  it("renders group by flattening children", () => {
    const slide = createMockSlide();
    const element = {
      type: "group", x: 0, y: 0, w: 10, h: 10,
      children: [
        { type: "text", x: 1, y: 1, w: 3, h: 2, text: "Child 1", paragraphs: [{ runs: [{ text: "Child 1" }] }] },
        { type: "shape", x: 5, y: 5, w: 2, h: 2, shapeType: "rect" },
      ],
    };
    renderElement(slide, element, {}, mockPptx);
    expect(slide.addText).toHaveBeenCalledTimes(1);
    expect(slide.addShape).toHaveBeenCalledTimes(1);
  });

  it("renders placeholder with dashed border when enabled", () => {
    const slide = createMockSlide();
    const element = {
      type: "placeholder", x: 1, y: 2, w: 5, h: 3,
      text: "[CHART] Chart1",
    };
    renderElement(slide, element, { renderPlaceholders: true }, mockPptx);
    expect(slide.addShape).toHaveBeenCalledTimes(1); // placeholder box
    expect(slide.addText).toHaveBeenCalledTimes(1); // text
  });

  it("applies solid background", () => {
    const slide = createMockSlide();
    applySlideBackground(slide, { type: "solid", color: "336699" });
    expect(slide.background).toEqual({ color: "336699" });
  });

  it("applies image background", () => {
    const slide = createMockSlide();
    applySlideBackground(slide, { type: "image", dataUri: "data:image/png;base64,abc" });
    expect(slide.background).toEqual({ data: "data:image/png;base64,abc" });
  });
});
