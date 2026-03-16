import { describe, it, expect } from "vitest";
const {
  extractRichParagraphs,
  flattenParagraphs,
  firstRunStyle,
  parseThemeColors,
  resolveSchemeColor,
  extractSolidFillColor,
  extractTransform,
  extractTable,
  extractChartDataFromXml,
  extractBackground,
  extractGroupTransform,
  extractShapeFallback,
  extractTextShape,
  mapPresetShape,
  normalizeColor,
  emuToInches,
  toNumber,
  asArray,
} = require("../src/pptx-reader");

// ── Utilities ───────────────────────────────────────────────────────────

describe("asArray", () => {
  it("wraps a single value in an array", () => expect(asArray("a")).toEqual(["a"]));
  it("returns an existing array as-is", () => expect(asArray([1, 2])).toEqual([1, 2]));
  it("returns [] for null", () => expect(asArray(null)).toEqual([]));
  it("returns [] for undefined", () => expect(asArray(undefined)).toEqual([]));
});

describe("toNumber", () => {
  it("parses a numeric string", () => expect(toNumber("42")).toBe(42));
  it("returns fallback for NaN", () => expect(toNumber("abc", 5)).toBe(5));
  it("returns fallback for undefined", () => expect(toNumber(undefined, 99)).toBe(99));
});

describe("emuToInches", () => {
  it("converts EMU to inches", () => expect(emuToInches(914400)).toBe(1));
  it("returns 0 for null", () => expect(emuToInches(null)).toBe(0));
});

describe("normalizeColor", () => {
  it("returns valid 6-char hex", () => expect(normalizeColor("FF3399")).toBe("FF3399"));
  it("strips # prefix", () => expect(normalizeColor("#aabbcc")).toBe("AABBCC"));
  it("returns null for invalid", () => expect(normalizeColor("blah")).toBe(null));
  it("returns null for null", () => expect(normalizeColor(null)).toBe(null));
});

// ── Theme Colors ────────────────────────────────────────────────────────

describe("parseThemeColors", () => {
  it("parses a theme document with srgb colors", () => {
    const doc = {
      "a:theme": {
        "a:themeElements": {
          "a:clrScheme": {
            "a:dk1": { "a:srgbClr": { "@_val": "000000" } },
            "a:lt1": { "a:srgbClr": { "@_val": "FFFFFF" } },
            "a:accent1": { "a:srgbClr": { "@_val": "4472C4" } },
          },
        },
      },
    };
    const map = parseThemeColors(doc);
    expect(map.get("dk1")).toBe("000000");
    expect(map.get("lt1")).toBe("FFFFFF");
    expect(map.get("accent1")).toBe("4472C4");
  });

  it("handles sysClr with lastClr", () => {
    const doc = {
      "a:theme": {
        "a:themeElements": {
          "a:clrScheme": {
            "a:dk1": { "a:sysClr": { "@_val": "windowText", "@_lastClr": "000000" } },
          },
        },
      },
    };
    const map = parseThemeColors(doc);
    expect(map.get("dk1")).toBe("000000");
  });

  it("returns empty map for missing theme", () => {
    expect(parseThemeColors({})).toEqual(new Map());
  });
});

describe("resolveSchemeColor", () => {
  it("resolves dk1", () => {
    const map = new Map([["dk1", "000000"]]);
    expect(resolveSchemeColor("dk1", map)).toBe("000000");
  });

  it("resolves tx1 as dk1", () => {
    const map = new Map([["dk1", "000000"]]);
    expect(resolveSchemeColor("tx1", map)).toBe("000000");
  });

  it("returns null for unknown scheme", () => {
    expect(resolveSchemeColor("unknown", new Map())).toBe(null);
  });
});

// ── extractSolidFillColor ───────────────────────────────────────────────

describe("extractSolidFillColor", () => {
  it("extracts srgb color", () => {
    const fill = { "a:srgbClr": { "@_val": "FF0000" } };
    expect(extractSolidFillColor(fill, new Map())).toBe("FF0000");
  });

  it("resolves scheme color with theme", () => {
    const fill = { "a:schemeClr": { "@_val": "accent1" } };
    const theme = new Map([["accent1", "4472C4"]]);
    expect(extractSolidFillColor(fill, theme)).toBe("4472C4");
  });

  it("returns null for missing fill", () => {
    expect(extractSolidFillColor(null, new Map())).toBe(null);
  });
});

// ── Transform ───────────────────────────────────────────────────────────

describe("extractTransform", () => {
  it("extracts position and size", () => {
    const xfrm = {
      "a:off": { "@_x": "914400", "@_y": "1828800" },
      "a:ext": { "@_cx": "4572000", "@_cy": "2743200" },
    };
    const result = extractTransform(xfrm);
    expect(result.x).toBeCloseTo(1);
    expect(result.y).toBeCloseTo(2);
    expect(result.w).toBeCloseTo(5);
    expect(result.h).toBeCloseTo(3);
  });

  it("returns null for missing xfrm", () => {
    expect(extractTransform(null)).toBe(null);
  });
});

// ── Rich Text ───────────────────────────────────────────────────────────

describe("extractRichParagraphs", () => {
  it("extracts multiple runs with formatting", () => {
    const txBody = {
      "a:p": [
        {
          "a:r": [
            { "a:t": "Hello ", "a:rPr": { "@_b": "1", "@_sz": "2400" } },
            { "a:t": "World", "a:rPr": { "@_i": "1", "@_sz": "1800" } },
          ],
          "a:pPr": { "@_algn": "ctr" },
        },
      ],
    };
    const result = extractRichParagraphs(txBody, new Map());
    expect(result).toHaveLength(1);
    expect(result[0].runs).toHaveLength(2);
    expect(result[0].runs[0].text).toBe("Hello ");
    expect(result[0].runs[0].bold).toBe(true);
    expect(result[0].runs[0].fontSizePt).toBe(24);
    expect(result[0].runs[1].text).toBe("World");
    expect(result[0].runs[1].italic).toBe(true);
    expect(result[0].runs[1].fontSizePt).toBe(18);
    expect(result[0].align).toBe("center");
  });

  it("handles empty paragraphs", () => {
    const txBody = { "a:p": { "a:r": [] } };
    const result = extractRichParagraphs(txBody, new Map());
    expect(result).toHaveLength(1);
    expect(result[0].runs).toHaveLength(0);
  });

  it("extracts fields", () => {
    const txBody = {
      "a:p": {
        "a:fld": { "a:t": "Field Value", "a:rPr": {} },
      },
    };
    const result = extractRichParagraphs(txBody, new Map());
    expect(result[0].runs[0].text).toBe("Field Value");
  });
});

describe("flattenParagraphs", () => {
  it("joins paragraph runs into text", () => {
    const paras = [
      { runs: [{ text: "Hello " }, { text: "World" }] },
      { runs: [{ text: "Line 2" }] },
    ];
    expect(flattenParagraphs(paras)).toBe("Hello World\nLine 2");
  });

  it("collapses excessive newlines", () => {
    const paras = [
      { runs: [{ text: "A" }] },
      { runs: [] },
      { runs: [] },
      { runs: [] },
      { runs: [{ text: "B" }] },
    ];
    expect(flattenParagraphs(paras)).toBe("A\n\nB");
  });
});

describe("firstRunStyle", () => {
  it("returns first non-empty run style", () => {
    const paras = [
      { runs: [{ text: "Hi", fontSizePt: 24, bold: true, color: "FF0000" }] },
    ];
    const style = firstRunStyle(paras);
    expect(style.fontSizePt).toBe(24);
    expect(style.bold).toBe(true);
    expect(style.color).toBe("FF0000");
  });

  it("returns defaults for empty paragraphs", () => {
    const style = firstRunStyle([{ runs: [] }]);
    expect(style.fontSizePt).toBe(18);
    expect(style.bold).toBe(false);
  });
});

// ── Shape Mapping ───────────────────────────────────────────────────────

describe("mapPresetShape", () => {
  it("returns known shapes", () => {
    expect(mapPresetShape("rect")).toBe("rect");
    expect(mapPresetShape("ellipse")).toBe("ellipse");
    expect(mapPresetShape("diamond")).toBe("diamond");
  });

  it("falls back to rect for unknown shapes", () => {
    expect(mapPresetShape("unknownShape")).toBe("rect");
  });

  it("handles null/undefined", () => {
    expect(mapPresetShape(null)).toBe("rect");
  });
});

// ── Text Shape ──────────────────────────────────────────────────────────

describe("extractTextShape", () => {
  it("extracts text shape with paragraphs", () => {
    const node = {
      "p:spPr": {
        "a:xfrm": {
          "a:off": { "@_x": "914400", "@_y": "914400" },
          "a:ext": { "@_cx": "4572000", "@_cy": "914400" },
        },
      },
      "p:txBody": {
        "a:p": {
          "a:r": { "a:t": "Test text", "a:rPr": { "@_sz": "2000", "@_b": "1" } },
        },
      },
      "p:nvSpPr": { "p:cNvPr": { "@_id": "5", "@_name": "TextBox 1" } },
    };
    const result = extractTextShape(node, new Map());
    expect(result.type).toBe("text");
    expect(result.text).toBe("Test text");
    expect(result.paragraphs).toHaveLength(1);
    expect(result.paragraphs[0].runs[0].bold).toBe(true);
    expect(result.name).toBe("TextBox 1");
  });

  it("returns null for shapes without text", () => {
    const node = {
      "p:spPr": {
        "a:xfrm": {
          "a:off": { "@_x": "0", "@_y": "0" },
          "a:ext": { "@_cx": "1000", "@_cy": "1000" },
        },
      },
      "p:txBody": { "a:p": {} },
    };
    expect(extractTextShape(node, new Map())).toBe(null);
  });
});

// ── Shape Fallback with Text ────────────────────────────────────────────

describe("extractShapeFallback", () => {
  it("extracts shape with text inside", () => {
    const node = {
      "p:spPr": {
        "a:xfrm": {
          "a:off": { "@_x": "914400", "@_y": "914400" },
          "a:ext": { "@_cx": "4572000", "@_cy": "914400" },
        },
        "a:prstGeom": { "@_prst": "roundRect" },
        "a:solidFill": { "a:srgbClr": { "@_val": "3366CC" } },
      },
      "p:txBody": {
        "a:p": { "a:r": { "a:t": "Button Label", "a:rPr": {} } },
      },
      "p:nvSpPr": { "p:cNvPr": { "@_id": "10", "@_name": "Btn1" } },
    };
    const result = extractShapeFallback(node, new Map());
    expect(result.type).toBe("shape");
    expect(result.shapeType).toBe("roundRect");
    expect(result.fillColor).toBe("3366CC");
    expect(result.shapeText).toBe("Button Label");
    expect(result.shapeParagraphs).toHaveLength(1);
  });

  it("extracts shape without text", () => {
    const node = {
      "p:spPr": {
        "a:xfrm": {
          "a:off": { "@_x": "0", "@_y": "0" },
          "a:ext": { "@_cx": "4572000", "@_cy": "914400" },
        },
        "a:prstGeom": { "@_prst": "rect" },
      },
      "p:nvSpPr": { "p:cNvPr": { "@_id": "11" } },
    };
    const result = extractShapeFallback(node, new Map());
    expect(result.type).toBe("shape");
    expect(result.shapeType).toBe("rect");
    expect(result.shapeParagraphs).toBeUndefined();
  });
});

// ── Table Extraction ────────────────────────────────────────────────────

describe("extractTable", () => {
  it("extracts structured table rows and cells", () => {
    const tableNode = {
      "a:tblGrid": {
        "a:gridCol": [{ "@_w": "1828800" }, { "@_w": "1828800" }],
      },
      "a:tr": [
        {
          "@_h": "457200",
          "a:tc": [
            { "a:txBody": { "a:p": { "a:r": { "a:t": "Name", "a:rPr": { "@_b": "1" } } } } },
            { "a:txBody": { "a:p": { "a:r": { "a:t": "Value", "a:rPr": {} } } } },
          ],
        },
        {
          "@_h": "457200",
          "a:tc": [
            { "a:txBody": { "a:p": { "a:r": { "a:t": "Foo", "a:rPr": {} } } } },
            { "a:txBody": { "a:p": { "a:r": { "a:t": "Bar", "a:rPr": {} } } } },
          ],
        },
      ],
    };
    const geo = { x: 0, y: 0, w: 4, h: 2 };
    const nvProps = { "@_id": "1", "@_name": "Table1" };

    const result = extractTable(tableNode, geo, nvProps, new Map());
    expect(result.type).toBe("table");
    expect(result.rows).toHaveLength(2);
    expect(result.rows[0].cells).toHaveLength(2);
    expect(result.rows[0].cells[0].text).toBe("Name");
    expect(result.rows[0].cells[0].bold).toBe(true);
    expect(result.rows[1].cells[1].text).toBe("Bar");
    expect(result.colWidths).toHaveLength(2);
  });
});

// ── Chart Data Extraction ───────────────────────────────────────────────

describe("extractChartDataFromXml", () => {
  it("extracts bar chart data", () => {
    const doc = {
      "c:chartSpace": {
        "c:chart": {
          "c:plotArea": {
            "c:barChart": {
              "c:barDir": { "@_val": "col" },
              "c:grouping": { "@_val": "clustered" },
              "c:ser": [
                {
                  "c:tx": { "c:strRef": { "c:strCache": { "c:pt": { "c:v": "Revenue" } } } },
                  "c:cat": {
                    "c:strRef": {
                      "c:strCache": {
                        "c:pt": [
                          { "@_idx": "0", "c:v": "Q1" },
                          { "@_idx": "1", "c:v": "Q2" },
                        ],
                      },
                    },
                  },
                  "c:val": {
                    "c:numRef": {
                      "c:numCache": {
                        "c:pt": [
                          { "@_idx": "0", "c:v": "100" },
                          { "@_idx": "1", "c:v": "120" },
                        ],
                      },
                    },
                  },
                },
              ],
            },
          },
          "c:title": {
            "c:tx": {
              "c:rich": {
                "a:p": { "a:r": { "a:t": "Revenue by Quarter" } },
              },
            },
          },
        },
      },
    };
    const result = extractChartDataFromXml(doc);
    expect(result.chartType).toBe("column_clustered");
    expect(result.categories).toEqual(["Q1", "Q2"]);
    expect(result.series).toHaveLength(1);
    expect(result.series[0].name).toBe("Revenue");
    expect(result.series[0].values).toEqual([100, 120]);
    expect(result.title).toBe("Revenue by Quarter");
  });

  it("handles pie chart", () => {
    const doc = {
      "c:chartSpace": {
        "c:chart": {
          "c:plotArea": {
            "c:pieChart": {
              "c:ser": [{
                "c:tx": { "c:v": "Market Share" },
                "c:cat": { "c:strRef": { "c:strCache": { "c:pt": [{ "@_idx": "0", "c:v": "A" }] } } },
                "c:val": { "c:numRef": { "c:numCache": { "c:pt": [{ "@_idx": "0", "c:v": "45" }] } } },
              }],
            },
          },
        },
      },
    };
    const result = extractChartDataFromXml(doc);
    expect(result.chartType).toBe("pie");
    expect(result.series[0].name).toBe("Market Share");
  });

  it("returns unknown for empty chart", () => {
    const doc = { "c:chartSpace": { "c:chart": { "c:plotArea": {} } } };
    const result = extractChartDataFromXml(doc);
    expect(result.chartType).toBe("unknown");
  });
});

// ── Background Extraction ───────────────────────────────────────────────

describe("extractBackground", () => {
  it("extracts solid fill background", async () => {
    const bgNode = {
      "p:bgPr": {
        "a:solidFill": { "a:srgbClr": { "@_val": "336699" } },
      },
    };
    const result = await extractBackground(bgNode, new Map(), {}, new Map());
    expect(result).toEqual({ type: "solid", color: "336699" });
  });

  it("extracts gradient background", async () => {
    const bgNode = {
      "p:bgPr": {
        "a:gradFill": {
          "a:gsLst": {
            "a:gs": [
              { "@_pos": "0", "a:srgbClr": { "@_val": "000000" } },
              { "@_pos": "100000", "a:srgbClr": { "@_val": "FFFFFF" } },
            ],
          },
          "a:lin": { "@_ang": "5400000" },
        },
      },
    };
    const result = await extractBackground(bgNode, new Map(), {}, new Map());
    expect(result.type).toBe("gradient");
    expect(result.stops).toHaveLength(2);
    expect(result.angle).toBe(90);
  });

  it("returns null for missing bg", async () => {
    expect(await extractBackground(null, new Map(), {}, new Map())).toBe(null);
  });
});

// ── Group Transform ─────────────────────────────────────────────────────

describe("extractGroupTransform", () => {
  it("extracts group and child offsets", () => {
    const grpSpPr = {
      "a:xfrm": {
        "a:off": { "@_x": "914400", "@_y": "914400" },
        "a:ext": { "@_cx": "4572000", "@_cy": "2743200" },
        "a:chOff": { "@_x": "0", "@_y": "0" },
        "a:chExt": { "@_cx": "4572000", "@_cy": "2743200" },
      },
    };
    const { offset, childOffset, childExt } = extractGroupTransform(grpSpPr);
    expect(offset.x).toBeCloseTo(1);
    expect(offset.y).toBeCloseTo(1);
    expect(childOffset.x).toBe(0);
    expect(childExt.cx).toBe(4572000);
  });

  it("handles missing transform", () => {
    const { offset } = extractGroupTransform({});
    expect(offset).toBe(null);
  });
});
