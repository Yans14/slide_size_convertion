const fs = require("fs/promises");
const path = require("path");
const JSZip = require("jszip");
const { XMLParser } = require("fast-xml-parser");

const EMU_PER_INCH = 914400;

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

// ── Utilities ───────────────────────────────────────────────────────────

function asArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function toNumber(value, fallback = 0) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function emuToInches(value) {
  return toNumber(value, 0) / EMU_PER_INCH;
}

function parseXml(xmlBuffer, label) {
  try {
    return xmlParser.parse(xmlBuffer.toString("utf8"));
  } catch (error) {
    throw new Error(`Failed to parse XML for ${label}: ${error.message}`);
  }
}

function normalizeZipPath(baseFile, target) {
  const baseDir = path.posix.dirname(baseFile);
  return path.posix.normalize(path.posix.join(baseDir, target));
}

function normalizeColor(color) {
  if (!color) return null;
  const upper = String(color).toUpperCase().replace(/^#/, "");
  if (/^[0-9A-F]{6}$/.test(upper)) return upper;
  return null;
}

// ── Theme Color Resolution ──────────────────────────────────────────────

const SCHEME_MAP = {
  dk1: "dk1", dk2: "dk2", lt1: "lt1", lt2: "lt2",
  accent1: "accent1", accent2: "accent2", accent3: "accent3",
  accent4: "accent4", accent5: "accent5", accent6: "accent6",
  hlink: "hlink", folHlink: "folHlink",
  tx1: "dk1", tx2: "dk2", bg1: "lt1", bg2: "lt2",
};

function parseThemeColors(themeDoc) {
  const colorMap = new Map();
  const clrScheme =
    themeDoc?.["a:theme"]?.["a:themeElements"]?.["a:clrScheme"];
  if (!clrScheme) return colorMap;

  const slotNames = [
    "a:dk1", "a:dk2", "a:lt1", "a:lt2",
    "a:accent1", "a:accent2", "a:accent3", "a:accent4",
    "a:accent5", "a:accent6", "a:hlink", "a:folHlink",
  ];

  for (const slotName of slotNames) {
    const slot = clrScheme[slotName];
    if (!slot) continue;
    const hex =
      slot["a:srgbClr"]?.["@_val"] ||
      slot["a:sysClr"]?.["@_lastClr"] ||
      slot["a:sysClr"]?.["@_val"] ||
      null;
    const key = slotName.replace("a:", "");
    const normalized = normalizeColor(hex);
    if (normalized) colorMap.set(key, normalized);
  }

  return colorMap;
}

function resolveSchemeColor(schemeVal, themeColors) {
  if (!schemeVal || !themeColors || themeColors.size === 0) return null;
  const mapped = SCHEME_MAP[schemeVal] || schemeVal;
  return themeColors.get(mapped) || null;
}

// ── Color Extraction ────────────────────────────────────────────────────

function extractSolidFillColor(fillNode, themeColors) {
  if (!fillNode) return null;
  const srgb = fillNode["a:srgbClr"]?.["@_val"];
  if (srgb) return normalizeColor(srgb);
  const scheme = fillNode["a:schemeClr"]?.["@_val"];
  if (scheme) return resolveSchemeColor(scheme, themeColors);
  return null;
}

function extractColorFromNode(node, themeColors) {
  if (!node) return null;
  const solidFill = node["a:solidFill"];
  if (solidFill) return extractSolidFillColor(solidFill, themeColors);
  return null;
}

// ── Relationships ───────────────────────────────────────────────────────

function parseRelationships(xmlDoc, baseFile) {
  const relRoot = xmlDoc?.Relationships;
  const rels = asArray(relRoot?.Relationship);
  const map = new Map();

  for (const rel of rels) {
    const id = rel?.["@_Id"];
    const target = rel?.["@_Target"];
    if (!id || !target) continue;
    map.set(id, normalizeZipPath(baseFile, target));
  }

  return map;
}

// ── Transform Extraction ────────────────────────────────────────────────

function extractTransform(xfrm) {
  if (!xfrm) return null;
  const off = xfrm["a:off"];
  const ext = xfrm["a:ext"];
  if (!off || !ext) return null;

  return {
    x: emuToInches(off["@_x"]),
    y: emuToInches(off["@_y"]),
    w: emuToInches(ext["@_cx"]),
    h: emuToInches(ext["@_cy"]),
  };
}

function extractGroupTransform(grpSpPr) {
  const xfrm = grpSpPr?.["a:xfrm"];
  if (!xfrm) return { offset: null, childOffset: null, childExt: null };

  const off = xfrm["a:off"];
  const ext = xfrm["a:ext"];
  const chOff = xfrm["a:chOff"];
  const chExt = xfrm["a:chExt"];

  return {
    offset: off && ext
      ? { x: emuToInches(off["@_x"]), y: emuToInches(off["@_y"]), w: emuToInches(ext["@_cx"]), h: emuToInches(ext["@_cy"]) }
      : null,
    childOffset: chOff
      ? { x: toNumber(chOff["@_x"], 0), y: toNumber(chOff["@_y"], 0) }
      : null,
    childExt: chExt
      ? { cx: toNumber(chExt["@_cx"], 0), cy: toNumber(chExt["@_cy"], 0) }
      : null,
  };
}

// ── Rich Text Extraction ────────────────────────────────────────────────

function extractRunFormatting(rPr, themeColors) {
  if (!rPr) return {};
  const fontSizePt = rPr["@_sz"] ? toNumber(rPr["@_sz"]) / 100 : undefined;
  const bold = rPr["@_b"] === "1" ? true : undefined;
  const italic = rPr["@_i"] === "1" ? true : undefined;
  const underline = rPr["@_u"] && rPr["@_u"] !== "none" ? true : undefined;
  const strikethrough = rPr["@_strike"] && rPr["@_strike"] !== "noStrike" ? true : undefined;

  const color = extractSolidFillColor(rPr["a:solidFill"], themeColors);

  const latin = rPr["a:latin"];
  const fontFace = latin?.["@_typeface"] || undefined;

  return {
    ...(fontSizePt !== undefined && { fontSizePt }),
    ...(bold !== undefined && { bold }),
    ...(italic !== undefined && { italic }),
    ...(underline !== undefined && { underline }),
    ...(strikethrough !== undefined && { strikethrough }),
    ...(color && { color }),
    ...(fontFace && { fontFace }),
  };
}

function extractParagraphAlignment(pPr) {
  if (!pPr) return undefined;
  const algn = pPr["@_algn"];
  if (algn === "ctr") return "center";
  if (algn === "r") return "right";
  if (algn === "just") return "justify";
  if (algn === "l") return "left";
  return undefined;
}

function extractRichParagraphs(txBody, themeColors) {
  const paragraphs = asArray(txBody?.["a:p"]);
  const result = [];

  for (const paragraph of paragraphs) {
    const pPr = paragraph?.["a:pPr"];
    const align = extractParagraphAlignment(pPr);
    const bulletLevel = pPr?.["@_lvl"] ? toNumber(pPr["@_lvl"]) : undefined;

    const runs = [];

    for (const run of asArray(paragraph?.["a:r"])) {
      const text = run?.["a:t"];
      if (typeof text !== "string") continue;
      const fmt = extractRunFormatting(run?.["a:rPr"], themeColors);
      runs.push({ text, ...fmt });
    }

    for (const fld of asArray(paragraph?.["a:fld"])) {
      const text = fld?.["a:t"];
      if (typeof text !== "string") continue;
      const fmt = extractRunFormatting(fld?.["a:rPr"], themeColors);
      runs.push({ text, ...fmt });
    }

    if (paragraph?.["a:br"]) {
      runs.push({ text: "\n" });
    }

    result.push({
      runs,
      ...(align && { align }),
      ...(bulletLevel !== undefined && { bulletLevel }),
    });
  }

  return result;
}

function flattenParagraphs(paragraphs) {
  const lines = [];
  for (const para of paragraphs) {
    const line = para.runs.map((r) => r.text).join("");
    lines.push(line);
  }
  return lines.join("\n").replace(/\n{3,}/g, "\n\n").trim();
}

function firstRunStyle(paragraphs) {
  for (const para of paragraphs) {
    for (const run of para.runs) {
      if (run.text && run.text.trim()) {
        return {
          fontSizePt: run.fontSizePt || 18,
          bold: run.bold || false,
          color: run.color || null,
        };
      }
    }
  }
  return { fontSizePt: 18, bold: false, color: null };
}

// ── MIME Type ───────────────────────────────────────────────────────────

function getMimeTypeFromPath(mediaPath) {
  const ext = path.extname(mediaPath).toLowerCase();
  if (ext === ".png") return "image/png";
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".gif") return "image/gif";
  if (ext === ".bmp") return "image/bmp";
  if (ext === ".svg") return "image/svg+xml";
  if (ext === ".emf") return "image/x-emf";
  if (ext === ".wmf") return "image/x-wmf";
  return "application/octet-stream";
}

// ── Text Shape Extraction ───────────────────────────────────────────────

function extractTextShape(shapeNode, themeColors) {
  const geometry = extractTransform(shapeNode?.["p:spPr"]?.["a:xfrm"]);
  if (!geometry || geometry.w <= 0 || geometry.h <= 0) return null;

  const txBody = shapeNode?.["p:txBody"];
  const paragraphs = extractRichParagraphs(txBody, themeColors);
  const text = flattenParagraphs(paragraphs);
  if (!text) return null;

  const { fontSizePt, bold, color } = firstRunStyle(paragraphs);
  const nvProps = shapeNode?.["p:nvSpPr"]?.["p:cNvPr"];

  return {
    type: "text",
    id: toNumber(nvProps?.["@_id"], 0),
    name: nvProps?.["@_name"] || "Text",
    ...geometry,
    text,
    fontSizePt,
    bold,
    color,
    paragraphs,
  };
}

// ── Shape Preset Mapping ────────────────────────────────────────────────

function mapPresetShape(preset) {
  const value = String(preset || "rect");
  const supported = new Set([
    "rect", "roundRect", "ellipse", "triangle", "rtTriangle",
    "diamond", "hexagon", "parallelogram", "trapezoid", "cloud",
    "star5", "star6", "heart", "lightningBolt", "pentagon",
    "octagon", "plus", "arrow", "chevron", "homePlate",
  ]);
  if (supported.has(value)) return value;
  return "rect";
}

// ── Shape Fallback Extraction (with text) ───────────────────────────────

function extractShapeFallback(shapeNode, themeColors) {
  const geometry = extractTransform(shapeNode?.["p:spPr"]?.["a:xfrm"]);
  if (!geometry || geometry.w <= 0 || geometry.h <= 0) return null;

  const nvProps = shapeNode?.["p:nvSpPr"]?.["p:cNvPr"];
  const shapeProps = shapeNode?.["p:spPr"] || {};
  const preset = shapeProps?.["a:prstGeom"]?.["@_prst"] || "rect";

  const hasNoFill = Boolean(shapeProps?.["a:noFill"]);
  const fillColor = hasNoFill ? null : extractSolidFillColor(shapeProps?.["a:solidFill"], themeColors);
  const lineColor = extractSolidFillColor(shapeProps?.["a:ln"]?.["a:solidFill"], themeColors);
  const lineWidthPt = toNumber(shapeProps?.["a:ln"]?.["@_w"], 9525) / 12700;

  // Extract text inside non-text shapes
  const txBody = shapeNode?.["p:txBody"];
  let shapeParagraphs = null;
  let shapeText = null;
  if (txBody) {
    const paras = extractRichParagraphs(txBody, themeColors);
    const flat = flattenParagraphs(paras);
    if (flat) {
      shapeParagraphs = paras;
      shapeText = flat;
    }
  }

  return {
    type: "shape",
    id: toNumber(nvProps?.["@_id"], 0),
    name: nvProps?.["@_name"] || "Shape",
    ...geometry,
    shapeType: mapPresetShape(preset),
    fillColor,
    lineColor,
    lineWidthPt: Number.isFinite(lineWidthPt) ? lineWidthPt : 0.75,
    ...(shapeParagraphs && { shapeParagraphs, shapeText }),
  };
}

// ── Picture Extraction ──────────────────────────────────────────────────

async function extractPicture(picNode, relsMap, zip) {
  const geometry = extractTransform(picNode?.["p:spPr"]?.["a:xfrm"]);
  if (!geometry || geometry.w <= 0 || geometry.h <= 0) return null;

  const embedId = picNode?.["p:blipFill"]?.["a:blip"]?.["@_r:embed"];
  const mediaPath = embedId ? relsMap.get(embedId) : null;

  let dataUri = null;
  let sourcePath = null;

  if (mediaPath) {
    const mediaFile = zip.file(mediaPath);
    if (mediaFile) {
      const base64 = await mediaFile.async("base64");
      const mimeType = getMimeTypeFromPath(mediaPath);
      dataUri = `data:${mimeType};base64,${base64}`;
      sourcePath = mediaPath;
    }
  }

  const nvProps = picNode?.["p:nvPicPr"]?.["p:cNvPr"];

  return {
    type: "image",
    id: toNumber(nvProps?.["@_id"], 0),
    name: nvProps?.["@_name"] || "Image",
    ...geometry,
    dataUri,
    sourcePath,
  };
}

// ── Group Shape Extraction ──────────────────────────────────────────────

async function extractGroupShape(grpSpNode, relsMap, zip, themeColors) {
  const grpSpPr = grpSpNode?.["p:grpSpPr"];
  const { offset, childOffset, childExt } = extractGroupTransform(grpSpPr);
  if (!offset || offset.w <= 0 || offset.h <= 0) return null;

  const nvProps = grpSpNode?.["p:nvGrpSpPr"]?.["p:cNvPr"];

  // Calculate scale factors for child → parent coordinate mapping
  let scaleX = 1;
  let scaleY = 1;
  if (childExt && childExt.cx > 0 && childExt.cy > 0) {
    scaleX = (offset.w * EMU_PER_INCH) / childExt.cx;
    scaleY = (offset.h * EMU_PER_INCH) / childExt.cy;
  }

  const chOffX = childOffset ? childOffset.x / EMU_PER_INCH : 0;
  const chOffY = childOffset ? childOffset.y / EMU_PER_INCH : 0;

  const children = [];

  // Extract child shapes
  for (const sp of asArray(grpSpNode?.["p:sp"])) {
    const shape = extractTextShape(sp, themeColors);
    if (shape) {
      applyGroupTransform(shape, offset, chOffX, chOffY, scaleX, scaleY);
      children.push(shape);
    } else {
      const fallback = extractShapeFallback(sp, themeColors);
      if (fallback) {
        applyGroupTransform(fallback, offset, chOffX, chOffY, scaleX, scaleY);
        children.push(fallback);
      }
    }
  }

  // Extract child pictures
  for (const pic of asArray(grpSpNode?.["p:pic"])) {
    const picture = await extractPicture(pic, relsMap, zip);
    if (picture) {
      applyGroupTransform(picture, offset, chOffX, chOffY, scaleX, scaleY);
      children.push(picture);
    }
  }

  // Recurse into nested groups
  for (const nestedGrp of asArray(grpSpNode?.["p:grpSp"])) {
    const nested = await extractGroupShape(nestedGrp, relsMap, zip, themeColors);
    if (nested) {
      // Flatten nested group children into this group
      for (const child of nested.children) {
        children.push(child);
      }
    }
  }

  // Extract child graphic frames (tables etc.)
  for (const frame of asArray(grpSpNode?.["p:graphicFrame"])) {
    const extracted = extractGraphicFrame(frame, relsMap, themeColors, zip);
    if (extracted) {
      applyGroupTransform(extracted, offset, chOffX, chOffY, scaleX, scaleY);
      children.push(extracted);
    }
  }

  return {
    type: "group",
    id: toNumber(nvProps?.["@_id"], 0),
    name: nvProps?.["@_name"] || "Group",
    ...offset,
    children,
  };
}

function applyGroupTransform(element, groupOffset, chOffX, chOffY, scaleX, scaleY) {
  element.x = (element.x - chOffX) * scaleX + groupOffset.x;
  element.y = (element.y - chOffY) * scaleY + groupOffset.y;
  element.w = element.w * scaleX;
  element.h = element.h * scaleY;
}

// ── Table Extraction ────────────────────────────────────────────────────

function extractTable(tableNode, geometry, nvProps, themeColors) {
  const tblGrid = asArray(tableNode?.["a:tblGrid"]?.["a:gridCol"]);
  const colWidths = tblGrid.map((col) => emuToInches(col?.["@_w"]));

  const aRows = asArray(tableNode?.["a:tr"]);
  const rows = [];

  for (const aRow of aRows) {
    const rowHeight = emuToInches(aRow?.["@_h"]);
    const cells = [];

    for (const aCell of asArray(aRow?.["a:tc"])) {
      const paragraphs = extractRichParagraphs(aCell?.["a:txBody"], themeColors);
      const text = flattenParagraphs(paragraphs);
      const { fontSizePt, bold, color } = firstRunStyle(paragraphs);

      const tcPr = aCell?.["a:tcPr"];
      const cellFill = extractSolidFillColor(tcPr?.["a:solidFill"], themeColors);

      const colspan = toNumber(aCell?.["@_gridSpan"], 1);
      const rowspan = toNumber(aCell?.["@_rowSpan"], 1);
      const hMerge = aCell?.["@_hMerge"] === "1";
      const vMerge = aCell?.["@_vMerge"] === "1";

      cells.push({
        text,
        paragraphs,
        fontSizePt,
        bold,
        color,
        fillColor: cellFill,
        ...(colspan > 1 && { colspan }),
        ...(rowspan > 1 && { rowspan }),
        ...(hMerge && { hMerge }),
        ...(vMerge && { vMerge }),
      });
    }

    rows.push({ cells, height: rowHeight });
  }

  return {
    type: "table",
    id: toNumber(nvProps?.["@_id"], 0),
    name: nvProps?.["@_name"] || "Table",
    ...geometry,
    rows,
    colWidths,
  };
}

// ── Chart Data Extraction ───────────────────────────────────────────────

const CHART_TYPE_MAP = {
  "c:barChart": "bar",
  "c:bar3DChart": "bar3D",
  "c:lineChart": "line",
  "c:line3DChart": "line3D",
  "c:pieChart": "pie",
  "c:pie3DChart": "pie3D",
  "c:doughnutChart": "doughnut",
  "c:areaChart": "area",
  "c:area3DChart": "area3D",
  "c:scatterChart": "scatter",
  "c:bubbleChart": "bubble",
  "c:radarChart": "radar",
  "c:stockChart": "stock",
  "c:surfaceChart": "surface",
  "c:surface3DChart": "surface3D",
};

function extractChartDataFromXml(chartDoc) {
  const chartSpace = chartDoc?.["c:chartSpace"];
  const chart = chartSpace?.["c:chart"];
  if (!chart) return null;

  const plotArea = chart["c:plotArea"];
  if (!plotArea) return null;

  // Detect chart type
  let chartType = "unknown";
  let chartNode = null;

  for (const [xmlTag, typeName] of Object.entries(CHART_TYPE_MAP)) {
    if (plotArea[xmlTag]) {
      chartType = typeName;
      chartNode = plotArea[xmlTag];
      break;
    }
  }

  if (!chartNode) return { chartType: "unknown", categories: [], series: [] };

  // Determine bar grouping/direction for more specific type
  if (chartType === "bar") {
    const dir = chartNode["c:barDir"]?.["@_val"];
    const grouping = chartNode["c:grouping"]?.["@_val"];
    if (dir === "col") {
      if (grouping === "stacked") chartType = "column_stacked";
      else if (grouping === "percentStacked") chartType = "column_stacked_100";
      else chartType = "column_clustered";
    } else {
      if (grouping === "stacked") chartType = "bar_stacked";
      else chartType = "bar_clustered";
    }
  }

  // Extract series
  const serNodes = asArray(chartNode["c:ser"]);
  const categories = [];
  const series = [];
  let categoriesExtracted = false;

  for (const ser of serNodes) {
    const name =
      ser["c:tx"]?.["c:strRef"]?.["c:strCache"]?.["c:pt"]?.["c:v"] ||
      ser["c:tx"]?.["c:v"] ||
      `Series ${series.length + 1}`;

    // Extract categories from the first series
    if (!categoriesExtracted) {
      const catRef = ser["c:cat"];
      if (catRef) {
        const catCache =
          catRef["c:strRef"]?.["c:strCache"] ||
          catRef["c:numRef"]?.["c:numCache"];
        if (catCache) {
          for (const pt of asArray(catCache["c:pt"])) {
            const idx = toNumber(pt?.["@_idx"], categories.length);
            categories[idx] = pt?.["c:v"] || "";
          }
        }
      }
      categoriesExtracted = true;
    }

    // Extract values
    const values = [];
    const valRef = ser["c:val"] || ser["c:yVal"];
    if (valRef) {
      const valCache = valRef["c:numRef"]?.["c:numCache"];
      if (valCache) {
        for (const pt of asArray(valCache["c:pt"])) {
          const idx = toNumber(pt?.["@_idx"], values.length);
          values[idx] = toNumber(pt?.["c:v"], 0);
        }
      }
    }

    series.push({ name: String(name), values });
  }

  // Title
  const titleNode = chart["c:title"];
  let title = null;
  if (titleNode) {
    const txBody = titleNode["c:tx"]?.["c:rich"];
    if (txBody) {
      const paras = asArray(txBody["a:p"]);
      const titleParts = [];
      for (const p of paras) {
        for (const r of asArray(p["a:r"])) {
          if (typeof r["a:t"] === "string") titleParts.push(r["a:t"]);
        }
      }
      title = titleParts.join("").trim() || null;
    }
  }

  return { chartType, categories, series, title };
}

async function extractChart(frameNode, relsMap, zip, geometry, nvProps) {
  const graphicData = frameNode?.["a:graphic"]?.["a:graphicData"];
  const chartRef = graphicData?.["c:chart"];
  if (!chartRef) return null;

  const embedId = chartRef["@_r:id"];
  const chartPath = embedId ? relsMap.get(embedId) : null;
  if (!chartPath) return null;

  const chartFile = zip.file(chartPath);
  if (!chartFile) return null;

  try {
    const chartXml = await chartFile.async("nodebuffer");
    const chartDoc = parseXml(chartXml, chartPath);
    const chartData = extractChartDataFromXml(chartDoc);
    if (!chartData) return null;

    return {
      type: "chart",
      id: toNumber(nvProps?.["@_id"], 0),
      name: nvProps?.["@_name"] || "Chart",
      ...geometry,
      ...chartData,
    };
  } catch {
    return null;
  }
}

// ── Graphic Frame Extraction ────────────────────────────────────────────

function extractGraphicFrameTransform(frameNode) {
  return extractTransform(frameNode?.["p:xfrm"]);
}

function determineGraphicKind(uri) {
  if (!uri) return "unknown";
  if (uri.includes("/table")) return "table";
  if (uri.includes("/chart")) return "chart";
  if (uri.includes("/diagram")) return "diagram";
  if (uri.includes("/oleObject")) return "ole";
  return "unknown";
}

function extractGraphicFrame(frameNode, relsMap, themeColors, zip) {
  const geometry = extractGraphicFrameTransform(frameNode);
  if (!geometry || geometry.w <= 0 || geometry.h <= 0) return null;

  const nvProps = frameNode?.["p:nvGraphicFramePr"]?.["p:cNvPr"];
  const name = nvProps?.["@_name"] || "GraphicFrame";
  const graphicData = frameNode?.["a:graphic"]?.["a:graphicData"];
  const uri = graphicData?.["@_uri"] || "";
  const kind = determineGraphicKind(uri);

  if (kind === "table") {
    const tableNode = graphicData?.["a:tbl"];
    if (tableNode) {
      return extractTable(tableNode, geometry, nvProps, themeColors);
    }
    return {
      type: "placeholder",
      placeholderKind: "table",
      id: toNumber(nvProps?.["@_id"], 0),
      name,
      ...geometry,
      text: `[Table] ${name}`,
    };
  }

  // Charts will be handled async in the main loop
  if (kind === "chart") {
    return {
      type: "_chart_pending",
      id: toNumber(nvProps?.["@_id"], 0),
      name,
      ...geometry,
      _frameNode: frameNode,
    };
  }

  const chartEmbedId = graphicData?.["c:chart"]?.["@_r:id"];
  const linkedPath = chartEmbedId ? relsMap.get(chartEmbedId) : null;

  return {
    type: "placeholder",
    placeholderKind: kind,
    id: toNumber(nvProps?.["@_id"], 0),
    name,
    ...geometry,
    text: `[${kind.toUpperCase()}] ${name}`,
    linkedPath: linkedPath || null,
  };
}

// ── Background Extraction ───────────────────────────────────────────────

async function extractBackground(bgNode, relsMap, zip, themeColors) {
  if (!bgNode) return null;

  const bgPr = bgNode["p:bgPr"];
  if (bgPr) {
    // Solid fill
    const solidFill = bgPr["a:solidFill"];
    if (solidFill) {
      const color = extractSolidFillColor(solidFill, themeColors);
      if (color) return { type: "solid", color };
    }

    // Gradient fill
    const gradFill = bgPr["a:gradFill"];
    if (gradFill) {
      const gsLst = asArray(gradFill["a:gsLst"]?.["a:gs"]);
      const stops = [];
      for (const gs of gsLst) {
        const pos = toNumber(gs?.["@_pos"], 0) / 100000;
        const color = extractSolidFillColor(gs, themeColors);
        if (color) stops.push({ position: pos, color });
      }
      if (stops.length > 0) {
        const angle = toNumber(gradFill["a:lin"]?.["@_ang"], 0) / 60000;
        return { type: "gradient", angle, stops };
      }
    }

    // Image fill
    const blipFill = bgPr["a:blipFill"];
    if (blipFill) {
      const embedId = blipFill["a:blip"]?.["@_r:embed"];
      const mediaPath = embedId ? relsMap.get(embedId) : null;
      if (mediaPath) {
        const mediaFile = zip.file(mediaPath);
        if (mediaFile) {
          const base64 = await mediaFile.async("base64");
          const mimeType = getMimeTypeFromPath(mediaPath);
          return { type: "image", dataUri: `data:${mimeType};base64,${base64}` };
        }
      }
    }
  }

  // bgRef (index-based, often references theme)
  const bgRef = bgNode["p:bgRef"];
  if (bgRef) {
    const color = extractSolidFillColor(bgRef, themeColors);
    if (color) return { type: "solid", color };
  }

  return null;
}

// ── Source Layout Extraction ────────────────────────────────────────────

function extractSourceLayout(presentationDoc) {
  const sldSz = presentationDoc?.["p:presentation"]?.["p:sldSz"];
  if (!sldSz) {
    return { width: 10, height: 7.5 };
  }

  return {
    width: emuToInches(sldSz["@_cx"]),
    height: emuToInches(sldSz["@_cy"]),
  };
}

// ── Slide Path Discovery ────────────────────────────────────────────────

function getSlidePaths(presentationDoc, presentationRelsDoc, zip) {
  const relMap = parseRelationships(presentationRelsDoc, "ppt/presentation.xml");
  const slideIds = asArray(
    presentationDoc?.["p:presentation"]?.["p:sldIdLst"]?.["p:sldId"]
  );

  const orderedPaths = [];
  for (const slideId of slideIds) {
    const relId = slideId?.["@_r:id"];
    if (!relId) continue;
    const slidePath = relMap.get(relId);
    if (!slidePath || !zip.file(slidePath)) continue;
    orderedPaths.push(slidePath);
  }

  if (orderedPaths.length > 0) return orderedPaths;

  return Object.keys(zip.files)
    .filter((fileName) => /^ppt\/slides\/slide\d+\.xml$/.test(fileName))
    .sort((a, b) => {
      const aNum = Number(a.match(/slide(\d+)\.xml/)?.[1] || "0");
      const bNum = Number(b.match(/slide(\d+)\.xml/)?.[1] || "0");
      return aNum - bNum;
    });
}

// ── Slide Layout/Master Background ──────────────────────────────────────

async function getSlideBackground(slideDoc, slideRelsMap, zip, themeColors) {
  // 1. Check slide-level background
  const slideBg = slideDoc?.["p:sld"]?.["p:cSld"]?.["p:bg"];
  const bg = await extractBackground(slideBg, slideRelsMap, zip, themeColors);
  if (bg) return bg;

  // 2. Check slide layout
  const layoutRelId = findRelByType(slideRelsMap, "slideLayout");
  if (layoutRelId) {
    const layoutFile = zip.file(layoutRelId);
    if (layoutFile) {
      try {
        const layoutDoc = parseXml(await layoutFile.async("nodebuffer"), layoutRelId);
        const layoutBg = layoutDoc?.["p:sldLayout"]?.["p:cSld"]?.["p:bg"];
        const layoutRelsPath = layoutRelId.replace("ppt/slideLayouts/", "ppt/slideLayouts/_rels/") + ".rels";
        const layoutRelsFile = zip.file(layoutRelsPath);
        const layoutRelsMap = layoutRelsFile
          ? parseRelationships(parseXml(await layoutRelsFile.async("nodebuffer"), layoutRelsPath), layoutRelId)
          : new Map();
        const lbg = await extractBackground(layoutBg, layoutRelsMap, zip, themeColors);
        if (lbg) return lbg;

        // 3. Check slide master
        const masterRelId = findRelByType(layoutRelsMap, "slideMaster");
        if (masterRelId) {
          const masterFile = zip.file(masterRelId);
          if (masterFile) {
            const masterDoc = parseXml(await masterFile.async("nodebuffer"), masterRelId);
            const masterBg = masterDoc?.["p:sldMaster"]?.["p:cSld"]?.["p:bg"];
            const masterRelsPath = masterRelId.replace("ppt/slideMasters/", "ppt/slideMasters/_rels/") + ".rels";
            const masterRelsFile = zip.file(masterRelsPath);
            const masterRelsMap = masterRelsFile
              ? parseRelationships(parseXml(await masterRelsFile.async("nodebuffer"), masterRelsPath), masterRelId)
              : new Map();
            const mbg = await extractBackground(masterBg, masterRelsMap, zip, themeColors);
            if (mbg) return mbg;
          }
        }
      } catch {
        // ignore layout/master parsing errors
      }
    }
  }

  return null;
}

function findRelByType(relsMap, typeFragment) {
  for (const [, target] of relsMap) {
    if (target.includes(typeFragment)) return target;
  }
  return null;
}

// ── Main Reader ─────────────────────────────────────────────────────────

async function readPptxToModel(inputPath) {
  const fileBuffer = await fs.readFile(inputPath);
  const zip = await JSZip.loadAsync(fileBuffer);

  const presentationFile = zip.file("ppt/presentation.xml");
  if (!presentationFile) {
    throw new Error("Invalid PPTX: missing ppt/presentation.xml");
  }
  const presentationDoc = parseXml(await presentationFile.async("nodebuffer"), "presentation.xml");

  const presentationRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const presentationRelsDoc = presentationRelsFile
    ? parseXml(await presentationRelsFile.async("nodebuffer"), "presentation.xml.rels")
    : {};

  // Parse theme colors
  let themeColors = new Map();
  const themeFile = zip.file("ppt/theme/theme1.xml");
  if (themeFile) {
    try {
      const themeDoc = parseXml(await themeFile.async("nodebuffer"), "theme1.xml");
      themeColors = parseThemeColors(themeDoc);
    } catch {
      // ignore theme parsing errors
    }
  }

  const sourceLayout = extractSourceLayout(presentationDoc);
  const slidePaths = getSlidePaths(presentationDoc, presentationRelsDoc, zip);

  const slides = [];
  let extractedGraphicFrames = 0;
  let placeholderGraphicFrames = 0;
  let unsupportedGraphicFrames = 0;
  let extractedShapeFallbacks = 0;
  let extractedTables = 0;
  let extractedCharts = 0;
  let extractedGroups = 0;

  for (let i = 0; i < slidePaths.length; i += 1) {
    const slidePath = slidePaths[i];
    const slideXml = await zip.file(slidePath).async("nodebuffer");
    const slideDoc = parseXml(slideXml, slidePath);

    const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    const relsFile = zip.file(relsPath);
    const relsMap = relsFile
      ? parseRelationships(parseXml(await relsFile.async("nodebuffer"), relsPath), slidePath)
      : new Map();

    const spTree = slideDoc?.["p:sld"]?.["p:cSld"]?.["p:spTree"];
    const shapeNodes = asArray(spTree?.["p:sp"]);
    const picNodes = asArray(spTree?.["p:pic"]);

    const elements = [];
    let unsupportedCount = 0;
    let recoveredGraphicCount = 0;
    let placeholderCount = 0;
    let recoveredShapeCount = 0;
    let tableCount = 0;
    let chartCount = 0;
    let groupCount = 0;

    // Extract shapes
    for (const shapeNode of shapeNodes) {
      const shape = extractTextShape(shapeNode, themeColors);
      if (shape) {
        elements.push(shape);
      } else {
        const fallback = extractShapeFallback(shapeNode, themeColors);
        if (fallback) {
          elements.push(fallback);
          recoveredShapeCount += 1;
          extractedShapeFallbacks += 1;
        } else {
          unsupportedCount += 1;
        }
      }
    }

    // Extract pictures
    for (const picNode of picNodes) {
      const picture = await extractPicture(picNode, relsMap, zip);
      if (picture) elements.push(picture);
      else unsupportedCount += 1;
    }

    // Extract group shapes
    for (const grpSpNode of asArray(spTree?.["p:grpSp"])) {
      const group = await extractGroupShape(grpSpNode, relsMap, zip, themeColors);
      if (group) {
        elements.push(group);
        groupCount += 1;
        extractedGroups += 1;
      } else {
        unsupportedCount += 1;
      }
    }

    // Extract graphic frames (tables, charts, etc.)
    const graphicFrames = asArray(spTree?.["p:graphicFrame"]);
    for (const frameNode of graphicFrames) {
      const extracted = extractGraphicFrame(frameNode, relsMap, themeColors, zip);
      if (!extracted) {
        unsupportedCount += 1;
        unsupportedGraphicFrames += 1;
        continue;
      }

      // Handle async chart extraction
      if (extracted.type === "_chart_pending") {
        const chart = await extractChart(frameNode, relsMap, zip, extracted, { "@_id": extracted.id, "@_name": extracted.name });
        if (chart) {
          elements.push(chart);
          chartCount += 1;
          extractedCharts += 1;
        } else {
          // Fall back to placeholder
          elements.push({
            type: "placeholder",
            placeholderKind: "chart",
            id: extracted.id,
            name: extracted.name,
            x: extracted.x, y: extracted.y, w: extracted.w, h: extracted.h,
            text: `[CHART] ${extracted.name}`,
          });
          placeholderCount += 1;
          placeholderGraphicFrames += 1;
        }
        continue;
      }

      if (extracted.type === "table") {
        elements.push(extracted);
        tableCount += 1;
        extractedTables += 1;
        recoveredGraphicCount += 1;
        extractedGraphicFrames += 1;
        continue;
      }

      elements.push(extracted);
      recoveredGraphicCount += 1;
      extractedGraphicFrames += 1;

      if (extracted.type === "placeholder") {
        placeholderCount += 1;
        placeholderGraphicFrames += 1;
      }
    }

    // Extract slide background
    const background = await getSlideBackground(slideDoc, relsMap, zip, themeColors);

    slides.push({
      index: i + 1,
      sourcePath: slidePath,
      elements: elements.sort((a, b) => (a.y - b.y) || (a.x - b.x)),
      unsupportedCount,
      recoveredGraphicCount,
      placeholderCount,
      recoveredShapeCount,
      tableCount,
      chartCount,
      groupCount,
      ...(background && { background }),
    });
  }

  return {
    sourcePath: inputPath,
    sourceLayout,
    slides,
    themeColors: Object.fromEntries(themeColors),
    parsingStats: {
      extractedGraphicFrames,
      placeholderGraphicFrames,
      unsupportedGraphicFrames,
      extractedShapeFallbacks,
      extractedTables,
      extractedCharts,
      extractedGroups,
    },
  };
}

module.exports = {
  readPptxToModel,
  // Exported for testing
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
};
