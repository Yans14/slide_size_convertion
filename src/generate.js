const fs = require("fs/promises");
const path = require("path");
const PptxGenJS = require("pptxgenjs");

function normalizeColor(color, fallback = "000000") {
  if (!color) return fallback;
  const value = String(color).replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{6}$/.test(value)) return value;
  return fallback;
}

async function ensureOutputDir(filePath) {
  const dirPath = path.dirname(filePath);
  await fs.mkdir(dirPath, { recursive: true });
}

// ── Rich Text Rendering ─────────────────────────────────────────────────

function buildTextProps(paragraphs) {
  const textProps = [];

  for (let pIdx = 0; pIdx < paragraphs.length; pIdx += 1) {
    const para = paragraphs[pIdx];

    if (para.runs.length === 0) {
      // Empty paragraph → line break
      textProps.push({ text: "\n", options: {} });
      continue;
    }

    for (let rIdx = 0; rIdx < para.runs.length; rIdx += 1) {
      const run = para.runs[rIdx];
      const opts = {};

      if (run.fontSizePt) opts.fontSize = run.fontSizePt;
      if (run.bold) opts.bold = true;
      if (run.italic) opts.italic = true;
      if (run.underline) opts.underline = { style: "sng" };
      if (run.strikethrough) opts.strike = "sngStrike";
      if (run.color) opts.color = normalizeColor(run.color);
      if (run.fontFace) opts.fontFace = run.fontFace;
      if (para.align) opts.align = para.align;
      if (para.bulletLevel !== undefined) opts.indentLevel = para.bulletLevel;

      // Add line break between paragraphs (after last run of each para except last)
      const isLastRunInPara = rIdx === para.runs.length - 1;
      const isLastPara = pIdx === paragraphs.length - 1;
      if (isLastRunInPara && !isLastPara) {
        opts.breakType = "break";
      }

      textProps.push({ text: run.text, options: opts });
    }
  }

  return textProps;
}

// ── Text-Like Element Rendering ─────────────────────────────────────────

function renderTextLikeElement(slide, element, options = {}) {
  if (element.type === "placeholder" && options.renderPlaceholders) {
    slide.addShape("rect", {
      x: element.x,
      y: element.y,
      w: element.w,
      h: element.h,
      line: { color: "888888", pt: 1, dash: "dash" },
      fill: { color: "F5F5F5", transparency: 15 },
      rectRadius: 0.03,
    });
  }

  if (element.type === "table-text" && options.renderTableBoxes) {
    slide.addShape("rect", {
      x: element.x,
      y: element.y,
      w: element.w,
      h: element.h,
      line: { color: "BBBBBB", pt: 0.5 },
      fill: { color: "FFFFFF", transparency: 0 },
    });
  }

  // Use rich text when paragraphs are available
  if (element.paragraphs && element.paragraphs.length > 0) {
    const textProps = buildTextProps(element.paragraphs);
    if (textProps.length > 0) {
      slide.addText(textProps, {
        x: element.x,
        y: element.y,
        w: element.w,
        h: element.h,
        valign: "top",
        fit: "shrink",
        margin: 2,
      });
      return;
    }
  }

  // Fallback: simple text
  const text = element.text || "";
  slide.addText(text, {
    x: element.x,
    y: element.y,
    w: element.w,
    h: element.h,
    fontSize: element.fontSizePt || 14,
    bold: Boolean(element.bold),
    italic: element.type === "placeholder",
    color: normalizeColor(element.color),
    valign: "top",
    fit: "shrink",
    margin: 2,
  });
}

// ── Shape Element Rendering ─────────────────────────────────────────────

function renderShapeElement(slide, element) {
  slide.addShape(element.shapeType || "rect", {
    x: element.x,
    y: element.y,
    w: element.w,
    h: element.h,
    line: {
      color: normalizeColor(element.lineColor, "666666"),
      pt: element.lineWidthPt || 0.75,
    },
    fill: element.fillColor
      ? { color: normalizeColor(element.fillColor, "FFFFFF"), transparency: 0 }
      : { color: "FFFFFF", transparency: 100 },
  });

  // Render text inside shapes
  if (element.shapeParagraphs && element.shapeParagraphs.length > 0) {
    const textProps = buildTextProps(element.shapeParagraphs);
    if (textProps.length > 0) {
      slide.addText(textProps, {
        x: element.x,
        y: element.y,
        w: element.w,
        h: element.h,
        valign: "middle",
        align: "center",
        fit: "shrink",
        margin: 4,
      });
    }
  } else if (element.shapeText) {
    slide.addText(element.shapeText, {
      x: element.x,
      y: element.y,
      w: element.w,
      h: element.h,
      fontSize: 10,
      valign: "middle",
      align: "center",
      fit: "shrink",
      margin: 4,
    });
  }
}

// ── Table Element Rendering ─────────────────────────────────────────────

function renderTableElement(slide, element) {
  const tableRows = [];

  for (const row of element.rows) {
    const tableRow = [];
    for (const cell of row.cells) {
      // Skip merged cells
      if (cell.hMerge || cell.vMerge) continue;

      const cellOpts = {};
      if (cell.fontSizePt) cellOpts.fontSize = cell.fontSizePt;
      if (cell.bold) cellOpts.bold = true;
      if (cell.color) cellOpts.color = normalizeColor(cell.color);
      if (cell.fillColor) cellOpts.fill = { color: normalizeColor(cell.fillColor) };
      if (cell.colspan && cell.colspan > 1) cellOpts.colspan = cell.colspan;
      if (cell.rowspan && cell.rowspan > 1) cellOpts.rowspan = cell.rowspan;
      cellOpts.valign = "middle";
      cellOpts.margin = 2;

      // Use rich text for cells with paragraphs
      if (cell.paragraphs && cell.paragraphs.length > 0) {
        const textProps = buildTextProps(cell.paragraphs);
        tableRow.push({ text: textProps, options: cellOpts });
      } else {
        tableRow.push({ text: cell.text || "", options: cellOpts });
      }
    }
    if (tableRow.length > 0) tableRows.push(tableRow);
  }

  if (tableRows.length === 0) return;

  const tableOpts = {
    x: element.x,
    y: element.y,
    w: element.w,
    h: element.h,
    border: { type: "solid", pt: 0.5, color: "CFCFCF" },
  };

  // Distribute column widths
  if (element.colWidths && element.colWidths.length > 0) {
    const totalColWidth = element.colWidths.reduce((sum, w) => sum + w, 0);
    if (totalColWidth > 0) {
      tableOpts.colW = element.colWidths.map((w) => (w / totalColWidth) * element.w);
    }
  }

  slide.addTable(tableRows, tableOpts);
}

// ── Chart Element Rendering ─────────────────────────────────────────────

const PPTXGENJS_CHART_MAP = {
  column_clustered: "bar",
  column_stacked: "bar",
  column_stacked_100: "bar",
  bar_clustered: "bar",
  bar_stacked: "bar",
  line: "line",
  line3D: "line",
  pie: "pie",
  pie3D: "pie",
  doughnut: "doughnut",
  area: "area",
  area3D: "area",
  scatter: "scatter",
  bubble: "bubble",
  radar: "radar",
  bar: "bar",
  bar3D: "bar",
  stock: "line",
  surface: "bar",
  surface3D: "bar",
};

function renderChartElement(slide, element, pptx) {
  const chartTypeName = PPTXGENJS_CHART_MAP[element.chartType] || "bar";

  // Get chart type enum from pptx instance
  const chartTypeEnum = pptx.ChartType?.[chartTypeName] || chartTypeName;

  const chartData = [];
  for (const ser of element.series || []) {
    chartData.push({
      name: ser.name,
      labels: element.categories || [],
      values: ser.values || [],
    });
  }

  if (chartData.length === 0) {
    // No data — render as placeholder
    renderTextLikeElement(slide, {
      ...element,
      type: "placeholder",
      text: `[CHART] ${element.name || "Chart"}`,
    });
    return;
  }

  const chartOpts = {
    x: element.x,
    y: element.y,
    w: element.w,
    h: element.h,
    showLegend: chartData.length > 1,
    legendPos: "b",
    showTitle: Boolean(element.title),
    title: element.title || undefined,
  };

  // Handle bar direction
  if (element.chartType?.startsWith("bar_")) {
    chartOpts.barDir = "bar";
  }

  // Handle stacked grouping
  if (element.chartType?.includes("stacked")) {
    chartOpts.barGrouping = element.chartType.includes("100") ? "percentStacked" : "stacked";
  }

  try {
    slide.addChart(chartTypeEnum, chartData, chartOpts);
  } catch {
    // Fallback to placeholder on chart rendering error
    renderTextLikeElement(slide, {
      ...element,
      type: "placeholder",
      text: `[CHART] ${element.name || "Chart"} (render error)`,
    });
  }
}

// ── Group Element Rendering ─────────────────────────────────────────────

function renderGroupChildren(slide, children, options, pptx) {
  for (const child of children) {
    renderElement(slide, child, options, pptx);
  }
}

// ── Background Rendering ────────────────────────────────────────────────

function applySlideBackground(slide, background) {
  if (!background) return;

  if (background.type === "solid" && background.color) {
    slide.background = { color: normalizeColor(background.color) };
    return;
  }

  if (background.type === "image" && background.dataUri) {
    slide.background = { data: background.dataUri };
    return;
  }

  if (background.type === "gradient" && background.stops?.length >= 2) {
    // PptxGenJS doesn't natively support gradient backgrounds,
    // so use a full-slide gradient shape as a workaround
    // (Note: only the first color is used as fallback)
    slide.background = { color: normalizeColor(background.stops[0].color) };
  }
}

// ── Unified Element Renderer ────────────────────────────────────────────

function renderElement(slide, element, options, pptx) {
  if (
    element.type === "text" ||
    element.type === "table-text" ||
    element.type === "placeholder"
  ) {
    renderTextLikeElement(slide, element, options);
    return;
  }

  if (element.type === "image" && element.dataUri) {
    slide.addImage({
      data: element.dataUri,
      x: element.x,
      y: element.y,
      w: element.w,
      h: element.h,
    });
    return;
  }

  if (element.type === "shape") {
    renderShapeElement(slide, element);
    return;
  }

  if (element.type === "table") {
    renderTableElement(slide, element);
    return;
  }

  if (element.type === "chart") {
    renderChartElement(slide, element, pptx);
    return;
  }

  if (element.type === "group") {
    renderGroupChildren(slide, element.children || [], options, pptx);
    return;
  }
}

// ── Main Writer ─────────────────────────────────────────────────────────

async function writePlanToPptx(plan, outputPath, rawOptions = {}) {
  const options = {
    renderPlaceholders: rawOptions.renderPlaceholders !== false,
    renderTableBoxes: rawOptions.renderTableBoxes !== false,
  };

  const pptx = new PptxGenJS();
  const layoutName = "CONVERTED_LAYOUT";

  pptx.defineLayout({
    name: layoutName,
    width: plan.targetLayout.width,
    height: plan.targetLayout.height,
  });
  pptx.layout = layoutName;
  pptx.author = "PptxGenJS Slide Converter Agent";
  pptx.subject = "Converted deck";
  pptx.title = "Slide format conversion";

  for (const plannedSlide of plan.slides) {
    const slide = pptx.addSlide();

    // Apply background
    if (plannedSlide.background) {
      applySlideBackground(slide, plannedSlide.background);
    }

    for (const element of plannedSlide.elements) {
      renderElement(slide, element, options, pptx);
    }
  }

  await ensureOutputDir(outputPath);
  await pptx.writeFile({ fileName: outputPath });
}

module.exports = {
  writePlanToPptx,
  // Exported for testing
  buildTextProps,
  normalizeColor,
  renderElement,
  applySlideBackground,
};
