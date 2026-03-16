#!/usr/bin/env node
const fs = require("fs/promises");
const path = require("path");
const { readPptxToModel } = require("./pptx-reader");
const { getTargetLayout } = require("./layouts");
const { buildConversionPlan } = require("./transform");
const { writePlanToPptx } = require("./generate");

function parseBool(value, defaultValue) {
  if (value === undefined) return defaultValue;
  const normalized = String(value).trim().toLowerCase();
  if (["1", "true", "yes", "y", "on"].includes(normalized)) return true;
  if (["0", "false", "no", "n", "off"].includes(normalized)) return false;
  return defaultValue;
}

function parseArgs(argv) {
  const args = {
    input: null,
    output: path.resolve(process.cwd(), "out/converted.pptx"),
    report: path.resolve(process.cwd(), "out/conversion-report.json"),
    target: "wide",
    allowSlideSplit: true,
    allowElementDeletion: false,
    maxSlidesGrowthPct: 200,
    readabilityMinFontPt: 12,
    reviewThreshold: 0.78,
    strictReview: true,
    renderPlaceholders: true,
    renderTableBoxes: true,
  };

  for (let i = 2; i < argv.length; i += 1) {
    const token = argv[i];
    if (!token.startsWith("--")) continue;

    const key = token.slice(2);
    const next = argv[i + 1];

    if (
      [
        "input",
        "output",
        "report",
        "target",
        "target-width",
        "target-height",
        "allow-slide-split",
        "allow-element-deletion",
        "max-slides-growth-pct",
        "readability-min-font-pt",
        "review-threshold",
        "strict-review",
        "render-placeholders",
        "render-table-boxes",
      ].includes(key)
    ) {
      if (next === undefined || next.startsWith("--")) {
        throw new Error(`Missing value for --${key}`);
      }
      i += 1;

      if (key === "input") args.input = path.resolve(next);
      if (key === "output") args.output = path.resolve(next);
      if (key === "report") args.report = path.resolve(next);
      if (key === "target") args.target = next;
      if (key === "target-width") args.targetWidth = Number(next);
      if (key === "target-height") args.targetHeight = Number(next);
      if (key === "allow-slide-split") args.allowSlideSplit = parseBool(next, true);
      if (key === "allow-element-deletion") args.allowElementDeletion = parseBool(next, false);
      if (key === "max-slides-growth-pct") args.maxSlidesGrowthPct = Number(next);
      if (key === "readability-min-font-pt") args.readabilityMinFontPt = Number(next);
      if (key === "review-threshold") args.reviewThreshold = Number(next);
      if (key === "strict-review") args.strictReview = parseBool(next, true);
      if (key === "render-placeholders") args.renderPlaceholders = parseBool(next, true);
      if (key === "render-table-boxes") args.renderTableBoxes = parseBool(next, true);
    }

    if (key === "help" || key === "h") {
      args.help = true;
    }
  }

  return args;
}

function helpText() {
  return [
    "PptxGenJS Slide Converter Agent",
    "",
    "Usage:",
    "  node src/convert.js --input <source.pptx> [options]",
    "",
    "Options:",
    "  --output <path>                    Output PPTX path (default: out/converted.pptx)",
    "  --report <path>                    Output JSON report path (default: out/conversion-report.json)",
    "  --target <wide|standard|a4|a4-portrait>",
    "  --target-width <inches>            Custom width in inches",
    "  --target-height <inches>           Custom height in inches",
    "  --allow-slide-split <true|false>   Split overflowing slides (default: true)",
    "  --allow-element-deletion <true|false>",
    "  --max-slides-growth-pct <number>   Max tolerated deck growth percentage",
    "  --readability-min-font-pt <number> Minimum text size after conversion",
    "  --review-threshold <0..1>          Confidence threshold for manual review",
    "  --strict-review <true|false>       Flag risk scenarios even above threshold",
    "  --render-placeholders <true|false> Render chart/diagram placeholder boxes",
    "  --render-table-boxes <true|false>  Render subtle table boundary boxes",
    "  --help                             Show help",
  ].join("\n");
}

async function writeJson(filePath, payload) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(payload, null, 2), "utf8");
}

async function run() {
  const args = parseArgs(process.argv);

  if (args.help || !args.input) {
    console.log(helpText());
    process.exit(args.help ? 0 : 1);
  }

  const sourceModel = await readPptxToModel(args.input);
  const targetLayout = getTargetLayout(args);

  const plan = buildConversionPlan(sourceModel, targetLayout, {
    allowSlideSplit: args.allowSlideSplit,
    allowElementDeletion: args.allowElementDeletion,
    maxSlidesGrowthPct: Number(args.maxSlidesGrowthPct),
    readabilityMinFontPt: Number(args.readabilityMinFontPt),
    reviewThreshold: Number(args.reviewThreshold),
    strictReview: args.strictReview,
  });

  await writePlanToPptx(plan, args.output, {
    renderPlaceholders: args.renderPlaceholders,
    renderTableBoxes: args.renderTableBoxes,
  });

  const report = {
    timestamp: new Date().toISOString(),
    source: {
      path: args.input,
      layout: sourceModel.sourceLayout,
      slides: sourceModel.slides.length,
      parsingStats: sourceModel.parsingStats,
    },
    target: {
      path: args.output,
      layout: plan.targetLayout,
    },
    policy: plan.policy,
    summary: plan.report,
  };

  await writeJson(args.report, report);

  console.log(`Converted ${sourceModel.slides.length} slide(s) -> ${plan.slides.length} slide(s)`);
  console.log(
    `Manual review queue: ${plan.report.manualReviewCount} slide(s), threshold=${plan.report.reviewThreshold}, strict=${plan.report.strictReview}`
  );
  console.log(`Output PPTX: ${args.output}`);
  console.log(`Report JSON: ${args.report}`);
  if (plan.report.maxGrowthExceeded) {
    console.log(
      `Warning: output slide count exceeded max allowed growth (${plan.report.outputSlides}/${plan.report.maxSlides}).`
    );
  }
}

run().catch((error) => {
  console.error(`Conversion failed: ${error.message}`);
  process.exit(1);
});
