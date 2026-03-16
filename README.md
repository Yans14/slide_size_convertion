# PptxGenJS Slide Converter Agent

This project converts a source PPTX deck from one slide size to another, then writes a new deck with PptxGenJS.

## What it does

- Parses slide XML from a source `.pptx`
- Extracts text boxes and images with geometry
- Extracts non-text shape geometry (fill/line best-effort)
- Extracts `graphicFrame` content:
  - table content is converted to text blocks
  - chart/diagram/OLE blocks are converted to explicit placeholders
- Converts coordinates to a target layout (`wide`, `standard`, `a4`, `a4-portrait`, or custom)
- Resolves overlaps and optionally splits overflowing content into additional slides
- Generates a conversion report (`JSON`) with slide-level actions, confidence scores, and manual-review queue

## Install

```bash
npm install
```

## Usage

```bash
npm run convert -- \
  --input "/absolute/path/to/source.pptx" \
  --output "./out/converted.pptx" \
  --report "./out/conversion-report.json" \
  --target wide \
  --allow-slide-split true \
  --allow-element-deletion false \
  --review-threshold 0.78 \
  --strict-review true

## Key options

- `--review-threshold <0..1>`: confidence floor for auto-accepting converted slides.
- `--strict-review <true|false>`: if true, force manual review for splits, dropped elements, placeholders, and unsupported content.
- `--render-placeholders <true|false>`: draw placeholder boxes for chart/diagram/OLE blocks.
- `--render-table-boxes <true|false>`: draw subtle table boundaries around extracted table text.
```

## Notes

- PptxGenJS is used to generate the output deck.
- `graphicFrame` extraction is best-effort. For complex charts/diagrams, placeholder rendering plus manual review is recommended.
- Non-text shape style fidelity is best-effort and may vary from the source theme.
