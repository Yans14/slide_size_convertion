const PRESET_LAYOUTS = {
  wide: { name: "WIDE_16_9", width: 13.333, height: 7.5 },
  standard: { name: "STANDARD_4_3", width: 10, height: 7.5 },
  a4: { name: "A4_LANDSCAPE", width: 11.693, height: 8.268 },
  "a4-portrait": { name: "A4_PORTRAIT", width: 8.268, height: 11.693 },
};

function getTargetLayout(options = {}) {
  if (options.targetWidth && options.targetHeight) {
    const width = Number(options.targetWidth);
    const height = Number(options.targetHeight);
    if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
      throw new Error("Invalid custom target size. Width and height must be positive numbers.");
    }
    return { name: "CUSTOM", width, height };
  }

  const preset = (options.target || "wide").toLowerCase();
  const layout = PRESET_LAYOUTS[preset];
  if (!layout) {
    throw new Error(
      `Unknown target layout \"${options.target}\". Valid presets: ${Object.keys(PRESET_LAYOUTS).join(", ")}`
    );
  }

  return layout;
}

module.exports = {
  PRESET_LAYOUTS,
  getTargetLayout,
};
