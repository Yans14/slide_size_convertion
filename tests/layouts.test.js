import { describe, it, expect } from "vitest";
const { PRESET_LAYOUTS, getTargetLayout } = require("../src/layouts");

describe("PRESET_LAYOUTS", () => {
  it("has wide, standard, a4, and a4-portrait", () => {
    expect(Object.keys(PRESET_LAYOUTS).sort()).toEqual(["a4", "a4-portrait", "standard", "wide"]);
  });

  it("wide is 13.333 x 7.5", () => {
    expect(PRESET_LAYOUTS.wide.width).toBe(13.333);
    expect(PRESET_LAYOUTS.wide.height).toBe(7.5);
  });

  it("standard is 10 x 7.5", () => {
    expect(PRESET_LAYOUTS.standard.width).toBe(10);
    expect(PRESET_LAYOUTS.standard.height).toBe(7.5);
  });
});

describe("getTargetLayout", () => {
  it("returns wide layout by default", () => {
    const result = getTargetLayout({});
    expect(result).toEqual(PRESET_LAYOUTS.wide);
  });

  it("returns wide layout when target is 'wide'", () => {
    const result = getTargetLayout({ target: "wide" });
    expect(result).toEqual(PRESET_LAYOUTS.wide);
  });

  it("returns standard layout", () => {
    const result = getTargetLayout({ target: "standard" });
    expect(result).toEqual(PRESET_LAYOUTS.standard);
  });

  it("returns a4-portrait layout", () => {
    const result = getTargetLayout({ target: "a4-portrait" });
    expect(result.width).toBeCloseTo(8.268);
    expect(result.height).toBeCloseTo(11.693);
  });

  it("is case-insensitive", () => {
    const result = getTargetLayout({ target: "WIDE" });
    expect(result).toEqual(PRESET_LAYOUTS.wide);
  });

  it("supports custom dimensions", () => {
    const result = getTargetLayout({ targetWidth: 16, targetHeight: 9 });
    expect(result).toEqual({ name: "CUSTOM", width: 16, height: 9 });
  });

  it("throws on invalid preset", () => {
    expect(() => getTargetLayout({ target: "invalid" })).toThrow(/Unknown target layout/);
  });

  it("throws on invalid custom dimensions", () => {
    expect(() => getTargetLayout({ targetWidth: -1, targetHeight: 9 })).toThrow(/Invalid custom target size/);
  });

  it("falls through to preset when width is zero (falsy)", () => {
    const result = getTargetLayout({ targetWidth: 0, targetHeight: 9 });
    expect(result).toEqual(PRESET_LAYOUTS.wide);
  });
});
