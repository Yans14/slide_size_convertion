function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function round(value, precision = 4) {
  const factor = 10 ** precision;
  return Math.round(value * factor) / factor;
}

function computeArea(item) {
  return item.w * item.h;
}

function overlapArea(a, b) {
  const xOverlap = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x));
  const yOverlap = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
  return xOverlap * yOverlap;
}

function isFlowElement(item) {
  return (
    item.type === "text" ||
    item.type === "table-text" ||
    item.type === "placeholder" ||
    item.type === "table" ||
    item.type === "chart"
  );
}

function isLowPriority(item, slideArea) {
  const areaRatio = slideArea > 0 ? computeArea(item) / slideArea : 0;

  if (item.type === "text" || item.type === "table-text") {
    const textLength = item.text ? item.text.replace(/\s+/g, "").length : 0;
    return textLength <= 3 || areaRatio < 0.02;
  }

  if (item.type === "image") {
    return areaRatio < 0.03;
  }

  return false;
}

function resolveOverlaps(items, slideHeight) {
  const minGap = 0.05;
  const sorted = [...items].sort((a, b) => (a.y - b.y) || (a.x - b.x));
  const actions = [];

  for (let i = 0; i < sorted.length; i += 1) {
    const current = sorted[i];
    if (!isFlowElement(current)) continue;

    for (let j = 0; j < i; j += 1) {
      const previous = sorted[j];
      if (!isFlowElement(previous)) continue;
      const overlap = overlapArea(current, previous);
      if (overlap <= 0.001) continue;

      const newY = previous.y + previous.h + minGap;
      if (newY > current.y) {
        actions.push({
          type: "reposition",
          reason: "overlap",
          elementName: current.name,
          from: { x: round(current.x), y: round(current.y) },
          to: { x: round(current.x), y: round(newY) },
        });
        current.y = newY;
      }
    }

    if (current.y + current.h > slideHeight) {
      current.overflow = true;
    }
  }

  return actions;
}

function transformElement(item, scale, offsetX, offsetY, policy) {
  const transformed = {
    ...item,
    x: round(item.x * scale + offsetX),
    y: round(item.y * scale + offsetY),
    w: round(item.w * scale),
    h: round(item.h * scale),
  };

  if (item.type === "text" || item.type === "table-text" || item.type === "placeholder") {
    const sourceFontSize = Number.isFinite(item.fontSizePt) ? item.fontSizePt : 14;
    transformed.fontSizePt = round(
      clamp(
        sourceFontSize * scale,
        policy.readabilityMinFontPt,
        Math.max(sourceFontSize, policy.readabilityMinFontPt)
      )
    );
  }

  // Scale table cell font sizes
  if (item.type === "table" && transformed.rows) {
    transformed.rows = item.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        fontSizePt: cell.fontSizePt
          ? round(clamp(cell.fontSizePt * scale, policy.readabilityMinFontPt, Math.max(cell.fontSizePt, policy.readabilityMinFontPt)))
          : cell.fontSizePt,
      })),
    }));
  }

  // Scale table column widths
  if (item.type === "table" && transformed.colWidths) {
    transformed.colWidths = item.colWidths.map((w) => round(w * scale));
  }

  // Handle group shapes — transform children
  if (item.type === "group" && item.children) {
    transformed.children = item.children.map((child) =>
      transformElement(child, scale, offsetX, offsetY, policy)
    );
  }

  return transformed;
}

function splitByHeight(items, dstLayout, policy) {
  const slides = [];
  const actions = [];
  const dropped = [];
  const topPadding = 0.1;
  const bottomLimit = dstLayout.height - 0.1;
  const usableHeight = bottomLimit - topPadding;
  const slideArea = dstLayout.width * dstLayout.height;

  const sorted = [...items].sort((a, b) => (a.y - b.y) || (a.x - b.x));

  function ensureSlide(slideIndex) {
    while (slides.length <= slideIndex) {
      slides.push([]);
    }
  }

  for (const item of sorted) {
    const itemCopy = { ...item };

    if (itemCopy.h > dstLayout.height - 0.2) {
      if (policy.allowElementDeletion && isLowPriority(itemCopy, slideArea)) {
        dropped.push({ elementName: itemCopy.name, reason: "oversized-low-priority" });
        continue;
      }

      const originalHeight = itemCopy.h;
      itemCopy.h = round(dstLayout.height - 0.2);
      itemCopy.y = topPadding;
      actions.push({
        type: "resize",
        reason: "oversized",
        elementName: itemCopy.name,
        from: { h: round(originalHeight) },
        to: { h: round(itemCopy.h) },
      });
    }
    if (itemCopy.y < topPadding) {
      itemCopy.y = topPadding;
    }

    const overflows = itemCopy.y + itemCopy.h > bottomLimit;
    if (!overflows) {
      ensureSlide(0);
      slides[0].push(itemCopy);
      continue;
    }

    if (!policy.allowSlideSplit) {
      if (policy.allowElementDeletion && isLowPriority(itemCopy, slideArea)) {
        dropped.push({ elementName: itemCopy.name, reason: "overflow-low-priority" });
        continue;
      }

      itemCopy.y = round(Math.max(topPadding, bottomLimit - itemCopy.h));
      actions.push({
        type: "reposition",
        reason: "overflow-clamped",
        elementName: itemCopy.name,
        to: { y: itemCopy.y },
      });
      ensureSlide(0);
      slides[0].push(itemCopy);
      continue;
    }

    let slideIndex = Math.max(1, Math.floor((itemCopy.y - topPadding) / usableHeight));
    let rebasedY = itemCopy.y - slideIndex * usableHeight;
    while (rebasedY + itemCopy.h > bottomLimit) {
      slideIndex += 1;
      rebasedY = itemCopy.y - slideIndex * usableHeight;
      if (slideIndex > 200) break;
    }

    itemCopy.y = round(clamp(rebasedY, topPadding, bottomLimit - itemCopy.h));
    ensureSlide(slideIndex);
    slides[slideIndex].push(itemCopy);
    actions.push({
      type: "split",
      reason: "overflow",
      elementName: itemCopy.name,
    });
  }

  const compactSlides = slides.filter((segment) => segment.length > 0);
  if (compactSlides.length === 0) {
    compactSlides.push([]);
  }

  return {
    slides: compactSlides,
    actions,
    dropped,
  };
}

function buildDefaultPolicy(policy = {}) {
  return {
    allowSlideSplit: policy.allowSlideSplit !== false,
    allowElementDeletion: policy.allowElementDeletion === true,
    maxSlidesGrowthPct: Number.isFinite(policy.maxSlidesGrowthPct)
      ? Number(policy.maxSlidesGrowthPct)
      : 200,
    readabilityMinFontPt: Number.isFinite(policy.readabilityMinFontPt)
      ? Number(policy.readabilityMinFontPt)
      : 12,
    reviewThreshold: Number.isFinite(policy.reviewThreshold)
      ? Number(policy.reviewThreshold)
      : 0.78,
    strictReview: policy.strictReview !== false,
  };
}

function scoreSlideQuality({ sourceSlide, actions, dropped, outputSlideCount }) {
  let score = 1;

  const splitCount = actions.filter((item) => item.type === "split").length;
  const overlapFixCount = actions.filter(
    (item) => item.type === "reposition" && item.reason === "overlap"
  ).length;
  const resizeCount = actions.filter((item) => item.type === "resize").length;

  score -= (sourceSlide.unsupportedCount || 0) * 0.14;
  score -= (sourceSlide.placeholderCount || 0) * 0.04;
  score -= dropped.length * 0.15;
  score -= splitCount * 0.2;
  score -= overlapFixCount * 0.03;
  score -= resizeCount * 0.05;
  score -= Math.max(0, outputSlideCount - 1) * 0.08;

  // Bonus for fully extracted tables and charts (higher fidelity)
  const tableBonus = (sourceSlide.tableCount || 0) * 0.02;
  const chartBonus = (sourceSlide.chartCount || 0) * 0.02;
  score += tableBonus + chartBonus;

  return round(clamp(score, 0, 1), 3);
}

function buildReviewDecision({ sourceSlide, actions, dropped, outputSlideIndexes, confidenceScore, policy }) {
  const reasons = [];
  const splitCount = actions.filter((item) => item.type === "split").length;

  if (confidenceScore < policy.reviewThreshold) {
    reasons.push("low-confidence");
  }

  if (sourceSlide.unsupportedCount > 0) {
    reasons.push("unsupported-elements");
  }

  if (dropped.length > 0) {
    reasons.push("deleted-elements");
  }

  if (splitCount > 0) {
    reasons.push("slide-split");
  }

  if (outputSlideIndexes.length > 1) {
    reasons.push("multi-output-slide");
  }

  if (sourceSlide.placeholderCount > 0) {
    reasons.push("placeholder-content");
  }

  const strictTrigger =
    sourceSlide.unsupportedCount > 0 ||
    dropped.length > 0 ||
    splitCount > 0 ||
    sourceSlide.placeholderCount > 0;

  const needsManualReview = policy.strictReview
    ? confidenceScore < policy.reviewThreshold || strictTrigger
    : confidenceScore < policy.reviewThreshold;

  return {
    needsManualReview,
    reasons: Array.from(new Set(reasons)),
  };
}

// Flatten group elements into the slide's element list
function flattenElements(elements) {
  const result = [];
  for (const el of elements) {
    if (el.type === "group" && el.children) {
      // Keep the group element itself for type tracking, but also add children
      result.push(el);
    } else {
      result.push(el);
    }
  }
  return result;
}

function buildConversionPlan(model, targetLayout, rawPolicy = {}) {
  const policy = buildDefaultPolicy(rawPolicy);
  const srcLayout = model.sourceLayout;

  const scale = Math.min(targetLayout.width / srcLayout.width, targetLayout.height / srcLayout.height);
  const offsetX = (targetLayout.width - srcLayout.width * scale) / 2;
  const offsetY = (targetLayout.height - srcLayout.height * scale) / 2;

  const plannedSlides = [];
  const sourceReports = [];
  const manualReviewQueue = [];

  for (const sourceSlide of model.slides) {
    const transformed = sourceSlide.elements.map((item) =>
      transformElement(item, scale, offsetX, offsetY, policy)
    );

    const overlapActions = resolveOverlaps(transformed, targetLayout.height);
    const { slides, actions: splitActions, dropped } = splitByHeight(transformed, targetLayout, policy);

    const outputIndexes = [];
    for (const segment of slides) {
      outputIndexes.push(plannedSlides.length + 1);
      plannedSlides.push({
        sourceSlideIndex: sourceSlide.index,
        elements: segment,
        ...(sourceSlide.background && { background: sourceSlide.background }),
      });
    }

    sourceReports.push({
      sourceSlideIndex: sourceSlide.index,
      sourceElementCount: sourceSlide.elements.length,
      unsupportedCount: sourceSlide.unsupportedCount,
      recoveredGraphicCount: sourceSlide.recoveredGraphicCount || 0,
      recoveredShapeCount: sourceSlide.recoveredShapeCount || 0,
      placeholderCount: sourceSlide.placeholderCount || 0,
      tableCount: sourceSlide.tableCount || 0,
      chartCount: sourceSlide.chartCount || 0,
      groupCount: sourceSlide.groupCount || 0,
      outputSlideIndexes: outputIndexes,
      actions: [...overlapActions, ...splitActions],
      dropped,
      confidenceScore: 0,
      needsManualReview: false,
      reviewReasons: [],
    });

    const slideReport = sourceReports[sourceReports.length - 1];
    const confidenceScore = scoreSlideQuality({
      sourceSlide,
      actions: slideReport.actions,
      dropped,
      outputSlideCount: outputIndexes.length,
    });

    const review = buildReviewDecision({
      sourceSlide,
      actions: slideReport.actions,
      dropped,
      outputSlideIndexes: outputIndexes,
      confidenceScore,
      policy,
    });

    slideReport.confidenceScore = confidenceScore;
    slideReport.needsManualReview = review.needsManualReview;
    slideReport.reviewReasons = review.reasons;

    if (review.needsManualReview) {
      manualReviewQueue.push({
        sourceSlideIndex: sourceSlide.index,
        outputSlideIndexes: outputIndexes,
        confidenceScore,
        reasons: review.reasons,
      });
    }
  }

  const maxSlides = Math.ceil(model.slides.length * (1 + policy.maxSlidesGrowthPct / 100));
  const maxGrowthExceeded = plannedSlides.length > maxSlides;

  return {
    sourceLayout: model.sourceLayout,
    targetLayout,
    policy,
    slides: plannedSlides,
    report: {
      sourceSlides: model.slides.length,
      outputSlides: plannedSlides.length,
      maxSlides,
      maxGrowthExceeded,
      reviewThreshold: policy.reviewThreshold,
      strictReview: policy.strictReview,
      manualReviewCount: manualReviewQueue.length,
      manualReviewQueue,
      slideReports: sourceReports,
    },
  };
}

module.exports = {
  buildConversionPlan,
  // Exported for testing
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
};
