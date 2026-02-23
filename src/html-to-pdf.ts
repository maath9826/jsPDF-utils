/**
 * jsPDF doc.html() utilities
 *
 * Helpers that prepare an HTML element for clean, paginated PDF output
 * via jsPDF's doc.html() renderer.
 */

import html2canvas from "html2canvas-pro";
import type { jsPDF } from "jspdf";

export interface Margin {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export type PageFormat =
  | "a0"
  | "a1"
  | "a2"
  | "a3"
  | "a4"
  | "a5"
  | "a6"
  | "letter"
  | "legal"
  | "tabloid";

export type MarginInput =
  | number
  | { top?: number; right?: number; bottom?: number; left?: number };

export interface PageOptions {
  unit: string;
  format: PageFormat;
  pageWidth: number;
  pageHeight: number;
  margin: Margin;
}

/** Input variant of PageOptions where margin accepts a number or partial sides. */
export type PageOptionsInput = Partial<Omit<PageOptions, "margin">> & {
  margin?: MarginInput;
};

/** Standard page dimensions in mm (portrait). */
const PAGE_SIZES: Record<PageFormat, [number, number]> = {
  a0: [841, 1189],
  a1: [594, 841],
  a2: [420, 594],
  a3: [297, 420],
  a4: [210, 297],
  a5: [148, 210],
  a6: [105, 148],
  letter: [215.9, 279.4],
  legal: [215.9, 355.6],
  tabloid: [279.4, 431.8],
};

/** Standard margins in mm per format. */
const PAGE_MARGINS: Record<PageFormat, number> = {
  a0: 40,
  a1: 35,
  a2: 30,
  a3: 25,
  a4: 25,
  a5: 20,
  a6: 12,
  letter: 25.4,
  legal: 25.4,
  tabloid: 25,
};

export interface Layout {
  renderedWidth: number;
  scale: number;
  contentWidthMm: number;
  pageContentPx: number;
}

export interface PrepareResult {
  clone: HTMLElement;
  layout: Layout;
  options: PageOptions;
  cleanup: () => void;
}

/** Create a Margin with the same value on all four sides. */
function createUniformMargin(value: number): Margin {
  return { top: value, right: value, bottom: value, left: value };
}

/** Resolve a MarginInput to a full Margin, falling back to the default for the format. */
function resolveMargin(
  input: MarginInput | undefined,
  format: PageFormat,
): Margin {
  const fallback = createUniformMargin(PAGE_MARGINS[format]);
  if (input == null) return fallback;
  if (typeof input === "number") return createUniformMargin(input);
  return {
    top: input.top ?? fallback.top,
    right: input.right ?? fallback.right,
    bottom: input.bottom ?? fallback.bottom,
    left: input.left ?? fallback.left,
  };
}

/** Resolve options: dimensions inferred from format unless explicitly provided. */
function resolveOptions(opts: PageOptionsInput = {}): PageOptions {
  const format = opts.format ?? "a4";
  const [defaultWidth, defaultHeight] = PAGE_SIZES[format];

  return {
    unit: opts.unit ?? "mm",
    format,
    pageWidth: opts.pageWidth ?? defaultWidth,
    pageHeight: opts.pageHeight ?? defaultHeight,
    margin: resolveMargin(opts.margin, format),
  };
}

/** Compute derived layout values from options. */
function computeLayout(container: HTMLElement, opts: PageOptions): Layout {
  const renderedWidth = container.offsetWidth;
  const contentWidthMm = opts.pageWidth - opts.margin.left - opts.margin.right;
  const scale = contentWidthMm / renderedWidth;
  const usableHeightMm = opts.pageHeight - opts.margin.top - opts.margin.bottom;
  // Floor to integer so page boundaries land on whole CSS pixels.
  // html2canvas renders at integer-pixel positions; fractional boundaries
  // cause sub-pixel rounding that accumulates across pages and eventually
  // cuts through text lines on later pages.
  const pageContentPx = Math.floor(usableHeightMm / scale);

  return { renderedWidth, scale, contentWidthMm, pageContentPx };
}

/**
 * CSS properties to snapshot from source computed styles onto the clone.
 * Resolves CSS custom properties (e.g. Tailwind v4's
 * `calc(var(--spacing) * 8)`) to concrete values that survive
 * html2canvas's internal cloning.
 *
 * Width/height/position properties are excluded so the clone can
 * reflow to the target page width.
 */
const SNAPSHOT_PROPERTIES = [
  // Spacing
  "padding-top",
  "padding-right",
  "padding-bottom",
  "padding-left",
  "margin-top",
  "margin-right",
  "margin-bottom",
  "margin-left",
  // Typography
  "font-family",
  "font-size",
  "font-weight",
  "font-style",
  "font-variant",
  "line-height",
  "letter-spacing",
  "word-spacing",
  "text-align",
  "text-decoration-line",
  "text-decoration-color",
  "text-decoration-style",
  "text-transform",
  "text-indent",
  "text-overflow",
  "text-shadow",
  "white-space",
  "word-break",
  "overflow-wrap",
  "direction",
  // Colors & backgrounds
  "color",
  "background-color",
  "background-image",
  "background-size",
  "background-position",
  "background-repeat",
  "background-clip",
  // Borders
  "border-top-width",
  "border-right-width",
  "border-bottom-width",
  "border-left-width",
  "border-top-style",
  "border-right-style",
  "border-bottom-style",
  "border-left-style",
  "border-top-color",
  "border-right-color",
  "border-bottom-color",
  "border-left-color",
  "border-top-left-radius",
  "border-top-right-radius",
  "border-bottom-left-radius",
  "border-bottom-right-radius",
  "outline-width",
  "outline-style",
  "outline-color",
  "outline-offset",
  // Layout
  "display",
  "flex-direction",
  "flex-wrap",
  "flex-grow",
  "flex-shrink",
  "flex-basis",
  "justify-content",
  "align-items",
  "align-self",
  "align-content",
  "order",
  "grid-template-columns",
  "grid-template-rows",
  "grid-column-start",
  "grid-column-end",
  "grid-row-start",
  "grid-row-end",
  "grid-auto-flow",
  "grid-auto-columns",
  "grid-auto-rows",
  "gap",
  "row-gap",
  "column-gap",
  // Box model
  "box-sizing",
  "overflow",
  "overflow-x",
  "overflow-y",
  // Effects
  "opacity",
  "box-shadow",
  "filter",
  "backdrop-filter",
  "mix-blend-mode",
  // Visibility
  "visibility",
  // Table
  "border-collapse",
  "border-spacing",
  "table-layout",
  // List
  "list-style-type",
  "list-style-position",
  // Vertical align
  "vertical-align",
];

/**
 * Snapshot computed styles from source elements onto matching clone elements.
 * Reads from the source (still in its original CSS context) so all CSS
 * variables and @layer rules are fully resolved, then writes concrete
 * inline values onto the clone.
 */
function snapshotElementStyles(source: HTMLElement, clone: HTMLElement): void {
  const computed = getComputedStyle(source);
  for (const prop of SNAPSHOT_PROPERTIES) {
    const value = computed.getPropertyValue(prop);
    if (value) {
      clone.style.setProperty(prop, value);
    }
  }
}

function snapshotComputedStyles(source: HTMLElement, clone: HTMLElement): void {
  snapshotElementStyles(source, clone);

  const sourceEls = source.querySelectorAll("*");
  const cloneEls = clone.querySelectorAll("*");
  const count = Math.min(sourceEls.length, cloneEls.length);

  for (let i = 0; i < count; i++) {
    const srcEl = sourceEls[i];
    const clnEl = cloneEls[i];
    if (!(srcEl instanceof HTMLElement) || !(clnEl instanceof HTMLElement))
      continue;
    snapshotElementStyles(srcEl, clnEl);
  }
}

/**
 * Clone an element and position it off-screen at print width for measurement.
 */
function createPrintClone(source: HTMLElement, pageWidth = 210): HTMLElement {
  const clone = source.cloneNode(true) as HTMLElement;
  snapshotComputedStyles(source, clone);
  Object.assign(clone.style, {
    position: "fixed",
    top: "0",
    left: "0",
    boxSizing: "border-box",
    width: pageWidth + "mm",
    opacity: "0.000001",
    pointerEvents: "none",
  });
  document.body.appendChild(clone);
  return clone;
}

/**
 * Wait for all images inside a container to finish loading so that
 * their bounding rects reflect the actual rendered dimensions.
 */
async function waitForImages(container: HTMLElement): Promise<void> {
  const images = Array.from(container.querySelectorAll("img"));
  await Promise.all(
    images.map((img) => {
      if (img.complete && img.naturalWidth > 0) return;
      return new Promise<void>((resolve) => {
        img.onload = () => resolve();
        img.onerror = () => resolve();
      });
    }),
  );
}

/**
 * Expand a container's height to encompass any descendant content that
 * overflows with `overflow: visible`. Without this, html2canvas clips
 * the capture area to the element's own dimensions and overflowed
 * content is lost.
 */
async function expandToFitOverflow(container: HTMLElement): Promise<void> {
  await waitForImages(container);
  const containerRect = container.getBoundingClientRect();
  let maxBottom = containerRect.bottom;
  for (const el of Array.from(container.querySelectorAll("*"))) {
    const bottom = (el as HTMLElement).getBoundingClientRect().bottom;
    if (bottom > maxBottom) maxBottom = bottom;
  }
  const requiredHeight = maxBottom - containerRect.top;
  if (requiredHeight > container.offsetHeight) {
    container.style.minHeight = requiredHeight + "px";
  }
}

/**
 * Downscale and compress images in a clone so that doc.html() doesn't embed
 * them at their full intrinsic resolution (which can be 10-100× larger than
 * the displayed size). Each image is redrawn at 2× its displayed size
 * (for reasonable print quality) and converted to a compressed JPEG data URL.
 */
async function compressCloneImages(clone: HTMLElement): Promise<void> {
  const images = Array.from(clone.querySelectorAll("img"));
  if (images.length === 0) return;

  await waitForImages(clone);

  const printScale = 2;
  for (const img of images) {
    if (!img.naturalWidth || !img.naturalHeight) continue;
    if (img.src.startsWith("data:image/svg")) continue;

    const displayW = img.offsetWidth || img.naturalWidth;
    const displayH = img.offsetHeight || img.naturalHeight;

    // Target: 2× display size, but never upscale beyond natural dimensions
    const targetW = Math.min(displayW * printScale, img.naturalWidth);
    const targetH = Math.min(displayH * printScale, img.naturalHeight);

    // Skip if already at or below target size
    if (
      img.naturalWidth <= targetW &&
      img.naturalHeight <= targetH &&
      img.src.startsWith("data:")
    )
      continue;

    const canvas = document.createElement("canvas");
    canvas.width = targetW;
    canvas.height = targetH;
    const ctx = canvas.getContext("2d");
    if (!ctx) continue;

    try {
      ctx.drawImage(img, 0, 0, targetW, targetH);
      // Preserve layout dimensions before replacing src
      img.style.width = displayW + "px";
      img.style.height = displayH + "px";
      img.src = canvas.toDataURL("image/png");
    } catch {
      // Cross-origin images can't be drawn to canvas; skip
    }
  }
}

/**
 * Inject a global style reset to counteract CSS framework base styles
 * (e.g., Tailwind's `img { display: block }`) that break jsPDF/html2canvas
 * rendering. Returns a cleanup function that removes the injected style.
 */
function injectRenderResetStyles(): () => void {
  const style = document.createElement("style");
  style.setAttribute("data-jspdf-utils", "");
  style.textContent = "img { display: inline !important; }";
  document.head.appendChild(style);
  return () => style.remove();
}

/**
 * Convert HTML table attributes (cellpadding, cellspacing, border) to
 * inline CSS so doc.html()'s renderer picks them up.
 */
function normalizeTableAttributes(container: HTMLElement): void {
  for (const table of Array.from(container.querySelectorAll("table"))) {
    const cellpadding = table.getAttribute("cellpadding");
    if (cellpadding) {
      for (const cell of Array.from(table.querySelectorAll("th, td"))) {
        if (!(cell as HTMLElement).style.padding) {
          (cell as HTMLElement).style.padding = cellpadding + "px";
        }
      }
      table.removeAttribute("cellpadding");
    }
  }
}

/**
 * Split tables that exceed one page height into smaller sub-tables,
 * repeating the header row in each chunk.
 *
 * Only operates on direct-child tables of `container`.
 */
function splitOversizedTables(
  container: HTMLElement,
  pageContentPx: number,
): void {
  for (const table of Array.from(
    container.querySelectorAll<HTMLTableElement>(":scope > table"),
  )) {
    if (table.offsetHeight <= pageContentPx) continue;

    const parsed = parseTableStructure(table);
    if (!parsed) continue;

    const { headerRow, bodyRows, headerHeight } = parsed;
    const maxRowsHeight = pageContentPx - headerHeight - 2;

    const groups: HTMLTableRowElement[][] = [];
    let group: HTMLTableRowElement[] = [];
    let groupHeight = 0;

    for (const row of bodyRows) {
      const rh = row.offsetHeight;
      if (groupHeight + rh > maxRowsHeight && group.length > 0) {
        groups.push(group);
        group = [];
        groupHeight = 0;
      }
      group.push(row);
      groupHeight += rh;
    }
    if (group.length > 0) groups.push(group);

    for (const g of groups) {
      const t = table.cloneNode(false) as HTMLTableElement;
      if (headerRow) t.appendChild(headerRow.cloneNode(true));
      for (const row of g) t.appendChild(row.cloneNode(true));
      table.parentNode!.insertBefore(t, table);
    }
    table.remove();
  }
}

/** Create a hidden off-screen element for height measurement. */
function createMeasureElement(
  tag: string,
  styleAttr: string,
  width: string,
  container: HTMLElement,
): HTMLElement {
  const measure = document.createElement(tag);
  measure.setAttribute("style", styleAttr);
  Object.assign(measure.style, {
    position: "absolute",
    visibility: "hidden",
    width,
  });
  container.appendChild(measure);
  return measure;
}

/** Check whether an element uses a white-space mode that preserves newlines. */
function hasPreformattedWhiteSpace(el: HTMLElement): boolean {
  const ws = getComputedStyle(el).whiteSpace;
  return (
    ws === "pre-line" ||
    ws === "pre-wrap" ||
    ws === "pre" ||
    ws === "break-spaces"
  );
}

/** Binary-search for the maximum word count (from startIndex) that fits within maxHeight. */
function binarySearchWordFit(
  measure: HTMLElement,
  words: string[],
  maxHeight: number,
  startIndex: number,
): number {
  let lo = startIndex + 1;
  let hi = words.length;

  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    measure.textContent = words.slice(startIndex, mid).join(" ");
    if (measure.getBoundingClientRect().height <= maxHeight) {
      lo = mid;
    } else {
      hi = mid - 1;
    }
  }

  return lo;
}

/** Parse a table's header/body row structure. Returns null if the table has no rows. */
function parseTableStructure(table: HTMLTableElement): {
  rows: HTMLTableRowElement[];
  hasHeader: boolean;
  headerRow: HTMLTableRowElement | null;
  bodyRows: HTMLTableRowElement[];
  headerHeight: number;
} | null {
  const rows = Array.from(table.rows);
  if (rows.length === 0) return null;

  const hasHeader = rows[0].querySelector("th") !== null;
  const headerRow = hasHeader ? rows[0] : null;
  const bodyRows = hasHeader ? rows.slice(1) : rows;
  const headerHeight = headerRow ? headerRow.offsetHeight : 0;

  return { rows, hasHeader, headerRow, bodyRows, headerHeight };
}

/**
 * Split direct-child elements (non-table) that are taller than one page
 * into word-boundary chunks using binary search.
 *
 * The original element is kept as a wrapper (preserving its padding,
 * margin, and other box-model styles) and the chunks are placed inside
 * it as plain <div> children.
 */
function splitOversizedText(
  container: HTMLElement,
  pageContentPx: number,
): void {
  for (const el of Array.from(container.querySelectorAll(":scope > *"))) {
    const htmlEl = el as HTMLElement;
    if (htmlEl.offsetHeight <= pageContentPx || htmlEl.tagName === "TABLE")
      continue;

    // If element has child elements, recurse into it instead of flattening
    // its subtree into plain text (which would destroy nested structure).
    if (htmlEl.children.length > 0) {
      splitOversizedText(htmlEl, pageContentPx);
      continue;
    }

    const styleAttr = htmlEl.getAttribute("style") || "";
    const width = getComputedStyle(htmlEl).width;
    const preformatted = hasPreformattedWhiteSpace(htmlEl);
    // For preformatted text, tokenize into words + \n markers so that
    // line breaks are preserved while still allowing word-level splitting.
    const words = preformatted
      ? (htmlEl.textContent || "").match(/\S+|\n/g) || []
      : (htmlEl.textContent || "").split(/\s+/).filter(Boolean);

    const measure = createMeasureElement(
      htmlEl.tagName,
      styleAttr,
      width,
      container,
    );
    measure.style.padding = "0";
    measure.style.margin = "0";

    let start = 0;
    // Clear original content and replace with chunk children.
    htmlEl.textContent = "";

    while (start < words.length) {
      const lo = binarySearchWordFit(measure, words, pageContentPx, start);

      const chunk = document.createElement("div");
      chunk.textContent = words.slice(start, lo).join(" ");
      htmlEl.appendChild(chunk);
      start = lo;
    }

    measure.remove();
  }
}

/**
 * Split a table at a page boundary so the rows that fit stay on the current
 * page and the remainder starts on the next page (with the header repeated).
 * Returns true if the split was performed.
 */
function splitTableAtBoundary(
  table: HTMLTableElement,
  container: HTMLElement,
  availableHeight: number,
): boolean {
  const parsed = parseTableStructure(table);
  if (!parsed) return false;

  const { headerRow, bodyRows, headerHeight } = parsed;

  if (bodyRows.length < 2) return false;

  const maxBodyHeight = availableHeight - headerHeight - 2;
  if (maxBodyHeight <= 0) return false;

  let fitCount = 0;
  let totalHeight = 0;
  for (const row of bodyRows) {
    if (totalHeight + row.offsetHeight > maxBodyHeight) break;
    totalHeight += row.offsetHeight;
    fitCount++;
  }

  if (fitCount === 0 || fitCount === bodyRows.length) return false;

  const firstTable = table.cloneNode(false) as HTMLTableElement;
  if (headerRow) firstTable.appendChild(headerRow.cloneNode(true));
  for (let i = 0; i < fitCount; i++) {
    firstTable.appendChild(bodyRows[i].cloneNode(true));
  }

  const secondTable = table.cloneNode(false) as HTMLTableElement;
  if (headerRow) secondTable.appendChild(headerRow.cloneNode(true));
  for (let i = fitCount; i < bodyRows.length; i++) {
    secondTable.appendChild(bodyRows[i].cloneNode(true));
  }

  container.insertBefore(firstTable, table);
  container.insertBefore(secondTable, table);
  table.remove();
  return true;
}

/**
 * Split a text element at a page boundary using word-boundary binary search.
 * Returns true if the split was performed.
 */
function splitTextAtBoundary(
  el: HTMLElement,
  container: HTMLElement,
  availableHeight: number,
): boolean {
  if (el.tagName === "TABLE" || el.tagName === "IMG") return false;
  // Don't flatten elements with child elements into plain text
  if (el.children.length > 0) return false;

  const preformatted = hasPreformattedWhiteSpace(el);
  // For preformatted text, tokenize into words + \n markers so that
  // line breaks are preserved while still allowing word-level splitting.
  const words = preformatted
    ? (el.textContent || "").match(/\S+|\n/g) || []
    : (el.textContent || "").split(/\s+/).filter(Boolean);
  if (words.length < 2) return false;

  const styleAttr = el.getAttribute("style") || "";
  const width = getComputedStyle(el).width;

  const measure = createMeasureElement(el.tagName, styleAttr, width, container);
  measure.style.padding = "0";
  measure.style.margin = "0";

  measure.textContent = words[0];
  if (measure.getBoundingClientRect().height > availableHeight) {
    measure.remove();
    return false;
  }

  const lo = binarySearchWordFit(measure, words, availableHeight, 0);

  measure.remove();

  if (lo >= words.length) return false;

  // Keep the original element as a wrapper (preserving its padding/margin)
  // and place the two halves inside it as plain <div> children.
  el.textContent = "";

  const first = document.createElement("div");
  first.textContent = words.slice(0, lo).join(" ");
  el.appendChild(first);

  const second = document.createElement("div");
  second.textContent = words.slice(lo).join(" ");
  el.appendChild(second);

  return true;
}

/**
 * Insert spacer divs so that no direct child straddles a page boundary.
 * For tables and text elements, attempts to split at the boundary first
 * so content fills the current page before flowing to the next.
 */
function insertPageBreakSpacers(
  container: HTMLElement,
  pageContentPx: number,
  originY?: number,
): void {
  if (originY === undefined) {
    originY = container.getBoundingClientRect().top;
  }

  let i = 0;
  while (i < container.children.length) {
    const child = container.children[i] as HTMLElement;
    const childRect = child.getBoundingClientRect();
    const childTop = childRect.top - originY;
    const childBottom = childRect.bottom - originY;
    const pageEnd = (Math.floor(childTop / pageContentPx) + 1) * pageContentPx;

    // Use a 0.5px tolerance to avoid false positives from sub-pixel rounding
    if (childBottom > pageEnd + 0.5) {
      const remainingSpace = pageEnd - childTop;

      // Try splitting at the boundary first
      if (child.tagName === "TABLE") {
        if (
          splitTableAtBoundary(
            child as HTMLTableElement,
            container,
            remainingSpace,
          )
        ) {
          continue; // Re-check same index (now holds the first part)
        }
      } else if (child.children.length > 0) {
        // Element has child elements — recurse to paginate its children
        // instead of flattening it as text.
        insertPageBreakSpacers(child, pageContentPx, originY);
        i++;
        continue;
      } else if (splitTextAtBoundary(child, container, remainingSpace)) {
        continue; // Re-check same index
      }

      // Fallback: push to next page with spacer
      if (childRect.height <= pageContentPx) {
        const spacer = document.createElement("div");
        spacer.style.height = Math.ceil(pageEnd - childTop) + "px";
        child.parentNode!.insertBefore(spacer, child);
        i++; // Skip past the spacer
      }
    }
    i++;
  }
}

/**
 * Prepare an HTML element for doc.html() rendering.
 *
 * Clones the element, splits oversized tables/text, and inserts page-break
 * spacers. Returns the ready-to-render clone and layout metadata.
 */
function prepare(
  source: HTMLElement,
  opts: PageOptionsInput = {},
): PrepareResult {
  const merged = resolveOptions(opts);

  const removeResetStyles = injectRenderResetStyles();
  const clone = createPrintClone(source, merged.pageWidth);
  normalizeTableAttributes(clone);
  const layout = computeLayout(clone, merged);

  splitOversizedTables(clone, layout.pageContentPx);
  splitOversizedText(clone, layout.pageContentPx);
  insertPageBreakSpacers(clone, layout.pageContentPx);

  return {
    clone,
    layout,
    options: merged,
    cleanup: () => {
      clone.remove();
      removeResetStyles();
    },
  };
}

/**
 * Render an HTML element to PDF using doc.html().
 */
async function generatePDF(
  doc: jsPDF,
  source: HTMLElement,
  opts: PageOptionsInput &
    Pick<
      ImagePDFOptions,
      "marginContent" | "forcedPageCount" | "textBorder" | "border"
    > = {},
): Promise<jsPDF> {
  const { clone, layout, options, cleanup } = prepare(source, opts);

  try {
    await compressCloneImages(clone);
    await new Promise<void>((resolve) => {
      doc.html(clone, {
        callback: () => resolve(),
        width: layout.contentWidthMm,
        windowWidth: layout.renderedWidth,
        margin: [
          options.margin.top,
          options.margin.right,
          options.margin.bottom,
          options.margin.left,
        ],
      });
    });
  } finally {
    cleanup();
  }

  if (opts.forcedPageCount) {
    trimDocumentToForcedPageCount(doc, opts.forcedPageCount);
  }

  if (opts.marginContent || opts.textBorder || opts.border) {
    const originalPageOpLengths = snapshotPageStreamLengths(doc);
    await addMarginContent(
      doc,
      opts.marginContent,
      opts,
      opts.textBorder,
      opts.border,
    );
    moveAddedPageOpsToBackground(doc, originalPageOpLengths);
  }

  return doc;
}

type MarginSlot = "top" | "right" | "bottom" | "left";
type MarginResult = HTMLElement | string | null | undefined | void;
type MarginFactory = (page: number, totalPages: number) => MarginResult;

/** Convert a factory result to an HTMLElement, or null to skip. */
function resolveMarginResult(value: MarginResult): HTMLElement | null {
  if (!value) return null;
  if (value instanceof HTMLElement) return value;
  const container = document.createElement("div");
  container.innerHTML = value;
  return container;
}

export interface Border {
  /** Stroke color (default: "#000000") */
  color?: string;
  /** Line width in mm (default: 0.3) */
  width?: number;
  /** Distance in mm from the page edge to the border (default: uses page margins). */
  margin?:
    | number
    | { top?: number; right?: number; bottom?: number; left?: number };
}

export interface TextBorder {
  /** The text to repeat along all four edges. */
  text: string;
  /** Text color (default: "#000000") */
  color?: string;
  /** Font size in mm (default: 2.5) */
  fontSize?: number;
  /** Font family (default: "Arial, sans-serif") */
  fontFamily?: string;
  /** Font weight (default: "normal") */
  fontWeight?: string;
  /** Gap between repetitions in mm (default: fontSize * 0.5) */
  gap?: number;
  /** Distance in mm from the page edge to the text border (default: uses page margins). */
  margin?:
    | number
    | { top?: number; right?: number; bottom?: number; left?: number };
}

export interface MarginContentInput {
  top?: HTMLElement | string | MarginFactory;
  right?: HTMLElement | string | MarginFactory;
  bottom?: HTMLElement | string | MarginFactory;
  left?: HTMLElement | string | MarginFactory;
  /** Distance in mm from the page edge to the margin content area (default: uses page margins). */
  margin?:
    | number
    | { top?: number; right?: number; bottom?: number; left?: number };
}

export interface ImagePDFOptions {
  imageFormat?: "JPEG" | "PNG";
  imageQuality?: number;
  scale?: number;
  marginContent?: MarginContentInput;
  /** Draw a rectangle border around the content area. */
  border?: Border;
  /** Draw a repeated-text border around the content area. */
  textBorder?: TextBorder;
  /**
   * Force output to the first N pages only.
   * Example: 1 means only page 1 is generated/exported.
   */
  forcedPageCount?: number;
}

function normalizeForcedPageCount(
  forcedPageCount: number | undefined,
): number | undefined {
  if (forcedPageCount == null || !Number.isFinite(forcedPageCount)) return;
  const normalized = Math.floor(forcedPageCount);
  if (normalized < 1) return;
  return normalized;
}

function resolveTotalPages(
  actualTotalPages: number,
  forcedPageCount: number | undefined,
): number {
  const forced = normalizeForcedPageCount(forcedPageCount);
  if (!forced) return actualTotalPages;
  return Math.min(actualTotalPages, forced);
}

function trimDocumentToForcedPageCount(
  doc: jsPDF,
  forcedPageCount: number | undefined,
): void {
  const forced = normalizeForcedPageCount(forcedPageCount);
  if (!forced) return;

  const currentTotal = doc.getNumberOfPages();
  for (let page = currentTotal; page > forced; page--) {
    doc.deletePage(page);
  }
}

function getSlotRect(
  slot: MarginSlot,
  opts: PageOptions,
  margin?: Margin,
): { x: number; y: number; width: number; height: number } {
  const m = margin ?? opts.margin;
  switch (slot) {
    case "top":
      return { x: 0, y: 0, width: opts.pageWidth, height: m.top };
    case "bottom":
      return {
        x: 0,
        y: opts.pageHeight - m.bottom,
        width: opts.pageWidth,
        height: m.bottom,
      };
    case "left":
      return { x: 0, y: 0, width: m.left, height: opts.pageHeight };
    case "right":
      return {
        x: opts.pageWidth - m.right,
        y: 0,
        width: m.right,
        height: opts.pageHeight,
      };
  }
}

async function renderSlotToCanvas(
  el: HTMLElement,
  widthMm: number,
  heightMm: number,
  scale: number,
): Promise<HTMLCanvasElement> {
  // Clone so the original element is never moved out of the DOM
  const source = el.cloneNode(true) as HTMLElement;
  if (el.isConnected) {
    snapshotComputedStyles(el, source);
  }
  // Reset any off-screen positioning (e.g. hidden via absolute/fixed + offsets)
  // so the clone renders in normal flow within the wrapper
  Object.assign(source.style, {
    left: "auto",
    right: "auto",
    top: "auto",
    bottom: "auto",
  });

  // Temporarily place in DOM so CSS variables and rules resolve
  const measure = document.createElement("div");
  Object.assign(measure.style, {
    position: "fixed",
    left: "-99999px",
    top: "0",
  });
  measure.appendChild(source);
  document.body.appendChild(measure);

  const renderEl = source.cloneNode(true) as HTMLElement;
  snapshotComputedStyles(source, renderEl);
  measure.remove();

  const wrapper = document.createElement("div");
  Object.assign(wrapper.style, {
    position: "fixed",
    left: "-99999px",
    top: "0",
    width: widthMm + "mm",
    height: heightMm + "mm",
    overflow: "hidden",
  });
  wrapper.appendChild(renderEl);
  document.body.appendChild(wrapper);

  try {
    return await html2canvas(wrapper, {
      scale,
      backgroundColor: null,
    });
  } finally {
    wrapper.remove();
  }
}

const MARGIN_SLOTS: MarginSlot[] = ["top", "right", "bottom", "left"];

/** Capture current page content-stream lengths so newly appended ops can be reordered. */
function snapshotPageStreamLengths(doc: jsPDF): number[] {
  const pages = (doc as unknown as { internal?: { pages?: unknown[] } })
    .internal?.pages;
  if (!Array.isArray(pages)) return [];

  const totalPages = doc.getNumberOfPages();
  const lengths: number[] = [];
  for (let page = 1; page <= totalPages; page++) {
    const stream = pages[page];
    lengths.push(Array.isArray(stream) ? stream.length : 0);
  }
  return lengths;
}

/**
 * Reorder newly appended page operations so they render beneath existing content.
 * jsPDF paints in command order: earlier ops are visually behind later ops.
 */
function moveAddedPageOpsToBackground(
  doc: jsPDF,
  originalLengths: number[],
): void {
  const pages = (doc as unknown as { internal?: { pages?: unknown[] } })
    .internal?.pages;
  if (!Array.isArray(pages)) return;

  const totalPages = Math.min(doc.getNumberOfPages(), originalLengths.length);
  for (let page = 1; page <= totalPages; page++) {
    const stream = pages[page];
    if (!Array.isArray(stream)) continue;

    const cutoff = originalLengths[page - 1] ?? stream.length;
    if (cutoff < 0 || cutoff >= stream.length) continue;

    const existingOps = stream.slice(0, cutoff);
    const addedOps = stream.slice(cutoff);
    pages[page] = [...addedOps, ...existingOps];
  }
}

/** Pre-render static (non-function) margin content once for reuse across pages. */
async function preRenderStaticSlots(
  content: MarginContentInput,
  opts: PageOptions,
  scale: number,
): Promise<Partial<Record<MarginSlot, HTMLCanvasElement>>> {
  const contentMargin = resolveMarginOverride(
    content.margin,
    opts,
    createUniformMargin(PAGE_MARGINS[opts.format]),
  );
  const cache: Partial<Record<MarginSlot, HTMLCanvasElement>> = {};
  for (const slot of MARGIN_SLOTS) {
    const val = content[slot];
    if (val && typeof val !== "function") {
      const rect = getSlotRect(slot, opts, contentMargin);
      if (rect.width <= 0 || rect.height <= 0) continue;
      const el = resolveMarginResult(val);
      if (el) {
        cache[slot] = await renderSlotToCanvas(
          val instanceof HTMLElement
            ? (val.cloneNode(true) as HTMLElement)
            : el,
          rect.width,
          rect.height,
          scale,
        );
      }
    }
  }
  return cache;
}

/** Resolve a margin override shared by Border and TextBorder. */
function resolveMarginOverride(
  m:
    | undefined
    | number
    | { top?: number; right?: number; bottom?: number; left?: number },
  opts: PageOptions,
  fallback?: Margin,
): Margin {
  const fb = fallback ?? opts.margin;
  if (m == null) return fb;
  if (typeof m === "number") {
    return createUniformMargin(m);
  }
  return {
    top: m.top ?? fb.top,
    right: m.right ?? fb.right,
    bottom: m.bottom ?? fb.bottom,
    left: m.left ?? fb.left,
  };
}

/** Draw a repeated-text rectangle on a canvas. */
const RTL_RE =
  /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF\u0590-\u05FF]/;

async function renderTextEdgeStrip(
  text: string,
  widthPx: number,
  heightPx: number,
  fontSizePx: number,
  fontFamily: string,
  fontWeight: string,
  color: string,
  gapPx: number,
  rtl: boolean = false,
): Promise<HTMLCanvasElement> {
  const wrapper = document.createElement("div");
  Object.assign(wrapper.style, {
    position: "fixed",
    left: "-99999px",
    top: "0",
    width: `${widthPx}px`,
    height: `${heightPx}px`,
    overflow: "hidden",
    whiteSpace: "nowrap",
    fontSize: `${fontSizePx}px`,
    fontFamily,
    fontWeight,
    color,
    display: "flex",
    alignItems: "center",
    gap: `${gapPx}px`,
    direction: rtl ? "rtl" : "ltr",
  });

  const measure = document.createElement("span");
  measure.textContent = text;
  Object.assign(measure.style, {
    position: "absolute",
    visibility: "hidden",
    whiteSpace: "nowrap",
    fontSize: `${fontSizePx}px`,
    fontFamily,
    fontWeight,
  });
  document.body.appendChild(measure);
  const singleWidth = measure.offsetWidth;
  measure.remove();

  const reps = Math.ceil(widthPx / (singleWidth + gapPx)) + 2;
  for (let i = 0; i < reps; i++) {
    const span = document.createElement("span");
    span.textContent = text;
    span.style.flexShrink = "0";
    wrapper.appendChild(span);
  }

  document.body.appendChild(wrapper);
  try {
    return await html2canvas(wrapper, {
      scale: 3,
      backgroundColor: null,
      width: Math.ceil(widthPx),
      height: Math.ceil(heightPx),
    });
  } finally {
    wrapper.remove();
  }
}

async function drawTextBorderOnCanvas(
  ctx: CanvasRenderingContext2D,
  tb: TextBorder,
  pxPerMm: number,
  rectX: number,
  rectY: number,
  rectW: number,
  rectH: number,
): Promise<void> {
  const {
    text,
    color = "#000000",
    fontSize = 2.5,
    fontFamily = "Arial, sans-serif",
    fontWeight = "normal",
  } = tb;
  const fontSizePx = fontSize * pxPerMm;
  const gapPx = (tb.gap ?? fontSize * 0.5) * pxPerMm;
  const cornerGap = fontSizePx * 0.5;
  const stripHeight = Math.ceil(fontSizePx * 2.5);

  const isRtl = RTL_RE.test(text);

  const hWidth = Math.round(rectW - cornerGap * 2);
  const vWidth = Math.round(rectH - cornerGap * 2);

  // For RTL text: horizontal edges clip from the left (RTL overflow),
  // vertical edges both use RTL so overflow clips from strip LEFT.
  // After rotation: left edge clips from page BOTTOM, right edge clips from page TOP.
  const [hStrip, vStrip] = await Promise.all([
    renderTextEdgeStrip(
      text,
      hWidth,
      stripHeight,
      fontSizePx,
      fontFamily,
      fontWeight,
      color,
      gapPx,
      isRtl,
    ),
    renderTextEdgeStrip(
      text,
      vWidth,
      stripHeight,
      fontSizePx,
      fontFamily,
      fontWeight,
      color,
      gapPx,
      isRtl,
    ),
  ]);

  const hOffsetY = Math.round(stripHeight / 2);

  // Top edge
  ctx.drawImage(
    hStrip,
    0,
    0,
    hStrip.width,
    hStrip.height,
    rectX + cornerGap,
    rectY - hOffsetY,
    hWidth,
    stripHeight,
  );

  // Bottom edge
  ctx.drawImage(
    hStrip,
    0,
    0,
    hStrip.width,
    hStrip.height,
    rectX + cornerGap,
    rectY + rectH - hOffsetY,
    hWidth,
    stripHeight,
  );

  // Left edge (bottom to top) — RTL strip so overflow clips at left = page bottom
  ctx.save();
  ctx.translate(rectX, rectY + rectH - cornerGap);
  ctx.rotate(-Math.PI / 2);
  ctx.drawImage(
    vStrip,
    0,
    0,
    vStrip.width,
    vStrip.height,
    0,
    -hOffsetY,
    vWidth,
    stripHeight,
  );
  ctx.restore();

  // Right edge (top to bottom) — RTL strip so overflow clips at left = page top
  ctx.save();
  ctx.translate(rectX + rectW, rectY + cornerGap);
  ctx.rotate(Math.PI / 2);
  ctx.drawImage(
    vStrip,
    0,
    0,
    vStrip.width,
    vStrip.height,
    0,
    -hOffsetY,
    vWidth,
    stripHeight,
  );
  ctx.restore();
}

function resolveBorderMargin(
  border: Border | TextBorder,
  opts: PageOptions,
): Margin {
  return resolveMarginOverride(border.margin, opts);
}

/** Render margin content for a single page onto a canvas context. */
async function drawMarginContentOnCanvas(
  ctx: CanvasRenderingContext2D,
  content: MarginContentInput | undefined,
  staticCache: Partial<Record<MarginSlot, HTMLCanvasElement>>,
  opts: PageOptions,
  pxPerMm: number,
  page: number,
  totalPages: number,
  scale: number,
  textBorder?: TextBorder,
  border?: Border,
): Promise<void> {
  if (content) {
    const contentMargin = resolveMarginOverride(
      content.margin,
      opts,
      createUniformMargin(PAGE_MARGINS[opts.format]),
    );
    for (const slot of MARGIN_SLOTS) {
      const val = content[slot];
      if (!val) continue;

      const rect = getSlotRect(slot, opts, contentMargin);
      if (rect.width <= 0 || rect.height <= 0) continue;

      let slotCanvas: HTMLCanvasElement;

      if (typeof val === "function") {
        const el = resolveMarginResult(val(page, totalPages));
        if (!el) continue;
        slotCanvas = await renderSlotToCanvas(
          el,
          rect.width,
          rect.height,
          scale,
        );
      } else {
        if (!staticCache[slot]) continue;
        slotCanvas = staticCache[slot]!;
      }

      ctx.drawImage(
        slotCanvas,
        0,
        0,
        slotCanvas.width,
        slotCanvas.height,
        Math.round(rect.x * pxPerMm),
        Math.round(rect.y * pxPerMm),
        Math.round(rect.width * pxPerMm),
        Math.round(rect.height * pxPerMm),
      );
    }
  }

  if (border) {
    const { color = "#000000", width = 0.3 } = border;
    const bm = resolveBorderMargin(border, opts);

    ctx.strokeStyle = color;
    ctx.lineWidth = width * pxPerMm;
    ctx.strokeRect(
      Math.round(bm.left * pxPerMm),
      Math.round(bm.top * pxPerMm),
      Math.round((opts.pageWidth - bm.left - bm.right) * pxPerMm),
      Math.round((opts.pageHeight - bm.top - bm.bottom) * pxPerMm),
    );
  }

  if (textBorder) {
    const bm = resolveBorderMargin(textBorder, opts);
    await drawTextBorderOnCanvas(
      ctx,
      textBorder,
      pxPerMm,
      Math.round(bm.left * pxPerMm),
      Math.round(bm.top * pxPerMm),
      Math.round((opts.pageWidth - bm.left - bm.right) * pxPerMm),
      Math.round((opts.pageHeight - bm.top - bm.bottom) * pxPerMm),
    );
  }
}

interface PageDimensions {
  contentWidthMm: number;
  contentHeightMm: number;
  contentWidthPx: number;
  contentHeightPx: number;
  pxPerMm: number;
  pageWidthPx: number;
  pageHeightPx: number;
  marginTopPx: number;
  marginLeftPx: number;
}

/** Compute all mm/px page dimensions from a rendered canvas and page options. */
function computePageDimensions(
  canvas: HTMLCanvasElement,
  opts: PageOptions,
  layout: Layout,
  html2canvasScale: number,
): PageDimensions {
  const contentWidthMm = opts.pageWidth - opts.margin.left - opts.margin.right;
  const contentHeightMm =
    opts.pageHeight - opts.margin.top - opts.margin.bottom;
  const contentWidthPx = canvas.width;
  // Use the html2canvas scale parameter directly — NOT canvas.width/renderedWidth.
  // html2canvas positions content at (cssPosition * scale) in the canvas.
  // canvas.width = ceil(getBoundingClientRect().width * scale) which can
  // differ from offsetWidth * scale by 1-2px.  Deriving the scale from
  // canvas.width introduces a per-page slice error that accumulates and
  // eventually cuts through text lines on later pages.
  const contentHeightPx = Math.round(layout.pageContentPx * html2canvasScale);
  const pxPerMm = contentWidthPx / contentWidthMm;

  return {
    contentWidthMm,
    contentHeightMm,
    contentWidthPx,
    contentHeightPx,
    pxPerMm,
    pageWidthPx: Math.round(opts.pageWidth * pxPerMm),
    pageHeightPx: Math.round(opts.pageHeight * pxPerMm),
    marginTopPx: Math.round(opts.margin.top * pxPerMm),
    marginLeftPx: Math.round(opts.margin.left * pxPerMm),
  };
}

/** Create a single-page canvas by slicing a content region from the source canvas. */
function createPageSliceCanvas(
  sourceCanvas: HTMLCanvasElement,
  pageIndex: number,
  dims: PageDimensions,
): {
  canvas: HTMLCanvasElement;
  ctx: CanvasRenderingContext2D;
  sliceHeight: number;
} {
  const sliceHeight = Math.min(
    dims.contentHeightPx,
    sourceCanvas.height - pageIndex * dims.contentHeightPx,
  );

  const pageCanvas = document.createElement("canvas");
  pageCanvas.width = dims.pageWidthPx;
  pageCanvas.height = dims.pageHeightPx;

  const ctx = pageCanvas.getContext("2d");
  if (!ctx) throw new Error("Could not get canvas context");

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, dims.pageWidthPx, dims.pageHeightPx);

  return { canvas: pageCanvas, ctx, sliceHeight };
}

/** Draw one page's content slice into the prepared page canvas. */
function drawContentSliceOnCanvas(
  ctx: CanvasRenderingContext2D,
  sourceCanvas: HTMLCanvasElement,
  pageIndex: number,
  dims: PageDimensions,
  sliceHeight: number,
): void {
  ctx.drawImage(
    sourceCanvas,
    0,
    pageIndex * dims.contentHeightPx,
    dims.contentWidthPx,
    sliceHeight,
    dims.marginLeftPx,
    dims.marginTopPx,
    dims.contentWidthPx,
    sliceHeight,
  );
}

/** Shared setup for image-based rendering (generateImagePDF & generateImages). */
async function prepareImageRenderClone(
  source: HTMLElement,
  merged: PageOptions,
): Promise<{ clone: HTMLElement; layout: Layout; cleanup: () => void }> {
  const removeResetStyles = injectRenderResetStyles();
  const clone = createPrintClone(source, merged.pageWidth);
  clone.style.opacity = "1";
  clone.style.left = "-99999px";
  normalizeTableAttributes(clone);
  const layout = computeLayout(clone, merged);

  splitOversizedTables(clone, layout.pageContentPx);
  splitOversizedText(clone, layout.pageContentPx);
  insertPageBreakSpacers(clone, layout.pageContentPx);
  await expandToFitOverflow(clone);

  return {
    clone,
    layout,
    cleanup: () => {
      clone.remove();
      removeResetStyles();
    },
  };
}

/**
 * Render an HTML element as an image-based PDF. Each page is a rasterized
 * screenshot — no selectable or extractable text in the output.
 */
async function generateImagePDF(
  source: HTMLElement,
  opts: PageOptionsInput & ImagePDFOptions = {},
): Promise<jsPDF> {
  const { imageFormat = "JPEG", imageQuality = 0.7, scale = 3 } = opts;
  const merged = resolveOptions(opts);
  const { clone, layout, cleanup } = await prepareImageRenderClone(
    source,
    merged,
  );

  try {
    const canvas = await html2canvas(clone, {
      scale,
      backgroundColor: null,
    });

    const { jsPDF: JsPDF } = await import("jspdf");

    const dims = computePageDimensions(canvas, merged, layout, scale);
    const actualTotalPages = Math.ceil(canvas.height / dims.contentHeightPx);
    const totalPages = resolveTotalPages(
      actualTotalPages,
      opts.forcedPageCount,
    );
    const orientation = merged.pageWidth > merged.pageHeight ? "l" : "p";
    const { marginContent, textBorder, border } = opts;
    const staticCache = marginContent
      ? await preRenderStaticSlots(marginContent, merged, scale)
      : {};

    const imagePDF = new JsPDF({
      orientation,
      unit: "mm",
      format: [merged.pageWidth, merged.pageHeight],
    });

    for (let i = 0; i < totalPages; i++) {
      const {
        canvas: pageCanvas,
        ctx,
        sliceHeight,
      } = createPageSliceCanvas(canvas, i, dims);

      if (marginContent || textBorder || border) {
        await drawMarginContentOnCanvas(
          ctx,
          marginContent,
          staticCache,
          merged,
          dims.pxPerMm,
          i + 1,
          totalPages,
          scale,
          textBorder,
          border,
        );
      }

      drawContentSliceOnCanvas(ctx, canvas, i, dims, sliceHeight);

      const imageData = pageCanvas.toDataURL(
        `image/${imageFormat.toLowerCase()}`,
        imageQuality,
      );

      if (i > 0) {
        imagePDF.addPage([merged.pageWidth, merged.pageHeight], orientation);
      }

      imagePDF.addImage(
        imageData,
        imageFormat,
        0,
        0,
        merged.pageWidth,
        merged.pageHeight,
        undefined,
        "SLOW",
      );
    }

    return imagePDF;
  } finally {
    cleanup();
  }
}

/**
 * Render an HTML element to an array of page images (data URLs).
 * Each image represents a full page with margins, matching the
 * visual output of renderImagePDF.
 */
async function generateImages(
  source: HTMLElement,
  opts: PageOptionsInput & ImagePDFOptions = {},
): Promise<string[]> {
  const { imageFormat = "PNG", imageQuality = 0.75, scale = 2 } = opts;
  const merged = resolveOptions(opts);
  const { clone, layout, cleanup } = await prepareImageRenderClone(
    source,
    merged,
  );

  try {
    const canvas = await html2canvas(clone, {
      scale,
      backgroundColor: null,
    });

    const dims = computePageDimensions(canvas, merged, layout, scale);
    const actualTotalPages = Math.ceil(canvas.height / dims.contentHeightPx);
    const totalPages = resolveTotalPages(
      actualTotalPages,
      opts.forcedPageCount,
    );
    const images: string[] = [];

    const { marginContent, textBorder, border } = opts;
    const staticCache = marginContent
      ? await preRenderStaticSlots(marginContent, merged, scale)
      : {};

    for (let i = 0; i < totalPages; i++) {
      const {
        canvas: pageCanvas,
        ctx,
        sliceHeight,
      } = createPageSliceCanvas(canvas, i, dims);

      if (marginContent || textBorder || border) {
        await drawMarginContentOnCanvas(
          ctx,
          marginContent,
          staticCache,
          merged,
          dims.pxPerMm,
          i + 1,
          totalPages,
          scale,
          textBorder,
          border,
        );
      }

      drawContentSliceOnCanvas(ctx, canvas, i, dims, sliceHeight);

      images.push(
        pageCanvas.toDataURL(
          `image/${imageFormat.toLowerCase()}`,
          imageQuality,
        ),
      );
    }

    return images;
  } finally {
    cleanup();
  }
}

/**
 * Render an HTML element as page images and inject them into a scrollable
 * container. Each image is sized to match the page format dimensions.
 */
async function previewImages(
  source: HTMLElement,
  container: HTMLElement,
  opts: PageOptionsInput & ImagePDFOptions = {},
): Promise<void> {
  const merged = resolveOptions(opts);
  const images = await generateImages(source, opts);

  container.innerHTML = "";
  Object.assign(container.style, {
    display: "flex",
    flexDirection: "column",
    direction: "ltr",
    width: "fit-content",
    height: merged.pageHeight + "mm",
    maxHeight: "100vh",
    overflowY: images.length > 1 ? "auto" : "hidden",
    background: "#e0e0e0",
  });

  for (let i = 0; i < images.length; i++) {
    const img = document.createElement("img");
    img.src = images[i];
    img.alt = `Page ${i + 1}`;
    Object.assign(img.style, {
      width: merged.pageWidth + "mm",
      maxWidth: "100%",
      height: "auto",
      boxSizing: "border-box",
      border: "1px solid #bbb",
      boxShadow: "0 2px 8px rgba(0,0,0,0.2)",
      marginBottom: "16px",
    });
    img.style.setProperty("display", "inline", "important");
    container.appendChild(img);
  }
}

/**
 * Add HTML content to the margin areas of each page in a jsPDF document.
 *
 * Each slot (top, right, bottom, left) accepts either a static HTMLElement
 * (rendered once and reused on every page) or a factory function that
 * receives `(page, totalPages)` and returns an HTMLElement per page
 * (useful for page numbers or dynamic content).
 */
async function addMarginContent(
  doc: jsPDF,
  content: MarginContentInput | undefined,
  opts: PageOptionsInput = {},
  textBorder?: TextBorder,
  border?: Border,
): Promise<jsPDF> {
  const merged = resolveOptions(opts);
  const totalPages = doc.getNumberOfPages();
  const scale = 4;
  const pxPerMm = scale * (96 / 25.4);
  const removeResetStyles = injectRenderResetStyles();
  const pageWidthPx = Math.round(merged.pageWidth * pxPerMm);
  const pageHeightPx = Math.round(merged.pageHeight * pxPerMm);

  const staticCache = content
    ? await preRenderStaticSlots(content, merged, scale)
    : {};

  // Pre-convert static slot canvases to data URLs (reused across pages via alias)
  const staticDataUrls: Partial<Record<MarginSlot, string>> = {};
  for (const slot of MARGIN_SLOTS) {
    if (staticCache[slot]) {
      staticDataUrls[slot] = staticCache[slot]!.toDataURL("image/png");
    }
  }

  // Pre-render text border once (identical on every page)
  let textBorderDataUrl: string | undefined;
  if (textBorder) {
    const tbCanvas = document.createElement("canvas");
    tbCanvas.width = pageWidthPx;
    tbCanvas.height = pageHeightPx;
    const tbCtx = tbCanvas.getContext("2d");
    if (tbCtx) {
      const bm = resolveBorderMargin(textBorder, merged);
      await drawTextBorderOnCanvas(
        tbCtx,
        textBorder,
        pxPerMm,
        Math.round(bm.left * pxPerMm),
        Math.round(bm.top * pxPerMm),
        Math.round((merged.pageWidth - bm.left - bm.right) * pxPerMm),
        Math.round((merged.pageHeight - bm.top - bm.bottom) * pxPerMm),
      );
      textBorderDataUrl = tbCanvas.toDataURL("image/png");
    }
  }

  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);

    // Render each margin slot as an individual small image
    if (content) {
      const contentMargin = resolveMarginOverride(
        content.margin,
        merged,
        createUniformMargin(PAGE_MARGINS[merged.format]),
      );
      for (const slot of MARGIN_SLOTS) {
        const val = content[slot];
        if (!val) continue;

        const rect = getSlotRect(slot, merged, contentMargin);
        if (rect.width <= 0 || rect.height <= 0) continue;

        let dataUrl: string;
        let alias: string | undefined;

        if (typeof val === "function") {
          const el = resolveMarginResult(val(i, totalPages));
          if (!el) continue;
          const slotCanvas = await renderSlotToCanvas(
            el,
            rect.width,
            rect.height,
            scale,
          );
          dataUrl = slotCanvas.toDataURL("image/png");
        } else {
          dataUrl = staticDataUrls[slot]!;
          alias = `margin-${slot}`;
        }

        doc.addImage(
          dataUrl,
          "PNG",
          rect.x,
          rect.y,
          rect.width,
          rect.height,
          alias,
          "SLOW",
        );
      }
    }

    // Draw content border natively using jsPDF vector commands (zero image overhead)
    if (border) {
      const { color = "#000000", width = 0.3 } = border;
      const bm = resolveBorderMargin(border, merged);
      doc.setDrawColor(color);
      doc.setLineWidth(width);
      doc.rect(
        bm.left,
        bm.top,
        merged.pageWidth - bm.left - bm.right,
        merged.pageHeight - bm.top - bm.bottom,
      );
    }

    // Add pre-rendered text border (reused across pages via alias)
    if (textBorderDataUrl) {
      doc.addImage(
        textBorderDataUrl,
        "PNG",
        0,
        0,
        merged.pageWidth,
        merged.pageHeight,
        "text-border",
        "SLOW",
      );
    }
  }

  removeResetStyles();
  return doc;
}

export {
  PAGE_SIZES,
  PAGE_MARGINS,
  generatePDF,
  generateImagePDF,
  generateImages,
  previewImages,
};
