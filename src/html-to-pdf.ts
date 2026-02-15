/**
 * jsPDF doc.html() utilities
 *
 * Helpers that prepare an HTML element for clean, paginated PDF output
 * via jsPDF's doc.html() renderer.
 */

import type { jsPDF } from "jspdf";
import html2canvas from "html2canvas";

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

export interface PageOptions {
  unit: string;
  format: PageFormat;
  pageWidth: number;
  pageHeight: number;
  margin: Margin;
}

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

/** Resolve options: dimensions inferred from format unless explicitly provided. */
function resolveOptions(opts: Partial<PageOptions> = {}): PageOptions {
  const format = opts.format ?? "a4";
  const [defaultWidth, defaultHeight] = PAGE_SIZES[format];
  const pageWidth = opts.pageWidth ?? defaultWidth;
  const pageHeight = opts.pageHeight ?? defaultHeight;

  const defaultMargin = PAGE_MARGINS[format];
  const defaultMarginObj: Margin = {
    top: defaultMargin,
    right: defaultMargin,
    bottom: defaultMargin,
    left: defaultMargin,
  };

  return {
    unit: opts.unit ?? "mm",
    format,
    pageWidth,
    pageHeight,
    margin: { ...defaultMarginObj, ...opts.margin },
  };
}

/** Compute derived layout values from options. */
function computeLayout(container: HTMLElement, opts: PageOptions): Layout {
  const renderedWidth = container.offsetWidth;
  const contentWidthMm = opts.pageWidth - opts.margin.left - opts.margin.right;
  const scale = contentWidthMm / renderedWidth;
  const usableHeightMm = opts.pageHeight - opts.margin.top - opts.margin.bottom;
  const pageContentPx = usableHeightMm / scale;

  return { renderedWidth, scale, contentWidthMm, pageContentPx };
}

/**
 * Clone an element and position it off-screen at print width for measurement.
 */
function createPrintClone(source: HTMLElement, pageWidth = 210): HTMLElement {
  const clone = source.cloneNode(true) as HTMLElement;
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
  for (const table of container.querySelectorAll("table")) {
    const cellpadding = table.getAttribute("cellpadding");
    if (cellpadding) {
      for (const cell of table.querySelectorAll("th, td")) {
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

    const rows = Array.from(table.rows);
    if (rows.length === 0) continue;

    const hasHeader = rows[0].querySelector("th") !== null;
    const headerRow = hasHeader ? rows[0] : null;
    const bodyRows = hasHeader ? rows.slice(1) : rows;
    const headerHeight = headerRow ? headerRow.offsetHeight : 0;
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

/**
 * Split direct-child elements (non-table) that are taller than one page
 * into word-boundary chunks using binary search.
 */
function splitOversizedText(
  container: HTMLElement,
  pageContentPx: number,
): void {
  for (const el of Array.from(container.querySelectorAll(":scope > *"))) {
    const htmlEl = el as HTMLElement;
    if (htmlEl.offsetHeight <= pageContentPx || htmlEl.tagName === "TABLE")
      continue;

    const tag = htmlEl.tagName;
    const styleAttr = htmlEl.getAttribute("style") || "";
    const width = getComputedStyle(htmlEl).width;
    const words = (htmlEl.textContent || "").split(/\s+/).filter(Boolean);

    const measure = document.createElement(tag);
    measure.setAttribute("style", styleAttr);
    Object.assign(measure.style, {
      position: "absolute",
      visibility: "hidden",
      width,
    });
    container.appendChild(measure);

    const chunks: HTMLElement[] = [];
    let start = 0;

    while (start < words.length) {
      let lo = start + 1;
      let hi = words.length;

      while (lo < hi) {
        const mid = Math.ceil((lo + hi) / 2);
        measure.textContent = words.slice(start, mid).join(" ");
        if (measure.offsetHeight <= pageContentPx) {
          lo = mid;
        } else {
          hi = mid - 1;
        }
      }

      const chunk = document.createElement(tag);
      chunk.setAttribute("style", styleAttr);
      chunk.textContent = words.slice(start, lo).join(" ");
      chunks.push(chunk);
      start = lo;
    }

    measure.remove();

    for (const chunk of chunks) {
      htmlEl.parentNode!.insertBefore(chunk, htmlEl);
    }
    htmlEl.remove();
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
  const rows = Array.from(table.rows);
  if (rows.length === 0) return false;

  const hasHeader = rows[0].querySelector("th") !== null;
  const headerRow = hasHeader ? rows[0] : null;
  const bodyRows = hasHeader ? rows.slice(1) : rows;
  const headerHeight = headerRow ? headerRow.offsetHeight : 0;

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
  const words = (el.textContent || "").split(/\s+/).filter(Boolean);
  if (words.length < 2) return false;

  const tag = el.tagName;
  const styleAttr = el.getAttribute("style") || "";
  const width = getComputedStyle(el).width;

  const measure = document.createElement(tag);
  measure.setAttribute("style", styleAttr);
  Object.assign(measure.style, {
    position: "absolute",
    visibility: "hidden",
    width,
  });
  container.appendChild(measure);

  measure.textContent = words[0];
  if (measure.offsetHeight > availableHeight) {
    measure.remove();
    return false;
  }

  let lo = 1;
  let hi = words.length;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    measure.textContent = words.slice(0, mid).join(" ");
    if (measure.offsetHeight <= availableHeight) {
      lo = mid;
    } else {
      hi = mid - 1;
    }
  }

  measure.remove();

  if (lo >= words.length) return false;

  const first = document.createElement(tag);
  first.setAttribute("style", styleAttr);
  first.textContent = words.slice(0, lo).join(" ");

  const second = document.createElement(tag);
  second.setAttribute("style", styleAttr);
  second.textContent = words.slice(lo).join(" ");

  container.insertBefore(first, el);
  container.insertBefore(second, el);
  el.remove();
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
): void {
  let i = 0;
  while (i < container.children.length) {
    const child = container.children[i] as HTMLElement;
    const childTop = child.offsetTop;
    const childBottom = childTop + child.offsetHeight;
    const pageEnd = (Math.floor(childTop / pageContentPx) + 1) * pageContentPx;

    if (childBottom > pageEnd) {
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
      } else if (splitTextAtBoundary(child, container, remainingSpace)) {
        continue; // Re-check same index
      }

      // Fallback: push to next page with spacer
      if (child.offsetHeight <= pageContentPx) {
        const spacer = document.createElement("div");
        spacer.style.height = pageEnd - childTop + 1 + "px";
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
  opts: Partial<PageOptions> = {},
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
async function renderHTML(
  doc: jsPDF,
  source: HTMLElement,
  opts: Partial<PageOptions> & Pick<ImagePDFOptions, "marginContent"> = {},
): Promise<jsPDF> {
  const { clone, layout, options, cleanup } = prepare(source, opts);

  try {
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

  if (opts.marginContent) {
    await addMarginContent(doc, opts.marginContent, opts);
  }

  return doc;
}

type MarginSlot = "top" | "right" | "bottom" | "left";
type MarginFactory = (page: number, totalPages: number) => HTMLElement;

export interface ContentBorder {
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
  /** Gap between repetitions in mm (default: fontSize * 0.5) */
  gap?: number;
  /** Distance in mm from the page edge to the text border (default: uses page margins). */
  margin?:
    | number
    | { top?: number; right?: number; bottom?: number; left?: number };
}

export interface MarginContentInput {
  top?: HTMLElement | MarginFactory;
  right?: HTMLElement | MarginFactory;
  bottom?: HTMLElement | MarginFactory;
  left?: HTMLElement | MarginFactory;
  /** Draw a rectangle border around the content area. */
  contentBorder?: ContentBorder;
  /** Draw a repeated-text border around the content area. */
  textBorder?: TextBorder;
}

export interface ImagePDFOptions {
  imageFormat?: "JPEG" | "PNG";
  imageQuality?: number;
  scale?: number;
  marginContent?: MarginContentInput;
}

function getSlotRect(
  slot: MarginSlot,
  opts: PageOptions,
): { x: number; y: number; width: number; height: number } {
  switch (slot) {
    case "top":
      return { x: 0, y: 0, width: opts.pageWidth, height: opts.margin.top };
    case "bottom":
      return {
        x: 0,
        y: opts.pageHeight - opts.margin.bottom,
        width: opts.pageWidth,
        height: opts.margin.bottom,
      };
    case "left":
      return { x: 0, y: 0, width: opts.margin.left, height: opts.pageHeight };
    case "right":
      return {
        x: opts.pageWidth - opts.margin.right,
        y: 0,
        width: opts.margin.right,
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
  const wrapper = document.createElement("div");
  Object.assign(wrapper.style, {
    position: "fixed",
    left: "-99999px",
    top: "0",
    width: widthMm + "mm",
    height: heightMm + "mm",
    overflow: "hidden",
  });
  wrapper.appendChild(el);
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

/** Pre-render static (non-function) margin content once for reuse across pages. */
async function preRenderStaticSlots(
  content: MarginContentInput,
  opts: PageOptions,
  scale: number,
): Promise<Partial<Record<MarginSlot, HTMLCanvasElement>>> {
  const cache: Partial<Record<MarginSlot, HTMLCanvasElement>> = {};
  for (const slot of MARGIN_SLOTS) {
    const val = content[slot];
    if (val && typeof val !== "function") {
      const rect = getSlotRect(slot, opts);
      cache[slot] = await renderSlotToCanvas(
        val.cloneNode(true) as HTMLElement,
        rect.width,
        rect.height,
        scale,
      );
    }
  }
  return cache;
}

/** Resolve a margin override shared by ContentBorder and TextBorder. */
function resolveMarginOverride(
  m:
    | undefined
    | number
    | { top?: number; right?: number; bottom?: number; left?: number },
  opts: PageOptions,
): Margin {
  if (m == null) return opts.margin;
  if (typeof m === "number") {
    return { top: m, right: m, bottom: m, left: m };
  }
  return {
    top: m.top ?? opts.margin.top,
    right: m.right ?? opts.margin.right,
    bottom: m.bottom ?? opts.margin.bottom,
    left: m.left ?? opts.margin.left,
  };
}

/** Draw a repeated-text rectangle on a canvas. */
function drawTextBorderOnCanvas(
  ctx: CanvasRenderingContext2D,
  tb: TextBorder,
  pxPerMm: number,
  rectX: number,
  rectY: number,
  rectW: number,
  rectH: number,
): void {
  const {
    text,
    color = "#000000",
    fontSize = 2.5,
    fontFamily = "Arial, sans-serif",
  } = tb;
  const fontSizePx = fontSize * pxPerMm;
  const gapPx = (tb.gap ?? fontSize * 0.5) * pxPerMm;

  ctx.save();
  ctx.fillStyle = color;
  ctx.font = `${fontSizePx}px ${fontFamily}`;
  ctx.textBaseline = "middle";

  const textWidth = ctx.measureText(text).width;
  const segmentWidth = textWidth + gapPx;
  const cornerGap = fontSizePx * 0.2;

  /** Draw repeated text along an edge, trimming the last segment at a
   *  whole-character boundary instead of clipping letters in half. */
  const drawEdge = (
    start: number,
    end: number,
    draw: (pos: number, t: string) => void,
  ) => {
    for (let pos = start; pos < end; pos += segmentWidth) {
      if (pos + textWidth <= end) {
        draw(pos, text);
      } else {
        // Find longest prefix that fits without cutting a letter
        for (let c = text.length - 1; c >= 1; c--) {
          const sub = text.substring(0, c);
          if (pos + ctx.measureText(sub).width <= end) {
            draw(pos, sub);
            break;
          }
        }
      }
    }
  };

  const hStart = rectX + cornerGap;
  const hEnd = rectX + rectW - cornerGap;

  // Top edge (left to right)
  drawEdge(hStart, hEnd, (pos, t) => ctx.fillText(t, pos, rectY));

  // Bottom edge (left to right)
  drawEdge(hStart, hEnd, (pos, t) => ctx.fillText(t, pos, rectY + rectH));

  // Left edge (bottom to top)
  ctx.save();
  ctx.translate(rectX, rectY + rectH);
  ctx.rotate(-Math.PI / 2);
  drawEdge(cornerGap, rectH - cornerGap, (pos, t) => ctx.fillText(t, pos, 0));
  ctx.restore();

  // Right edge (top to bottom)
  ctx.save();
  ctx.translate(rectX + rectW, rectY);
  ctx.rotate(Math.PI / 2);
  drawEdge(cornerGap, rectH - cornerGap, (pos, t) => ctx.fillText(t, pos, 0));
  ctx.restore();

  ctx.restore();
}

function resolveBorderMargin(
  border: ContentBorder | TextBorder,
  opts: PageOptions,
): Margin {
  return resolveMarginOverride(border.margin, opts);
}

/** Render margin content for a single page onto a canvas context. */
async function drawMarginContentOnCanvas(
  ctx: CanvasRenderingContext2D,
  content: MarginContentInput,
  staticCache: Partial<Record<MarginSlot, HTMLCanvasElement>>,
  opts: PageOptions,
  pxPerMm: number,
  page: number,
  totalPages: number,
  scale: number,
): Promise<void> {
  for (const slot of MARGIN_SLOTS) {
    const val = content[slot];
    if (!val) continue;

    const rect = getSlotRect(slot, opts);
    let slotCanvas: HTMLCanvasElement;

    if (typeof val === "function") {
      slotCanvas = await renderSlotToCanvas(
        val(page, totalPages),
        rect.width,
        rect.height,
        scale,
      );
    } else {
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

  if (content.contentBorder) {
    const { color = "#000000", width = 0.3 } = content.contentBorder;
    const bm = resolveBorderMargin(content.contentBorder, opts);

    ctx.strokeStyle = color;
    ctx.lineWidth = width * pxPerMm;
    ctx.strokeRect(
      Math.round(bm.left * pxPerMm),
      Math.round(bm.top * pxPerMm),
      Math.round((opts.pageWidth - bm.left - bm.right) * pxPerMm),
      Math.round((opts.pageHeight - bm.top - bm.bottom) * pxPerMm),
    );
  }

  if (content.textBorder) {
    const bm = resolveBorderMargin(content.textBorder, opts);
    drawTextBorderOnCanvas(
      ctx,
      content.textBorder,
      pxPerMm,
      Math.round(bm.left * pxPerMm),
      Math.round(bm.top * pxPerMm),
      Math.round((opts.pageWidth - bm.left - bm.right) * pxPerMm),
      Math.round((opts.pageHeight - bm.top - bm.bottom) * pxPerMm),
    );
  }
}

/**
 * Render an HTML element as an image-based PDF. Each page is a rasterized
 * screenshot — no selectable or extractable text in the output.
 */
async function renderImagePDF(
  source: HTMLElement,
  opts: Partial<PageOptions> & ImagePDFOptions = {},
): Promise<jsPDF> {
  const { imageFormat = "JPEG", imageQuality = 1, scale = 2 } = opts;
  const merged = resolveOptions(opts);

  const removeResetStyles = injectRenderResetStyles();
  const clone = createPrintClone(source, merged.pageWidth);
  clone.style.opacity = "1";
  clone.style.left = "-99999px";
  normalizeTableAttributes(clone);
  const layout = computeLayout(clone, merged);

  splitOversizedTables(clone, layout.pageContentPx);
  splitOversizedText(clone, layout.pageContentPx);
  insertPageBreakSpacers(clone, layout.pageContentPx);

  try {
    const canvas = await html2canvas(clone, {
      scale,
      backgroundColor: "#ffffff",
    });

    const { jsPDF: JsPDF } = await import("jspdf");

    const contentWidthMm =
      merged.pageWidth - merged.margin.left - merged.margin.right;
    const contentHeightMm =
      merged.pageHeight - merged.margin.top - merged.margin.bottom;

    const contentWidthPx = canvas.width;
    const contentHeightPx = (contentHeightMm / contentWidthMm) * contentWidthPx;

    const totalPages = Math.ceil(canvas.height / contentHeightPx);
    const orientation = merged.pageWidth > merged.pageHeight ? "l" : "p";

    const imagePDF = new JsPDF({
      orientation,
      unit: "mm",
      format: [merged.pageWidth, merged.pageHeight],
    });

    for (let i = 0; i < totalPages; i++) {
      const sliceHeight = Math.min(
        contentHeightPx,
        canvas.height - i * contentHeightPx,
      );

      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = contentWidthPx;
      pageCanvas.height = sliceHeight;

      const ctx = pageCanvas.getContext("2d");
      if (!ctx) throw new Error("Could not get canvas context");

      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, contentWidthPx, sliceHeight);

      ctx.drawImage(
        canvas,
        0,
        i * contentHeightPx,
        contentWidthPx,
        sliceHeight,
        0,
        0,
        contentWidthPx,
        sliceHeight,
      );

      const imageData = pageCanvas.toDataURL(
        `image/${imageFormat.toLowerCase()}`,
        imageQuality,
      );

      if (i > 0) {
        imagePDF.addPage([merged.pageWidth, merged.pageHeight], orientation);
      }

      const sliceHeightMm = (sliceHeight / contentWidthPx) * contentWidthMm;

      imagePDF.addImage(
        imageData,
        imageFormat,
        merged.margin.left,
        merged.margin.top,
        contentWidthMm,
        sliceHeightMm,
        undefined,
        "FAST",
      );
    }

    if (opts.marginContent) {
      await addMarginContent(imagePDF, opts.marginContent, opts);
    }

    return imagePDF;
  } finally {
    clone.remove();
    removeResetStyles();
  }
}

/**
 * Render an HTML element to an array of page images (data URLs).
 * Each image represents a full page with margins, matching the
 * visual output of renderImagePDF.
 */
async function renderPageImages(
  source: HTMLElement,
  opts: Partial<PageOptions> & ImagePDFOptions = {},
): Promise<string[]> {
  const { imageFormat = "PNG", imageQuality = 1, scale = 2 } = opts;
  const merged = resolveOptions(opts);

  const removeResetStyles = injectRenderResetStyles();
  const clone = createPrintClone(source, merged.pageWidth);
  clone.style.opacity = "1";
  clone.style.left = "-99999px";
  normalizeTableAttributes(clone);
  const layout = computeLayout(clone, merged);

  splitOversizedTables(clone, layout.pageContentPx);
  splitOversizedText(clone, layout.pageContentPx);
  insertPageBreakSpacers(clone, layout.pageContentPx);

  try {
    const canvas = await html2canvas(clone, {
      scale,
      backgroundColor: "#ffffff",
    });

    const contentWidthMm =
      merged.pageWidth - merged.margin.left - merged.margin.right;
    const contentHeightMm =
      merged.pageHeight - merged.margin.top - merged.margin.bottom;

    const contentWidthPx = canvas.width;
    const contentHeightPx = (contentHeightMm / contentWidthMm) * contentWidthPx;

    // Compute full page dimensions in pixels (including margins)
    const pxPerMm = contentWidthPx / contentWidthMm;
    const pageWidthPx = Math.round(merged.pageWidth * pxPerMm);
    const pageHeightPx = Math.round(merged.pageHeight * pxPerMm);
    const marginTopPx = Math.round(merged.margin.top * pxPerMm);
    const marginLeftPx = Math.round(merged.margin.left * pxPerMm);

    const totalPages = Math.ceil(canvas.height / contentHeightPx);
    const images: string[] = [];

    const { marginContent } = opts;
    const staticCache = marginContent
      ? await preRenderStaticSlots(marginContent, merged, scale)
      : {};

    for (let i = 0; i < totalPages; i++) {
      const sliceHeight = Math.min(
        contentHeightPx,
        canvas.height - i * contentHeightPx,
      );

      const pageCanvas = document.createElement("canvas");
      pageCanvas.width = pageWidthPx;
      pageCanvas.height = pageHeightPx;

      const ctx = pageCanvas.getContext("2d");
      if (!ctx) throw new Error("Could not get canvas context");

      // Fill with white (the full page background)
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, pageWidthPx, pageHeightPx);

      // Draw the content slice at the margin offset
      ctx.drawImage(
        canvas,
        0,
        i * contentHeightPx,
        contentWidthPx,
        sliceHeight,
        marginLeftPx,
        marginTopPx,
        contentWidthPx,
        sliceHeight,
      );

      // Draw margin content (headers, footers, borders)
      if (marginContent) {
        await drawMarginContentOnCanvas(
          ctx,
          marginContent,
          staticCache,
          merged,
          pxPerMm,
          i + 1,
          totalPages,
          scale,
        );
      }

      images.push(
        pageCanvas.toDataURL(
          `image/${imageFormat.toLowerCase()}`,
          imageQuality,
        ),
      );
    }

    return images;
  } finally {
    clone.remove();
    removeResetStyles();
  }
}

/**
 * Render an HTML element as page images and inject them into a scrollable
 * container. Each image is sized to match the page format dimensions.
 */
async function previewPageImages(
  source: HTMLElement,
  container: HTMLElement,
  opts: Partial<PageOptions> & ImagePDFOptions = {},
): Promise<void> {
  const merged = resolveOptions(opts);
  const images = await renderPageImages(source, opts);

  container.innerHTML = "";
  Object.assign(container.style, {
    direction: "ltr",
    width: "fit-content",
    height: merged.pageHeight + "mm",
    maxHeight: "100vh",
    overflowY: "auto",
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
      display: "block",
      border: "1px solid #bbb",
      boxShadow: "0 2px 8px rgba(0,0,0,0.2)",
      marginBottom: "16px",
    });
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
  content: MarginContentInput,
  opts: Partial<PageOptions> = {},
): Promise<jsPDF> {
  const merged = resolveOptions(opts);
  const totalPages = doc.getNumberOfPages();
  const scale = 2;
  const pxPerMm = scale * (96 / 25.4);
  const pageWidthPx = Math.round(merged.pageWidth * pxPerMm);
  const pageHeightPx = Math.round(merged.pageHeight * pxPerMm);

  const staticCache = await preRenderStaticSlots(content, merged, scale);

  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);

    const pageCanvas = document.createElement("canvas");
    pageCanvas.width = pageWidthPx;
    pageCanvas.height = pageHeightPx;
    const ctx = pageCanvas.getContext("2d");
    if (!ctx) continue;

    await drawMarginContentOnCanvas(
      ctx,
      content,
      staticCache,
      merged,
      pxPerMm,
      i,
      totalPages,
      scale,
    );

    doc.addImage(
      pageCanvas.toDataURL("image/png"),
      "PNG",
      0,
      0,
      merged.pageWidth,
      merged.pageHeight,
    );
  }

  return doc;
}

export {
  PAGE_SIZES,
  PAGE_MARGINS,
  computeLayout,
  createPrintClone,
  normalizeTableAttributes,
  splitOversizedTables,
  splitOversizedText,
  insertPageBreakSpacers,
  prepare,
  renderHTML,
  renderImagePDF,
  renderPageImages,
  previewPageImages,
  addMarginContent,
};
