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
  | "a0" | "a1" | "a2" | "a3" | "a4" | "a5" | "a6"
  | "letter" | "legal" | "tabloid";

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
 * Reset framework/global CSS on table elements inside the clone so that
 * jsPDF's doc.html() renderer produces consistent output regardless of
 * the host page's CSS environment (e.g. Tailwind preflight, CSS resets).
 *
 * Uses `all: revert` to roll back author-stylesheet properties to
 * user-agent defaults while preserving inline styles set by the caller.
 *
 * Returns a cleanup function that removes the injected style element.
 */
function resetFrameworkCSS(clone: HTMLElement): () => void {
  const uid = "__jspdf_" + Math.random().toString(36).slice(2, 8);
  clone.dataset.jspdfClone = uid;

  const sel = `[data-jspdf-clone="${uid}"]`;
  const tags = ["table", "thead", "tbody", "tfoot", "tr", "td", "th"];
  const rule = tags.map((t) => `${sel} ${t}`).join(",") + "{all:revert}";

  const style = document.createElement("style");
  style.textContent = rule;
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

/** Insert spacer divs so that no direct child straddles a page boundary. */
function insertPageBreakSpacers(
  container: HTMLElement,
  pageContentPx: number,
): void {
  const children = Array.from(container.children) as HTMLElement[];
  for (const child of children) {
    const childTop = child.offsetTop;
    const childBottom = childTop + child.offsetHeight;
    const pageEnd = (Math.floor(childTop / pageContentPx) + 1) * pageContentPx;

    if (childBottom > pageEnd && child.offsetHeight <= pageContentPx) {
      const spacer = document.createElement("div");
      spacer.style.height = pageEnd - childTop + 1 + "px";
      child.parentNode!.insertBefore(spacer, child);
    }
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

  const clone = createPrintClone(source, merged.pageWidth);
  const cleanupCSS = resetFrameworkCSS(clone);
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
      cleanupCSS();
      clone.remove();
    },
  };
}

/**
 * Render an HTML element to PDF using doc.html().
 */
async function renderHTML(
  doc: jsPDF,
  source: HTMLElement,
  opts: Partial<PageOptions> = {},
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

  return doc;
}

export interface ImagePDFOptions {
  imageFormat?: "JPEG" | "PNG";
  imageQuality?: number;
  scale?: number;
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

    return imagePDF;
  } finally {
    clone.remove();
  }
}

export {
  PAGE_SIZES,
  PAGE_MARGINS,
  computeLayout,
  createPrintClone,
  resetFrameworkCSS,
  normalizeTableAttributes,
  splitOversizedTables,
  splitOversizedText,
  insertPageBreakSpacers,
  prepare,
  renderHTML,
  renderImagePDF,
};
