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

export interface PageOptions {
  unit: string;
  format: string;
  pageWidth: number;
  pageHeight: number;
  margin: Margin;
}

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

/** Default page layout (A4, millimetres). */
const DEFAULT_OPTIONS: PageOptions = {
  unit: "mm",
  format: "a4",
  pageWidth: 210,
  pageHeight: 297,
  margin: { top: 20, right: 20, bottom: 20, left: 20 },
};

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
function createPrintClone(source: HTMLElement): HTMLElement {
  const clone = source.cloneNode(true) as HTMLElement;
  Object.assign(clone.style, {
    position: "fixed",
    top: "0",
    left: "0",
    boxSizing: "border-box",
    width: "210mm",
    opacity: "0.000001",
    pointerEvents: "none",
  });
  document.body.appendChild(clone);
  return clone;
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
  const merged: PageOptions = {
    ...DEFAULT_OPTIONS,
    ...opts,
    margin: { ...DEFAULT_OPTIONS.margin, ...opts.margin },
  };

  const clone = createPrintClone(source);
  normalizeTableAttributes(clone);
  const layout = computeLayout(clone, merged);

  splitOversizedTables(clone, layout.pageContentPx);
  splitOversizedText(clone, layout.pageContentPx);
  insertPageBreakSpacers(clone, layout.pageContentPx);

  return {
    clone,
    layout,
    options: merged,
    cleanup: () => clone.remove(),
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

  const merged: PageOptions = {
    ...DEFAULT_OPTIONS,
    ...opts,
    margin: { ...DEFAULT_OPTIONS.margin, ...opts.margin },
  };

  const clone = createPrintClone(source);
  clone.style.opacity = "1";
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
  DEFAULT_OPTIONS,
  computeLayout,
  createPrintClone,
  normalizeTableAttributes,
  splitOversizedTables,
  splitOversizedText,
  insertPageBreakSpacers,
  prepare,
  renderHTML,
  renderImagePDF,
};
