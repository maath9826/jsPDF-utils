/**
 * jsPDF doc.html() utilities
 *
 * Helpers that prepare an HTML element for clean, paginated PDF output
 * via jsPDF's doc.html() renderer.
 */

/**
 * Default page layout (A4, millimetres).
 */
const DEFAULT_OPTIONS = {
  unit: "mm",
  format: "a4",
  pageWidth: 210,
  pageHeight: 297,
  margin: { top: 20, right: 20, bottom: 20, left: 20 },
};

/**
 * Compute derived layout values from options.
 *
 * @param {HTMLElement} container - The positioned container to measure.
 * @param {object}      opts     - Merged options (DEFAULT_OPTIONS + user).
 * @returns {{ renderedWidth: number, scale: number, contentWidthMm: number, pageContentPx: number }}
 */
function computeLayout(container, opts) {
  const renderedWidth = container.offsetWidth;
  const contentWidthMm =
    opts.pageWidth - opts.margin.left - opts.margin.right;
  const scale = contentWidthMm / renderedWidth;
  const usableHeightMm = opts.pageHeight - opts.margin.top - opts.margin.bottom;
  const pageContentPx = usableHeightMm / scale;

  return { renderedWidth, scale, contentWidthMm, pageContentPx };
}

/**
 * Clone an element and position it off-screen at print width for measurement.
 *
 * @param {HTMLElement} source - The element to clone.
 * @returns {HTMLElement} The positioned clone (appended to document.body).
 */
function createPrintClone(source) {
  const clone = source.cloneNode(true);
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
 *
 * @param {HTMLElement} container - The print container.
 */
function normalizeTableAttributes(container) {
  for (const table of container.querySelectorAll("table")) {
    const cellpadding = table.getAttribute("cellpadding");
    if (cellpadding) {
      for (const cell of table.querySelectorAll("th, td")) {
        if (!cell.style.padding) {
          cell.style.padding = cellpadding + "px";
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
 *
 * @param {HTMLElement} container     - The print container.
 * @param {number}      pageContentPx - Usable page height in CSS pixels.
 */
function splitOversizedTables(container, pageContentPx) {
  for (const table of Array.from(
    container.querySelectorAll(":scope > table")
  )) {
    if (table.offsetHeight <= pageContentPx) continue;

    const rows = Array.from(table.rows);
    if (rows.length === 0) continue;

    const hasHeader = rows[0].querySelector("th") !== null;
    const headerRow = hasHeader ? rows[0] : null;
    const bodyRows = hasHeader ? rows.slice(1) : rows;
    const headerHeight = headerRow ? headerRow.offsetHeight : 0;
    const maxRowsHeight = pageContentPx - headerHeight - 2;

    const groups = [];
    let group = [];
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
      const t = table.cloneNode(false);
      if (headerRow) t.appendChild(headerRow.cloneNode(true));
      for (const row of g) t.appendChild(row.cloneNode(true));
      table.parentNode.insertBefore(t, table);
    }
    table.remove();
  }
}

/**
 * Split direct-child elements (non-table) that are taller than one page
 * into word-boundary chunks using binary search.
 *
 * @param {HTMLElement} container     - The print container.
 * @param {number}      pageContentPx - Usable page height in CSS pixels.
 */
function splitOversizedText(container, pageContentPx) {
  for (const el of Array.from(container.querySelectorAll(":scope > *"))) {
    if (el.offsetHeight <= pageContentPx || el.tagName === "TABLE") continue;

    const tag = el.tagName;
    const styleAttr = el.getAttribute("style") || "";
    const width = getComputedStyle(el).width;
    const words = el.textContent.split(/\s+/).filter(Boolean);

    const measure = document.createElement(tag);
    measure.setAttribute("style", styleAttr);
    Object.assign(measure.style, {
      position: "absolute",
      visibility: "hidden",
      width,
    });
    container.appendChild(measure);

    const chunks = [];
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
      el.parentNode.insertBefore(chunk, el);
    }
    el.remove();
  }
}

/**
 * Insert spacer divs so that no direct child straddles a page boundary.
 *
 * @param {HTMLElement} container     - The print container.
 * @param {number}      pageContentPx - Usable page height in CSS pixels.
 */
function insertPageBreakSpacers(container, pageContentPx) {
  const children = Array.from(container.children);
  for (const child of children) {
    const childTop = child.offsetTop;
    const childBottom = childTop + child.offsetHeight;
    const pageEnd =
      (Math.floor(childTop / pageContentPx) + 1) * pageContentPx;

    if (childBottom > pageEnd && child.offsetHeight <= pageContentPx) {
      const spacer = document.createElement("div");
      spacer.style.height = pageEnd - childTop + 1 + "px";
      child.parentNode.insertBefore(spacer, child);
    }
  }
}

/**
 * Prepare an HTML element for doc.html() rendering.
 *
 * Clones the element, splits oversized tables/text, and inserts page-break
 * spacers. Returns the ready-to-render clone and layout metadata.
 *
 * @param {HTMLElement} source  - The source element to prepare.
 * @param {object}      [opts] - Override any key from DEFAULT_OPTIONS.
 * @returns {{ clone: HTMLElement, layout: object, cleanup: () => void }}
 */
function prepare(source, opts = {}) {
  const merged = {
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
 *
 * @param {jsPDF}       doc     - An existing jsPDF instance.
 * @param {HTMLElement}  source - The HTML element to render.
 * @param {object}       [opts] - Override any key from DEFAULT_OPTIONS.
 * @returns {Promise<jsPDF>} Resolves with the jsPDF instance once rendering is done.
 */
async function renderHTML(doc, source, opts = {}) {
  const { clone, layout, options, cleanup } = prepare(source, opts);

  try {
    await new Promise((resolve) => {
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
};
