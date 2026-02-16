# jsPDF Utils

Utilities for rendering HTML into paginated PDF output with `jsPDF` and
`html2canvas-pro`.

## Installation

```bash
npm install jspdf-utils jspdf html2canvas-pro
```

## Exported API

- `generatePDF(doc, source, opts)`
- `generateImagePDF(source, opts)`
- `generateImages(source, opts)`
- `previewImages(source, container, opts)`
- `PAGE_SIZES`
- `PAGE_MARGINS`
- Types: `PageOptions`, `PageOptionsInput`, `ImagePDFOptions`

Type import example:

```ts
import type { PageOptionsInput } from "jspdf-utils";
```

## Quick Start

### 1) HTML -> vector/text PDF (`doc.html`)

```ts
import jsPDF from "jspdf";
import { generatePDF } from "jspdf-utils";

const target = document.getElementById("print-section");
if (!target) throw new Error("Missing #print-section");

const doc = new jsPDF({ unit: "mm" });

// Optional for Arabic/RTL text:
// doc.addFont("/fonts/arial.ttf", "arial", "normal");
// doc.addFont("/fonts/arial-bold.ttf", "arial", "bold");

await generatePDF(doc, target, {
  format: "a4",
  margin: { top: 20, right: 20, bottom: 20, left: 20 },
  forcedPageCount: 1,
});

doc.save("output.pdf");
```

### 2) HTML -> image-based PDF (raster pages)

```ts
import { generateImagePDF } from "jspdf-utils";

const target = document.getElementById("print-section");
if (!target) throw new Error("Missing #print-section");

const imagePDF = await generateImagePDF(target, {
  format: "a5",
  imageFormat: "PNG",
  forcedPageCount: 1,
});

imagePDF.save("output-image.pdf");
```

### 3) Preview pages as images in a container

```ts
import { previewImages } from "jspdf-utils";

const target = document.getElementById("print-section");
const preview = document.getElementById("preview-container");
if (!target || !preview) throw new Error("Missing preview elements");

await previewImages(target, preview, {
  format: "a5",
  forcedPageCount: 1,
});
```

## Options

### `PageOptionsInput`

- `unit?: string` (default: `"mm"`)
- `format?: "a0" | "a1" | "a2" | "a3" | "a4" | "a5" | "a6" | "letter" | "legal" | "tabloid"` (default: `"a4"`)
- `pageWidth?: number` (default comes from `format`)
- `pageHeight?: number` (default comes from `format`)
- `margin?: number | { top?: number; right?: number; bottom?: number; left?: number }`

Important:

- `generatePDF`, `generateImagePDF`, `generateImages`, and `previewImages` use
  page sizing from their `opts` (`format` / `pageWidth` / `pageHeight`).
- Do not rely on `new jsPDF({ format: ... })` to control layout in
  `generatePDF`; pass `format` in `opts` instead.

### `ImagePDFOptions`

- `imageFormat?: "JPEG" | "PNG"`
- `imageQuality?: number`
- `scale?: number`
- `marginContent?: MarginContentInput`
- `forcedPageCount?: number`

`forcedPageCount` behavior:

- Forces output to the first `N` pages only.
- `generatePDF`: trims extra pages after `doc.html` rendering.
- `generateImagePDF`: only rasterizes and writes first `N` pages.
- `generateImages` and `previewImages`: only returns/displays first `N` pages.
- Invalid values (`<= 0`, `NaN`, `Infinity`) are ignored.

## Margin Content and Borders

`marginContent` supports:

- `top`, `right`, `bottom`, `left` as:
  - `HTMLElement`, or
  - `(page: number, totalPages: number) => HTMLElement`
- `contentBorder` (vector rectangle)
- `textBorder` (repeated text around page edges)

Rendering order:

- Margin content and borders are rendered beneath page content.
- Main document content stays visually above borders/text borders.

## Development

```bash
npm install
npm run dev
```

Open [http://localhost:5173](http://localhost:5173).

## License

MIT
