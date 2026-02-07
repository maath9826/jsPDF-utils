# jsPDF Utils

HTML to PDF utilities for jsPDF with support for Arabic text, tables, and automatic page breaking.

## Installation

```bash
npm install jspdf-utils jspdf html2canvas
```

## Usage

```javascript
import jsPDF from "jspdf";
import { renderHTML } from "jspdf-utils";

const doc = new jsPDF({ unit: "mm", format: "a4" });

// Add fonts if needed for Arabic/RTL text
doc.addFont("path/to/arial.ttf", "arial", "normal", "normal");
doc.addFont("path/to/arial-bold.ttf", "arial", "normal", "bold");

// Render HTML element to PDF
const element = document.getElementById("content");
await renderHTML(doc, element);

doc.save("output.pdf");
```

## Features

- **Automatic page breaking**: Prevents tables and text from being split awkwardly
- **Table splitting**: Large tables are split across pages with repeated headers
- **Text wrapping**: Long text blocks are intelligently broken at word boundaries
- **RTL/Arabic support**: Works with right-to-left languages when proper fonts are loaded

## Development

### Running the Example

```bash
npm install
npm run dev
# Open http://localhost:5173
```

### Project Structure

```
├── src/
│   └── html-to-pdf.js    # Main utility functions
├── index.html             # Example/demo page
└── package.json
```

## API

### `renderHTML(doc, source, opts)`

Renders an HTML element to PDF.

- **doc**: jsPDF instance
- **source**: HTML element to render
- **opts**: Optional configuration (overrides defaults)

### `prepare(source, opts)`

Prepares an HTML element for rendering (used internally by renderHTML).

Returns: `{ clone, layout, options, cleanup }`

### Default Options

```javascript
{
  unit: 'mm',
  format: 'a4',
  pageWidth: 210,
  pageHeight: 297,
  margin: { top: 20, right: 20, bottom: 20, left: 20 }
}
```

## License

MIT
