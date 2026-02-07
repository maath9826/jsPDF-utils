import { defineConfig } from "vite";

export default defineConfig({
  build: {
    lib: {
      entry: "src/html-to-pdf.ts",
      formats: ["es"],
      fileName: "html-to-pdf",
    },
    rollupOptions: {
      external: ["jspdf", "html2canvas"],
    },
  },
});
