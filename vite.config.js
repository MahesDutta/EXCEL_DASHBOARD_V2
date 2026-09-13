import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  base: "/EXCEL_DASHBOARD_V2/",
  build: {
    outDir: "dist",
    sourcemap: false,
    minify: "terser"
  },
  server: {
    open: true,
    strictPort: false
  }
});