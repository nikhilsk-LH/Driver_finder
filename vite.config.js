import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: "./" keeps asset paths relative so the built site works both locally
// and when served from a subpath (e.g. GitHub Pages at /Driver-finder/).
export default defineConfig({
  plugins: [react()],
  base: "./",
});
