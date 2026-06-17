import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";
import runtimeErrorOverlay from "@replit/vite-plugin-runtime-error-modal";

const port = Number(process.env.PORT);
const basePath = process.env.BASE_PATH ?? "/";

export default defineConfig({
  base: basePath,
  plugins: [react(), tailwindcss(), runtimeErrorOverlay()],
  resolve: {
    alias: { "@": path.resolve(import.meta.dirname, "src") },
    dedupe: ["react", "react-dom"],
  },
  server: { port, strictPort: true, host: "0.0.0.0", allowedHosts: true },
  preview: { port, host: "0.0.0.0", allowedHosts: true },
  build: { outDir: path.resolve(import.meta.dirname, "dist/public"), emptyOutDir: true },
});
