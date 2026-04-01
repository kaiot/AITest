import { defineConfig } from "vite";
import basicSsl from "@vitejs/plugin-basic-ssl";

export default defineConfig({
  plugins: [basicSsl()],
  server: {
    https: true,
    port: 5173,
    proxy: {
      // Proxy WebSocket to backend
      "/ws": {
        target: "wss://localhost:8340",
        ws: true,
        secure: false,
      },
      // Proxy REST API to backend
      "/api": {
        target: "https://localhost:8340",
        secure: false,
        changeOrigin: true,
      },
    },
  },
});
