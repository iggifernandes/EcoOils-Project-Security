import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// IMPORTANT: set `base` to your repo name for GitHub Pages project sites,
// e.g. if your repo is github.com/yourname/security-timesheet, keep this as
// "/security-timesheet/". If you deploy to a custom domain or a "yourname.github.io"
// root repo instead, change this to "/".
export default defineConfig({
  plugins: [react()],
  base: process.env.VITE_BASE || "/",
});
