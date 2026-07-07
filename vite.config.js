import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// If you deploy to https://<username>.github.io/<repo-name>/ (a normal project
// repo), set REPO_NAME below to match your repository name exactly.
// If you deploy to a user/org page (https://<username>.github.io/) or a custom
// domain, set REPO_NAME to "" instead.
const REPO_NAME = "delivery_dashboard";

export default defineConfig({
  plugins: [react()],
  base: process.env.NODE_ENV === "production" && REPO_NAME ? `/${REPO_NAME}/` : "/",
});
