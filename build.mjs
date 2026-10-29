import { build } from "esbuild";
await build({
  entryPoints: ["src/main.js"],
  bundle: true,
  minify: true,
  format: "iife",
  target: ["es2020"],
  outfile: "dist/app.js",
  legalComments: "none",
});
console.log("built dist/app.js");
