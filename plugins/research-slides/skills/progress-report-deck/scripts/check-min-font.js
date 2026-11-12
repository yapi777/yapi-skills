#!/usr/bin/env node
/**
 * Scan every ppt/slides/slideN.xml in an unpacked .pptx directory for
 * explicit font sizes below a floor (default 16pt) and report them.
 *
 * This only catches EXPLICIT sz="" attributes on runs — text that inherits
 * its size from the slide layout/master (no sz attribute on the run) is not
 * checked here. In this template, body/title/footer text almost always
 * carries an explicit sz, so this catches the real risk: someone shrinking
 * a citation, caption, or overflowing box below 16pt to make it fit.
 * If a slide relies on <a:normAutofit fontScale="..."/> to shrink text,
 * check that slide by eye too — a large fontScale reduction can push
 * inherited text under the floor even though no run-level sz looks small.
 *
 * Usage:
 *   node check-min-font.js <unpacked-dir> [minPoints]
 *
 * Exit code 0 = all clear, 1 = violations found (or bad usage).
 */
const fs = require("fs");
const path = require("path");

const root = process.argv[2];
const minPt = parseFloat(process.argv[3] || "16");
if (!root) {
  console.error("usage: node check-min-font.js <unpacked-dir> [minPoints]");
  process.exit(1);
}
const minSz = Math.round(minPt * 100);

const slidesDir = path.join(root, "ppt", "slides");
const files = fs.readdirSync(slidesDir)
  .filter((f) => /^slide\d+\.xml$/.test(f))
  .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));

let violations = 0;
for (const file of files) {
  const xml = fs.readFileSync(path.join(slidesDir, file), "utf8");
  // sz="NNNN" appears on <a:rPr>/<a:endParaRPr>/<a:defRPr> — all font-size carriers
  const matches = [...xml.matchAll(/sz="(\d+)"/g)];
  for (const m of matches) {
    const sz = parseInt(m[1], 10);
    if (sz < minSz) {
      // grab a bit of surrounding text for context (best-effort)
      const start = m.index;
      const tail = xml.slice(start, start + 400);
      const textMatch = tail.match(/<a:t>([^<]*)<\/a:t>/);
      const snippet = textMatch ? textMatch[1].slice(0, 40) : "(no text found nearby)";
      console.log(`${file}: sz=${sz} (${(sz / 100).toFixed(1)}pt) near "${snippet}"`);
      violations++;
    }
  }
}

if (violations === 0) {
  console.log(`OK: no explicit font size below ${minPt}pt across ${files.length} slides.`);
  process.exit(0);
} else {
  console.log(`\n${violations} run(s) below ${minPt}pt. Bump these to at least ${minPt}pt (reflow/split the slide if it no longer fits).`);
  process.exit(1);
}
