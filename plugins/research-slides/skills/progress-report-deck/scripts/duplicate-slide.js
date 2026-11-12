#!/usr/bin/env node
/**
 * Duplicate a slide inside an unpacked .pptx directory, including all the
 * package bookkeeping PowerPoint needs (content types, relationships,
 * sldIdLst). This machine has no python-pptx / LibreOffice, so this
 * replaces the pptx skill's python-based add_slide.py.
 *
 * Usage:
 *   node duplicate-slide.js <unpacked-dir> <sourceSlideFile> [--after <slideFile>]
 *
 * Example:
 *   node duplicate-slide.js unpacked slide9.xml --after slide9.xml
 *
 * Prints the path of the newly created slide XML file on success.
 *
 * Do all structural work (duplicate, delete, reorder) BEFORE editing any
 * slide's new content — duplicating after editing clones the edited text
 * onto both slides, since it's a plain file copy.
 */
const fs = require("fs");
const path = require("path");

function fail(msg) {
  console.error("Error: " + msg);
  process.exit(1);
}

const args = process.argv.slice(2);
if (args.length < 2) {
  fail("usage: node duplicate-slide.js <unpacked-dir> <sourceSlideFile> [--after <slideFile>]");
}
const root = args[0];
const sourceFile = args[1];
let afterFile = null;
const afterIdx = args.indexOf("--after");
if (afterIdx !== -1) afterFile = args[afterIdx + 1];

const slidesDir = path.join(root, "ppt", "slides");
const relsDir = path.join(slidesDir, "_rels");
const sourcePath = path.join(slidesDir, sourceFile);
if (!fs.existsSync(sourcePath)) fail(`source slide not found: ${sourcePath}`);

// --- find next free slideN.xml number ---
const existing = fs.readdirSync(slidesDir)
  .filter((f) => /^slide\d+\.xml$/.test(f))
  .map((f) => parseInt(f.match(/^slide(\d+)\.xml$/)[1], 10));
const nextNum = Math.max(...existing) + 1;
const newFile = `slide${nextNum}.xml`;
const newPath = path.join(slidesDir, newFile);

// --- 1. copy slide XML ---
fs.copyFileSync(sourcePath, newPath);

// --- 2. copy the .rels file (layout link, and any image/chart rels) ---
const sourceRels = path.join(relsDir, sourceFile + ".rels");
const newRels = path.join(relsDir, newFile + ".rels");
if (fs.existsSync(sourceRels)) {
  fs.copyFileSync(sourceRels, newRels);
}

// --- 3. register the new part in [Content_Types].xml ---
const ctPath = path.join(root, "[Content_Types].xml");
let ct = fs.readFileSync(ctPath, "utf8");
const ctEntry = `<Override PartName="/ppt/slides/${newFile}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
if (!ct.includes(ctEntry)) {
  ct = ct.replace("</Types>", ctEntry + "</Types>");
  fs.writeFileSync(ctPath, ct, "utf8");
}

// --- 4. add relationship in ppt/_rels/presentation.xml.rels ---
const presRelsPath = path.join(root, "ppt", "_rels", "presentation.xml.rels");
let presRels = fs.readFileSync(presRelsPath, "utf8");
const usedRIds = [...presRels.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
const nextRId = "rId" + (Math.max(...usedRIds) + 1);
const relEntry = `<Relationship Id="${nextRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${newFile}"/>`;
presRels = presRels.replace("</Relationships>", relEntry + "</Relationships>");
fs.writeFileSync(presRelsPath, presRels, "utf8");

// --- 5. insert <p:sldId> into ppt/presentation.xml's <p:sldIdLst> ---
const presPath = path.join(root, "ppt", "presentation.xml");
let pres = fs.readFileSync(presPath, "utf8");
const usedSldIds = [...pres.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => parseInt(m[1], 10));
const nextSldId = (usedSldIds.length ? Math.max(...usedSldIds) : 255) + 1;
const sldIdEntry = `<p:sldId id="${nextSldId}" r:id="${nextRId}"/>`;

const listMatch = pres.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
if (!listMatch) fail("could not find <p:sldIdLst> in presentation.xml");
let inner = listMatch[1];

if (afterFile) {
  // find the rId presentation.xml.rels maps to afterFile, then find that sldId entry
  const afterRelMatch = [...presRels.matchAll(/<Relationship Id="(rId\d+)"[^>]*Target="slides\/([^"]+)"/g)]
    .find((m) => m[2] === afterFile);
  if (!afterRelMatch) fail(`could not find a relationship for --after target ${afterFile}`);
  const afterRId = afterRelMatch[1];
  const afterSldIdRe = new RegExp(`<p:sldId id="\\d+" r:id="${afterRId}"/>`);
  if (!afterSldIdRe.test(inner)) fail(`--after slide ${afterFile} is not in <p:sldIdLst> (rId ${afterRId})`);
  inner = inner.replace(afterSldIdRe, (m) => m + sldIdEntry);
} else {
  inner = inner + sldIdEntry;
}

pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${inner}</p:sldIdLst>`);
fs.writeFileSync(presPath, pres, "utf8");

console.log(`Created ${path.join("ppt", "slides", newFile)} from ${path.join("ppt", "slides", sourceFile)}`);
console.log(`  content-type entry: ${ctEntry}`);
console.log(`  relationship: ${nextRId} -> slides/${newFile}`);
console.log(`  sldId ${nextSldId} inserted ${afterFile ? "after " + afterFile : "at end"}`);
