import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  allowBooleanAttributes: true,
  parseTagValue: false,
  trimValues: false,
});

const SEVERITY_WEIGHT = { critical: 12, serious: 7, moderate: 3, minor: 1 };
const SEVERITY_ORDER = { critical: 0, serious: 1, moderate: 2, minor: 3 };

const asArray = (value) => (value == null ? [] : Array.isArray(value) ? value : [value]);
const clean = (value) => String(value == null ? "" : value).replace(/\s+/g, " ").trim();

function each(node, key, fn) {
  if (!node || typeof node !== "object") return;
  if (Array.isArray(node)) {
    node.forEach((item) => each(item, key, fn));
    return;
  }
  for (const [k, value] of Object.entries(node)) {
    if (k === key) {
      asArray(value).forEach(fn);
    }
    each(value, key, fn);
  }
}

function first(node, key) {
  let found = null;
  each(node, key, (value) => {
    if (found == null) found = value;
  });
  return found;
}

function collectText(node) {
  let out = "";
  if (node == null) return out;
  if (typeof node === "string") return node;
  if (typeof node === "number" || typeof node === "boolean") return String(node);
  if (Array.isArray(node)) {
    node.forEach((item) => { out += collectText(item); });
    return out;
  }
  if (typeof node !== "object") return out;
  for (const [key, value] of Object.entries(node)) {
    if (key === "t") {
      asArray(value).forEach((item) => {
        if (typeof item === "string") out += item;
        else if (item && typeof item === "object" && item["#text"]) out += item["#text"];
      });
    } else if (!key.startsWith("@_") && key !== "rPr" && key !== "pPr" && key !== "endParaRPr") {
      out += collectText(value);
    }
  }
  return out;
}

function issue(id, severity, title, wcag, why, fix, locations) {
  return { id, severity, title, wcag, why, fix, locations: locations.slice(0, 8), count: Math.max(locations.length, 1) };
}

function isBadAlt(alt) {
  if (!alt) return true;
  const value = clean(alt).toLowerCase();
  if (value.length < 3) return true;
  if (["image", "picture", "photo", "graphic", "logo", "icon", "chart", "figure", "screenshot"].includes(value)) return true;
  if (/\.(png|jpe?g|gif|webp|bmp|svg)$/.test(value)) return true;
  if (/^(img|image|photo|picture|screen ?shot|untitled)[-_ ]?\d*$/.test(value)) return true;
  return false;
}

function scoreIssues(issues) {
  let penalty = 0;
  for (const item of issues) {
    penalty += Math.min(15, SEVERITY_WEIGHT[item.severity] * Math.min(item.count, 2));
  }
  return Math.max(0, Math.min(100, Math.round(100 - penalty)));
}

function gradeFor(score) {
  return score >= 90 ? "Good" : score >= 70 ? "Needs work" : "High risk";
}

async function readXml(zip, path) {
  const file = zip.file(path);
  if (!file) return null;
  return parser.parse(await file.async("string"));
}

async function readRels(zip, path) {
  const data = await readXml(zip, path);
  const map = {};
  for (const rel of asArray(data && data.Relationships && data.Relationships.Relationship)) {
    map[rel["@_Id"]] = { target: rel["@_Target"], type: rel["@_Type"] || "", external: rel["@_TargetMode"] === "External" };
  }
  return map;
}

function validateZip(zip) {
  let total = 0;
  let count = 0;
  zip.forEach((path, file) => {
    if (file.dir) return;
    count += 1;
    total += file._data ? file._data.uncompressedSize || 0 : 0;
    if (path.endsWith(".xml") || path.endsWith(".rels")) {
      // XML entity declarations are checked after decompression, see scanDocx/scanPptx.
    }
  });
  if (count > 5000) throw new Error("This file contains too many embedded parts to scan safely.");
  if (total > 200 * 1024 * 1024) throw new Error("This file expands to more than 200 MB and cannot be scanned.");
}

async function checkUnsafeXml(zip) {
  for (const path of Object.keys(zip.files)) {
    if (!path.endsWith(".xml") && !path.endsWith(".rels")) continue;
    const data = await zip.file(path).async("string");
    if (data.includes("<!DOCTYPE") || data.includes("<!ENTITY")) {
      throw new Error("This file contains an unsafe XML declaration.");
    }
  }
}

function docxHas(zip, path) {
  return Boolean(zip.file(path));
}

async function scanDocx(zip, filename, bytes) {
  await checkUnsafeXml(zip);
  const document = await readXml(zip, "word/document.xml");
  if (!document || !document.document || !document.document.body) {
    throw new Error("Word could not read this DOCX. It may be encrypted or corrupted.");
  }
  const core = (await readXml(zip, "docProps/core.xml")) || {};
  const coreProps = core.coreProperties || {};
  const title = clean(collectText(coreProps.title));
  const rels = await readRels(zip, "word/_rels/document.xml.rels");
  const body = document.document.body;
  const issues = [];

  if (!title) {
    issues.push(issue("docx.title.missing", "moderate", "Document title is missing", "2.4.2 Page Titled",
      "A title helps screen-reader users and navigation tools identify the document.",
      "File > Info > Properties > Title. Use a short, descriptive title (not the file name).", ["Document properties"]));
  }

  const langXml = await Promise.all(["word/document.xml", "word/styles.xml"].map(async (p) => {
    const f = zip.file(p);
    return f ? f.async("string") : "";
  }));
  if (!langXml.some((data) => /<w:lang\b[^>]*w:val="[A-Za-z]{2,3}/.test(data))) {
    issues.push(issue("docx.language.missing", "serious", "Document language is not set", "3.1.1 Language of Page",
      "Screen readers may pronounce the whole document with the wrong language and voice.",
      "Review > Language > Set Proofing Language, and set the document default language.", ["Document defaults"]));
  }

  const headings = [];
  each(body, "p", (p) => {
    const pPr = p.pPr || {};
    const style = clean((pPr.pStyle && pPr.pStyle["@_val"]) || "").toLowerCase();
    let level = 0;
    const match = style.match(/^heading\s*(\d)$/);
    if (match) level = Number(match[1]);
    else if (pPr.outlineLvl && pPr.outlineLvl["@_val"] != null) level = Number(pPr.outlineLvl["@_val"]) + 1;
    const text = clean(collectText(p));
    if (level && text) headings.push({ level, text });
  });

  if (!headings.length && clean(collectText(body))) {
    issues.push(issue("docx.headings.missing", "serious", "No heading structure detected", "1.3.1 Info and Relationships; 2.4.6 Headings and Labels",
      "Screen-reader users navigate long documents by headings. Bold text is not a heading.",
      "Apply Heading 1/2/3 styles to section titles. Do not use bold paragraphs as fake headings.", ["Whole document"]));
  } else if (headings.length) {
    if (headings[0].level > 1) {
      issues.push(issue("docx.headings.start", "moderate", "First heading is not Heading 1", "1.3.1 Info and Relationships",
        "A skipped first level makes the outline start in the middle of the hierarchy.",
        "Make the document title the first Heading 1, then use Heading 2 for sections.", [`First heading: ${headings[0].text.slice(0, 60)}`]));
    }
    const skipped = [];
    let previous = 0;
    for (const h of headings) {
      if (previous && h.level > previous + 1) skipped.push(`H${previous} -> H${h.level}: ${h.text.slice(0, 50)}`);
      previous = h.level;
    }
    if (skipped.length) {
      issues.push(issue("docx.headings.skip", "moderate", "Heading levels are skipped", "1.3.1 Info and Relationships",
        "A broken outline confuses screen-reader navigation and document maps.",
        "Use heading levels in order (H1 -> H2 -> H3). Restyle instead of skipping levels.", skipped));
    }
    const h1 = headings.filter((h) => h.level === 1).length;
    if (h1 > 1) {
      issues.push(issue("docx.headings.multiple_h1", "minor", "More than one Heading 1", "1.3.1 Info and Relationships",
        "Multiple top-level headings make the document outline ambiguous.", "Keep one document-level H1; use H2/H3 for sections.", [`${h1} Heading 1 paragraphs`]));
    }
  }

  const missingAlt = [];
  const badAlt = [];
  const altSeen = {};
  each(body, "drawing", (drawing) => {
    if (!first(drawing, "blip")) return;
    const docPr = first(drawing, "docPr");
    if (!docPr) return;
    const name = clean(docPr["@_name"]) || "image";
    const alt = clean(docPr["@_descr"] || docPr["@_title"]);
    if (!alt) missingAlt.push(name);
    else {
      if (isBadAlt(alt)) badAlt.push(`${name}: "${alt.slice(0, 60)}"`);
      altSeen[alt.toLowerCase()] = (altSeen[alt.toLowerCase()] || 0) + 1;
    }
  });
  if (missingAlt.length) {
    issues.push(issue("docx.image.alt_missing", "critical", "Images are missing alternative text", "1.1.1 Non-text Content",
      "Screen readers announce nothing useful for these images, so the information is lost.",
      "Right-click image > Edit Alt Text. Describe the purpose in 1-2 sentences; mark decorative images as decorative.", missingAlt));
  }
  if (badAlt.length) {
    issues.push(issue("docx.image.alt_poor", "serious", "Alternative text is missing or not descriptive", "1.1.1 Non-text Content",
      "File names, \"image\", or generic labels do not convey the image's purpose.",
      "Replace with a short description of what the image communicates in this context.", badAlt));
  }
  const dupAlt = Object.entries(altSeen).filter(([, n]) => n > 2).map(([text, n]) => `${n} images share alt text: "${text.slice(0, 50)}"`);
  if (dupAlt.length) {
    issues.push(issue("docx.image.alt_duplicate", "moderate", "The same alternative text is reused on several images", "1.1.1 Non-text Content",
      "Repeated alt text usually means the description is not specific to each image.",
      "Review each image and describe its distinct purpose; mark purely decorative images as decorative.", dupAlt));
  }

  const tableHeaders = [];
  each(body, "tbl", (tbl) => {
    const rows = asArray(tbl.tr);
    const cols = asArray(tbl.tblGrid && tbl.tblGrid.gridCol);
    if (rows.length < 2 || cols.length < 2) return;
    const trPr = rows[0].trPr;
    if (!trPr || trPr.tblHeader == null) tableHeaders.push(`Table (${rows.length}x${cols.length})`);
  });
  if (tableHeaders.length) {
    issues.push(issue("docx.table.header_missing", "serious", "Data tables do not repeat a header row", "1.3.1 Info and Relationships",
      "Screen readers need header cells to announce which column a value belongs to.",
      "Select the header row > Table Layout > Repeat Header Rows; use a real table, not tabs.", tableHeaders));
  }

  const ambiguous = new Set(["click here", "here", "more", "link", "this", "read more", "learn more"]);
  const badLinks = [];
  each(body, "hyperlink", (link) => {
    const text = clean(collectText(link));
    if (!text || ambiguous.has(text.toLowerCase())) {
      const rel = rels[link["@_id"]] || {};
      badLinks.push(`"${text || "(empty)"}" -> ${(rel.target || "").slice(0, 60)}`);
    }
  });
  if (badLinks.length) {
    issues.push(issue("docx.link.ambiguous", "serious", "Links have empty or ambiguous text", "2.4.4 Link Purpose (In Context)",
      "Screen-reader users often jump link-to-link; \"click here\" gives no destination context.",
      "Use descriptive link text that names the destination or action.", badLinks));
  }

  const textBoxes = [];
  each(body, "txbxContent", (tb) => textBoxes.push(tb));
  if (textBoxes.length) {
    issues.push(issue("docx.textbox.reading_order", "moderate", "Text boxes may create reading-order problems", "1.3.2 Meaningful Sequence",
      "Floating text boxes can be read in the wrong order by assistive technology.",
      "Confirm the reading order with a screen reader, or move the content into the main document flow.", [`${textBoxes.length} text box(es)`]));
  }

  const tiny = [];
  each(body, "r", (r) => {
    const size = r.rPr && r.rPr.sz && r.rPr.sz["@_val"];
    if (size != null && Number(size) < 18 && clean(collectText(r))) tiny.push(`${Number(size) / 2} pt`);
  });
  if (tiny.length) {
    issues.push(issue("docx.text.tiny", "minor", "Some text is set below 9 pt", "1.4.4 Resize Text (review)",
      "Very small text is hard to read and may not resize cleanly.",
      "Use at least 11-12 pt for body text; 9 pt only for footnotes.", tiny));
  }

  if (first(body, "ins") || first(body, "del")) {
    issues.push(issue("docx.tracked_changes", "moderate", "Tracked changes are present", "1.3.1 Info and Relationships",
      "Insertions/deletions can be read inconsistently and may expose draft content.",
      "Accept or reject all changes before publishing, then re-run the preflight.", ["Document contains tracked changes"]));
  }
  if (docxHas(zip, "word/comments.xml")) {
    issues.push(issue("docx.comments", "minor", "Comments are present", "1.3.1 Info and Relationships",
      "Review comments may contain internal notes or draft content not meant for learners.",
      "Resolve and delete comments before publishing.", ["Document contains comments"]));
  }

  return finalize(issues, filename, "docx", bytes, { paragraphs: headings.length, tables: tableHeaders.length, slides: 0 });
}

function shapeAlt(shape) {
  const holders = ["nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"];
  for (const key of holders) {
    const holder = shape[key];
    if (holder && holder.cNvPr) return { alt: clean(holder.cNvPr["@_descr"] || holder.cNvPr["@_title"]), name: clean(holder.cNvPr["@_name"]) };
  }
  return { alt: "", name: "" };
}

function shapeText(shape) {
  return clean(collectText(shape.txBody || {}));
}

function shapeOffset(shape) {
  const xfrm = (shape.spPr && shape.spPr.xfrm) || shape.xfrm;
  if (!xfrm || !xfrm.off) return null;
  return { x: Number(xfrm.off["@_x"] || 0), y: Number(xfrm.off["@_y"] || 0) };
}

function isTitlePlaceholder(shape) {
  const ph = shape.nvSpPr && shape.nvSpPr.nvPr && shape.nvSpPr.nvPr.ph;
  return Boolean(ph && ["title", "ctrTitle"].includes(ph["@_type"]));
}

function hasTable(shape) {
  return Boolean(shape.graphic && shape.graphic.graphicData && shape.graphic.graphicData.tbl);
}

function hasChart(shape) {
  const uri = shape.graphic && shape.graphic.graphicData && shape.graphic.graphicData["@_uri"];
  return Boolean(uri && String(uri).includes("/chart"));
}

function hasMedia(shape) {
  return Boolean(first(shape, "videoFile") || first(shape, "audioFile"));
}

async function scanPptx(zip, filename, bytes) {
  await checkUnsafeXml(zip);
  const presentation = await readXml(zip, "ppt/presentation.xml");
  if (!presentation) throw new Error("PowerPoint could not read this PPTX. It may be encrypted or corrupted.");

  const slideFiles = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a, b) => Number(a.match(/slide(\d+)/)[1]) - Number(b.match(/slide(\d+)/)[1]));
  if (!slideFiles.length) throw new Error("No slides were found in this PPTX.");

  const issues = [];
  const titles = {};
  const missingTitles = [];
  const missingAlt = [];
  const badAlt = [];
  const tableHeaders = [];
  const links = [];
  const tiny = [];
  const hidden = [];
  const media = [];
  const readingOrder = [];
  let shapeCount = 0;

  for (let i = 0; i < slideFiles.length; i += 1) {
    const slideNo = i + 1;
    const slide = await readXml(zip, slideFiles[i]);
    if (!slide || !slide.sld) continue;
    const rels = await readRels(zip, slideFiles[i].replace("slides/", "slides/_rels/") + ".rels");
    if (slide.sld["@_show"] === "0") hidden.push(`Slide ${slideNo}`);
    const tree = slide.sld.cSld && slide.sld.cSld.spTree;
    if (!tree) continue;

    let title = "";
    const shapes = [];
    for (const key of ["sp", "pic", "graphicFrame", "grpSp", "cxnSp"]) {
      asArray(tree[key]).forEach((node) => shapes.push({ type: key, node }));
    }
    shapeCount += shapes.length;

    for (const entry of shapes) {
      const shape = entry.node;
      if (entry.type === "sp" && isTitlePlaceholder(shape)) title = shapeText(shape);
      const altInfo = shapeAlt(shape);
      const name = altInfo.name || `shape on slide ${slideNo}`;
      if (entry.type === "pic" || hasChart(shape)) {
        if (!altInfo.alt) missingAlt.push(`${name} (slide ${slideNo})`);
        else if (isBadAlt(altInfo.alt)) badAlt.push(`${name}: "${altInfo.alt.slice(0, 50)}" (slide ${slideNo})`);
      }
      if (entry.type === "grpSp") {
        if (!altInfo.alt) badAlt.push(`Grouped object without alt text (slide ${slideNo})`);
      }
      if (hasTable(shape)) {
        const table = shape.graphic.graphicData.tbl;
        const rows = asArray(table.tr);
        const cols = asArray(table.tblGrid && table.tblGrid.gridCol);
        const firstRow = table.tblPr && String(table.tblPr["@_firstRow"]) === "1";
        if (rows.length >= 2 && cols.length >= 2 && !firstRow) tableHeaders.push(`Table on slide ${slideNo}`);
      }
      each(shape.txBody || {}, "hlinkClick", (link) => {
        const text = clean(collectText(shape.txBody || {}));
        if (!text || ["click here", "here", "more", "link", "read more"].includes(text.toLowerCase())) {
          const rel = rels[link["@_id"]] || {};
          links.push(`Slide ${slideNo}: "${text || "(empty)"}" -> ${(rel.target || "").slice(0, 50)}`);
        }
      });
      each(shape.txBody || {}, "r", (run) => {
        const size = run.rPr && run.rPr["@_sz"];
        if (size != null && Number(size) < 900 && clean(collectText(run))) tiny.push(`Slide ${slideNo}: ${Number(size) / 100} pt`);
      });
      if (hasMedia(shape)) media.push(`Slide ${slideNo}: ${name}`);
    }

    if (!title) missingTitles.push(`Slide ${slideNo}`);
    else titles[title.toLowerCase()] = (titles[title.toLowerCase()] || []).concat([slideNo]);

    const positioned = [];
    shapes.forEach((entry, index) => {
      const shape = entry.node;
      const offset = shapeOffset(shape);
      if (offset && (shape.txBody || entry.type === "pic" || hasChart(shape))) positioned.push({ id: index, ...offset });
    });
    if (positioned.length >= 3) {
      const xmlOrder = positioned.map((p) => p.id);
      const visualOrder = positioned.slice().sort((a, b) => a.y - b.y || a.x - b.x).map((p) => p.id);
      if (xmlOrder.join(",") !== visualOrder.join(",")) readingOrder.push(`Slide ${slideNo}: shape order differs from top-left order`);
    }
  }

  if (missingTitles.length) {
    issues.push(issue("pptx.title.missing", "critical", "Slides have no title", "2.4.2 Page Titled; 1.3.1 Info and Relationships",
      "Titles are the primary way screen-reader users navigate slides and understand context.",
      "Use the Title placeholder (not a plain text box). Give every slide a unique, descriptive title.", missingTitles));
  }
  const duplicates = Object.entries(titles).filter(([, nums]) => nums.length > 1).map(([text, nums]) => `Slides ${nums.join(", ")}: "${text.slice(0, 50)}"`);
  if (duplicates.length) {
    issues.push(issue("pptx.title.duplicate", "moderate", "Several slides share the same title", "2.4.2 Page Titled",
      "Duplicate titles make it hard to know which slide you are on.", "Give each slide a distinct title that describes its specific content.", duplicates));
  }
  if (missingAlt.length) {
    issues.push(issue("pptx.image.alt_missing", "critical", "Pictures or charts are missing alternative text", "1.1.1 Non-text Content",
      "Charts and images often carry the main message; without alt text, that message is lost.",
      "Select the object > Alt Text. Describe the takeaway, not the pixels. Mark decorative objects as decorative.", missingAlt));
  }
  if (badAlt.length) {
    issues.push(issue("pptx.image.alt_poor", "serious", "Alternative text is missing, generic, or missing on a group", "1.1.1 Non-text Content",
      "Generic labels and unlabelled groups do not communicate the object's purpose.",
      "Write a short, specific description; ungroup and label meaningful child objects where needed.", badAlt));
  }
  if (tableHeaders.length) {
    issues.push(issue("pptx.table.header_missing", "serious", "Tables do not have a header row enabled", "1.3.1 Info and Relationships",
      "Screen readers use header cells to associate data with the correct column.",
      "Select the table > Table Design > Header Row. Keep the header text short and specific.", tableHeaders));
  }
  if (links.length) {
    issues.push(issue("pptx.link.ambiguous", "serious", "Links have empty or ambiguous text", "2.4.4 Link Purpose (In Context)",
      "Screen-reader users navigate by link list; \"click here\" is not meaningful.", "Use link text that names the destination or action.", links));
  }
  if (hidden.length) {
    issues.push(issue("pptx.hidden_slides", "minor", "Hidden slides are present", "1.3.1 Info and Relationships (review)",
      "Hidden slides may still be exported, printed, or read by some tools.",
      "Delete slides that are not part of the published deck, or confirm the export excludes them.", hidden));
  }
  if (readingOrder.length) {
    issues.push(issue("pptx.reading_order", "moderate", "Reading order may not match visual order", "1.3.2 Meaningful Sequence",
      "Screen readers follow the object order in the file, which may differ from what sighted users see.",
      "Use the Selection Pane to reorder objects top-to-bottom, left-to-right; verify with a screen reader.", readingOrder));
  }
  if (tiny.length) {
    issues.push(issue("pptx.text.tiny", "minor", "Some text is set below 9 pt", "1.4.4 Resize Text (review)",
      "Very small text is hard to read and may not resize cleanly.", "Use at least 12 pt for slide body text; 9 pt only for source notes.", tiny));
  }
  if (media.length) {
    issues.push(issue("pptx.media.captions", "serious", "Embedded audio/video needs captions or a transcript", "1.2.2 Captions (Prerecorded)",
      "Deaf and hard-of-hearing learners cannot access audio-only content.",
      "Add captions and a transcript; confirm the player supports captions.", media));
  }

  return finalize(issues, filename, "pptx", bytes, { paragraphs: 0, tables: tableHeaders.length, slides: slideFiles.length });
}

function finalize(rawIssues, filename, type, bytes, stats) {
  const issues = rawIssues
    .map((item) => ({ ...item, count: Math.max(item.locations.length, 1) }))
    .sort((a, b) => (SEVERITY_ORDER[a.severity] ?? 9) - (SEVERITY_ORDER[b.severity] ?? 9) || a.title.localeCompare(b.title));
  const counts = { critical: 0, serious: 0, moderate: 0, minor: 0 };
  issues.forEach((item) => { counts[item.severity] = (counts[item.severity] || 0) + 1; });
  const score = scoreIssues(issues);
  const summary = buildSummary(type, score, counts, issues);
  return {
    ok: true,
    file: { name: filename, type, bytes },
    score,
    grade: gradeFor(score),
    counts,
    issues,
    stats,
    ai: {
      enabled: false,
      model: "browser-rules",
      summary,
      priorities: issues.slice(0, 3).map((item) => `${item.title} — ${item.fix}`),
      limitations: "This browser-only build performs rule-based checks. It cannot verify visual reading order, color contrast in every context, screen-reader pronunciation, or whether alt text is meaningful. It is a preflight, not a certification or legal advice.",
    },
    meta: { engine: "browser-rules-v0.1", generated_at: new Date().toISOString(), wcag: "WCAG 2.1/2.2 A/AA references" },
  };
}

function buildSummary(type, score, counts, issues) {
  if (!issues.length) {
    return `This ${type.toUpperCase()} scored ${score}/100 with no automatically detectable issues. Automated checks cannot confirm full WCAG conformance, so a manual review is still recommended.`;
  }
  const top = issues.slice(0, 3).map((item) => item.title.toLowerCase()).join("; ");
  return `This ${type.toUpperCase()} scored ${score}/100, with ${counts.critical || 0} critical and ${counts.serious || 0} serious findings. The highest-impact items are: ${top}.`;
}

export async function scanFile(file) {
  const name = file.name || "document";
  const ext = (name.split(".").pop() || "").toLowerCase();
  if (!["docx", "pptx"].includes(ext)) throw new Error("Only .docx and .pptx files are supported in this version.");
  if (!file.size) throw new Error("That file is empty.");
  if (file.size > 15 * 1024 * 1024) throw new Error("That file is larger than 15 MB.");
  const buffer = await file.arrayBuffer();
  let zip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (err) {
    throw new Error("This is not a valid DOCX/PPTX file (the ZIP container is damaged).");
  }
  validateZip(zip);
  if (ext === "docx") return scanDocx(zip, name, file.size);
  return scanPptx(zip, name, file.size);
}
