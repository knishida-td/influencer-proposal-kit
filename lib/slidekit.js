const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

// =============================================================
// Image Utilities
// =============================================================

/** アスペクト比を維持してmaxW/maxHに収める */
function fitImage(origW, origH, maxW, maxH) {
  if (!origH || !origW) return { w: maxW, h: maxH };
  const ratio = origW / origH;
  let w = maxW, h = maxW / ratio;
  if (h > maxH) { h = maxH; w = maxH * ratio; }
  return { w, h };
}

/** ファイルからBase64読み込み */
function loadLocalImage(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return null;
  const ext = path.extname(filePath).toLowerCase();
  const mime = ext === ".png" ? "image/png" : "image/jpeg";
  const buf = fs.readFileSync(filePath);
  return { data: `${mime};base64,${buf.toString("base64")}`, path: filePath };
}

function resolveAssetPath(filePath, baseDir) {
  if (!filePath) return filePath;
  if (path.isAbsolute(filePath)) return filePath;
  return baseDir ? path.join(baseDir, filePath) : filePath;
}

/** sharpで画像の実ピクセルを取得（クロスプラットフォーム） */
async function getImageDimensions(filePath) {
  try {
    const meta = await sharp(filePath).metadata();
    return { w: meta.width || 800, h: meta.height || 600 };
  } catch {
    return { w: 800, h: 600 };
  }
}

// =============================================================
// Design System (Dark Gold)
// =============================================================
const BASE_COLORS = {
  bg: "111111", title: "FFFFFF", body: "E0E0E0", sub: "AAAAAA", muted: "666666",
  kmBg: "1F1A14", sep: "333333", divider: "333333", white: "FFFFFF", rowAlt: "1A1A1A",
};

const PALETTES = {
  gold:     { primary: "C9A96E", secondary: "E8D5B0" },
  feminine: { primary: "D4829C", secondary: "F0C4D4" },
  natural:  { primary: "7BA05B", secondary: "B8D4A0" },
  cool:     { primary: "5B8FB9", secondary: "A0C4E0" },
};

function buildColors(tone) {
  const palette = PALETTES[tone] || PALETTES.gold;
  return { ...BASE_COLORS, ...palette };
}
const FONT = "Hiragino Kaku Gothic Pro W3";
const SW = 10, SH = 5.625;

// レイアウト定数
const LAYOUT = {
  hdr: { x: 0.5, y: 0.39, w: 9.0, h: 0.45 },
  hdrLineY: 0.857,
  hdrLine1: { x: 0.5, w: 4.25 },
  hdrLine2: { x: 4.75, w: 4.75 },
  lineH: 0.035,
  km: { x: 0.5, y: 4.837, w: 9.0, h: 0.4 },
  pn: { x: 9.2, y: 5.337, w: 0.5, h: 0.25 },
  bodyTop: 0.893,
  bodyBot: 4.837,
  sepH: 0.015,
  margin: 0.5,
  defBar: { w: 0.06, h: 0.7 },
  bullet: { w: 0.14, h: 0.14 },
  bulletSpacing: 0.4,
  featureSpacing: 0.43,
  strategySpacing: 0.8,
  variationSpacing: 0.65,
  maxPoints: 4,
  maxFeatures: 6,
  maxCompRows: 6,
  maxSchedRows: 8,
  maxProjRows: 9,
  maxVariations: 4,
  maxStrategies: 3,
};
const BODY_H = LAYOUT.bodyBot - LAYOUT.bodyTop;

function centerY(contentH) {
  return LAYOUT.bodyTop + (BODY_H - contentH) / 2;
}

// =============================================================
// SVG Badge Generator
// =============================================================
async function makeBadge(svg) {
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// =============================================================
// JSON Validation
// =============================================================
const REQUIRED_KEYS = ["title", "images", "imageDir", "cover", "concept", "overview",
  "materials", "quality", "comparison", "production", "design", "variations",
  "sales", "schedule", "projection"];

function validateData(data) {
  const missing = REQUIRED_KEYS.filter(k => !data[k]);
  if (missing.length) throw new Error(`JSON missing required keys: ${missing.join(", ")}`);
}

// =============================================================
// Shared Table Renderer (C passed as parameter)
// =============================================================
function renderTable(s, { pres, C, colX, colW, headers, rows, hdrH, rowH, top, cellColorFn }) {
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
  headers.forEach((h, i) => {
    s.addText(h, { x: colX[i], y: top, w: colW[i], h: hdrH, fontSize: 12, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
  });
  rows.forEach((row, ri) => {
    const ry = top + hdrH + ri * rowH;
    if (ri % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: C.rowAlt } });
    row.forEach((cell, ci) => {
      const color = cellColorFn ? cellColorFn(ri, ci) : C.body;
      const bold = cellColorFn ? ci <= 1 : ci === 0;
      s.addText(cell, { x: colX[ci], y: ry, w: colW[ci], h: rowH, fontSize: 11, fontFace: FONT, color, bold, align: "center", valign: "middle", margin: 0 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry + rowH - 0.01, w: 9, h: LAYOUT.sepH, fill: { color: C.sep } });
  });
  return top + hdrH + rows.length * rowH;
}

// =============================================================
// Slide Context: helpers shared across single & bundle modes
// =============================================================
function createSlideContext(pres, C, logoData) {
  function addHeader(s, titleText) {
    s.background = { color: C.bg };
    s.addText(titleText, { x: LAYOUT.hdr.x, y: LAYOUT.hdr.y, w: LAYOUT.hdr.w, h: LAYOUT.hdr.h, fontSize: 22, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: LAYOUT.hdrLine1.x, y: LAYOUT.hdrLineY, w: LAYOUT.hdrLine1.w, h: LAYOUT.lineH, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: LAYOUT.hdrLine2.x, y: LAYOUT.hdrLineY, w: LAYOUT.hdrLine2.w, h: LAYOUT.lineH, fill: { color: C.secondary } });
  }
  function addKeyMsg(s, text) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: LAYOUT.km.x, y: LAYOUT.km.y, w: LAYOUT.km.w, h: LAYOUT.km.h, fill: { color: C.kmBg }, rectRadius: 0.05 });
    s.addText(text, { x: LAYOUT.km.x, y: LAYOUT.km.y, w: LAYOUT.km.w, h: LAYOUT.km.h, fontSize: 18, fontFace: FONT, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  }
  function addPageNum(s, num) {
    s.addText(String(num), { x: LAYOUT.pn.x, y: LAYOUT.pn.y, w: LAYOUT.pn.w, h: LAYOUT.pn.h, fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0 });
  }
  function addLogo(s) {
    if (!logoData) return;
    s.addImage({ data: logoData, x: 9.1, y: 0.15, w: 0.45, h: 0.58 });
  }
  function addSep(s, x, y, w) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: LAYOUT.sepH, fill: { color: C.sep } });
  }
  function addDefBlock(s, x, y, heading, body, w) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: LAYOUT.defBar.w, h: LAYOUT.defBar.h, fill: { color: C.primary } });
    s.addText(heading, { x: x + 0.2, y, w: w - 0.2, h: 0.28, fontSize: 16, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    s.addText(body, { x: x + 0.2, y: y + 0.3, w: w - 0.2, h: 0.4, fontSize: 14, fontFace: FONT, color: C.body, align: "left", valign: "top", margin: 0 });
  }
  function addImgFit(s, imgData, imgDims, x, y, maxW, maxH) {
    const fit = fitImage(imgDims.w, imgDims.h, maxW, maxH);
    s.addImage({ data: imgData, x: x + (maxW - fit.w) / 2, y: y + (maxH - fit.h) / 2, w: fit.w, h: fit.h });
  }
  return { addHeader, addKeyMsg, addPageNum, addLogo, addSep, addDefBlock, addImgFit, C, pres };
}

// =============================================================
// Load product assets (images + badges)
// =============================================================
async function loadProductAssets(data) {
  const imageDir = (data.baseDir && !path.isAbsolute(data.imageDir))
    ? path.join(data.baseDir, data.imageDir)
    : data.imageDir;
  const imgs = {}, dims = {};
  for (const [key, file] of Object.entries(data.images)) {
    const fp = path.join(imageDir, file);
    const loaded = loadLocalImage(fp);
    if (loaded) {
      imgs[key] = loaded.data;
      dims[key] = await getImageDimensions(fp);
    }
  }
  const badges = {};
  for (const [key, svg] of Object.entries(data.badges || {})) {
    badges[key] = await makeBadge(svg);
  }
  return { imgs, dims, badges };
}

// =============================================================
// Product Slides (11 slides: Concept〜Projection, no cover/end)
// =============================================================
function addProductSlides(ctx, data, assets, pgOffset) {
  const { addHeader, addKeyMsg, addPageNum, addLogo, addSep, addDefBlock, addImgFit, C, pres } = ctx;
  const { imgs, dims, badges } = assets;
  let pg = pgOffset;

  // Slide: Concept
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "コンセプト");
    const top = centerY(3.1);
    const cn = data.concept;
    s.addText(cn.headline, { x: 0.5, y: top, w: 9, h: 0.5, fontSize: 24, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    s.addText(cn.body, { x: 0.5, y: top + 0.65, w: 9, h: 0.4, fontSize: 16, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0 });
    addSep(s, 0.5, top + 1.2, 9);
    s.addText(cn.matchLabel, { x: 0.5, y: top + 1.35, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0 });
    cn.points.slice(0, LAYOUT.maxPoints).forEach((p, i) => {
      const py = top + 1.85 + i * LAYOUT.bulletSpacing;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: py, w: LAYOUT.defBar.w, h: 0.28, fill: { color: C.primary } });
      s.addText(p, { x: 0.95, y: py, w: 8.5, h: 0.3, fontSize: 15, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
    });
    addKeyMsg(s, cn.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Product Overview
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "商品概要");
    const top = centerY(3.3);
    const ov = data.overview;
    s.addText(ov.headline, { x: 0.5, y: top, w: 9, h: 0.4, fontSize: 20, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    if (imgs[ov.leftImg]) addImgFit(s, imgs[ov.leftImg], dims[ov.leftImg], 0.5, top + 0.55, 2.0, 2.8);
    if (imgs[ov.rightImg]) addImgFit(s, imgs[ov.rightImg], dims[ov.rightImg], 2.3, top + 0.7, 2.5, 2.5);
    ov.features.slice(0, LAYOUT.maxFeatures).forEach((f, i) => {
      const fy = top + 0.65 + i * LAYOUT.featureSpacing;
      s.addShape(pres.shapes.OVAL, { x: 5.0, y: fy + 0.06, w: LAYOUT.bullet.w, h: LAYOUT.bullet.h, fill: { color: C.primary } });
      s.addText(f, { x: 5.3, y: fy, w: 4.2, h: 0.35, fontSize: 14, fontFace: FONT, color: C.body, bold: true, align: "left", valign: "middle", margin: 0 });
    });
    addKeyMsg(s, ov.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Materials / Features (3-column)
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, data.materials.header);
    const top = centerY(3.2);
    s.addText(data.materials.headline, { x: 0.5, y: top, w: 9, h: 0.4, fontSize: 20, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    const cardW = 2.8, cardGap = 0.3;
    const startX = (SW - (cardW * 3 + cardGap * 2)) / 2;
    data.materials.cards.slice(0, 3).forEach((c, i) => {
      const cx = startX + i * (cardW + cardGap), cy = top + 0.5;
      s.addText(c.title, { x: cx, y: cy, w: cardW, h: 0.25, fontSize: 13, fontFace: FONT, color: C.title, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addText(c.sub, { x: cx, y: cy + 0.25, w: cardW, h: 0.2, fontSize: 11, fontFace: FONT, color: C.sub, align: "center", valign: "middle", margin: 0 });
      if (imgs[c.imgKey]) {
        const fit = fitImage(dims[c.imgKey].w, dims[c.imgKey].h, cardW - 0.2, 1.3);
        s.addImage({ data: imgs[c.imgKey], x: cx + (cardW - fit.w) / 2, y: cy + 0.5, w: fit.w, h: fit.h });
      }
      s.addText(c.desc, { x: cx + 0.05, y: cy + 1.85, w: cardW - 0.1, h: 0.8, fontSize: 10, fontFace: FONT, color: C.body, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3 });
    });
    s.addText("※仕様変更可能", { x: 7.5, y: top + 3.2, w: 2, h: 0.2, fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0 });
    addKeyMsg(s, data.materials.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Quality Certification
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "品質証明");
    const top = centerY(3.2);
    const q = data.quality;
    s.addText("信頼性を保証するエビデンス", { x: 0.5, y: top, w: 7, h: 0.4, fontSize: 20, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    addDefBlock(s, 0.5, top + 0.5, q.cert1.title, q.cert1.body, 6.0);
    if (badges[q.cert1.badgeKey]) s.addImage({ data: badges[q.cert1.badgeKey], x: 7.3, y: top + 0.45, w: 1.1, h: 1.1 });
    addSep(s, 0.5, top + 1.6, 6.5);
    addDefBlock(s, 0.5, top + 1.75, q.cert2.title, q.cert2.body, 6.0);
    if (badges[q.cert2.badgeKey]) s.addImage({ data: badges[q.cert2.badgeKey], x: 7.3, y: top + 1.7, w: 1.1, h: 1.1 });
    addKeyMsg(s, q.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Competitor Comparison
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "競合比較");
    const top = centerY(2.8);
    const cmp = data.comparison;
    const tableBottom = renderTable(s, {
      pres, C, top, hdrH: 0.42, rowH: 0.4,
      colX: [0.5, 2.5, 5.0, 7.25], colW: [2.0, 2.5, 2.25, 2.25],
      headers: cmp.headers,
      rows: cmp.rows.slice(0, LAYOUT.maxCompRows),
      cellColorFn: (ri, ci) => ci === 1 ? C.primary : C.body,
    });
    s.addText("※独自比較", { x: 7.5, y: tableBottom + 0.05, w: 2, h: 0.2, fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0 });
    addKeyMsg(s, cmp.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Production Background
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "生産背景");
    const top = centerY(2.9);
    const pr = data.production;
    s.addText(pr.headline, { x: 0.5, y: top, w: 9, h: 0.4, fontSize: 20, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    addDefBlock(s, 0.5, top + 0.6, pr.block1.title, pr.block1.body, 5.5);
    addDefBlock(s, 0.5, top + 1.45, pr.block2.title, pr.block2.body, 5.5);
    if (imgs[pr.imgKey]) addImgFit(s, imgs[pr.imgKey], dims[pr.imgKey], 6.3, top + 0.5, 3.2, 2.0);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top + 2.4, w: 9, h: 0.5, fill: { color: C.kmBg } });
    s.addText(pr.callout, { x: 0.7, y: top + 2.4, w: 8.6, h: 0.5, fontSize: 15, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    addKeyMsg(s, pr.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Design Details
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, data.design.header);
    const top = centerY(3.4);
    const ds = data.design;
    s.addText(ds.benchmark, { x: 0.5, y: top, w: 4.0, h: 0.3, fontSize: 16, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
    if (imgs[ds.imgKey]) addImgFit(s, imgs[ds.imgKey], dims[ds.imgKey], 0.5, top + 0.4, 3.5, 2.9);
    const specX = 4.5, specW = 5.0;
    let sy = top + 0.4;
    ds.specs.forEach((spec) => {
      if (sy > LAYOUT.bodyBot - 0.8) return;
      s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: LAYOUT.bullet.w, h: LAYOUT.bullet.h, fill: { color: C.primary } });
      s.addText(spec.text, { x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3, fontSize: 13, fontFace: FONT, color: C.body, bold: true, align: "left", valign: "middle", margin: 0 });
      sy += 0.35;
      if (spec.sub) {
        spec.sub.forEach((sub) => {
          s.addText(sub, { x: specX + 0.5, y: sy, w: specW - 0.5, h: 0.25, fontSize: 12, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0 });
          sy += 0.25;
        });
        sy += 0.1;
      }
      if (spec.note) {
        s.addText(spec.note, { x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3, fontSize: 11, fontFace: FONT, color: C.sub, align: "left", valign: "middle", margin: 0 });
        sy += 0.4;
      }
    });
    s.addText("※デザイン変更可能", { x: 7.5, y: LAYOUT.bodyBot - 0.3, w: 2, h: 0.2, fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0 });
    addKeyMsg(s, ds.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Variations
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, data.variations.header);
    const top = centerY(3.0);
    const vr = data.variations;
    if (imgs[vr.imgKey]) addImgFit(s, imgs[vr.imgKey], dims[vr.imgKey], 0.5, top + 0.1, 4.5, 2.9);
    const infoX = 5.3;
    s.addShape(pres.shapes.OVAL, { x: infoX, y: top + 0.15, w: 0.13, h: 0.13, fill: { color: C.primary } });
    s.addText(vr.summary, { x: infoX + 0.25, y: top + 0.08, w: 4, h: 0.3, fontSize: 13, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0 });
    addSep(s, infoX, top + 0.55, 4.2);
    vr.items.slice(0, LAYOUT.maxVariations).forEach((c, i) => {
      const cy = top + 0.75 + i * LAYOUT.variationSpacing;
      s.addShape(pres.shapes.RECTANGLE, { x: infoX, y: cy, w: LAYOUT.defBar.w, h: 0.5, fill: { color: C.primary } });
      s.addText(c.label, { x: infoX + 0.2, y: cy, w: 4, h: 0.25, fontSize: 14, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(c.detail, { x: infoX + 0.2, y: cy + 0.25, w: 4, h: 0.22, fontSize: 12, fontFace: FONT, color: C.sub, align: "left", valign: "middle", margin: 0 });
    });
    addKeyMsg(s, vr.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Sales Strategy
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売戦略");
    const top = centerY(2.7);
    const st = data.sales;
    s.addText(st.headline, { x: 0.5, y: top, w: 6, h: 0.4, fontSize: 18, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    st.strategies.slice(0, LAYOUT.maxStrategies).forEach((item, i) => {
      const sy = top + 0.6 + i * LAYOUT.strategySpacing;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: sy, w: LAYOUT.defBar.w, h: 0.65, fill: { color: C.primary } });
      s.addText(item.title, { x: 0.75, y: sy, w: 5.5, h: 0.28, fontSize: 14, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(item.desc, { x: 0.75, y: sy + 0.3, w: 5.5, h: 0.28, fontSize: 12, fontFace: FONT, color: C.sub, align: "left", valign: "middle", margin: 0 });
    });
    const imgX = 6.5, imgW = 3.0;
    if (imgs[st.topImg]) addImgFit(s, imgs[st.topImg], dims[st.topImg], imgX, top + 0.1, imgW, 1.5);
    if (imgs[st.bottomImg]) addImgFit(s, imgs[st.bottomImg], dims[st.bottomImg], imgX, top + 1.7, imgW, 1.3);
    addKeyMsg(s, st.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Schedule
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "スケジュール");
    const top = centerY(3.0);
    const sc = data.schedule;
    const tableBottom = renderTable(s, {
      pres, C, top, hdrH: 0.38, rowH: 0.32,
      colX: [0.5, 2.7, 5.1], colW: [2.2, 2.4, 4.4],
      headers: ["時期", "進行管理", "詳細"],
      rows: sc.rows.slice(0, LAYOUT.maxSchedRows),
      cellColorFn: (ri, ci) => ci === 0 ? C.primary : C.body,
    });
    s.addText(sc.note, { x: 0.5, y: tableBottom + 0.08, w: 9, h: 0.3, fontSize: 11, fontFace: FONT, color: C.primary, align: "left", valign: "middle", margin: 0 });
    addKeyMsg(s, sc.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  // Slide: Sales Projection
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売イメージ");
    const top = centerY(2.8);
    const sp = data.projection;
    renderTable(s, {
      pres, C, top, hdrH: 0.38, rowH: 0.3,
      colX: [0.5, 2.7, 5.1], colW: [2.2, 2.4, 4.4],
      headers: ["項目", "内容", "備考"],
      rows: sp.rows.slice(0, LAYOUT.maxProjRows),
      cellColorFn: (ri, ci) => C.body,
    });
    addKeyMsg(s, sp.keyMsg);
    addPageNum(s, pg + 1);
    addLogo(s);
  }

  return pg;
}

// =============================================================
// Logo loader
// =============================================================
function loadLogo(logoField, baseDir) {
  if (!logoField) return null;
  const loaded = loadLocalImage(resolveAssetPath(logoField, baseDir));
  return loaded ? loaded.data : null;
}

// =============================================================
// Single Product Generator (後方互換)
// =============================================================
async function generate(data, outPath) {
  validateData(data);

  const C = buildColors(data.tone);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "提案チーム";
  pres.title = data.title + " ご提案資料";

  const date = data.date || new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long" });
  const logoData = loadLogo(data.logo, data.baseDir);
  const ctx = createSlideContext(pres, C, logoData);
  const assets = await loadProductAssets(data);
  const { addLogo } = ctx;

  let pg = 0;

  // ═══════════════════════════════════════════════
  // Slide 1: Cover
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 2.8) / 2;
    const cv = data.cover;
    if (assets.imgs[cv.leftImg]) ctx.addImgFit(s, assets.imgs[cv.leftImg], assets.dims[cv.leftImg], 0.2, (SH - 4.8) / 2, 2.8, 4.8);
    if (assets.imgs[cv.rightImg]) ctx.addImgFit(s, assets.imgs[cv.rightImg], assets.dims[cv.rightImg], SW - 3.0, (SH - 4.8) / 2, 2.8, 4.8);
    s.addText(cv.title, { x: 2.5, y: bTop, w: 5.0, h: 1.4, fontSize: 34, fontFace: FONT, color: C.title, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: bTop + 1.5, w: 1.25, h: LAYOUT.lineH, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: bTop + 1.5, w: 1.75, h: LAYOUT.lineH, fill: { color: C.secondary } });
    s.addText("ご提案資料", { x: 3.0, y: bTop + 1.7, w: 4.0, h: 0.4, fontSize: 16, fontFace: FONT, color: C.sub, align: "center", valign: "middle", margin: 0 });
    s.addText(date, { x: 3.0, y: bTop + 2.2, w: 4.0, h: 0.3, fontSize: 12, fontFace: FONT, color: C.muted, align: "center", valign: "middle", margin: 0 });
    addLogo(s);
  }

  // Product slides (11 slides)
  pg = addProductSlides(ctx, data, assets, pg);

  // ═══════════════════════════════════════════════
  // Slide 13: End
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const mid = SH / 2 - 0.5;
    s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: mid - 0.3, w: 3.25, h: LAYOUT.lineH, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: mid - 0.3, w: 3.75, h: LAYOUT.lineH, fill: { color: C.secondary } });
    s.addText(data.title + "\nご提案資料", { x: 1.5, y: mid, w: 7.0, h: 1.0, fontSize: 30, fontFace: FONT, color: C.title, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    s.addText("ご検討のほど、よろしくお願いいたします。", { x: 2.0, y: mid + 1.2, w: 6.0, h: 0.4, fontSize: 14, fontFace: FONT, color: C.sub, align: "center", valign: "middle", margin: 0 });
    s.addText(date, { x: 2.0, y: mid + 1.7, w: 6.0, h: 0.3, fontSize: 12, fontFace: FONT, color: C.muted, align: "center", valign: "middle", margin: 0 });
    addLogo(s);
  }

  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides)`);
}

// =============================================================
// Bundle Generator (複数商品を1資料にまとめる)
// =============================================================
async function generateBundle(bundleData, outPath) {
  const tone = bundleData.tone || "gold";
  const C = buildColors(tone);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "提案チーム";
  pres.title = bundleData.influencer + " 様 ご提案資料";

  const baseDir = bundleData.baseDir || ".";
  const date = bundleData.date || new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long" });
  const logoData = loadLogo(bundleData.logo, baseDir);
  const ctx = createSlideContext(pres, C, logoData);
  const { addHeader, addLogo } = ctx;

  // Load all product JSONs
  const products = bundleData.products.map(p => {
    const jsonPath = path.isAbsolute(p) ? p : path.join(baseDir, p);
    return JSON.parse(fs.readFileSync(jsonPath, "utf8"));
  });

  // Validate all
  products.forEach(p => validateData(p));

  let pg = 0;

  // ═══════════════════════════════════════════════
  // Slide 1: Bundle Cover
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 3.2) / 2;
    s.addText(bundleData.influencer + " 様", { x: 1.5, y: bTop, w: 7.0, h: 0.8, fontSize: 30, fontFace: FONT, color: C.title, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: bTop + 1.0, w: 1.25, h: LAYOUT.lineH, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: bTop + 1.0, w: 1.75, h: LAYOUT.lineH, fill: { color: C.secondary } });
    s.addText("商品企画ご提案資料", { x: 2.0, y: bTop + 1.2, w: 6.0, h: 0.5, fontSize: 20, fontFace: FONT, color: C.sub, align: "center", valign: "middle", margin: 0 });
    const productNames = products.map(p => p.title).join("  /  ");
    s.addText(productNames, { x: 1.0, y: bTop + 1.9, w: 8.0, h: 0.4, fontSize: 14, fontFace: FONT, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
    s.addText(date, { x: 3.0, y: bTop + 2.6, w: 4.0, h: 0.3, fontSize: 12, fontFace: FONT, color: C.muted, align: "center", valign: "middle", margin: 0 });
    addLogo(s);
  }

  // ═══════════════════════════════════════════════
  // Slide 2: Agenda (目次)
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "目次");
    const top = centerY(products.length * 0.7 + 0.5);
    products.forEach((p, i) => {
      const iy = top + i * 0.7;
      const num = String(i + 1).padStart(2, "0");
      s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: iy, w: LAYOUT.defBar.w, h: 0.5, fill: { color: C.primary } });
      s.addText(num, { x: 1.8, y: iy, w: 0.6, h: 0.5, fontSize: 24, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(p.title, { x: 2.5, y: iy, w: 5.0, h: 0.5, fontSize: 20, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
    });
    ctx.addPageNum(s, pg + 1);
    addLogo(s);
  }

  // ═══════════════════════════════════════════════
  // Product sections
  // ═══════════════════════════════════════════════
  for (let i = 0; i < products.length; i++) {
    const product = products[i];
    const assets = await loadProductAssets(product);

    // Section divider slide
    {
      pg++;
      const s = pres.addSlide();
      s.background = { color: C.bg };
      const mid = SH / 2 - 0.4;
      const num = String(i + 1).padStart(2, "0");
      s.addText(num, { x: 3.0, y: mid - 0.6, w: 4.0, h: 0.5, fontSize: 36, fontFace: FONT, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
      s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: mid + 0.05, w: 1.25, h: LAYOUT.lineH, fill: { color: C.primary } });
      s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: mid + 0.05, w: 1.75, h: LAYOUT.lineH, fill: { color: C.secondary } });
      s.addText(product.title, { x: 2.0, y: mid + 0.2, w: 6.0, h: 0.7, fontSize: 28, fontFace: FONT, color: C.title, bold: true, align: "center", valign: "middle", margin: 0 });
      ctx.addPageNum(s, pg + 1);
      addLogo(s);
    }

    // 11 product content slides
    pg = addProductSlides(ctx, product, assets, pg);
  }

  // ═══════════════════════════════════════════════
  // End slide
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const mid = SH / 2 - 0.5;
    s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: mid - 0.3, w: 3.25, h: LAYOUT.lineH, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: mid - 0.3, w: 3.75, h: LAYOUT.lineH, fill: { color: C.secondary } });
    s.addText(bundleData.influencer + " 様\n商品企画ご提案資料", { x: 1.5, y: mid, w: 7.0, h: 1.0, fontSize: 28, fontFace: FONT, color: C.title, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    s.addText("ご検討のほど、よろしくお願いいたします。", { x: 2.0, y: mid + 1.2, w: 6.0, h: 0.4, fontSize: 14, fontFace: FONT, color: C.sub, align: "center", valign: "middle", margin: 0 });
    s.addText(date, { x: 2.0, y: mid + 1.7, w: 6.0, h: 0.3, fontSize: 12, fontFace: FONT, color: C.muted, align: "center", valign: "middle", margin: 0 });
    addLogo(s);
  }

  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides, ${products.length} products)`);
}

module.exports = { generate, generateBundle, fitImage, loadLocalImage, getImageDimensions, makeBadge, buildColors, PALETTES, BASE_COLORS, FONT };
