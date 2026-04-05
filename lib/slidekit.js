const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

// =============================================================
// Image Utilities
// =============================================================
function fitImage(origW, origH, maxW, maxH) {
  const ratio = origW / origH;
  let w = maxW, h = maxW / ratio;
  if (h > maxH) { h = maxH; w = maxH * ratio; }
  return { w, h };
}

function loadLocalImage(filePath) {
  if (!filePath || !fs.existsSync(filePath)) return null;
  const ext = path.extname(filePath).toLowerCase();
  const mime = ext === ".png" ? "image/png" : "image/jpeg";
  const buf = fs.readFileSync(filePath);
  return { data: `${mime};base64,${buf.toString("base64")}`, path: filePath };
}

function getImageDimensions(filePath) {
  try {
    const { execSync } = require("child_process");
    const out = execSync(`sips -g pixelWidth -g pixelHeight "${filePath}" 2>/dev/null`, { encoding: "utf8" });
    const wm = out.match(/pixelWidth:\s*(\d+)/);
    const hm = out.match(/pixelHeight:\s*(\d+)/);
    if (wm && hm) return { w: parseInt(wm[1]), h: parseInt(hm[1]) };
  } catch {}
  return { w: 800, h: 600 };
}

// =============================================================
// Design System (Dark Gold)
// =============================================================
const C = {
  bg: "111111", title: "FFFFFF", body: "E0E0E0", sub: "AAAAAA", muted: "666666",
  primary: "C9A96E", secondary: "E8D5B0", kmBg: "1F1A14", sep: "333333",
  divider: "333333", white: "FFFFFF", rowAlt: "1A1A1A",
};
const FONT = "Hiragino Kaku Gothic Pro W3";
const SW = 10, SH = 5.625;

const HDR = {
  titleX: 0.5, titleY: 0.39, titleW: 9.0, titleH: 0.45,
  redLineX: 0.5, redLineY: 0.857, redLineW: 4.25, redLineH: 0.035,
  yellowLineX: 4.75, yellowLineY: 0.857, yellowLineW: 4.75, yellowLineH: 0.035,
};
const KM = { x: 0.5, y: 4.837, w: 9.0, h: 0.4 };
const PN = { x: 9.2, y: 5.337, w: 0.5, h: 0.25 };
const BODY_TOP = 0.893;
const BODY_BOT = 4.837;
const BODY_H = BODY_BOT - BODY_TOP;

function centerY(contentH) {
  return BODY_TOP + (BODY_H - contentH) / 2;
}

// =============================================================
// SVG Badge Generator
// =============================================================
async function makeBadge(svg) {
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// =============================================================
// Main Generator
// =============================================================
async function generate(data, outPath) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "提案チーム";
  pres.title = data.title + " ご提案資料";

  // ── Helpers (bound to pres) ──
  function addHeader(s, titleText) {
    s.background = { color: C.bg };
    s.addText(titleText, {
      x: HDR.titleX, y: HDR.titleY, w: HDR.titleW, h: HDR.titleH,
      fontSize: 22, fontFace: FONT, color: C.title, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, { x: HDR.redLineX, y: HDR.redLineY, w: HDR.redLineW, h: HDR.redLineH, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: HDR.yellowLineX, y: HDR.yellowLineY, w: HDR.yellowLineW, h: HDR.yellowLineH, fill: { color: C.secondary } });
  }

  function addKeyMsg(s, text) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: KM.x, y: KM.y, w: KM.w, h: KM.h, fill: { color: C.kmBg }, rectRadius: 0.05 });
    s.addText(text, { x: KM.x, y: KM.y, w: KM.w, h: KM.h, fontSize: 18, fontFace: FONT, color: C.primary, bold: true, align: "center", valign: "middle", margin: 0 });
  }

  function addPageNum(s, num) {
    s.addText(String(num), { x: PN.x, y: PN.y, w: PN.w, h: PN.h, fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0 });
  }

  function addSep(s, x, y, w) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.015, fill: { color: C.sep } });
  }

  function addDefBlock(s, x, y, heading, body, w) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 0.7, fill: { color: C.primary } });
    s.addText(heading, { x: x + 0.2, y, w: w - 0.2, h: 0.28, fontSize: 16, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    s.addText(body, { x: x + 0.2, y: y + 0.3, w: w - 0.2, h: 0.4, fontSize: 14, fontFace: FONT, color: C.body, align: "left", valign: "top", margin: 0 });
  }

  function addImgFit(s, imgData, imgDims, x, y, maxW, maxH) {
    const fit = fitImage(imgDims.w, imgDims.h, maxW, maxH);
    s.addImage({ data: imgData, x: x + (maxW - fit.w) / 2, y: y + (maxH - fit.h) / 2, w: fit.w, h: fit.h });
  }

  // ── Load images ──
  const imgs = {}, dims = {};
  for (const [key, file] of Object.entries(data.images)) {
    const fp = path.join(data.imageDir, file);
    const loaded = loadLocalImage(fp);
    if (loaded) {
      imgs[key] = loaded.data;
      dims[key] = getImageDimensions(fp);
    }
  }

  // ── SVG Badges ──
  const badges = {};
  for (const [key, svg] of Object.entries(data.badges || {})) {
    badges[key] = await makeBadge(svg);
  }

  let pg = 0;

  // ═══════════════════════════════════════════════
  // Slide 1: Cover
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 2.8) / 2;
    const cv = data.cover;
    if (imgs[cv.leftImg]) addImgFit(s, imgs[cv.leftImg], dims[cv.leftImg], 0.2, (SH - 4.8) / 2, 2.8, 4.8);
    if (imgs[cv.rightImg]) addImgFit(s, imgs[cv.rightImg], dims[cv.rightImg], SW - 3.0, (SH - 4.8) / 2, 2.8, 4.8);
    s.addText(cv.title, { x: 2.5, y: bTop, w: 5.0, h: 1.4, fontSize: 34, fontFace: FONT, color: C.title, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15 });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: bTop + 1.5, w: 1.25, h: 0.035, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: bTop + 1.5, w: 1.75, h: 0.035, fill: { color: C.secondary } });
    s.addText("ご提案資料", { x: 3.0, y: bTop + 1.7, w: 4.0, h: 0.4, fontSize: 16, fontFace: FONT, color: C.sub, align: "center", valign: "middle", margin: 0 });
    s.addText("2026年4月", { x: 3.0, y: bTop + 2.2, w: 4.0, h: 0.3, fontSize: 12, fontFace: FONT, color: C.muted, align: "center", valign: "middle", margin: 0 });
  }

  // ═══════════════════════════════════════════════
  // Slide 2: Concept
  // ═══════════════════════════════════════════════
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
    cn.points.forEach((p, i) => {
      const py = top + 1.85 + i * 0.4;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: py, w: 0.06, h: 0.28, fill: { color: C.primary } });
      s.addText(p, { x: 0.95, y: py, w: 8.5, h: 0.3, fontSize: 15, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
    });
    addKeyMsg(s, cn.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 3: Product Overview
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "商品概要");
    const top = centerY(3.3);
    const ov = data.overview;
    s.addText(ov.headline, { x: 0.5, y: top, w: 9, h: 0.4, fontSize: 20, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    if (imgs[ov.leftImg]) addImgFit(s, imgs[ov.leftImg], dims[ov.leftImg], 0.5, top + 0.55, 2.0, 2.8);
    if (imgs[ov.rightImg]) addImgFit(s, imgs[ov.rightImg], dims[ov.rightImg], 2.3, top + 0.7, 2.5, 2.5);
    ov.features.forEach((f, i) => {
      const fy = top + 0.65 + i * 0.43;
      s.addShape(pres.shapes.OVAL, { x: 5.0, y: fy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
      s.addText(f, { x: 5.3, y: fy, w: 4.2, h: 0.35, fontSize: 14, fontFace: FONT, color: C.body, bold: true, align: "left", valign: "middle", margin: 0 });
    });
    addKeyMsg(s, ov.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 4: Materials / Features (3-column)
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, data.materials.header);
    const top = centerY(3.2);
    s.addText(data.materials.headline, { x: 0.5, y: top, w: 9, h: 0.4, fontSize: 20, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    const cardW = 2.8, cardGap = 0.3;
    const startX = (SW - (cardW * 3 + cardGap * 2)) / 2;
    data.materials.cards.forEach((c, i) => {
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
  }

  // ═══════════════════════════════════════════════
  // Slide 5: Quality Certification
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "品質証明");
    const top = centerY(3.2);
    const q = data.quality;
    s.addText("信頼性を保証するエビデンス", { x: 0.5, y: top, w: 7, h: 0.4, fontSize: 20, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    // cert1: top+0.45 ~ top+1.55 (badge 1.1inch)
    addDefBlock(s, 0.5, top + 0.5, q.cert1.title, q.cert1.body, 6.0);
    if (badges[q.cert1.badgeKey]) s.addImage({ data: badges[q.cert1.badgeKey], x: 7.3, y: top + 0.45, w: 1.1, h: 1.1 });
    // sep: top+1.6
    addSep(s, 0.5, top + 1.6, 6.5);
    // cert2: top+1.7 ~ top+2.8 (badge 1.1inch)
    addDefBlock(s, 0.5, top + 1.75, q.cert2.title, q.cert2.body, 6.0);
    if (badges[q.cert2.badgeKey]) s.addImage({ data: badges[q.cert2.badgeKey], x: 7.3, y: top + 1.7, w: 1.1, h: 1.1 });
    addKeyMsg(s, q.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 6: Competitor Comparison
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "競合比較");
    const top = centerY(2.8);
    const cmp = data.comparison;
    const colX = [0.5, 2.5, 5.0, 7.25], colW = [2.0, 2.5, 2.25, 2.25];
    const hdrH = 0.42, rowH = 0.4;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
    cmp.headers.forEach((h, i) => {
      s.addText(h, { x: colX[i], y: top, w: colW[i], h: hdrH, fontSize: 12, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    });
    cmp.rows.forEach((row, ri) => {
      const ry = top + hdrH + ri * rowH;
      if (ri % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: C.rowAlt } });
      row.forEach((cell, ci) => {
        s.addText(cell, { x: colX[ci], y: ry, w: colW[ci], h: rowH, fontSize: 11, fontFace: FONT, color: ci === 1 ? C.primary : C.body, bold: ci <= 1, align: "center", valign: "middle", margin: 0 });
      });
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });
    s.addText("※独自比較", { x: 7.5, y: top + hdrH + cmp.rows.length * rowH + 0.05, w: 2, h: 0.2, fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0 });
    addKeyMsg(s, cmp.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 7: Production Background
  // ═══════════════════════════════════════════════
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
  }

  // ═══════════════════════════════════════════════
  // Slide 8: Design Details
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, data.design.header);
    const photoTop = BODY_TOP + 0.1;
    const ds = data.design;
    s.addText(ds.benchmark, { x: 0.5, y: photoTop, w: 4.0, h: 0.3, fontSize: 16, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
    if (imgs[ds.imgKey]) addImgFit(s, imgs[ds.imgKey], dims[ds.imgKey], 0.5, photoTop + 0.4, 3.5, 3.2);
    const specX = 4.5, specW = 5.0;
    let sy = photoTop + 0.4;
    ds.specs.forEach((spec) => {
      s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
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
    s.addText("※デザイン変更可能", { x: 7.5, y: BODY_BOT - 0.3, w: 2, h: 0.2, fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0 });
    addKeyMsg(s, ds.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 9: Variations
  // ═══════════════════════════════════════════════
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
    vr.items.forEach((c, i) => {
      const cy = top + 0.75 + i * 0.65;
      s.addShape(pres.shapes.RECTANGLE, { x: infoX, y: cy, w: 0.06, h: 0.5, fill: { color: C.primary } });
      s.addText(c.label, { x: infoX + 0.2, y: cy, w: 4, h: 0.25, fontSize: 14, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(c.detail, { x: infoX + 0.2, y: cy + 0.25, w: 4, h: 0.22, fontSize: 12, fontFace: FONT, color: C.sub, align: "left", valign: "middle", margin: 0 });
    });
    addKeyMsg(s, vr.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 10: Sales Strategy
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売戦略");
    const top = centerY(2.7);
    const st = data.sales;
    s.addText(st.headline, { x: 0.5, y: top, w: 6, h: 0.4, fontSize: 18, fontFace: FONT, color: C.primary, bold: true, align: "left", valign: "middle", margin: 0 });
    st.strategies.forEach((item, i) => {
      const sy = top + 0.6 + i * 0.8;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: sy, w: 0.06, h: 0.65, fill: { color: C.primary } });
      s.addText(item.title, { x: 0.75, y: sy, w: 5.5, h: 0.28, fontSize: 14, fontFace: FONT, color: C.title, bold: true, align: "left", valign: "middle", margin: 0 });
      s.addText(item.desc, { x: 0.75, y: sy + 0.3, w: 5.5, h: 0.28, fontSize: 12, fontFace: FONT, color: C.sub, align: "left", valign: "middle", margin: 0 });
    });
    const imgX = 6.5, imgW = 3.0;
    if (imgs[st.topImg]) addImgFit(s, imgs[st.topImg], dims[st.topImg], imgX, top + 0.1, imgW, 1.5);
    if (imgs[st.bottomImg]) addImgFit(s, imgs[st.bottomImg], dims[st.bottomImg], imgX, top + 1.7, imgW, 1.3);
    addKeyMsg(s, st.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 11: Schedule
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "スケジュール");
    const top = centerY(3.0);
    const sc = data.schedule;
    const colX = [0.5, 2.7, 5.1], colW = [2.2, 2.4, 4.4];
    const hdrH = 0.38, rowH = 0.32;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
    ["時期", "進行管理", "詳細"].forEach((h, i) => {
      s.addText(h, { x: colX[i], y: top, w: colW[i], h: hdrH, fontSize: 12, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    });
    sc.rows.forEach((row, ri) => {
      const ry = top + hdrH + ri * rowH;
      if (ri % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: C.rowAlt } });
      row.forEach((cell, ci) => {
        s.addText(cell, { x: colX[ci], y: ry, w: colW[ci], h: rowH, fontSize: 11, fontFace: FONT, color: ci === 0 ? C.primary : C.body, bold: ci <= 1, align: "center", valign: "middle", margin: 0 });
      });
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });
    s.addText(sc.note, { x: 0.5, y: top + hdrH + sc.rows.length * rowH + 0.08, w: 9, h: 0.3, fontSize: 11, fontFace: FONT, color: C.primary, align: "left", valign: "middle", margin: 0 });
    addKeyMsg(s, sc.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 12: Sales Projection
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売イメージ");
    const top = centerY(2.8);
    const sp = data.projection;
    const colX = [0.5, 2.7, 5.1], colW = [2.2, 2.4, 4.4];
    const hdrH = 0.38, rowH = 0.3;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
    ["項目", "内容", "備考"].forEach((h, i) => {
      s.addText(h, { x: colX[i], y: top, w: colW[i], h: hdrH, fontSize: 12, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    });
    sp.rows.forEach((row, ri) => {
      const ry = top + hdrH + ri * rowH;
      if (ri % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: C.rowAlt } });
      row.forEach((cell, ci) => {
        s.addText(cell, { x: colX[ci], y: ry, w: colW[ci], h: rowH, fontSize: 11, fontFace: FONT, color: C.body, bold: ci === 0, align: "center", valign: "middle", margin: 0 });
      });
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });
    addKeyMsg(s, sp.keyMsg);
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 13: End
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const mid = SH / 2 - 0.5;
    s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: mid - 0.3, w: 3.25, h: 0.035, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: mid - 0.3, w: 3.75, h: 0.035, fill: { color: C.secondary } });
    s.addText(data.title + "\nご提案資料", { x: 1.5, y: mid, w: 7.0, h: 1.0, fontSize: 30, fontFace: FONT, color: C.title, bold: true, align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2 });
    s.addText("ご検討のほど、よろしくお願いいたします。", { x: 2.0, y: mid + 1.2, w: 6.0, h: 0.4, fontSize: 14, fontFace: FONT, color: C.sub, align: "center", valign: "middle", margin: 0 });
    s.addText("2026年4月", { x: 2.0, y: mid + 1.7, w: 6.0, h: 0.3, fontSize: 12, fontFace: FONT, color: C.muted, align: "center", valign: "middle", margin: 0 });
  }

  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides)`);
}

module.exports = { generate, fitImage, loadLocalImage, getImageDimensions, makeBadge, C, FONT };
