const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

// =============================================================
// Image Utilities (v1と同一)
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
// SlideKit Design System
// =============================================================
const C = {
  bg: "F5F5F5", title: "222222", body: "333333", sub: "666666", muted: "AAAAAA",
  primary: "EF4823", secondary: "FCBF17", kmBg: "FFF5F0", sep: "EEEEEE",
  divider: "DDDDDD", white: "FFFFFF",
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

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "提案チーム";
  pres.title = "オリジナル万能調味料 ご提案資料";

  // ── Helpers ──
  function addHeader(s, titleText) {
    s.background = { color: C.bg };
    s.addText(titleText, {
      x: HDR.titleX, y: HDR.titleY, w: HDR.titleW, h: HDR.titleH,
      fontSize: 22, fontFace: FONT, color: C.title, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: HDR.redLineX, y: HDR.redLineY, w: HDR.redLineW, h: HDR.redLineH,
      fill: { color: C.primary }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: HDR.yellowLineX, y: HDR.yellowLineY, w: HDR.yellowLineW, h: HDR.yellowLineH,
      fill: { color: C.secondary }
    });
  }

  function addKeyMsg(s, text) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: KM.x, y: KM.y, w: KM.w, h: KM.h,
      fill: { color: C.kmBg }, rectRadius: 0.05
    });
    s.addText(text, {
      x: KM.x, y: KM.y, w: KM.w, h: KM.h,
      fontSize: 18, fontFace: FONT, color: C.primary, bold: true,
      align: "center", valign: "middle", margin: 0
    });
  }

  function addPageNum(s, num) {
    s.addText(String(num), {
      x: PN.x, y: PN.y, w: PN.w, h: PN.h,
      fontSize: 9, fontFace: FONT, color: C.muted,
      align: "right", valign: "middle", margin: 0
    });
  }

  function addSep(s, x, y, w) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.015, fill: { color: C.sep } });
  }

  function addDefBlock(s, x, y, heading, body, w) {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h: 0.7, fill: { color: C.primary } });
    s.addText(heading, {
      x: x + 0.2, y, w: w - 0.2, h: 0.28,
      fontSize: 16, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addText(body, {
      x: x + 0.2, y: y + 0.3, w: w - 0.2, h: 0.4,
      fontSize: 14, fontFace: FONT, color: C.body,
      align: "left", valign: "top", margin: 0
    });
  }

  function addImgFit(s, imgData, imgDims, x, y, maxW, maxH) {
    const fit = fitImage(imgDims.w, imgDims.h, maxW, maxH);
    const cx = x + (maxW - fit.w) / 2;
    const cy = y + (maxH - fit.h) / 2;
    s.addImage({ data: imgData, x: cx, y: cy, w: fit.w, h: fit.h });
  }

  // ── Load images ──
  const IMG_DIR = "/tmp/seasoning-images";
  const imgFiles = {
    spiceJars:    "spice-jars.jpg",      // 800x1063 portrait
    cookingPan:   "cooking-pan.jpg",     // 800x533 landscape
    spicesTable:  "spices-table.jpg",    // 800x1198 portrait
    momCooking:   "mom-cooking.jpg",     // 800x534 landscape
    foodPlate:    "food-plate.jpg",      // 800x533 landscape
    herbsFresh:   "herbs-fresh.jpg",     // 800x533 landscape
    kitchenPrep:  "kitchen-prep.jpg",    // 800x534 landscape
  };

  const imgs = {};
  const dims = {};
  for (const [key, file] of Object.entries(imgFiles)) {
    const fp = path.join(IMG_DIR, file);
    const loaded = loadLocalImage(fp);
    if (loaded) {
      imgs[key] = loaded.data;
      dims[key] = getImageDimensions(fp);
    }
  }

  // 国内GMP工場認証 (SVG生成 - industrial blue badge)
  const gmpSvg = `<svg width="300" height="300" xmlns="http://www.w3.org/2000/svg">
    <rect width="300" height="300" rx="150" fill="#1B3A5C"/>
    <circle cx="150" cy="150" r="120" fill="none" stroke="#A8C4E0" stroke-width="4"/>
    <circle cx="150" cy="150" r="108" fill="none" stroke="#A8C4E0" stroke-width="2"/>
    <text x="150" y="115" text-anchor="middle" fill="#A8C4E0" font-family="Helvetica,Arial" font-size="22" font-weight="bold">GMP</text>
    <text x="150" y="145" text-anchor="middle" fill="#E8F0F8" font-family="Helvetica,Arial" font-size="16" font-weight="bold">認定工場</text>
    <text x="150" y="175" text-anchor="middle" fill="#A8C4E0" font-family="Helvetica,Arial" font-size="12">CERTIFIED</text>
    <text x="150" y="200" text-anchor="middle" fill="#A8C4E0" font-family="Helvetica,Arial" font-size="11">MANUFACTURING</text>
  </svg>`;
  const gmpImg = "image/png;base64," + (await sharp(Buffer.from(gmpSvg)).png().toBuffer()).toString("base64");

  // 無添加認証 (SVG生成 - green/natural clean mark)
  const mutenkaSvg = `<svg width="280" height="340" xmlns="http://www.w3.org/2000/svg">
    <rect width="280" height="340" rx="10" fill="#F0F8F0" stroke="#2E7D32" stroke-width="3"/>
    <rect x="30" y="25" width="220" height="45" rx="5" fill="#2E7D32"/>
    <text x="140" y="55" text-anchor="middle" fill="#F0F8F0" font-family="Helvetica,Arial" font-size="18" font-weight="bold">無添加認証</text>
    <circle cx="140" cy="130" r="45" fill="none" stroke="#4CAF50" stroke-width="3"/>
    <text x="140" y="125" text-anchor="middle" fill="#2E7D32" font-family="Helvetica,Arial" font-size="28" font-weight="bold">✓</text>
    <text x="140" y="150" text-anchor="middle" fill="#2E7D32" font-family="Helvetica,Arial" font-size="10">ADDITIVE FREE</text>
    <line x1="50" y1="190" x2="230" y2="190" stroke="#A5D6A7" stroke-width="1"/>
    <text x="50" y="215" fill="#2E7D32" font-family="Helvetica,Arial" font-size="11">化学調味料</text>
    <text x="200" y="215" text-anchor="end" fill="#2E7D32" font-family="Helvetica,Arial" font-size="12" font-weight="bold">不使用</text>
    <text x="50" y="245" fill="#2E7D32" font-family="Helvetica,Arial" font-size="11">保存料</text>
    <text x="200" y="245" text-anchor="end" fill="#2E7D32" font-family="Helvetica,Arial" font-size="12" font-weight="bold">不使用</text>
    <text x="50" y="275" fill="#2E7D32" font-family="Helvetica,Arial" font-size="11">着色料</text>
    <text x="200" y="275" text-anchor="end" fill="#2E7D32" font-family="Helvetica,Arial" font-size="12" font-weight="bold">不使用</text>
    <text x="140" y="320" text-anchor="middle" fill="#4CAF50" font-family="Helvetica,Arial" font-size="9">NATURAL INGREDIENTS ONLY</text>
  </svg>`;
  const mutenkaImg = "image/png;base64," + (await sharp(Buffer.from(mutenkaSvg)).png().toBuffer()).toString("base64");

  let pg = 0;

  // ═══════════════════════════════════════════════
  // Slide 1: 表紙
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 2.8) / 2;

    if (imgs.spiceJars) {
      const fit = fitImage(dims.spiceJars.w, dims.spiceJars.h, 2.8, 4.8);
      s.addImage({ data: imgs.spiceJars, x: 0.2, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }
    if (imgs.foodPlate) {
      const fit = fitImage(dims.foodPlate.w, dims.foodPlate.h, 2.8, 4.8);
      s.addImage({ data: imgs.foodPlate, x: SW - fit.w - 0.2, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }

    s.addText("オリジナル\n万能調味料", {
      x: 2.5, y: bTop, w: 5.0, h: 1.4,
      fontSize: 34, fontFace: FONT, color: C.title, bold: true,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15
    });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.5, y: bTop + 1.5, w: 1.25, h: 0.035, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: bTop + 1.5, w: 1.75, h: 0.035, fill: { color: C.secondary } });
    s.addText("ご提案資料", {
      x: 3.0, y: bTop + 1.7, w: 4.0, h: 0.4,
      fontSize: 16, fontFace: FONT, color: C.sub,
      align: "center", valign: "middle", margin: 0
    });
    s.addText("2026年4月", {
      x: 3.0, y: bTop + 2.2, w: 4.0, h: 0.3,
      fontSize: 12, fontFace: FONT, color: C.muted,
      align: "center", valign: "middle", margin: 0
    });
  }

  // ═══════════════════════════════════════════════
  // Slide 2: コンセプト
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "コンセプト");
    const top = centerY(3.1);

    s.addText("「栄養士ママの味」をご家庭に", {
      x: 0.5, y: top, w: 9, h: 0.5,
      fontSize: 24, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addText(
      "栄養士資格を持つ3児のママが本気で作った、毎日の料理を格上げする万能調味料。",
      { x: 0.5, y: top + 0.65, w: 9, h: 0.4, fontSize: 16, fontFace: FONT, color: C.body,
        align: "left", valign: "middle", margin: 0 }
    );
    addSep(s, 0.5, top + 1.2, 9);

    s.addText("成瀬愛里さんとの相性:", {
      x: 0.5, y: top + 1.35, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: C.body,
      align: "left", valign: "middle", margin: 0
    });

    const points = [
      "栄養士・料理専門家の資格を最大限に活用",
      "32万フォロワーへのレシピ投稿との自然な連動",
      "3児のママが実感する「時短」「栄養」「おいしい」",
    ];
    points.forEach((p, i) => {
      const py = top + 1.85 + i * 0.4;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: py, w: 0.06, h: 0.28, fill: { color: C.primary } });
      s.addText(p, {
        x: 0.95, y: py, w: 8.5, h: 0.3,
        fontSize: 15, fontFace: FONT, color: C.title, bold: true,
        align: "left", valign: "middle", margin: 0
      });
    });

    addKeyMsg(s, "栄養士ママ x 時短レシピ x 万能調味料");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 3: 商品概要
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "商品概要");
    const top = centerY(3.3);

    s.addText("「振るだけ」で料理がプロの味に", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真
    if (imgs.spicesTable) {
      const fit = fitImage(dims.spicesTable.w, dims.spicesTable.h, 2.0, 2.8);
      s.addImage({ data: imgs.spicesTable, x: 0.5, y: top + 0.55, w: fit.w, h: fit.h });
    }
    if (imgs.herbsFresh) {
      const fit = fitImage(dims.herbsFresh.w, dims.herbsFresh.h, 2.5, 2.5);
      s.addImage({ data: imgs.herbsFresh, x: 2.3, y: top + 0.7, w: fit.w, h: fit.h });
    }

    // 右: 箇条書き
    const features = [
      "厳選した天然素材のみを使用",
      "化学調味料・保存料 無添加",
      "肉・魚・野菜・パスタ何にでも合う",
      "1本で味が決まるから時短調理に最適",
      "子供も安心して食べられる優しい味",
      "栄養士が栄養バランスも考慮した配合",
    ];
    features.forEach((f, i) => {
      const fy = top + 0.65 + i * 0.43;
      s.addShape(pres.shapes.OVAL, { x: 5.0, y: fy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
      s.addText(f, {
        x: 5.3, y: fy, w: 4.2, h: 0.35,
        fontSize: 14, fontFace: FONT, color: C.body, bold: true,
        align: "left", valign: "middle", margin: 0
      });
    });

    addKeyMsg(s, "無添加・天然素材の万能調味料");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 4: 原材料・機能性 — 3カラム
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "原材料・機能性");
    const top = centerY(3.2);

    s.addText("こだわり抜いた天然素材だけを使用", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const cards = [
      { title: "国産天日塩", sub: "ミネラル豊富", key: "spiceJars",
        desc: "海水を天日で結晶化\nまろやかで深みのある塩味\n精製塩では出せない旨味" },
      { title: "厳選スパイスブレンド", sub: "黄金比率", key: "herbsFresh",
        desc: "ブラックペッパー・ガーリック等\n栄養士が配合バランスを監修\nどの料理にも合う黄金比率" },
      { title: "国産かつお・昆布エキス", sub: "旨味の土台", key: "cookingPan",
        desc: "和の旨味をベースに\n洋食にも和食にも合う\n化学調味料に頼らない自然な旨味" },
    ];

    const cardW = 2.8, cardGap = 0.3;
    const startX = (SW - (cardW * 3 + cardGap * 2)) / 2;

    cards.forEach((c, i) => {
      const cx = startX + i * (cardW + cardGap);
      const cy = top + 0.5;

      s.addText(c.title, {
        x: cx, y: cy, w: cardW, h: 0.25,
        fontSize: 13, fontFace: FONT, color: C.title, bold: true,
        align: "center", valign: "middle", margin: 0
      });
      s.addText(c.sub, {
        x: cx, y: cy + 0.25, w: cardW, h: 0.2,
        fontSize: 11, fontFace: FONT, color: C.sub,
        align: "center", valign: "middle", margin: 0
      });

      if (imgs[c.key] && dims[c.key]) {
        const fit = fitImage(dims[c.key].w, dims[c.key].h, cardW - 0.2, 1.3);
        const ix = cx + (cardW - fit.w) / 2;
        s.addImage({ data: imgs[c.key], x: ix, y: cy + 0.5, w: fit.w, h: fit.h });
      }

      s.addText(c.desc, {
        x: cx + 0.05, y: cy + 1.85, w: cardW - 0.1, h: 0.8,
        fontSize: 10, fontFace: FONT, color: C.body,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3
      });
    });

    addKeyMsg(s, "天日塩+黄金比スパイス+和の旨味エキス");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 5: 品質証明
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "品質証明");
    const top = centerY(2.8);

    s.addText("信頼性を保証するエビデンス", {
      x: 0.5, y: top, w: 7, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addDefBlock(s, 0.5, top + 0.6, "国内GMP工場で製造",
      "食品衛生法に基づくGMP認定工場\n原料入荷から出荷まで一貫した品質管理", 6.0);
    s.addImage({ data: gmpImg, x: 7.0, y: top + 0.4, w: 1.5, h: 1.5 });

    addSep(s, 0.5, top + 1.5, 9);

    addDefBlock(s, 0.5, top + 1.7, "無添加処方",
      "化学調味料・保存料・着色料 不使用\n子供から大人まで安心して使える", 6.0);
    s.addImage({ data: mutenkaImg, x: 7.0, y: top + 1.5, w: 1.2, h: 1.7 });

    addKeyMsg(s, "GMP工場製造+完全無添加で安心品質");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 6: 競合比較テーブル
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "競合比較");
    const top = centerY(2.8);

    const headers = ["スペック項目", "本企画", "ほりにし", "マキシマム"];
    const dataRows = [
      ["無添加", "○(完全無添加)", "×(一部添加物)", "×(一部添加物)"],
      ["栄養士監修", "○", "×", "×"],
      ["和洋対応", "○(万能)", "△(洋食寄り)", "△(洋食寄り)"],
      ["容量", "80g", "100g", "140g"],
      ["販売価格", "800~1,500円", "900円前後", "500円前後"],
    ];

    const colX = [0.5, 2.5, 5.0, 7.25];
    const colW = [2.0, 2.5, 2.25, 2.25];
    const hdrH = 0.42, rowH = 0.4;

    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
    headers.forEach((h, i) => {
      s.addText(h, {
        x: colX[i], y: top, w: colW[i], h: hdrH,
        fontSize: 12, fontFace: FONT, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    });

    dataRows.forEach((row, ri) => {
      const ry = top + hdrH + ri * rowH;
      if (ri % 2 === 0) {
        s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: C.white } });
      }
      row.forEach((cell, ci) => {
        s.addText(cell, {
          x: colX[ci], y: ry, w: colW[ci], h: rowH,
          fontSize: 11, fontFace: FONT,
          color: ci === 1 ? C.primary : C.body,
          bold: ci <= 1,
          align: "center", valign: "middle", margin: 0
        });
      });
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });

    s.addText("※独自比較", {
      x: 7.5, y: top + hdrH + dataRows.length * rowH + 0.05, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "栄養士監修×完全無添加は本企画だけ");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 7: 生産背景
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "生産背景");
    const top = centerY(2.9);

    s.addText("国内GMP食品工場で一貫製造", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addDefBlock(s, 0.5, top + 0.6, "製造工場",
      "国内のGMP認定食品工場\nスパイス・調味料の製造実績豊富", 5.5);

    addDefBlock(s, 0.5, top + 1.45, "品質管理",
      "原料の残留農薬検査を全ロット実施\n微生物検査・金属検出も実施", 5.5);

    // 右: 工場写真
    if (imgs.kitchenPrep) {
      const fit = fitImage(dims.kitchenPrep.w, dims.kitchenPrep.h, 3.2, 2.0);
      s.addImage({ data: imgs.kitchenPrep, x: 6.3, y: top + 0.5, w: fit.w, h: fit.h });
    }

    addSep(s, 0.5, top + 2.25, 9);

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top + 2.4, w: 9, h: 0.5, fill: { color: C.kmBg }
    });
    s.addText("全ロットの検査成績書を発行可能です", {
      x: 0.7, y: top + 2.4, w: 8.6, h: 0.5,
      fontSize: 15, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "国内GMP工場で製造。全ロット検査済み");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 8: パッケージデザイン
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "パッケージデザイン");
    const photoTop = BODY_TOP + 0.1;

    s.addText("ほりにしをベンチマーク", {
      x: 0.5, y: photoTop, w: 4.0, h: 0.3,
      fontSize: 16, fontFace: FONT, color: C.title, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真
    if (imgs.spiceJars) {
      const fit = fitImage(dims.spiceJars.w, dims.spiceJars.h, 3.5, 3.2);
      s.addImage({ data: imgs.spiceJars, x: 0.5, y: photoTop + 0.4, w: fit.w, h: fit.h });
    }

    // 右: 仕様箇条書き
    const specX = 4.5, specW = 5.0;
    let sy = photoTop + 0.4;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("片手で振れるワンタッチキャップ", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.35;

    const subItems = ["1. 料理中でも片手操作OK", "2. 出し過ぎ防止の穴サイズ設計"];
    subItems.forEach((item) => {
      s.addText(item, {
        x: specX + 0.5, y: sy, w: specW - 0.5, h: 0.25,
        fontSize: 12, fontFace: FONT, color: C.body,
        align: "left", valign: "middle", margin: 0
      });
      sy += 0.25;
    });
    sy += 0.1;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("キッチンに置いても映えるナチュラルデザイン", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.3;
    s.addText("（SNS投稿時の背景にもなじむ）", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 11, fontFace: FONT, color: C.sub,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.4;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("湿気を防ぐ内蓋付きで長期保存可能", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    s.addText("※デザイン変更可能", {
      x: 7.5, y: BODY_BOT - 0.3, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "ワンタッチキャップ+映えるデザイン");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 9: ラインナップ
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "ラインナップ");
    const top = centerY(3.0);

    // 左: 写真
    if (imgs.foodPlate) {
      const fit = fitImage(dims.foodPlate.w, dims.foodPlate.h, 4.5, 2.9);
      s.addImage({ data: imgs.foodPlate, x: 0.5, y: top + 0.1, w: fit.w, h: fit.h });
    }

    const infoX = 5.3;
    s.addShape(pres.shapes.OVAL, { x: infoX, y: top + 0.15, w: 0.13, h: 0.13, fill: { color: C.primary } });
    s.addText("3種展開想定", {
      x: infoX + 0.25, y: top + 0.08, w: 4, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0
    });

    addSep(s, infoX, top + 0.55, 4.2);

    const types = [
      { label: "オリジナル(80g)", detail: "肉・魚・野菜なんでも合う万能タイプ" },
      { label: "ガーリック(80g)", detail: "にんにく好きに。炒め物・パスタに最適" },
      { label: "ゆず胡椒(60g)", detail: "和の香り。鍋・焼き鳥・刺身にも" },
    ];
    types.forEach((c, i) => {
      const cy = top + 0.75 + i * 0.65;
      s.addShape(pres.shapes.RECTANGLE, { x: infoX, y: cy, w: 0.06, h: 0.5, fill: { color: C.primary } });
      s.addText(c.label, {
        x: infoX + 0.2, y: cy, w: 4, h: 0.25,
        fontSize: 14, fontFace: FONT, color: C.title, bold: true,
        align: "left", valign: "middle", margin: 0
      });
      s.addText(c.detail, {
        x: infoX + 0.2, y: cy + 0.25, w: 4, h: 0.22,
        fontSize: 12, fontFace: FONT, color: C.sub,
        align: "left", valign: "middle", margin: 0
      });
    });

    addKeyMsg(s, "オリジナル x ガーリック x ゆず胡椒の3種");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 10: 販売戦略
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売戦略");
    const top = centerY(2.7);

    s.addText("レシピ投稿との連動で認知拡大", {
      x: 0.5, y: top, w: 6, h: 0.4,
      fontSize: 18, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const strategies = [
      { title: "毎日のレシピ投稿で調味料を自然露出", desc: "普段の料理投稿にさりげなくボトルを映す" },
      { title: "「この調味料だけで味が決まる」Reels", desc: "時短レシピ×万能調味料の訴求動画" },
      { title: "購入特典: レシピカード10枚セット", desc: "成瀬さんオリジナルレシピを同封" },
    ];

    strategies.forEach((st, i) => {
      const sy = top + 0.6 + i * 0.8;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: sy, w: 0.06, h: 0.65, fill: { color: C.primary } });
      s.addText(st.title, {
        x: 0.75, y: sy, w: 5.5, h: 0.28,
        fontSize: 14, fontFace: FONT, color: C.title, bold: true,
        align: "left", valign: "middle", margin: 0
      });
      s.addText(st.desc, {
        x: 0.75, y: sy + 0.3, w: 5.5, h: 0.28,
        fontSize: 12, fontFace: FONT, color: C.sub,
        align: "left", valign: "middle", margin: 0
      });
    });

    const imgX = 6.5, imgW = 3.0;
    if (imgs.momCooking) {
      const fit = fitImage(dims.momCooking.w, dims.momCooking.h, imgW, 1.5);
      s.addImage({ data: imgs.momCooking, x: imgX + (imgW - fit.w) / 2, y: top + 0.1, w: fit.w, h: fit.h });
    }
    if (imgs.foodPlate) {
      const fit = fitImage(dims.foodPlate.w, dims.foodPlate.h, imgW, 1.3);
      s.addImage({ data: imgs.foodPlate, x: imgX + (imgW - fit.w) / 2, y: top + 1.7, w: fit.w, h: fit.h });
    }

    addKeyMsg(s, "レシピ投稿+Reels+レシピカード特典で拡販");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 11: スケジュール
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "スケジュール");
    const top = centerY(3.0);

    const headers = ["時期", "進行管理", "詳細"];
    const dataRows = [
      ["4月", "企画・レシピ開発", "調味料コンセプト・配合方針の確定"],
      ["5月", "試作・味の調整", "栄養士監修のもと試作と改良"],
      ["6月", "パッケージデザイン確定", "ラベル・ボトル形状の最終決定"],
      ["7月", "量産", "GMP工場にて量産スタート"],
      ["8月", "製品納品・検品", "全ロット検査後、倉庫に納品"],
      ["9月", "プロモーション撮影", "レシピ動画・商品写真の撮影"],
      ["10月", "リリース", "EC+SNSプロモーション開始"],
    ];

    const colX = [0.5, 2.7, 5.1];
    const colW = [2.2, 2.4, 4.4];
    const hdrH = 0.38, rowH = 0.32;

    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
    headers.forEach((h, i) => {
      s.addText(h, {
        x: colX[i], y: top, w: colW[i], h: hdrH,
        fontSize: 12, fontFace: FONT, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    });

    dataRows.forEach((row, ri) => {
      const ry = top + hdrH + ri * rowH;
      if (ri % 2 === 0) {
        s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: C.white } });
      }
      row.forEach((cell, ci) => {
        s.addText(cell, {
          x: colX[ci], y: ry, w: colW[ci], h: rowH,
          fontSize: 11, fontFace: FONT,
          color: ci === 0 ? C.primary : C.body,
          bold: ci <= 1,
          align: "center", valign: "middle", margin: 0
        });
      });
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });

    s.addText("※秋の食欲シーズンに合わせたスケジュール", {
      x: 0.5, y: top + hdrH + dataRows.length * rowH + 0.08, w: 9, h: 0.3,
      fontSize: 11, fontFace: FONT, color: C.primary,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "4月企画 → 10月リリースの製造スケジュール");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 12: 販売イメージ
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売イメージ");
    const top = centerY(2.8);

    const headers = ["項目", "内容", "備考"];
    const dataRows = [
      ["販売価格", "800-1,500円", "タイプによって変動"],
      ["販売時期", "10月", "秋の食欲シーズンに合わせて"],
      ["原材料", "天然素材のみ", "化学調味料・保存料無添加"],
      ["品質", "GMP工場+全ロット検査", "栄養士監修の配合"],
      ["販売戦略", "レシピ投稿連動", "32万フォロワーへの自然露出"],
      ["ラインナップ", "3種", "オリジナル・ガーリック・ゆず胡椒"],
      ["生産数", "5,000-20,000個", "初回ロット想定"],
      ["売上想定", "400万-3,000万円", "販売価格800-1500円で算出"],
    ];

    const colX = [0.5, 2.7, 5.1];
    const colW = [2.2, 2.4, 4.4];
    const hdrH = 0.38, rowH = 0.3;

    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
    headers.forEach((h, i) => {
      s.addText(h, {
        x: colX[i], y: top, w: colW[i], h: hdrH,
        fontSize: 12, fontFace: FONT, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    });

    dataRows.forEach((row, ri) => {
      const ry = top + hdrH + ri * rowH;
      if (ri % 2 === 0) {
        s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: C.white } });
      }
      row.forEach((cell, ci) => {
        s.addText(cell, {
          x: colX[ci], y: ry, w: colW[ci], h: rowH,
          fontSize: 11, fontFace: FONT,
          color: C.body,
          bold: ci === 0,
          align: "center", valign: "middle", margin: 0
        });
      });
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });

    addKeyMsg(s, "売上想定400万〜3,000万円のポテンシャル");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 13: エンドスライド
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const mid = SH / 2 - 0.5;

    s.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: mid - 0.3, w: 3.25, h: 0.035, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: mid - 0.3, w: 3.75, h: 0.035, fill: { color: C.secondary } });

    s.addText("オリジナル万能調味料\nご提案資料", {
      x: 1.5, y: mid, w: 7.0, h: 1.0,
      fontSize: 30, fontFace: FONT, color: C.title, bold: true,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.2
    });
    s.addText("ご検討のほど、よろしくお願いいたします。", {
      x: 2.0, y: mid + 1.2, w: 6.0, h: 0.4,
      fontSize: 14, fontFace: FONT, color: C.sub,
      align: "center", valign: "middle", margin: 0
    });
    s.addText("2026年4月", {
      x: 2.0, y: mid + 1.7, w: 6.0, h: 0.3,
      fontSize: 12, fontFace: FONT, color: C.muted,
      align: "center", valign: "middle", margin: 0
    });
  }

  // ── Save ──
  const outPath = "/tmp/seasoning-proposal-v1.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides)`);
}

main().catch(e => { console.error(e); process.exit(1); });
