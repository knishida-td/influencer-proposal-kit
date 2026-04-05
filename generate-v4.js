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
  pres.title = "プレミアムヘアオイル ご提案資料";

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
  const IMG_DIR = "/tmp/hairoil-images";
  const imgFiles = {
    serumBottle:    "serum-bottle.jpg",      // 800x533
    womanHair:      "woman-hair.jpg",        // 800x1200
    hairBack:       "hair-back.jpg",         // 800x534
    cosmeticFlatlay:"cosmetic-flatlay.jpg",  // 800x1067
    naturalOil:     "natural-oil.jpg",       // 800x533
    labCosmetic:    "lab-cosmetic.jpg",      // 800x533
    dropper:        "dropper.jpg",           // 800x1200
    arganOil:       "argan-oil.jpg",         // 800x1200
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

  // COSMOS ORGANIC認証マーク (SVG生成)
  const cosmosSvg = `<svg width="300" height="300" xmlns="http://www.w3.org/2000/svg">
    <circle cx="150" cy="150" r="145" fill="#2E7D32" stroke="#1B5E20" stroke-width="4"/>
    <circle cx="150" cy="150" r="120" fill="none" stroke="#A5D6A7" stroke-width="2"/>
    <text x="150" y="100" text-anchor="middle" fill="#FFFFFF" font-family="Helvetica,Arial" font-size="22" font-weight="bold">COSMOS</text>
    <text x="150" y="135" text-anchor="middle" fill="#C8E6C9" font-family="Helvetica,Arial" font-size="18" font-weight="bold">ORGANIC</text>
    <text x="150" y="170" text-anchor="middle" fill="#A5D6A7" font-family="Helvetica,Arial" font-size="13">CERTIFIED</text>
    <path d="M120,195 Q150,215 180,195" fill="none" stroke="#A5D6A7" stroke-width="2"/>
    <circle cx="130" cy="210" r="5" fill="#66BB6A"/>
    <circle cx="150" cy="220" r="5" fill="#81C784"/>
    <circle cx="170" cy="210" r="5" fill="#66BB6A"/>
    <text x="150" y="260" text-anchor="middle" fill="#C8E6C9" font-family="Helvetica,Arial" font-size="10">NATURAL COSMETICS</text>
  </svg>`;
  const cosmosImg = "image/png;base64," + (await sharp(Buffer.from(cosmosSvg)).png().toBuffer()).toString("base64");

  // ノンシリコン認証マーク (SVG生成)
  const nonSiliSvg = `<svg width="280" height="340" xmlns="http://www.w3.org/2000/svg">
    <rect width="280" height="340" rx="10" fill="#FAFAFA" stroke="#37474F" stroke-width="3"/>
    <rect x="30" y="25" width="220" height="45" rx="5" fill="#37474F"/>
    <text x="140" y="55" text-anchor="middle" fill="#FFFFFF" font-family="Helvetica,Arial" font-size="16" font-weight="bold">NON-SILICONE</text>
    <text x="140" y="95" text-anchor="middle" fill="#37474F" font-family="Helvetica,Arial" font-size="14" font-weight="bold">CERTIFIED</text>
    <line x1="50" y1="110" x2="230" y2="110" stroke="#B0BEC5" stroke-width="1"/>
    <text x="50" y="140" fill="#546E7A" font-family="Helvetica,Arial" font-size="11">Product:</text>
    <text x="50" y="160" fill="#333333" font-family="Helvetica,Arial" font-size="12">Premium Hair Oil</text>
    <text x="50" y="190" fill="#546E7A" font-family="Helvetica,Arial" font-size="11">Ingredients:</text>
    <text x="50" y="210" fill="#333333" font-family="Helvetica,Arial" font-size="12">100% Silicone Free</text>
    <text x="50" y="240" fill="#546E7A" font-family="Helvetica,Arial" font-size="11">Standard:</text>
    <text x="50" y="260" fill="#333333" font-family="Helvetica,Arial" font-size="12">Natural Origin 95%+</text>
    <circle cx="140" cy="305" r="22" fill="none" stroke="#37474F" stroke-width="2"/>
    <text x="140" y="310" text-anchor="middle" fill="#37474F" font-family="Helvetica,Arial" font-size="9" font-weight="bold">VERIFIED</text>
  </svg>`;
  const nonSiliImg = "image/png;base64," + (await sharp(Buffer.from(nonSiliSvg)).png().toBuffer()).toString("base64");

  let pg = 0;

  // ═══════════════════════════════════════════════
  // Slide 1: 表紙
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 2.8) / 2;

    if (imgs.dropper) {
      const fit = fitImage(dims.dropper.w, dims.dropper.h, 2.8, 4.8);
      s.addImage({ data: imgs.dropper, x: 0.2, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }
    if (imgs.womanHair) {
      const fit = fitImage(dims.womanHair.w, dims.womanHair.h, 2.8, 4.8);
      s.addImage({ data: imgs.womanHair, x: SW - fit.w - 0.2, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }

    s.addText("プレミアム\nヘアオイル", {
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

    s.addText("「サロン帰りの艶」を毎日のケアで", {
      x: 0.5, y: top, w: 9, h: 0.5,
      fontSize: 24, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addText(
      "シャンプー月7000人定期購入の成功実績に続くヘアケア第2弾。",
      { x: 0.5, y: top + 0.65, w: 9, h: 0.4, fontSize: 16, fontFace: FONT, color: C.body,
        align: "left", valign: "middle", margin: 0 }
    );
    addSep(s, 0.5, top + 1.2, 9);

    s.addText("鈴木亜美さんとの相性:", {
      x: 0.5, y: top + 1.35, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: C.body,
      align: "left", valign: "middle", margin: 0
    });

    const points = [
      "シャンプー月7000人定期の実績をヘアオイルに横展開",
      "セレブ感のあるライフスタイル発信との親和性",
      "30代ママ層が求める時短ケアに合致",
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

    addKeyMsg(s, "シャンプー第2弾 x サロン品質 x 時短ヘアケア");
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

    s.addText("「艶」と「まとまり」を1プッシュで", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真
    if (imgs.dropper) {
      const fit = fitImage(dims.dropper.w, dims.dropper.h, 2.0, 2.8);
      s.addImage({ data: imgs.dropper, x: 0.5, y: top + 0.55, w: fit.w, h: fit.h });
    }
    if (imgs.cosmeticFlatlay) {
      const fit = fitImage(dims.cosmeticFlatlay.w, dims.cosmeticFlatlay.h, 2.5, 2.5);
      s.addImage({ data: imgs.cosmeticFlatlay, x: 2.3, y: top + 0.7, w: fit.w, h: fit.h });
    }

    // 右: 箇条書き
    const features = [
      "天然由来成分95%以上配合",
      "アルガンオイル+ホホバオイルのW処方",
      "ヒートプロテクト処方でドライヤー熱から保護",
      "ノンシリコンで髪に優しい",
      "1プッシュで毛先まで艶とまとまり",
      "華やかなフローラルムスクの香り",
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

    addKeyMsg(s, "天然由来95%のノンシリコンヘアオイル");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 4: 成分・機能性 — 3カラム
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "成分・機能性");
    const top = centerY(3.2);

    s.addText("サロン品質の成分を惜しみなく配合", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const cards = [
      { title: "アルガンオイル", sub: "モロッコ産", key: "arganOil",
        desc: "ビタミンEが豊富\nダメージ補修+保湿\n「液体の黄金」と呼ばれる希少オイル" },
      { title: "ホホバオイル", sub: "オーガニック認証", key: "naturalOil",
        desc: "人の皮脂に近い構造\n髪に素早く浸透\nベタつかないサラサラ仕上げ" },
      { title: "ヒートプロテクト処方", sub: "ドライヤー対応", key: "labCosmetic",
        desc: "180度までの熱から髪を保護\n毎日のドライヤーでも安心\nカラーの退色も防止" },
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

    s.addText("※処方変更可能", {
      x: 7.5, y: top + 3.2, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "アルガン+ホホバ+ヒートプロテクトの三重処方");
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

    addDefBlock(s, 0.5, top + 0.6, "COSMOS ORGANIC認証",
      "国際的なオーガニック化粧品認証\n天然由来成分95%以上の配合を第三者機関が保証", 6.0);
    s.addImage({ data: cosmosImg, x: 7.0, y: top + 0.4, w: 1.5, h: 1.5 });

    addSep(s, 0.5, top + 1.5, 9);

    addDefBlock(s, 0.5, top + 1.7, "ノンシリコン認証",
      "シリコン成分を一切含まない処方を認証\n髪と頭皮への負担を最小限に抑える設計", 6.0);
    s.addImage({ data: nonSiliImg, x: 7.0, y: top + 1.5, w: 1.2, h: 1.7 });

    addKeyMsg(s, "オーガニック認証+ノンシリコンで安心品質");
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

    const headers = ["スペック項目", "本企画", "N. (エヌドット)", "MOROCCANOIL"];
    const dataRows = [
      ["主成分", "アルガン+ホホバ", "シアバター", "アルガン"],
      ["ノンシリコン", "○", "×", "×"],
      ["容量", "100ml", "150ml", "100ml"],
      ["ヒートプロテクト", "○", "×", "○"],
      ["販売価格", "3千~5千", "3千~4千", "4千~5千"],
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

    addKeyMsg(s, "N.・モロッカンオイル同等品質を同価格帯で");
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

    s.addText("国内GMP工場で製造", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addDefBlock(s, 0.5, top + 0.6, "製造工場",
      "国内のGMP認定化粧品工場\n大手メーカーのOEM実績多数", 5.5);

    addDefBlock(s, 0.5, top + 1.45, "品質管理",
      "全ロット検査+安定性試験を実施\nアレルギーテスト済み", 5.5);

    // 右: 工場写真
    if (imgs.labCosmetic) {
      const fit = fitImage(dims.labCosmetic.w, dims.labCosmetic.h, 3.2, 2.0);
      s.addImage({ data: imgs.labCosmetic, x: 6.3, y: top + 0.5, w: fit.w, h: fit.h });
    }

    addSep(s, 0.5, top + 2.25, 9);

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top + 2.4, w: 9, h: 0.5, fill: { color: C.kmBg }
    });
    s.addText("工場見学・製造工程の撮影が可能です", {
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

    s.addText("N.をベンチマーク", {
      x: 0.5, y: photoTop, w: 4.0, h: 0.3,
      fontSize: 16, fontFace: FONT, color: C.title, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真
    if (imgs.dropper) {
      const fit = fitImage(dims.dropper.w, dims.dropper.h, 3.5, 3.2);
      s.addImage({ data: imgs.dropper, x: 0.5, y: photoTop + 0.4, w: fit.w, h: fit.h });
    }

    // 右: 仕様箇条書き
    const specX = 4.5, specW = 5.0;
    let sy = photoTop + 0.4;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("スポイト式で1回の使用量を調整しやすい", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.35;

    const subItems = ["1. ショートヘア: 1プッシュ", "2. ロングヘア: 2-3プッシュ"];
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
    s.addText("高級感のあるマットガラスボトル", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.3;
    s.addText("（洗面台に置いても映えるデザイン）", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 11, fontFace: FONT, color: C.sub,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.4;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("コンパクトサイズで旅行にも持ち運び可能", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    s.addText("※デザイン変更可能", {
      x: 7.5, y: BODY_BOT - 0.3, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "スポイト式+マットガラスで高見えデザイン");
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

    // 左: 商品写真
    if (imgs.serumBottle) {
      const fit = fitImage(dims.serumBottle.w, dims.serumBottle.h, 4.5, 2.9);
      s.addImage({ data: imgs.serumBottle, x: 0.5, y: top + 0.1, w: fit.w, h: fit.h });
    }

    const infoX = 5.3;
    s.addShape(pres.shapes.OVAL, { x: infoX, y: top + 0.15, w: 0.13, h: 0.13, fill: { color: C.primary } });
    s.addText("3タイプ展開", {
      x: infoX + 0.25, y: top + 0.08, w: 4, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0
    });

    addSep(s, infoX, top + 0.55, 4.2);

    const types = [
      { label: "スタンダード(100ml)", detail: "万能タイプ。朝晩のケアに" },
      { label: "ライト(100ml)", detail: "細い髪・猫っ毛向けの軽い仕上がり" },
      { label: "リッチ(50ml)", detail: "ダメージ毛向けの集中補修タイプ" },
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

    addKeyMsg(s, "3タイプ展開で髪質に合わせて選べる");
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

    s.addText("シャンプーとのクロスセルで売上最大化", {
      x: 0.5, y: top, w: 6, h: 0.4,
      fontSize: 18, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const strategies = [
      { title: "ヘアケアルーティン動画", desc: "朝のスタイリングにオイルを使う様子をReelsで" },
      { title: "シャンプーとのセット販売", desc: "既存シャンプーユーザーにセット割引で訴求" },
      { title: "購入特典: ミニサイズ(10ml)3本セット", desc: "お試しサイズで3タイプ全て試せる" },
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
    if (imgs.womanHair) {
      const fit = fitImage(dims.womanHair.w, dims.womanHair.h, imgW, 1.5);
      s.addImage({ data: imgs.womanHair, x: imgX + (imgW - fit.w) / 2, y: top + 0.1, w: fit.w, h: fit.h });
    }
    if (imgs.cosmeticFlatlay) {
      const fit = fitImage(dims.cosmeticFlatlay.w, dims.cosmeticFlatlay.h, imgW, 1.3);
      s.addImage({ data: imgs.cosmeticFlatlay, x: imgX + (imgW - fit.w) / 2, y: top + 1.7, w: fit.w, h: fit.h });
    }

    addKeyMsg(s, "シャンプーとのセット販売+Reels連動で拡販");
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
      ["4月", "企画・処方確定", "成分・香り・テクスチャーの最終決定"],
      ["5月", "試作・安定性試験", "処方の安定性試験を実施"],
      ["6月", "パッケージデザイン確定", "ボトル・箱・ラベルの最終決定"],
      ["7月", "量産", "GMP工場にて量産スタート"],
      ["8月", "製品納品・検品", "全ロット検査後、倉庫に納品"],
      ["9月", "プロモーション撮影", "ヘアケアルーティン動画撮影"],
      ["10月", "リリース", "EC+セット販売で販売開始"],
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

    s.addText("※シャンプーとの同時プロモーションを想定", {
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
      ["販売価格", "3,000-5,000円", "タイプ・容量によって変動"],
      ["販売時期", "10月", "秋冬ヘアケア需要に合わせて"],
      ["主成分", "アルガン+ホホバ", "天然由来95%以上"],
      ["品質", "ノンシリコン+オーガニック", "GMP工場製造"],
      ["販売戦略", "シャンプーとのセット販売", "既存顧客への横展開"],
      ["ラインナップ", "3タイプ", "スタンダード・ライト・リッチ"],
      ["生産数", "3,000-10,000本", "初回ロット想定"],
      ["売上想定", "900万-5,000万円", "販売価格3千-5千円で算出"],
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

    addKeyMsg(s, "売上想定900万〜5,000万円のポテンシャル");
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

    s.addText("プレミアムヘアオイル\nご提案資料", {
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
  const outPath = "/tmp/hairoil-proposal-v1.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides)`);
}

main().catch(e => { console.error(e); process.exit(1); });
