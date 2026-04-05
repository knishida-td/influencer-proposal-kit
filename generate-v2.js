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
  pres.title = "高品質イタリアンレザーミニバッグ ご提案資料";

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
  const IMG_DIR = "/tmp/bag-proposal-images";
  const imgFiles = {
    coverLeft:     "woman-brown-coat.jpg",   // コート+黒バッグの女性 (800x1200)
    coverRight:    "woman-bag-elegant.jpg",  // ブラウンミニバッグの女性 (800x1200)
    portraitBag:   "woman-portrait-bag.jpg", // グリーンミニショルダー (800x1120)
    miniBag:       "woman-holding-bag2.jpg", // 黒クロコミニバッグ (800x1067)
    flatLay:       "flat-lay-bag.jpg",       // バッグ+小物フラットレイ (800x533)
    bagCloseup:    "bag-closeup.jpg",        // レザーバッグクローズアップ (800x533)
    leatherCraft:  "leather-craft.jpg",      // 職人リベット打ち (800x533)
    leatherCutting:"leather-cutting.jpg",    // レザー裁断工房 (800x533)
    giftPkg:       "gift-packaging.jpg",     // ギフト包装 (800x1208)
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

  // Vera Pelle認証マーク (SVG生成)
  const veraSvg = `<svg width="300" height="300" xmlns="http://www.w3.org/2000/svg">
    <rect width="300" height="300" rx="150" fill="#8B4513"/>
    <rect x="50" y="80" width="200" height="140" rx="10" fill="none" stroke="#F5DEB3" stroke-width="4"/>
    <text x="150" y="135" text-anchor="middle" fill="#F5DEB3" font-family="Georgia,serif" font-size="28" font-weight="bold">VERA</text>
    <text x="150" y="170" text-anchor="middle" fill="#F5DEB3" font-family="Georgia,serif" font-size="28" font-weight="bold">PELLE</text>
    <text x="150" y="205" text-anchor="middle" fill="#D2B48C" font-family="Helvetica,Arial" font-size="14">ITALIAN LEATHER</text>
    <text x="150" y="250" text-anchor="middle" fill="#D2B48C" font-family="Helvetica,Arial" font-size="11">CERTIFIED</text>
  </svg>`;
  const veraImg = "image/png;base64," + (await sharp(Buffer.from(veraSvg)).png().toBuffer()).toString("base64");

  // Consorzio認証 (SVG生成)
  const consSvg = `<svg width="280" height="340" xmlns="http://www.w3.org/2000/svg">
    <rect width="280" height="340" rx="10" fill="#FFFAF0" stroke="#8B4513" stroke-width="3"/>
    <rect x="30" y="25" width="220" height="45" rx="5" fill="#8B4513"/>
    <text x="140" y="55" text-anchor="middle" fill="#FFFAF0" font-family="Georgia,serif" font-size="16" font-weight="bold">CONSORZIO VERA PELLE</text>
    <text x="140" y="95" text-anchor="middle" fill="#8B4513" font-family="Georgia,serif" font-size="14" font-weight="bold">CERTIFICATO</text>
    <line x1="50" y1="105" x2="230" y2="105" stroke="#D2B48C" stroke-width="1"/>
    <text x="50" y="130" fill="#5C3317" font-family="Helvetica,Arial" font-size="11">Prodotto:</text>
    <text x="50" y="150" fill="#333333" font-family="Helvetica,Arial" font-size="12">Mini Shoulder Bag</text>
    <text x="50" y="180" fill="#5C3317" font-family="Helvetica,Arial" font-size="11">Materiale:</text>
    <text x="50" y="200" fill="#333333" font-family="Helvetica,Arial" font-size="12">Vegetable Tanned Leather</text>
    <text x="50" y="230" fill="#5C3317" font-family="Helvetica,Arial" font-size="11">Origine:</text>
    <text x="50" y="250" fill="#333333" font-family="Helvetica,Arial" font-size="12">Toscana, Italia</text>
    <circle cx="140" cy="300" r="22" fill="none" stroke="#8B4513" stroke-width="2"/>
    <text x="140" y="305" text-anchor="middle" fill="#8B4513" font-family="Helvetica,Arial" font-size="9" font-weight="bold">VERIFIED</text>
  </svg>`;
  const consImg = "image/png;base64," + (await sharp(Buffer.from(consSvg)).png().toBuffer()).toString("base64");

  let pg = 0;

  // ═══════════════════════════════════════════════
  // Slide 1: 表紙
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 2.8) / 2;

    if (imgs.coverLeft) {
      const fit = fitImage(dims.coverLeft.w, dims.coverLeft.h, 2.8, 4.8);
      s.addImage({ data: imgs.coverLeft, x: 0.2, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }
    if (imgs.coverRight) {
      const fit = fitImage(dims.coverRight.w, dims.coverRight.h, 2.8, 4.8);
      s.addImage({ data: imgs.coverRight, x: SW - fit.w - 0.2, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }

    s.addText("高品質\nレザーミニバッグ", {
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

    s.addText("「毎日持ちたい」上質を手元に", {
      x: 0.5, y: top, w: 9, h: 0.5,
      fontSize: 24, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addText(
      "ダウンジャケットに続く第2弾。冬のコーディネートを完成させるミニバッグ。",
      { x: 0.5, y: top + 0.65, w: 9, h: 0.4, fontSize: 16, fontFace: FONT, color: C.body,
        align: "left", valign: "middle", margin: 0 }
    );
    addSep(s, 0.5, top + 1.2, 9);

    s.addText("鈴木亜美さんとの相性:", {
      x: 0.5, y: top + 1.35, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: C.body,
      align: "left", valign: "middle", margin: 0
    });

    const points = [
      "着圧ソックス2万足即完売の販売力をバッグでも",
      "ダウンジャケットとのセットコーデ提案が可能",
      "30代ママ層が求める「上質だけど実用的」に合致",
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

    addKeyMsg(s, "ダウンJK第2弾 x 上質レザー x 即戦力ミニバッグ");
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

    s.addText("「映える」と「使える」を両立するミニバッグ", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真
    if (imgs.portraitBag) {
      const fit = fitImage(dims.portraitBag.w, dims.portraitBag.h, 2.0, 2.8);
      s.addImage({ data: imgs.portraitBag, x: 0.5, y: top + 0.55, w: fit.w, h: fit.h });
    }
    if (imgs.flatLay) {
      const fit = fitImage(dims.flatLay.w, dims.flatLay.h, 2.5, 2.5);
      s.addImage({ data: imgs.flatLay, x: 2.3, y: top + 0.7, w: fit.w, h: fit.h });
    }

    // 右: 箇条書き
    const features = [
      "イタリア・トスカーナ産の本革を使用",
      "ショルダー+ハンドの2WAY仕様",
      "スマホ・財布・鍵がちょうど入るサイズ",
      "約350gで一日中持っても疲れない軽さ",
      "アンティークゴールドの上品な金具",
      "使い込むほど味が出るエイジング素材",
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

    addKeyMsg(s, "トスカーナレザーの2WAYミニバッグ、約350g");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 4: 素材・機能性 — 3カラム
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "素材・機能性");
    const top = centerY(3.2);

    s.addText("ハイブランドと同等の素材と製法", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const cards = [
      { title: "ベジタブルタンニンレザー", sub: "トスカーナ伝統製法", key: "bagCloseup",
        desc: "植物由来のタンニンで鞣した\n環境に優しい天然皮革\n使い込むほど色艶が深まる" },
      { title: "アンティークゴールド金具", sub: "真鍮ベース", key: "flatLay",
        desc: "高級感のある落ち着いたゴールド\n経年変化でヴィンテージ感が出る\nニッケルフリーで肌にも安心" },
      { title: "職人の手縫い仕上げ", sub: "耐久性+上品さ", key: "leatherCraft",
        desc: "熟練職人による手縫いと\nリベット打ちで耐久性を確保\nブランドロゴの刻印入り" },
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

    s.addText("※仕様変更可能", {
      x: 7.5, y: top + 3.2, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "トスカーナ伝統製法の天然皮革+真鍮金具");
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

    addDefBlock(s, 0.5, top + 0.6, "Vera Pelle認証: イタリア本革の証明",
      "イタリア政府公認の本革認証マーク\n合成皮革やPUレザーではない本物の革であることを保証", 6.0);
    s.addImage({ data: veraImg, x: 7.0, y: top + 0.4, w: 1.5, h: 1.5 });

    addSep(s, 0.5, top + 1.5, 9);

    addDefBlock(s, 0.5, top + 1.7, "Consorzio Vera Pelle Italiana認証",
      "トスカーナのベジタブルタンニンレザー協会による\n製造工程と原産地の認証", 6.0);
    s.addImage({ data: consImg, x: 7.0, y: top + 1.5, w: 1.2, h: 1.7 });

    addKeyMsg(s, "イタリア公認の本革認証で品質を担保");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 6: 競合比較テーブル
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "他ハイブランドとの比較");
    const top = centerY(2.8);

    const headers = ["スペック項目", "本企画", "CELINE", "LOEWE"];
    const dataRows = [
      ["素材", "ベジタブルタンニン", "カーフスキン", "クラシックカーフ"],
      ["重量", "約350g", "約450g", "約500g"],
      ["仕様", "2WAY", "2WAY", "5WAY"],
      ["サイズ", "W22xH15xD8cm", "W20xH14xD7cm", "W18xH12xD8cm"],
      ["販売価格", "5万 ~ 8万", "30万 ~ 40万", "20万 ~ 30万"],
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

    addKeyMsg(s, "ハイブランド同等の品質を1/4以下の価格で実現");
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

    s.addText("ハイブランドOEM工場との直接取引", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addDefBlock(s, 0.5, top + 0.6, "製造工房",
      "トスカーナ州フィレンツェ近郊の革工房\nCHANEL、GUCCI等メゾンブランドのOEM実績あり", 5.5);

    addDefBlock(s, 0.5, top + 1.45, "皮革タンナー",
      "ベジタブルタンニンなめし専門のConceria\n環境配慮型の伝統製法を継続", 5.5);

    // 右: 工房写真
    if (imgs.leatherCutting) {
      const fit = fitImage(dims.leatherCutting.w, dims.leatherCutting.h, 3.2, 2.0);
      s.addImage({ data: imgs.leatherCutting, x: 6.3, y: top + 0.5, w: fit.w, h: fit.h });
    }

    addSep(s, 0.5, top + 2.25, 9);

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top + 2.4, w: 9, h: 0.5, fill: { color: C.kmBg }
    });
    s.addText("サンプル製造時に工房見学 / 撮影が可能です", {
      x: 0.7, y: top + 2.4, w: 8.6, h: 0.5,
      fontSize: 15, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "メゾンブランドOEM工房で製造。現地撮影も可能");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 8: デザインイメージ
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "デザインイメージ");
    const photoTop = BODY_TOP + 0.1;

    s.addText("CELINEをベンチマーク", {
      x: 0.5, y: photoTop, w: 4.0, h: 0.3,
      fontSize: 16, fontFace: FONT, color: C.title, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真
    if (imgs.miniBag) {
      const fit = fitImage(dims.miniBag.w, dims.miniBag.h, 3.5, 3.2);
      s.addImage({ data: imgs.miniBag, x: 0.5, y: photoTop + 0.4, w: fit.w, h: fit.h });
    }

    // 右: 仕様箇条書き
    const specX = 4.5, specW = 5.0;
    let sy = photoTop + 0.4;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("ショルダーストラップ着脱可能で2WAY仕様", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.35;

    const subItems = ["1. ショルダーバッグとして", "2. ハンドバッグとして"];
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
    s.addText("内部にファスナーポケット1+オープンポケット2", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.3;
    s.addText("（スマホ・リップ・鍵など小物が迷子にならない設計）", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 11, fontFace: FONT, color: C.sub,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.4;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("開口部はマグネットボタンで片手開閉", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.35;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("底鋲付きで自立する安定構造", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    s.addText("※デザイン変更可能", {
      x: 7.5, y: BODY_BOT - 0.3, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "2WAYストラップ+3ポケットで実用性を両立");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 9: カラー展開
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "カラー展開");
    const top = centerY(3.0);

    // 左: カラーバリエーション写真
    if (imgs.bagCloseup) {
      const fit = fitImage(dims.bagCloseup.w, dims.bagCloseup.h, 4.5, 2.9);
      s.addImage({ data: imgs.bagCloseup, x: 0.5, y: top + 0.1, w: fit.w, h: fit.h });
    }

    const infoX = 5.3;
    s.addShape(pres.shapes.OVAL, { x: infoX, y: top + 0.15, w: 0.13, h: 0.13, fill: { color: C.primary } });
    s.addText("3色展開想定", {
      x: infoX + 0.25, y: top + 0.08, w: 4, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0
    });

    addSep(s, infoX, top + 0.55, 4.2);

    const colors = [
      { label: "キャメル", detail: "定番カラー。エイジングが最も映える" },
      { label: "ブラック", detail: "どんなコーデにも合う万能カラー" },
      { label: "グレージュ", detail: "トレンド感のあるニュアンスカラー" },
    ];
    colors.forEach((c, i) => {
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

    addKeyMsg(s, "キャメル x ブラック x グレージュの3色展開");
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

    s.addText("ダウンJKとのクロスセルで売上最大化", {
      x: 0.5, y: top, w: 6, h: 0.4,
      fontSize: 18, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const strategies = [
      { title: "ダウンJK+バッグのセットコーデ撮影", desc: "第1弾商品との世界観を繋ぐプロモーション撮影" },
      { title: "ポップアップイベントでの先行販売", desc: "実物の革の質感に触れる機会を提供し購買を促進" },
      { title: "購入特典: レザーケアキット", desc: "オリジナルのレザーケアセット(ミニクリーム+クロス)" },
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
    if (imgs.coverRight) {
      const fit = fitImage(dims.coverRight.w, dims.coverRight.h, imgW, 1.5);
      s.addImage({ data: imgs.coverRight, x: imgX + (imgW - fit.w) / 2, y: top + 0.1, w: fit.w, h: fit.h });
    }
    if (imgs.giftPkg) {
      const fit = fitImage(dims.giftPkg.w, dims.giftPkg.h, imgW, 1.3);
      s.addImage({ data: imgs.giftPkg, x: imgX + (imgW - fit.w) / 2, y: top + 1.7, w: fit.w, h: fit.h });
    }

    addKeyMsg(s, "ダウンJKとのセットコーデ+先行販売で話題化");
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
      ["4月", "企画・デザイン確定", "素材・金具・カラーの最終決定"],
      ["5月", "サンプル製作", "イタリア工房にてサンプル制作"],
      ["6月", "サンプルアップ・確認", "素材感・使用感の確認と修正指示"],
      ["7月", "量産発注", "2026AW向けの量産スタート"],
      ["9月", "製品納品", "検品後、倉庫に納品"],
      ["10月", "プロモーション撮影", "ダウンJKとのセットコーデ撮影"],
      ["11月", "リリース", "EC+ポップアップで販売開始"],
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

    s.addText("※ダウンJKと同時リリースのためタイトなスケジュールです。", {
      x: 0.5, y: top + hdrH + dataRows.length * rowH + 0.08, w: 9, h: 0.3,
      fontSize: 11, fontFace: FONT, color: C.primary,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "4月企画 → 11月リリースの製造スケジュール");
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
      ["販売価格", "5-8万円", "デザイン・仕様によって変動あり"],
      ["販売時期", "11月中旬ごろ", "ダウンJKと同時リリース想定"],
      ["素材", "イタリアンVTレザー", "トスカーナ産ベジタブルタンニン"],
      ["デザイン性", "2WAY+3ポケット", "実用性とデザインの両立"],
      ["販売戦略", "セットコーデ訴求", "ダウンJKとのクロスセル"],
      ["カラー展開", "3色", "キャメル/ブラック/グレージュ"],
      ["生産数", "300-800個", "価格とプロモーションによって変動"],
      ["売上想定", "1,500万-6,400万円", "販売価格5-8万円想定で算出"],
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

    addKeyMsg(s, "売上想定1,500万〜6,400万円のポテンシャル");
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

    s.addText("高品質レザーミニバッグ\nご提案資料", {
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
  const outPath = "/tmp/leather-bag-proposal-v1.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides)`);
}

main().catch(e => { console.error(e); process.exit(1); });
