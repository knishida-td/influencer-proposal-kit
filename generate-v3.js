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
  bg: "111111", title: "FFFFFF", body: "E0E0E0", sub: "AAAAAA", muted: "666666",
  primary: "C9A96E", secondary: "E8D5B0", kmBg: "1F1A14", sep: "333333",
  divider: "333333", white: "FFFFFF",
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
  pres.title = "軽量撥水マルチトートバッグ ご提案資料";

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
  const IMG_DIR = "/tmp/tote-proposal-images";
  const imgFiles = {
    personTote:    "person-tote.jpg",        // 白トートバッグ+ピンクBG (800x1200)
    womanShopping: "woman-shopping2.jpg",    // 女性+紙袋カジュアル (800x1200)
    toteLifestyle: "tote-lifestyle.jpg",     // 女性がトートを選ぶ (800x533)
    blackTextile:  "black-textile.jpg",      // 黒ファブリック (800x533)
    fabricColor:   "fabric-handmade.jpg",    // カラフル織物 (800x1067)
    womanBagHold:  "woman-bag-hold.jpg",     // 女性+赤バッグ (800x533)
    zipperDetail:  "zipper-detail.jpg",      // ジッパー (800x534)
    giftPkg:       "gift-wrap.jpg",          // ギフト包装 (800x1208)
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

  // CORDURA認証マーク (SVG生成)
  const corduraSvg = `<svg width="300" height="300" xmlns="http://www.w3.org/2000/svg">
    <rect width="300" height="300" rx="20" fill="#2C3E50"/>
    <text x="150" y="120" text-anchor="middle" fill="#ECF0F1" font-family="Helvetica,Arial" font-size="32" font-weight="bold">CORDURA</text>
    <line x1="60" y1="140" x2="240" y2="140" stroke="#3498DB" stroke-width="3"/>
    <text x="150" y="175" text-anchor="middle" fill="#3498DB" font-family="Helvetica,Arial" font-size="18">FABRIC</text>
    <text x="150" y="210" text-anchor="middle" fill="#BDC3C7" font-family="Helvetica,Arial" font-size="13">High Performance</text>
    <text x="150" y="235" text-anchor="middle" fill="#BDC3C7" font-family="Helvetica,Arial" font-size="13">Nylon Material</text>
    <text x="150" y="275" text-anchor="middle" fill="#95A5A6" font-family="Helvetica,Arial" font-size="11">CERTIFIED DURABILITY</text>
  </svg>`;
  const corduraImg = "image/png;base64," + (await sharp(Buffer.from(corduraSvg)).png().toBuffer()).toString("base64");

  // YKK品質認証 (SVG生成)
  const ykkSvg = `<svg width="280" height="280" xmlns="http://www.w3.org/2000/svg">
    <rect width="280" height="280" rx="140" fill="#E8E8E8"/>
    <text x="140" y="110" text-anchor="middle" fill="#333333" font-family="Helvetica,Arial" font-size="48" font-weight="bold">YKK</text>
    <line x1="70" y1="125" x2="210" y2="125" stroke="#666666" stroke-width="2"/>
    <text x="140" y="160" text-anchor="middle" fill="#555555" font-family="Helvetica,Arial" font-size="16">QUALITY</text>
    <text x="140" y="185" text-anchor="middle" fill="#555555" font-family="Helvetica,Arial" font-size="16">ZIPPER</text>
    <text x="140" y="220" text-anchor="middle" fill="#888888" font-family="Helvetica,Arial" font-size="11">JAPAN STANDARD</text>
  </svg>`;
  const ykkImg = "image/png;base64," + (await sharp(Buffer.from(ykkSvg)).png().toBuffer()).toString("base64");

  let pg = 0;

  // ═══════════════════════════════════════════════════════════════
  // ここから成瀬愛里さん向け「軽量撥水マルチトートバッグ」全スライド
  // ═══════════════════════════════════════════════════════════════

  // ═══════════════════════════════════════════════
  // Slide 1: 表紙
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 2.8) / 2;

    if (imgs.personTote) {
      const fit = fitImage(dims.personTote.w, dims.personTote.h, 2.8, 4.8);
      s.addImage({ data: imgs.personTote, x: 0.2, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }
    if (imgs.womanShopping) {
      const fit = fitImage(dims.womanShopping.w, dims.womanShopping.h, 2.8, 4.8);
      s.addImage({ data: imgs.womanShopping, x: SW - fit.w - 0.2, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }

    s.addText("軽量撥水\nマルチトートバッグ", {
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

    s.addText("「プチプラ高見え」を毎日のバッグでも", {
      x: 0.5, y: top, w: 9, h: 0.5,
      fontSize: 24, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addText(
      "UNIQLO/GU/ZARAコーデに合わせても高見えする、ママの毎日を支えるトートバッグ。",
      { x: 0.5, y: top + 0.65, w: 9, h: 0.4, fontSize: 16, fontFace: FONT, color: C.body,
        align: "left", valign: "middle", margin: 0 }
    );
    addSep(s, 0.5, top + 1.2, 9);

    s.addText("成瀬愛里さんとの相性:", {
      x: 0.5, y: top + 1.35, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: C.body,
      align: "left", valign: "middle", margin: 0
    });

    const points = [
      "32万フォロワーのプチプラコーデ発信力を活用",
      "UNIQLO/GU/ZARAとの着回し投稿との親和性",
      "3児のママが実感する「軽い・丈夫・洗える」訴求",
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

    addKeyMsg(s, "プチプラ高見え x ママの実用性 x 軽量撥水トート");
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

    s.addText("「高見え」と「ママの味方」を両立するトートバッグ", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真
    if (imgs.personTote) {
      const fit = fitImage(dims.personTote.w, dims.personTote.h, 2.0, 2.8);
      s.addImage({ data: imgs.personTote, x: 0.5, y: top + 0.55, w: fit.w, h: fit.h });
    }
    if (imgs.toteLifestyle) {
      const fit = fitImage(dims.toteLifestyle.w, dims.toteLifestyle.h, 2.5, 2.5);
      s.addImage({ data: imgs.toteLifestyle, x: 2.3, y: top + 0.7, w: fit.w, h: fit.h });
    }

    // 右: 箇条書き
    const features = [
      "コーデュラナイロン採用で軽くて丈夫",
      "撥水加工で急な雨や子供のこぼしにも安心",
      "A4サイズ対応で保育園書類もすっぽり",
      "約380gで一日中持っても疲れない超軽量",
      "内部ポケット7つで荷物が迷子にならない",
      "自立構造で置いた時もくたっとしない",
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

    addKeyMsg(s, "コーデュラナイロンの撥水トート、約380g");
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

    s.addText("プチプラ価格でも妥協しない素材と品質", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const cards = [
      { title: "コーデュラナイロン", sub: "軍用規格の耐久素材", key: "blackTextile",
        desc: "通常のナイロンの7倍の強度\n摩擦・引き裂きに強い\n子供が引っ張っても安心" },
      { title: "YKKファスナー", sub: "日本品質の信頼性", key: "zipperDetail",
        desc: "世界シェアNo.1のYKK製\nスムーズな開閉で片手操作OK\n壊れにくさが段違い" },
      { title: "撥水加工", sub: "急な雨でも安心", key: "fabricColor",
        desc: "表面に撥水コーティング処理\n水滴がコロコロ転がる\n汚れも付きにくく手入れ簡単" },
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

    addKeyMsg(s, "コーデュラナイロン+YKK+撥水の三拍子");
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

    addDefBlock(s, 0.5, top + 0.6, "CORDURA認定ファブリック",
      "INVISTA社のコーデュラブランド認定素材\n軍用・アウトドア用品にも採用される高耐久ナイロン", 6.0);
    s.addImage({ data: corduraImg, x: 7.0, y: top + 0.4, w: 1.5, h: 1.5 });

    addSep(s, 0.5, top + 1.5, 9);

    addDefBlock(s, 0.5, top + 1.7, "YKKファスナー: 日本品質の証明",
      "世界シェア45%のYKK製ファスナーを全箇所に採用\nスムーズな開閉と耐久性を両立", 6.0);
    s.addImage({ data: ykkImg, x: 7.0, y: top + 1.5, w: 1.5, h: 1.5 });

    addKeyMsg(s, "CORDURA+YKKで品質と耐久性を担保");
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

    const headers = ["スペック", "本企画", "Longchamp", "L.L.Bean"];
    const dataRows = [
      ["素材", "コーデュラナイロン", "ナイロンツイル", "ナイロン"],
      ["重量", "約380g", "約300g", "約450g"],
      ["撥水", "撥水加工あり", "撥水加工あり", "なし"],
      ["ポケット数", "内外7つ", "内1つ", "内外4つ"],
      ["販売価格", "6千 ~ 1.2万", "1.5万 ~ 3万", "5千 ~ 1万"],
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
        s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: "1A1A1A" } });
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

    addKeyMsg(s, "Longchamp品質をプチプラ価格帯で実現");
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

    s.addText("国内検品体制で品質を担保", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addDefBlock(s, 0.5, top + 0.6, "製造工場",
      "中国広州の縫製工場(有名SPAブランドOEM実績あり)\n国内大手アパレルの品質基準をクリア", 5.5);

    addDefBlock(s, 0.5, top + 1.45, "国内検品",
      "日本国内の検品センターで全数検査\n縫製不良・汚れ・金具の動作確認を実施", 5.5);

    // 右: ライフスタイル写真
    if (imgs.toteLifestyle) {
      const fit = fitImage(dims.toteLifestyle.w, dims.toteLifestyle.h, 3.2, 2.0);
      s.addImage({ data: imgs.toteLifestyle, x: 6.3, y: top + 0.5, w: fit.w, h: fit.h });
    }

    addSep(s, 0.5, top + 2.25, 9);

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top + 2.4, w: 9, h: 0.5, fill: { color: C.kmBg }
    });
    s.addText("サンプル到着後、実物での撮影が可能です", {
      x: 0.7, y: top + 2.4, w: 8.6, h: 0.5,
      fontSize: 15, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "OEM工場製造+国内全数検品で品質を担保");
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

    s.addText("Longchampをベンチマーク", {
      x: 0.5, y: photoTop, w: 4.0, h: 0.3,
      fontSize: 16, fontFace: FONT, color: C.title, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真
    if (imgs.personTote) {
      const fit = fitImage(dims.personTote.w, dims.personTote.h, 3.5, 3.2);
      s.addImage({ data: imgs.personTote, x: 0.5, y: photoTop + 0.4, w: fit.w, h: fit.h });
    }

    // 右: 仕様箇条書き
    const specX = 4.5, specW = 5.0;
    let sy = photoTop + 0.4;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("トート+ショルダーの2WAY仕様", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.35;

    const subItems = ["1. トートバッグとして(通勤・買い物)", "2. ショルダーバッグとして(子連れ外出)"];
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
    s.addText("内外7ポケットで荷物整理がラク", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.3;
    s.addText("（スマホ・鍵・水筒・おむつポーチなど分類収納）", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 11, fontFace: FONT, color: C.sub,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.4;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("天ファスナー+マグネットで中身が見えない安心設計", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    sy += 0.35;

    s.addShape(pres.shapes.OVAL, { x: specX, y: sy + 0.06, w: 0.14, h: 0.14, fill: { color: C.primary } });
    s.addText("底板+底鋲付きで自立する安定構造", {
      x: specX + 0.25, y: sy, w: specW - 0.25, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    s.addText("※デザイン変更可能", {
      x: 7.5, y: BODY_BOT - 0.3, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "2WAY+7ポケット+自立構造でママの味方");
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
    if (imgs.personTote) {
      const fit = fitImage(dims.personTote.w, dims.personTote.h, 4.5, 2.9);
      s.addImage({ data: imgs.personTote, x: 0.5, y: top + 0.1, w: fit.w, h: fit.h });
    }

    const infoX = 5.3;
    s.addShape(pres.shapes.OVAL, { x: infoX, y: top + 0.15, w: 0.13, h: 0.13, fill: { color: C.primary } });
    s.addText("3色展開想定", {
      x: infoX + 0.25, y: top + 0.08, w: 4, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0
    });

    addSep(s, infoX, top + 0.55, 4.2);

    const colors = [
      { label: "ブラック", detail: "どんなプチプラコーデにも合う万能色" },
      { label: "グレージュ", detail: "きれいめコーデに映えるニュアンス色" },
      { label: "ネイビー", detail: "通勤にも使えるきちんと感のある色" },
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

    addKeyMsg(s, "ブラック x グレージュ x ネイビーの3色展開");
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

    s.addText("プチプラコーデ投稿との連動で認知拡大", {
      x: 0.5, y: top, w: 6, h: 0.4,
      fontSize: 18, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const strategies = [
      { title: "着回し投稿でのバッグ露出", desc: "普段のプチプラコーデ投稿にバッグを自然に組込み" },
      { title: "Instagram Reelsで「中身紹介」", desc: "7ポケットの使い方を動画で見せて実用性を訴求" },
      { title: "購入特典: ミニポーチ付き", desc: "おむつ替えシートも入るミニポーチをセットで" },
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
    if (imgs.womanBagHold) {
      const fit = fitImage(dims.womanBagHold.w, dims.womanBagHold.h, imgW, 1.5);
      s.addImage({ data: imgs.womanBagHold, x: imgX + (imgW - fit.w) / 2, y: top + 0.1, w: fit.w, h: fit.h });
    }
    if (imgs.giftPkg) {
      const fit = fitImage(dims.giftPkg.w, dims.giftPkg.h, imgW, 1.3);
      s.addImage({ data: imgs.giftPkg, x: imgX + (imgW - fit.w) / 2, y: top + 1.7, w: fit.w, h: fit.h });
    }

    addKeyMsg(s, "着回し投稿+Reels+ミニポーチ特典で話題化");
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
      ["4月", "企画・デザイン確定", "素材・カラー・ポケット仕様の確定"],
      ["5月", "サンプル製作", "中国工場にてサンプル制作"],
      ["5月下旬", "サンプル確認", "実物確認・修正指示・撮影テスト"],
      ["6月", "量産発注", "2026AW向けの量産スタート"],
      ["8月", "製品納品・検品", "国内検品センターで全数検査"],
      ["9月", "プロモーション撮影", "成瀬さんのコーデ撮影"],
      ["10月", "リリース", "EC販売+Instagram連動開始"],
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
        s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: "1A1A1A" } });
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

    s.addText("※秋冬シーズンに合わせたスケジュールです。", {
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
      ["販売価格", "6,000-12,000円", "プチプラ層が手を出せる価格帯"],
      ["販売時期", "10月", "秋冬コーデシーズンに合わせて"],
      ["素材", "コーデュラナイロン", "撥水加工+YKKファスナー"],
      ["デザイン性", "2WAY+7ポケット", "ママの実用性を最優先"],
      ["販売戦略", "着回し投稿連動", "成瀬さんのコーデ投稿にバッグ露出"],
      ["カラー展開", "3色", "ブラック/グレージュ/ネイビー"],
      ["生産数", "1,000-3,000個", "価格とプロモーションによって変動"],
      ["売上想定", "600万-3,600万円", "販売価格6千-1.2万円想定で算出"],
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
        s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: ry, w: 9.0, h: rowH, fill: { color: "1A1A1A" } });
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

    addKeyMsg(s, "売上想定600万〜3,600万円のポテンシャル");
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

    s.addText("軽量撥水マルチトートバッグ\nご提案資料", {
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
  const outPath = "/tmp/tote-bag-proposal-v1.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides)`);
}

main().catch(e => { console.error(e); process.exit(1); });
