const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

// =============================================================
// Image Utilities
// =============================================================

// アスペクト比を維持してmaxW/maxHに収める
function fitImage(origW, origH, maxW, maxH) {
  const ratio = origW / origH;
  let w = maxW, h = maxW / ratio;
  if (h > maxH) { h = maxH; w = maxH * ratio; }
  return { w, h };
}

// ファイルからBase64読み込み + sipsでピクセル取得
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
  return { w: 600, h: 824 }; // fallback
}

// sharpでラベル付きプレースホルダーを生成
async function makePlaceholder(color, label, w, h) {
  const pw = Math.round(w * 150), ph = Math.round(h * 150);
  const svg = `<svg width="${pw}" height="${ph}" xmlns="http://www.w3.org/2000/svg">
    <rect width="100%" height="100%" fill="#${color}"/>
    <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle"
          fill="#ffffff" font-family="sans-serif" font-size="${Math.round(ph * 0.07)}px">${label}</text>
  </svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
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
  pres.title = "高品質ダウンジャケット ご提案資料";

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

  // アスペクト比を維持して配置するヘルパー
  function addImgFit(s, imgData, dims, x, y, maxW, maxH) {
    const fit = fitImage(dims.w, dims.h, maxW, maxH);
    const cx = x + (maxW - fit.w) / 2; // 水平中央
    const cy = y + (maxH - fit.h) / 2; // 垂直中央
    s.addImage({ data: imgData, x: cx, y: cy, w: fit.w, h: fit.h });
  }

  // ── Load images ──
  const IMG_DIR = "/tmp/proposal-images";
  const imgFiles = {
    coverA:   "cover-black.jpg",    // 白ジャケット・カフェ (600x824)
    coverB:   "cover-white.jpg",    // 白ジャケット・フード (600x824)
    black:    "product-white.jpg",  // 黒ジャケット着用 (600x824)
    white:    "product-black.jpg",  // 白ジャケット・買い物 (600x824)
    blackFull:"detail1.jpg",        // 黒ジャケット全身 (600x824)
    beltDetail:"fur.jpg",           // ベルトディテール (1000x984)
    detail2:  "detail2.jpg",        // ディテール (600x824)
    modelPromo:"model-promo.jpg",   // モデルプロモ (1364x1156)
    infoPage: "info-detail.jpg",    // 商品詳細ページ (1000x1750)
    sheepWool:"sheep-wool.jpg",     // 羊毛 (600x906)
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

  // 素材プレースホルダー（ダウンロード不可だった画像）
  const phDown = await makePlaceholder("D4C8B0", "ダウン素材", 3, 2);
  const phDownBall = await makePlaceholder("E8DCC0", "ダウンボール", 3, 2);
  const phFur = await makePlaceholder("C0B4A0", "シープファー", 3, 2);
  const phRDS = await makePlaceholder("2196F3", "RDS認証", 2, 2);
  const phTC = await makePlaceholder("607D8B", "TC証明書", 2, 3);
  const phNovelty = await makePlaceholder("D4C8B0", "ノベルティ", 2, 2);

  let pg = 0;

  // ═══════════════════════════════════════════════
  // Slide 1: 表紙 — 商品写真2枚 + 中央タイトル
  // 元PDF: 黒JK左 + 白JK右 + 大タイトル中央
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 2.8) / 2;

    // 左写真（アスペクト比維持）
    if (imgs.black) {
      const fit = fitImage(dims.black.w, dims.black.h, 2.8, 4.5);
      s.addImage({ data: imgs.black, x: 0.1, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }
    // 右写真
    if (imgs.coverB) {
      const fit = fitImage(dims.coverB.w, dims.coverB.h, 2.8, 4.5);
      s.addImage({ data: imgs.coverB, x: SW - fit.w - 0.1, y: (SH - fit.h) / 2, w: fit.w, h: fit.h });
    }

    // 中央タイトル
    s.addText("高品質\nダウンジャケット", {
      x: 2.8, y: bTop, w: 4.4, h: 1.4,
      fontSize: 34, fontFace: FONT, color: C.title, bold: true,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 3.5, y: bTop + 1.5, w: 1.25, h: 0.035, fill: { color: C.primary }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 4.75, y: bTop + 1.5, w: 1.75, h: 0.035, fill: { color: C.secondary }
    });
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
  // 元PDF: 「一生モノ」+ 3つの共通ポイント
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "コンセプト");

    const contentH = 3.1;
    const top = centerY(contentH);

    s.addText("「一生モノ」の価値を届ける", {
      x: 0.5, y: top, w: 9, h: 0.5,
      fontSize: 24, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });
    s.addText(
      "トレンドを追いすぎず、定番で誰もが着れて飽きがこない「冬の主役」アイテム。",
      { x: 0.5, y: top + 0.65, w: 9, h: 0.4, fontSize: 16, fontFace: FONT, color: C.body,
        align: "left", valign: "middle", margin: 0 }
    );

    addSep(s, 0.5, top + 1.2, 9);

    s.addText("スニーカーを含む靴のアイテムとも共通している:", {
      x: 0.5, y: top + 1.35, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: C.body,
      align: "left", valign: "middle", margin: 0
    });

    const points = [
      "定番で取り入れやすいシンプルなデザイン",
      "高品質かつ手に取りやすい価格帯",
      "スタイルアップが叶う",
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

    addKeyMsg(s, "定番 x 高品質 x スタイルアップ の三拍子");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 3: 商品概要 — 写真左2枚 + 箇条書き右
  // 元PDF: 黒JK + 白JK の全身写真左、機能箇条書き右
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "商品概要");

    const top = centerY(3.3);

    s.addText("「盛れる」と「高品質」を両立する設計", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 2枚の着用写真（アスペクト比維持）
    const photoMaxW = 2.0, photoMaxH = 2.7;
    if (imgs.blackFull) {
      const fit = fitImage(dims.blackFull.w, dims.blackFull.h, photoMaxW, photoMaxH);
      s.addImage({ data: imgs.blackFull, x: 0.5, y: top + 0.55, w: fit.w, h: fit.h });
    }
    if (imgs.white) {
      const fit = fitImage(dims.white.w, dims.white.h, photoMaxW, photoMaxH);
      s.addImage({ data: imgs.white, x: 0.5 + photoMaxW + 0.15, y: top + 0.55, w: fit.w, h: fit.h });
    }

    // 右: 機能箇条書き
    const features = [
      "軽くて暖かい",
      "ウエストマークで細見え、スタイルアップ",
      "ボリュームファーで小顔効果",
      "ショート/ロングどちらにも合う丈感",
      "軽くて疲れない着心地",
      "トレンドに左右されず一生着られる安心感",
    ];
    features.forEach((f, i) => {
      const fy = top + 0.65 + i * 0.43;
      s.addShape(pres.shapes.OVAL, { x: 4.9, y: fy + 0.05, w: 0.15, h: 0.15, fill: { color: C.primary } });
      s.addText(f, {
        x: 5.2, y: fy, w: 4.3, h: 0.35,
        fontSize: 14, fontFace: FONT, color: C.body, bold: true,
        align: "left", valign: "middle", margin: 0
      });
    });

    addKeyMsg(s, "見た目の華やかさと実用性を兼ね備えた一着");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 4: 機能性 — 3カラム素材カード
  // 元PDF: FP800 / ダウン90% / ファー の3カード + 素材写真
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "機能性");

    const top = centerY(3.2);

    s.addText("ハイブランドと同等のスペックを実現", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const cards = [
      { title: "フィルパワー (FP) 800", sub: "圧倒的な軽さ", img: phDown,
        desc: "スペイン産ホワイトダック使用予定\n驚きの軽さを実現" },
      { title: "ダウンとフェザーの黄金割合", sub: "90% / 10%", img: phDownBall,
        desc: "ダウン=柔らかくて軽い\nフェザー=弾力があり形を整える" },
      { title: "華やかなファーで", sub: "デザイン性をプラス", img: imgs.sheepWool ? imgs.sheepWool : phFur,
        desc: "シープ(羊)ファーを使用予定\n小顔効果や映えにも有効" },
    ];

    const cardW = 2.7, cardH = 2.5, cardGap = 0.35;
    const startX = (SW - (cardW * 3 + cardGap * 2)) / 2;

    cards.forEach((c, i) => {
      const cx = startX + i * (cardW + cardGap);
      const cy = top + 0.6;

      s.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: cardW, h: cardH, fill: { color: C.white } });

      s.addText(c.title, {
        x: cx + 0.1, y: cy + 0.1, w: cardW - 0.2, h: 0.25,
        fontSize: 12, fontFace: FONT, color: C.title, bold: true,
        align: "center", valign: "middle", margin: 0
      });
      s.addText(c.sub, {
        x: cx + 0.1, y: cy + 0.35, w: cardW - 0.2, h: 0.2,
        fontSize: 11, fontFace: FONT, color: C.sub,
        align: "center", valign: "middle", margin: 0
      });

      // 素材写真（sheep-woolは実画像、他はプレースホルダー）
      if (i === 2 && imgs.sheepWool) {
        const fit = fitImage(dims.sheepWool.w, dims.sheepWool.h, cardW - 0.4, 0.9);
        const ix = cx + (cardW - fit.w) / 2;
        s.addImage({ data: imgs.sheepWool, x: ix, y: cy + 0.65, w: fit.w, h: fit.h });
      } else {
        s.addImage({ data: c.img, x: cx + 0.35, y: cy + 0.65, w: 2.0, h: 0.9 });
      }

      s.addText(c.desc, {
        x: cx + 0.1, y: cy + 1.65, w: cardW - 0.2, h: 0.75,
        fontSize: 10, fontFace: FONT, color: C.body,
        align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3
      });
    });

    s.addText("※仕様変更可能", {
      x: 7.5, y: top + 3.15, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "FP800・ダウン90%・シープファーのハイスペック");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 5: 品質証明 — RDS + TC
  // 元PDF: RDS認証マーク右 + TC証明書イメージ右
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

    // RDS
    addDefBlock(s, 0.5, top + 0.6, "RDS認証: 取得予定（動物福祉の国際認証）",
      "強制的な飼育や生きたまま羽毛を刈り取っていない\n農場から仕入れていることを証明する国際認証マーク", 6.0);
    s.addImage({ data: phRDS, x: 7.3, y: top + 0.5, w: 1.5, h: 1.0 });

    addSep(s, 0.5, top + 1.5, 9);

    // TC
    addDefBlock(s, 0.5, top + 1.7, "TC（Transaction Certificate）: 発行可能",
      "素材のパスポートのような証明書。偽物ではなく\nクリーンなルートで届いた本物である証明", 6.0);
    s.addImage({ data: phTC, x: 7.3, y: top + 1.6, w: 1.0, h: 1.2 });

    addKeyMsg(s, "国際認証で品質と倫理性を担保");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 6: 他ハイブランドとの比較
  // 元PDF: 4列比較テーブル
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "他ハイブランドとの比較");

    const top = centerY(2.8);

    const headers = ["スペック項目", "本企画", "MONCLER", "PRADA等"];
    const dataRows = [
      ["フィルパワー(FP)", "800FP", "710~800", "650~750(推定)"],
      ["ダウン混合率", "90% / 10%", "90% / 10%", "90% / 10%"],
      ["シルエット設計", "美くびれ・スタイルUP", "曲線美・ラグジュアリー", "ミニマル・モード"],
      ["ファーの品質", "シープ", "シープ / フォックス", "なし / 合成 / シープ"],
      ["販売価格帯", "15万 ~ 30万", "25万 ~ 50万+", "30万 ~ 60万+"],
    ];

    const colX = [0.5, 2.5, 5.0, 7.25];
    const colW = [2.0, 2.5, 2.25, 2.25];
    const hdrH = 0.42, rowH = 0.4;

    // ヘッダー行
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
    headers.forEach((h, i) => {
      s.addText(h, {
        x: colX[i], y: top, w: colW[i], h: hdrH,
        fontSize: 12, fontFace: FONT, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    });

    // データ行
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

    addKeyMsg(s, "ハイブランド同等スペックを半額以下で実現");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 7: 生産背景
  // 元PDF: 工場実績 + コールアウトボックス（写真なし）
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "生産背景");

    const top = centerY(2.9);

    s.addText("ハイブランド実績のある工場", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addDefBlock(s, 0.5, top + 0.6, "主要実績", "Max Mara, Acne Studios, THE NORTH FACE", 9);
    addDefBlock(s, 0.5, top + 1.45, "過去実績", "Canada Goose", 9);

    addSep(s, 0.5, top + 2.25, 9);

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top + 2.4, w: 9, h: 0.5, fill: { color: C.kmBg }
    });
    s.addText("工場は中国にあり、製造時に工場見学 / 撮影が可能です", {
      x: 0.7, y: top + 2.4, w: 8.6, h: 0.5,
      fontSize: 15, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "ハイブランドOEM工場で製造。現地撮影も可能");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 8: デザインイメージ — 商品写真左 + スペック右
  // 元PDF: MONCLER風JK写真左 + 3WAY/2WAY箇条書き右
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "デザインイメージ");

    const top = centerY(2.6);

    s.addText("MONCLERをベンチマーク", {
      x: 0.5, y: top, w: 9, h: 0.35,
      fontSize: 16, fontFace: FONT, color: C.title, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真（アスペクト比維持）
    if (imgs.black) {
      const fit = fitImage(dims.black.w, dims.black.h, 3.5, 2.6);
      s.addImage({ data: imgs.black, x: 0.5, y: top + 0.5, w: fit.w, h: fit.h });
    }

    // 右: スペック
    const specs = [
      { bullet: true, text: "フードとファーは取り外し可能な3WAY構造" },
      { bullet: false, text: "  1. ファー+フード付き\n  2. ファーなしフード付き\n  3. ファーとフードなし" },
      { bullet: true, text: "ウエストベルトも取り外し可能な2WAY構造" },
      { bullet: true, text: "160cmの女性でお尻が隠れるくらいの着丈" },
    ];

    let sy = top + 0.5;
    specs.forEach((sp) => {
      const isSub = !sp.bullet;
      const h = isSub ? 0.6 : 0.3;
      if (sp.bullet) {
        s.addShape(pres.shapes.OVAL, { x: 4.5, y: sy + 0.07, w: 0.13, h: 0.13, fill: { color: C.primary } });
      }
      s.addText(sp.text, {
        x: isSub ? 4.8 : 4.75, y: sy, w: 4.75, h,
        fontSize: isSub ? 12 : 13, fontFace: FONT, color: C.body,
        bold: sp.bullet, align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3
      });
      sy += h + 0.08;
    });

    s.addText("※デザイン変更可能", {
      x: 7.5, y: top + 3.0, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted, align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "3WAYフード + 2WAYベルトで着回し自在");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 9: カラー・サイズ展開 — 着用写真2枚左 + 情報右
  // 元PDF: 2着用写真左 + カラー情報右 + 商品切り抜き右下
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "カラー・サイズ展開");

    const top = centerY(3.0);

    // 左: 着用写真2枚（アスペクト比維持）
    const pMaxW = 2.2, pMaxH = 2.8;
    if (imgs.coverA) {
      const fit = fitImage(dims.coverA.w, dims.coverA.h, pMaxW, pMaxH);
      s.addImage({ data: imgs.coverA, x: 0.5, y: top + 0.1, w: fit.w, h: fit.h });
    }
    if (imgs.coverB) {
      const fit = fitImage(dims.coverB.w, dims.coverB.h, pMaxW, pMaxH);
      s.addImage({ data: imgs.coverB, x: 0.5 + pMaxW + 0.15, y: top + 0.1, w: fit.w, h: fit.h });
    }

    // 右上: サイズ
    s.addShape(pres.shapes.OVAL, { x: 5.2, y: top + 0.15, w: 0.13, h: 0.13, fill: { color: C.primary } });
    s.addText("フリーサイズ想定", {
      x: 5.45, y: top + 0.08, w: 4, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0
    });
    s.addShape(pres.shapes.OVAL, { x: 5.2, y: top + 0.5, w: 0.13, h: 0.13, fill: { color: C.primary } });
    s.addText("ブラックとホワイトの2色展開想定", {
      x: 5.45, y: top + 0.43, w: 4, h: 0.3,
      fontSize: 13, fontFace: FONT, color: C.body, align: "left", valign: "middle", margin: 0
    });

    addSep(s, 5.2, top + 0.85, 4.3);

    // カラー詳細
    const colors = [
      { label: "ブラック本体", detail: "ファーはブラウンベージュ" },
      { label: "ホワイト本体", detail: "ファーはミルクベージュ" },
    ];
    colors.forEach((c, i) => {
      const cy = top + 1.05 + i * 0.55;
      s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: cy, w: 0.06, h: 0.42, fill: { color: C.primary } });
      s.addText(c.label, {
        x: 5.4, y: cy, w: 4, h: 0.22,
        fontSize: 13, fontFace: FONT, color: C.title, bold: true,
        align: "left", valign: "middle", margin: 0
      });
      s.addText(c.detail, {
        x: 5.4, y: cy + 0.22, w: 4, h: 0.2,
        fontSize: 12, fontFace: FONT, color: C.sub,
        align: "left", valign: "middle", margin: 0
      });
    });

    // 右下: ベルトディテール写真
    if (imgs.beltDetail) {
      const fit = fitImage(dims.beltDetail.w, dims.beltDetail.h, 2.5, 1.0);
      s.addImage({ data: imgs.beltDetail, x: 5.2 + (4.3 - fit.w) / 2, y: top + 2.2, w: fit.w, h: fit.h });
    }

    addKeyMsg(s, "ブラック x ホワイトの定番2色 / フリーサイズ");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 10: 販売戦略 — テキスト左 + 写真右
  // 元PDF: 3施策左 + モデル写真+ノベルティ写真右
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売戦略");

    const top = centerY(2.7);

    s.addText("タレント/モデル起用の撮影とノベルティ", {
      x: 0.5, y: top, w: 6, h: 0.4,
      fontSize: 18, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    const strategies = [
      { title: "モデル起用のプロモーション撮影", desc: "知名度のあるモデル・タレントを起用し洗練された撮影" },
      { title: "ポップアップや展示会の開催", desc: "高価格帯アパレルのため実際に触れる機会を提供" },
      { title: "販売記念のノベルティ", desc: "ファーと同素材のチャーム等、購入特典を検討中" },
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

    // 右: プロモ写真（アスペクト比維持）
    if (imgs.modelPromo) {
      const fit = fitImage(dims.modelPromo.w, dims.modelPromo.h, 2.8, 1.8);
      s.addImage({ data: imgs.modelPromo, x: 6.8, y: top + 0.1, w: fit.w, h: fit.h });
    }
    // ノベルティイメージ
    s.addImage({ data: phNovelty, x: 7.3, y: top + 2.0, w: 1.5, h: 0.9 });

    addKeyMsg(s, "撮影・展示会・ノベルティの三本柱で販売加速");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 11: スケジュール
  // 元PDF: タイムラインテーブル + 注記ボックス
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "スケジュール");

    const top = centerY(3.0);

    const headers = ["時期", "進行管理", "詳細"];
    const dataRows = [
      ["4月上旬", "1stサンプル発注", "ベンチマークサンプルの発注"],
      ["4月下旬", "1stサンプルアップ", "デザイン・仕様の確認、2ndサンプル進行"],
      ["5月中旬〜下旬", "2ndサンプルアップ", "デザイン・仕様の最終確認"],
      ["5月下旬", "量産発注", "2026AWに向けて発注"],
      ["10月下旬〜11月上旬", "製品納品", "製品が倉庫に納品"],
      ["11月中旬", "リリース", "製品本販売開始"],
    ];

    const colX = [0.5, 2.7, 5.1];
    const colW = [2.2, 2.4, 4.4];
    const hdrH = 0.38, rowH = 0.35;

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
          fontSize: 11, fontFace: FONT, color: C.body,
          bold: ci === 0, align: "center", valign: "middle", margin: 0
        });
      });
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });

    // 注記
    const noteY = top + hdrH + dataRows.length * rowH + 0.1;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: noteY, w: 9, h: 0.38, fill: { color: C.kmBg } });
    s.addText("※優先的に進行するためタイトなスケジュールです。", {
      x: 0.7, y: noteY, w: 8.6, h: 0.38,
      fontSize: 12, fontFace: FONT, color: C.primary,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "4月発注 → 11月リリースの製造スケジュール");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 12: 販売イメージ
  // 元PDF: 9行サマリーテーブル
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売イメージ");

    const top = centerY(3.4);

    const colX = [0.5, 2.5, 5.2];
    const colW = [2.0, 2.7, 4.3];
    const hdrH = 0.34, rowH = 0.33;

    const headers = ["項目", "内容", "備考"];
    const dataRows = [
      ["販売価格", "15~30万円", "デザイン・仕様によって変動あり"],
      ["販売時期", "11月中旬ごろ", "案件状況により調整"],
      ["機能性", "ハイブランド同等の高品質", "800FP、ダウン90%、実績ある工場"],
      ["デザイン性", "スタイルアップ&小顔効果", "くびれライン+ボリュームファー"],
      ["販売戦略", "モデルやタレント起用", "ノベルティ施策も検討中"],
      ["サイズ展開", "フリーサイズ", "サイズ展開可"],
      ["カラー展開", "ブラック、ホワイト 2色", "カラー展開可"],
      ["生産数", "300~1,000着", "価格とプロモーションで変動"],
      ["売上想定", "4,500万~1.5億円", "販売価格15万円想定で算出"],
    ];

    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: top, w: 9.0, h: hdrH, fill: { color: C.primary } });
    headers.forEach((h, i) => {
      s.addText(h, {
        x: colX[i], y: top, w: colW[i], h: hdrH,
        fontSize: 11, fontFace: FONT, color: C.white, bold: true,
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
          fontSize: ci === 0 ? 11 : 10, fontFace: FONT,
          color: ci === 0 ? C.title : C.body, bold: ci === 0,
          align: "center", valign: "middle", margin: 0
        });
      });
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });

    addKeyMsg(s, "売上想定4,500万〜1.5億円のポテンシャル");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 13: エンド
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 1.8) / 2;

    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: bTop, w: 4.25, h: 0.035, fill: { color: C.primary } });
    s.addShape(pres.shapes.RECTANGLE, { x: 4.75, y: bTop, w: 4.75, h: 0.035, fill: { color: C.secondary } });
    s.addText("高品質ダウンジャケット\nご提案資料", {
      x: 0.5, y: bTop + 0.2, w: 9, h: 0.8,
      fontSize: 28, fontFace: FONT, color: C.title, bold: true,
      align: "center", valign: "middle", margin: 0, lineSpacingMultiple: 1.15
    });
    s.addText("ご検討のほど、よろしくお願いいたします。", {
      x: 0.5, y: bTop + 1.1, w: 9, h: 0.35,
      fontSize: 14, fontFace: FONT, color: C.sub,
      align: "center", valign: "middle", margin: 0
    });
    s.addText("2026年4月", {
      x: 0.5, y: bTop + 1.55, w: 9, h: 0.3,
      fontSize: 12, fontFace: FONT, color: C.muted,
      align: "center", valign: "middle", margin: 0
    });
  }

  const outPath = "/tmp/influencer-proposal-v1.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides)`);
}

main().catch(e => { console.error(e); process.exit(1); });
