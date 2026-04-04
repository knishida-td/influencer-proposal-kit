const pptxgen = require("pptxgenjs");
const sharp = require("sharp");
const fs = require("fs");
const path = require("path");

// ── Image helper: placeholder if file missing ──
async function loadImg(filePath, fallbackColor, fallbackText, w, h) {
  if (filePath && fs.existsSync(filePath)) {
    const ext = path.extname(filePath).toLowerCase();
    const mime = ext === ".png" ? "image/png" : "image/jpeg";
    const buf = fs.readFileSync(filePath);
    return { data: `${mime};base64,${buf.toString("base64")}` };
  }
  const pw = Math.round(w * 150), ph = Math.round(h * 150);
  const svg = `<svg width="${pw}" height="${ph}" xmlns="http://www.w3.org/2000/svg">
    <rect width="100%" height="100%" fill="#${fallbackColor}"/>
    <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle"
          fill="#ffffff" font-family="sans-serif" font-size="${Math.round(ph * 0.06)}px">${fallbackText}</text>
  </svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return { data: "image/png;base64," + buf.toString("base64") };
}

// =============================================================
// SlideKit Design System (共通)
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

  // ── Image directory ──
  const IMG_DIR = "/tmp/proposal-images";

  // ── Load images (placeholder if missing) ──
  // 画像マッピング（DHOLICからDL済み）
  const imgCover1 = await loadImg(path.join(IMG_DIR, "cover-black.jpg"), "E8E0D8", "白ジャケット・カフェ", 4, 5);
  const imgCover2 = await loadImg(path.join(IMG_DIR, "cover-white.jpg"), "E8E0D8", "白ジャケット・フード", 4, 5);
  const imgBlack = await loadImg(path.join(IMG_DIR, "product-white.jpg"), "2A2A2A", "黒ジャケット着用", 3, 4);
  const imgWhite = await loadImg(path.join(IMG_DIR, "product-black.jpg"), "E8E0D8", "白ジャケット着用", 3, 4);
  const imgDetail = await loadImg(path.join(IMG_DIR, "fur.jpg"), "D4C5B8", "ベルトディテール", 3, 2);
  const imgBlackFull = await loadImg(path.join(IMG_DIR, "detail1.jpg"), "2A2A2A", "黒ジャケット全身", 3, 4);
  const imgDetail2 = await loadImg(path.join(IMG_DIR, "detail2.jpg"), "D4C5B8", "ディテール", 3, 3);
  const imgModel = await loadImg(path.join(IMG_DIR, "model-promo.jpg"), "E8E0D8", "モデル撮影イメージ", 3, 3);
  const imgInfo = await loadImg(path.join(IMG_DIR, "info-detail.jpg"), "D4C5B8", "商品詳細", 3, 4);
  // 素材写真はプレースホルダー（実物写真は差し替え前提）
  const imgFeather = await loadImg(null, "E8D8C0", "ダウン素材写真", 3, 2);
  const imgDownFill = await loadImg(null, "F0E8D8", "ダウンボール写真", 3, 2);
  const imgFurMat = await loadImg(null, "C8B8A0", "ファー素材写真", 3, 2);
  const imgRDS = await loadImg(null, "2196F3", "RDS認証マーク", 2, 2);

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
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w, h: 0.015, fill: { color: C.sep }
    });
  }

  function addDefBlock(s, x, y, heading, body, w) {
    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: 0.06, h: 0.7, fill: { color: C.primary }
    });
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

  let pg = 0;

  // ═══════════════════════════════════════════════
  // Slide 1: 表紙 (Type A)
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 2.8) / 2;

    // 左右に商品写真（中央にタイトル用の空きを確保）
    s.addImage({ ...imgCover1, x: 0.0, y: 0.15, w: 3.0, h: 4.2 });
    s.addImage({ ...imgCover2, x: 7.0, y: 0.15, w: 3.0, h: 4.2 });

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
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "コンセプト");

    const contentH = 3.2;
    const top = centerY(contentH);

    s.addText("「一生モノ」の価値を届ける", {
      x: 0.5, y: top, w: 9, h: 0.5,
      fontSize: 24, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    s.addText(
      "トレンドを追いすぎず、定番で誰もが着れて飽きがこない「冬の主役」アイテム。",
      {
        x: 0.5, y: top + 0.7, w: 9, h: 0.4,
        fontSize: 16, fontFace: FONT, color: C.body,
        align: "left", valign: "middle", margin: 0
      }
    );

    addSep(s, 0.5, top + 1.25, 9);

    s.addText(
      "スニーカーを含む靴のアイテムとも共通している、\nインフルエンサーの世界観にマッチするポイント:",
      {
        x: 0.5, y: top + 1.4, w: 9, h: 0.55,
        fontSize: 14, fontFace: FONT, color: C.body,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3
      }
    );

    const points = [
      "定番で取り入れやすいシンプルなデザイン",
      "高品質かつ手に取りやすい価格帯",
      "スタイルアップが叶う",
    ];
    points.forEach((p, i) => {
      const py = top + 2.1 + i * 0.35;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.7, y: py, w: 0.06, h: 0.25, fill: { color: C.primary }
      });
      s.addText(p, {
        x: 0.95, y: py, w: 8.5, h: 0.28,
        fontSize: 15, fontFace: FONT, color: C.title, bold: true,
        align: "left", valign: "middle", margin: 0
      });
    });

    addKeyMsg(s, "定番 x 高品質 x スタイルアップ の三拍子");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 3: 商品概要（写真左 + 箇条書き右）
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "商品概要");

    const top = centerY(3.4);

    s.addText("「盛れる」と「高品質」を両立する設計", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品写真2枚
    s.addImage({ ...imgBlack, x: 0.5, y: top + 0.6, w: 2.2, h: 2.7 });
    s.addImage({ ...imgWhite, x: 2.8, y: top + 0.6, w: 2.2, h: 2.7 });

    // 右: セールスポイント
    const features = [
      "軽くて暖かい",
      "ウエストマークで細見え、スタイルアップ",
      "ボリュームファーで小顔効果",
      "ショート/ロングどちらにも合う丈感",
      "軽くて疲れない着心地",
      "トレンドに左右されず一生着られる安心感",
    ];
    features.forEach((f, i) => {
      const fy = top + 0.7 + i * 0.42;
      s.addShape(pres.shapes.OVAL, {
        x: 5.3, y: fy + 0.05, w: 0.15, h: 0.15, fill: { color: C.primary }
      });
      s.addText(f, {
        x: 5.6, y: fy, w: 3.9, h: 0.35,
        fontSize: 14, fontFace: FONT, color: C.body, bold: true,
        align: "left", valign: "middle", margin: 0
      });
    });

    addKeyMsg(s, "見た目の華やかさと実用性を兼ね備えた一着");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 4: 機能性（3カラム素材カード）
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
      { title: "フィルパワー (FP) 800", sub: "圧倒的な軽さ", img: imgFeather, desc: "スペイン産ホワイトダック を使用予定\n驚きの軽さを実現" },
      { title: "ダウンとフェザーの黄金割合", sub: "90% / 10%", img: imgDownFill, desc: "ダウン=柔らかくて軽い\nフェザー=弾力があり形を整える" },
      { title: "華やかなファーで", sub: "デザイン性をプラス", img: imgFurMat, desc: "シープ(羊)ファーを使用予定\n小顔効果や映えにも有効" },
    ];

    const cardW = 2.7, cardH = 2.5, cardGap = 0.3;
    const startX = (SW - (cardW * 3 + cardGap * 2)) / 2;

    cards.forEach((c, i) => {
      const cx = startX + i * (cardW + cardGap);
      const cy = top + 0.6;

      s.addShape(pres.shapes.RECTANGLE, {
        x: cx, y: cy, w: cardW, h: cardH, fill: { color: C.white }
      });

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

      // 素材写真
      s.addImage({ ...c.img, x: cx + 0.35, y: cy + 0.65, w: 2.0, h: 0.9 });

      s.addText(c.desc, {
        x: cx + 0.1, y: cy + 1.65, w: cardW - 0.2, h: 0.75,
        fontSize: 10, fontFace: FONT, color: C.body,
        align: "center", valign: "top", margin: 0, lineSpacingMultiple: 1.3
      });
    });

    s.addText("※仕様変更可能", {
      x: 7.5, y: top + 3.15, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted,
      align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "FP800・ダウン90%・シープファーのハイスペック構成");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 5: 品質証明
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "品質証明");

    const top = centerY(2.3);

    s.addText("信頼性を保証するエビデンス", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // RDS認証
    addDefBlock(s, 0.5, top + 0.6, "RDS認証: 取得予定（動物福祉に関する国際認証）",
      "強制的な飼育や生きたまま羽毛を刈り取っていない農場で\n仕入れていることを証明する国際認証マーク", 6.5);

    // 認証マーク画像
    s.addImage({ ...imgRDS, x: 7.5, y: top + 0.5, w: 1.8, h: 1.0 });

    addSep(s, 0.5, top + 1.5, 9);

    // TC
    addDefBlock(s, 0.5, top + 1.7, "TC（Transaction Certificate）: 発行可能",
      "素材のパスポートのような証明書。偽物ではなく、\nクリーンなルートで届いた本物である証明書", 6.5);

    addKeyMsg(s, "国際認証で品質と倫理性を担保");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 6: 他ハイブランドとの比較（テーブル）
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "他ハイブランドとの比較");

    const top = centerY(3.0);

    const rows = [
      ["スペック項目", "本企画", "MONCLER", "PRADA等"],
      ["フィルパワー(FP)", "800FP", "710~800", "650~750(推定)"],
      ["ダウン混合率", "90% / 10%", "90% / 10%", "90% / 10%"],
      ["シルエット設計", "美くびれ・スタイルUP", "曲線美・ラグジュアリー", "ミニマル・モード"],
      ["ファーの品質", "シープ", "シープ / フォックス", "なし / 合成 / シープ"],
      ["販売価格帯", "15万 ~ 30万", "25万 ~ 50万+", "30万 ~ 60万+"],
    ];

    s.addTable(rows, {
      x: 0.5, y: top, w: 9.0,
      colW: [2.0, 2.5, 2.25, 2.25],
      border: { pt: 0.5, color: C.sep },
      rowH: [0.45, 0.4, 0.4, 0.4, 0.4, 0.4],
      fontFace: FONT,
      fontSize: 12,
      color: C.body,
      align: "center",
      valign: "middle",
      autoPage: false,
    });

    // ヘッダー行（個別セルで配置）
    const colX6 = [0.5, 2.5, 5.0, 7.25];
    const colW6 = [2.0, 2.5, 2.25, 2.25];
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top, w: 9.0, h: 0.45,
      fill: { color: C.primary }
    });
    rows[0].forEach((cell, i) => {
      s.addText(cell, {
        x: colX6[i], y: top, w: colW6[i], h: 0.45,
        fontSize: 12, fontFace: FONT, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    });

    s.addText("※独自比較", {
      x: 7.5, y: top + 2.65, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted,
      align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "ハイブランド同等スペックを半額以下で実現");
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

    s.addText("ハイブランド実績のある工場", {
      x: 0.5, y: top, w: 9, h: 0.4,
      fontSize: 20, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addDefBlock(s, 0.5, top + 0.6, "主要実績",
      "Max Mara, Acne Studios, THE NORTH FACE", 9);

    addDefBlock(s, 0.5, top + 1.45, "過去実績",
      "Canada Goose", 9);

    addSep(s, 0.5, top + 2.25, 9);

    // コールアウトボックス
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top + 2.4, w: 9, h: 0.45,
      fill: { color: C.kmBg }
    });
    s.addText("工場は中国にあり、製造時に工場見学 / 撮影が可能（コンテンツ化できる）", {
      x: 0.7, y: top + 2.4, w: 8.6, h: 0.45,
      fontSize: 14, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "ハイブランドOEMの工場で製造。現地撮影も可能");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 8: デザインイメージ（写真左 + スペック右）
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "デザインイメージ");

    const top = centerY(3.2);

    s.addText("MONCLERをベンチマーク", {
      x: 0.5, y: top, w: 9, h: 0.35,
      fontSize: 16, fontFace: FONT, color: C.title, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 左: 商品画像
    s.addImage({ ...imgBlackFull, x: 0.5, y: top + 0.5, w: 3.5, h: 2.7 });

    // 右: スペックリスト
    const specs = [
      "フードとファーはそれぞれ取り外し可能な設計で3WAY構造",
      "  1. ファー+フード付き\n  2. ファーなしフード付き\n  3. ファーとフードなし",
      "ウエストベルトも取り外し可能な2WAY構造",
      "160cmの女性でお尻が隠れるくらいの着丈",
    ];

    let sy = top + 0.5;
    specs.forEach((sp, i) => {
      const isSubList = sp.startsWith("  ");
      const h = isSubList ? 0.65 : 0.3;
      if (!isSubList) {
        s.addShape(pres.shapes.OVAL, {
          x: 4.5, y: sy + 0.07, w: 0.13, h: 0.13, fill: { color: C.primary }
        });
      }
      s.addText(sp, {
        x: isSubList ? 4.8 : 4.75, y: sy, w: 4.75, h,
        fontSize: isSubList ? 12 : 13, fontFace: FONT, color: C.body,
        bold: !isSubList,
        align: "left", valign: "top", margin: 0, lineSpacingMultiple: 1.3
      });
      sy += h + 0.08;
    });

    s.addText("※デザイン変更可能", {
      x: 7.5, y: top + 3.0, w: 2, h: 0.2,
      fontSize: 9, fontFace: FONT, color: C.muted,
      align: "right", valign: "middle", margin: 0
    });

    addKeyMsg(s, "3WAYフード + 2WAYベルトで着回し自在");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 9: カラー・サイズ展開
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "カラー・サイズ展開");

    const top = centerY(3.2);

    // 左: 着用写真2枚
    s.addImage({ ...imgBlack, x: 0.5, y: top, w: 2.2, h: 3.0 });
    s.addImage({ ...imgWhite, x: 2.8, y: top, w: 2.2, h: 3.0 });

    // 右上: サイズ
    const items = [
      "フリーサイズ想定",
      "取り入れやすい、ブラックとホワイトの2色展開想定",
    ];
    items.forEach((item, i) => {
      const iy = top + 0.1 + i * 0.4;
      s.addShape(pres.shapes.OVAL, {
        x: 5.3, y: iy + 0.05, w: 0.13, h: 0.13, fill: { color: C.primary }
      });
      s.addText(item, {
        x: 5.55, y: iy, w: 4.0, h: 0.35,
        fontSize: 13, fontFace: FONT, color: C.body,
        align: "left", valign: "middle", margin: 0
      });
    });

    addSep(s, 5.3, top + 1.0, 4.2);

    // カラー詳細
    const colors = [
      { label: "ブラック本体", detail: "ファーはブラウンベージュ" },
      { label: "ホワイト本体", detail: "ファーはミルクベージュ" },
    ];
    colors.forEach((c, i) => {
      const cy = top + 1.2 + i * 0.5;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 5.3, y: cy, w: 0.06, h: 0.4, fill: { color: C.primary }
      });
      s.addText(c.label, {
        x: 5.5, y: cy, w: 4.0, h: 0.22,
        fontSize: 13, fontFace: FONT, color: C.title, bold: true,
        align: "left", valign: "middle", margin: 0
      });
      s.addText(c.detail, {
        x: 5.5, y: cy + 0.22, w: 4.0, h: 0.2,
        fontSize: 12, fontFace: FONT, color: C.sub,
        align: "left", valign: "middle", margin: 0
      });
    });

    addKeyMsg(s, "ブラック x ホワイトの定番2色 / フリーサイズ");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 10: 販売戦略
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売戦略");

    const top = centerY(3.0);

    s.addText("タレント/モデル起用の撮影とノベルティ", {
      x: 0.5, y: top, w: 6, h: 0.4,
      fontSize: 18, fontFace: FONT, color: C.primary, bold: true,
      align: "left", valign: "middle", margin: 0
    });

    // 右: プロモ写真
    s.addImage({ ...imgModel, x: 7.0, y: top, w: 2.5, h: 2.8 });

    const strategies = [
      { title: "モデル起用のプロモーション撮影", desc: "知名度のあるモデル・タレントを起用し洗練された撮影を実施" },
      { title: "ポップアップや展示会の開催", desc: "高価格帯のアパレル商材のため、実際に触れる機会を提供" },
      { title: "販売記念のノベルティ", desc: "ファーと同素材のチャーム等、購入特典を検討中" },
    ];

    strategies.forEach((st, i) => {
      const sy = top + 0.6 + i * 0.8;
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.5, y: sy, w: 0.06, h: 0.65, fill: { color: C.primary }
      });
      s.addText(st.title, {
        x: 0.75, y: sy, w: 5.8, h: 0.28,
        fontSize: 15, fontFace: FONT, color: C.title, bold: true,
        align: "left", valign: "middle", margin: 0
      });
      s.addText(st.desc, {
        x: 0.75, y: sy + 0.3, w: 5.8, h: 0.3,
        fontSize: 13, fontFace: FONT, color: C.sub,
        align: "left", valign: "middle", margin: 0
      });
    });

    addKeyMsg(s, "撮影・展示会・ノベルティの三本柱で販売を加速");
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

    const schedule = [
      ["時期", "進行管理", "詳細"],
      ["4月上旬", "1stサンプル発注", "ベンチマークサンプルの発注"],
      ["4月下旬", "1stサンプルアップ", "デザイン・仕様の確認、2ndサンプル進行"],
      ["5月中旬〜下旬", "2ndサンプルアップ", "デザイン・仕様の最終確認"],
      ["5月下旬", "量産発注", "2026AWに向けて発注"],
      ["10月下旬〜11月上旬", "製品納品", "製品が倉庫に納品"],
      ["11月中旬", "リリース", "製品本販売開始"],
    ];

    s.addTable(schedule, {
      x: 0.5, y: top, w: 9.0,
      colW: [2.2, 2.4, 4.4],
      border: { pt: 0.5, color: C.sep },
      rowH: [0.38, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
      fontFace: FONT,
      fontSize: 12,
      color: C.body,
      align: "center",
      valign: "middle",
      autoPage: false,
    });

    // ヘッダー行
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top, w: 9.0, h: 0.38,
      fill: { color: C.primary }
    });
    const hdrCols = [
      { x: 0.5, w: 2.2 }, { x: 2.7, w: 2.4 }, { x: 5.1, w: 4.4 }
    ];
    schedule[0].forEach((cell, i) => {
      s.addText(cell, {
        x: hdrCols[i].x, y: top, w: hdrCols[i].w, h: 0.38,
        fontSize: 12, fontFace: FONT, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    });

    // 注記
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top + 2.65, w: 9, h: 0.35,
      fill: { color: C.kmBg }
    });
    s.addText("※優先的に進行するためにもタイトなスケジュールとなっています。", {
      x: 0.7, y: top + 2.65, w: 8.6, h: 0.35,
      fontSize: 12, fontFace: FONT, color: C.primary,
      align: "left", valign: "middle", margin: 0
    });

    addKeyMsg(s, "4月発注開始 → 11月リリースの製造スケジュール");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 12: 販売イメージ（サマリーテーブル）
  // ═══════════════════════════════════════════════
  {
    pg++;
    const s = pres.addSlide();
    addHeader(s, "販売イメージ");

    const top = centerY(3.2);

    const summary = [
      ["販売価格", "15~30万円", "デザイン・仕様によって変動あり"],
      ["販売時期", "11月中旬ごろ", "案件状況により調整"],
      ["機能性ポイント", "ハイブランドと同等の高品質", "800FP、ダウン90%、ハイブランド実績工場"],
      ["デザイン性ポイント", "スタイルアップ&小顔効果", "ウエストくびれライン+ボリュームファー"],
      ["販売戦略", "モデルやタレント起用", "大々的に打ち出し、ノベルティ施策も検討"],
      ["サイズ展開", "フリーサイズ", "サイズ展開可"],
      ["カラー展開", "ブラック、ホワイト 2色", "カラー展開可"],
      ["生産数", "2色合わせて300~1,000着", "販売価格とプロモーションによって変動"],
      ["売上想定", "4,500万~1.5億円", "販売価格15万円想定で算出"],
    ];

    const colX = [0.5, 2.5, 5.2];
    const colW2 = [2.0, 2.7, 4.3];
    const rowH = 0.34;

    // ヘッダー行
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: top, w: 9.0, h: rowH,
      fill: { color: C.primary }
    });
    ["項目", "内容", "備考"].forEach((h, i) => {
      s.addText(h, {
        x: colX[i], y: top, w: colW2[i], h: rowH,
        fontSize: 11, fontFace: FONT, color: C.white, bold: true,
        align: "center", valign: "middle", margin: 0
      });
    });

    // データ行
    summary.forEach((row, ri) => {
      const ry = top + rowH + ri * rowH;
      // 背景: 交互色
      if (ri % 2 === 0) {
        s.addShape(pres.shapes.RECTANGLE, {
          x: 0.5, y: ry, w: 9.0, h: rowH,
          fill: { color: C.white }
        });
      }
      row.forEach((cell, ci) => {
        s.addText(cell, {
          x: colX[ci], y: ry, w: colW2[ci], h: rowH,
          fontSize: ci === 0 ? 11 : 10, fontFace: FONT,
          color: ci === 0 ? C.title : C.body,
          bold: ci === 0,
          align: "center", valign: "middle", margin: 0
        });
      });
      // 行区切り線
      addSep(s, 0.5, ry + rowH - 0.01, 9);
    });

    addKeyMsg(s, "売上想定4,500万〜1.5億円のポテンシャル");
    addPageNum(s, pg + 1);
  }

  // ═══════════════════════════════════════════════
  // Slide 13: エンド (Type D)
  // ═══════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    const bTop = (SH - 1.8) / 2;

    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: bTop, w: 4.25, h: 0.035, fill: { color: C.primary }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 4.75, y: bTop, w: 4.75, h: 0.035, fill: { color: C.secondary }
    });
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

  // ── Save ──
  const outPath = "/tmp/influencer-proposal-v1.pptx";
  await pres.writeFile({ fileName: outPath });
  console.log(`Done: ${outPath} (${pres.slides.length} slides)`);
}

main().catch(e => { console.error(e); process.exit(1); });
