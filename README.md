# influencer-proposal-kit

インフルエンサー向け商品コラボ提案資料の自動生成ツール。
Claude Codeに「この人に○○を提案して」と話しかけるだけで、リサーチから完成Google Slidesまで一気通貫で実行。

## インストール

```bash
curl -fsSL https://raw.githubusercontent.com/knishida-td/influencer-proposal-kit/main/install.sh | bash
```

## 使い方

Claude Code（Web / Desktop / CLI）を開いて話しかけるだけ:

```
「https://www.instagram.com/ami927s/ にヘアオイル提案して」
「成瀬愛里さんに万能調味料を提案して」
「@xxx にファッション以外で何か提案して」
```

→ リサーチ → 商材企画 → PPTX生成 → Google Slides アップロードまで自動実行。

## 提案書の構成（13スライド固定）

| # | スライド | 内容 |
|---|---------|------|
| 1 | 表紙 | 商品ビジュアル + タイトル |
| 2 | コンセプト | なぜこの商品をこの人に提案するか |
| 3 | 商品概要 | 商品写真 + セールスポイント |
| 4 | 機能性/素材 | スペック詳細（3カラム） |
| 5 | 品質証明 | 認証・エビデンス |
| 6 | 競合比較 | 他ブランドとのスペック比較表 |
| 7 | 生産背景 | 工場実績・製造クオリティ |
| 8 | デザイン詳細 | 商品デザインの仕様 |
| 9 | バリエーション | カラー・フレーバー等の展開 |
| 10 | 販売戦略 | プロモーション計画 |
| 11 | スケジュール | 製造〜リリースのタイムライン |
| 12 | 販売イメージ | 価格・数量・売上サマリー |
| 13 | エンドスライド | クロージング |

## 対応ジャンル

ファッションに限らず、あらゆる商材カテゴリに対応:

| ジャンル | 実績例 |
|---------|-------|
| ファッション | ダウンジャケット、レザーミニバッグ、ナイロントート |
| ビューティー | プレミアムヘアオイル |
| 食品 | オリジナル万能調味料 |

## デザインシステム

SlideKit準拠（claude-pptx-mcp と同一のデザインシステム）。

## 前提条件

- Claude Code アカウント
- Node.js 18+
- pptx MCPサーバー（`curl -fsSL https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main/install.sh | bash`）

### Google Slides自動アップロード（任意）

GAS Webアプリをデプロイすると、生成したPPTXを自動でGoogle Slidesにアップロードできます。

1. `gas/upload-to-slides.gs` をGoogle Apps Scriptにコピー
2. Drive APIサービスを有効化（エディタ → サービス → Drive API）
3. デプロイ → ウェブアプリ → アクセス: 全員 → デプロイ
4. URLを設定: `echo 'GAS_UPLOAD_URL="https://script.google.com/macros/s/xxxxx/exec"' > ~/.config/influencer-proposal-kit/config`

未設定でも提案書の生成自体は可能です（手動でGoogle Driveにアップロード）。
