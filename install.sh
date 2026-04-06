#!/usr/bin/env bash
# influencer-proposal-kit ワンライナーインストール
# 使い方: curl -fsSL https://raw.githubusercontent.com/knishida-td/influencer-proposal-kit/main/install.sh | bash
set -euo pipefail

REPO_URL="https://github.com/knishida-td/influencer-proposal-kit.git"
PROJECT_DIR="$HOME/Projects/influencer-proposal-kit"
SKILL_DIR="$HOME/.claude/skills/influencer-proposal"
CLAUDE_MD="$HOME/.claude/CLAUDE.md"

echo "=== influencer-proposal-kit インストール ==="

# ── 1. Node.js チェック（なければ自動インストール） ──
if ! command -v node &>/dev/null; then
  echo "  Node.js が見つかりません。インストールします..."
  if command -v brew &>/dev/null; then
    brew install node
  elif command -v apt-get &>/dev/null; then
    sudo apt-get update && sudo apt-get install -y nodejs npm
  else
    echo "  Homebrewもapt-getもありません。Node.jsを手動インストールしてください:"
    echo "  https://nodejs.org/"
    exit 1
  fi
  echo "  Node.js をインストールしました: $(node -v)"
fi

# ── 2. Git clone / pull（常に最新を取得） ──
if [ -d "$PROJECT_DIR/.git" ]; then
  echo "  リポジトリを最新に更新中..."
  (cd "$PROJECT_DIR" && git pull --ff-only origin main 2>/dev/null) || \
    (cd "$PROJECT_DIR" && git fetch origin && git reset --hard origin/main)
  echo "  最新版に更新しました"
else
  echo "  リポジトリをクローン中..."
  mkdir -p "$(dirname "$PROJECT_DIR")"
  git clone "$REPO_URL" "$PROJECT_DIR"
  echo "  クローン完了"
fi

# ── 3. npm依存パッケージ ──
if [ ! -d "$PROJECT_DIR/node_modules" ]; then
  echo "  npm install を実行中..."
  (cd "$PROJECT_DIR" && npm install --silent 2>/dev/null)
  echo "  依存パッケージをインストールしました"
else
  echo "  依存パッケージはインストール済みです"
fi

# ── 4. スキルファイルのシンボリックリンク ──
mkdir -p "$SKILL_DIR"
ln -sf "$PROJECT_DIR/skill/SKILL.md" "$SKILL_DIR/SKILL.md"
echo "  SKILL.md をリンクしました"

# ── 5. CLAUDE.md にルールを追記 ──
PROPOSAL_RULE='インフルエンサーへの商品提案依頼では、Skill「influencer-proposal」を使用すること。リサーチ→商材企画→PPTX生成→Google Slidesアップロードを一気通貫で実行する。'

if [ ! -f "$CLAUDE_MD" ]; then
  echo "$PROPOSAL_RULE" > "$CLAUDE_MD"
  echo "  CLAUDE.md を新規作成しました"
elif ! grep -q "influencer-proposal" "$CLAUDE_MD" 2>/dev/null; then
  printf '\n%s\n' "$PROPOSAL_RULE" >> "$CLAUDE_MD"
  echo "  CLAUDE.md にルールを追記しました"
else
  echo "  CLAUDE.md にルールは設定済みです"
fi

# ── 6. アップロードスクリプトの実行権限 ──
chmod +x "$PROJECT_DIR/scripts/upload-to-gslides.sh" 2>/dev/null || true

# ── 7. clasp (Google Slides アップロード用) の確認 ──
if command -v clasp &>/dev/null; then
  echo "  clasp はインストール済みです"
else
  echo "  clasp をインストール中..."
  npm install -g @google/clasp 2>/dev/null || echo "  ⚠ clasp のインストールに失敗しました（npm install -g @google/clasp を手動実行してください）"
fi

if [ -f "$HOME/.clasprc.json" ]; then
  echo "  clasp 認証済みです"
else
  echo ""
  echo "  Google Slides自動アップロードを使うには clasp login を実行してください（任意）"
  echo "  ※未設定でも提案書の生成自体は可能です"
fi

# ── 8. pptx MCP サーバーの確認 ──
SETTINGS_FILE="$HOME/.claude/settings.json"
if [ -f "$SETTINGS_FILE" ] && grep -q "claude-pptx-mcp\|pptx" "$SETTINGS_FILE" 2>/dev/null; then
  echo "  pptx MCPサーバーは設定済みです"
else
  echo ""
  echo "  PPTX生成にはpptx MCPサーバーも必要です。未設定の場合:"
  echo "  curl -fsSL https://raw.githubusercontent.com/knishida-td/claude-pptx-mcp/main/install.sh | bash"
fi

echo ""
echo "=== インストール完了 ==="
echo ""
echo "使い方:"
echo "  Claude Code を開いて以下のように話しかけるだけ:"
echo ""
echo '  「https://www.instagram.com/ami927s/ にヘアオイル提案して」'
echo '  「成瀬愛里さんに万能調味料を提案して」'
echo '  「@xxx にファッション以外で何か提案して」'
echo ""
echo "  → リサーチ → 企画 → PPTX生成 → Google Slides で完成します"
echo ""
echo "更新方法:"
echo "  同じコマンドを再実行するだけで最新版に更新されます"
