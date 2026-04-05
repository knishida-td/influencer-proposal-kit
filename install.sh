#!/usr/bin/env bash
# influencer-proposal-kit ワンライナーインストール
# 使い方: curl -fsSL https://raw.githubusercontent.com/knishida-td/influencer-proposal-kit/main/install.sh | bash
set -euo pipefail

REPO="knishida-td/influencer-proposal-kit"
REPO_RAW="https://raw.githubusercontent.com/${REPO}/main"
SKILL_DIR="$HOME/.claude/skills/influencer-proposal"
CLAUDE_MD="$HOME/.claude/CLAUDE.md"

echo "=== influencer-proposal-kit インストール ==="

# ── 1. Node.js チェック ──
if ! command -v node &>/dev/null; then
  echo "  Node.js が見つかりません。先にインストールしてください:"
  echo "  https://nodejs.org/"
  exit 1
fi

# ── 2. npm依存パッケージのグローバルインストール確認 ──
# pptxgenjs と sharp はスクリプト実行時にローカルで使うため、
# プロジェクトディレクトリにインストールする
PROJECT_DIR="$HOME/Projects/influencer-proposal-kit"
mkdir -p "$PROJECT_DIR"

if [ ! -f "$PROJECT_DIR/package.json" ]; then
  cat > "$PROJECT_DIR/package.json" << 'PKGJSON'
{
  "name": "influencer-proposal-kit",
  "private": true,
  "dependencies": {
    "pptxgenjs": "^4.0.1",
    "sharp": "^0.34.5"
  }
}
PKGJSON
  echo "  package.json を作成しました"
fi

if [ ! -d "$PROJECT_DIR/node_modules" ]; then
  echo "  npm install を実行中..."
  (cd "$PROJECT_DIR" && npm install --silent 2>/dev/null)
  echo "  依存パッケージをインストールしました"
else
  echo "  依存パッケージはインストール済みです"
fi

# ── 3. スキルファイルのインストール ──
mkdir -p "$SKILL_DIR"

install_file() {
  local name="$1"
  local repo_path="$2"
  local dest="$3"

  # ローカル実行時（git clone後）
  local script_dir=""
  script_dir="$(cd "$(dirname "$0")" 2>/dev/null && pwd)" || true
  if [ -n "$script_dir" ] && [ -f "$script_dir/$repo_path" ]; then
    cp "$script_dir/$repo_path" "$dest"
    echo "  $name をインストールしました（ローカル）"
    return 0
  fi

  # curl | bash 時: GitHubからダウンロード
  if curl -fsSL "$REPO_RAW/$repo_path" -o "$dest" 2>/dev/null; then
    echo "  $name をインストールしました（GitHub）"
    return 0
  fi

  echo "  $name のインストールに失敗しました"
  return 1
}

install_file "SKILL.md" "skill/SKILL.md" "$SKILL_DIR/SKILL.md"

# ── 4. コアファイルのインストール ──
install_file "lib/slidekit.js" "lib/slidekit.js" "$PROJECT_DIR/lib/slidekit.js"
install_file "generate.js" "generate.js" "$PROJECT_DIR/generate.js"
mkdir -p "$PROJECT_DIR/templates"
install_file "templates/example.json" "templates/example.json" "$PROJECT_DIR/templates/example.json"

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

# ── 6. アップロードスクリプトのインストール ──
install_file "upload-to-gslides.sh" "scripts/upload-to-gslides.sh" "$PROJECT_DIR/scripts/upload-to-gslides.sh" || true
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
