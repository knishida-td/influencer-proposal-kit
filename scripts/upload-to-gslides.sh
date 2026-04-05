#!/usr/bin/env bash
# PPTXファイルをGoogle Slidesにアップロードするスクリプト
# 使い方: ./upload-to-gslides.sh /tmp/proposal.pptx "提案資料名"
#
# 前提: GAS_UPLOAD_URL 環境変数にデプロイ済みGAS WebアプリURLを設定
#   export GAS_UPLOAD_URL="https://script.google.com/macros/s/xxxxx/exec"
#
# 設定ファイル: ~/.config/influencer-proposal-kit/config
set -euo pipefail

PPTX_PATH="${1:?使い方: $0 <pptx-path> [name]}"
NAME="${2:-$(basename "$PPTX_PATH" .pptx)}"

# 設定読み込み
CONFIG_FILE="$HOME/.config/influencer-proposal-kit/config"
if [ -f "$CONFIG_FILE" ]; then
  # shellcheck source=/dev/null
  source "$CONFIG_FILE"
fi

if [ -z "${GAS_UPLOAD_URL:-}" ]; then
  echo "GAS_UPLOAD_URL が未設定です。"
  echo ""
  echo "設定方法:"
  echo "  1. gas/upload-to-slides.gs をGoogle Apps Scriptにデプロイ"
  echo "  2. 以下を実行:"
  echo "     mkdir -p ~/.config/influencer-proposal-kit"
  echo '     echo '\''GAS_UPLOAD_URL="https://script.google.com/macros/s/xxxxx/exec"'\'' > ~/.config/influencer-proposal-kit/config'
  echo ""
  echo "代替: ファイルを直接開きます..."
  open "$PPTX_PATH"
  exit 0
fi

# ファイルサイズチェック (GAS制限: 50MB)
FILE_SIZE=$(stat -f%z "$PPTX_PATH" 2>/dev/null || stat -c%s "$PPTX_PATH" 2>/dev/null)
if [ "$FILE_SIZE" -gt 52428800 ]; then
  echo "ファイルサイズが50MBを超えています。手動でアップロードしてください。"
  open "$PPTX_PATH"
  exit 1
fi

echo "アップロード中: $NAME ..."

# Base64エンコード + JSON送信
BASE64=$(base64 -i "$PPTX_PATH")
RESPONSE=$(curl -fsSL -X POST "$GAS_UPLOAD_URL" \
  -H "Content-Type: application/json" \
  -d "$(python3 -c "
import json, sys
print(json.dumps({
    'name': sys.argv[1],
    'file': sys.argv[2]
}))
" "$NAME.pptx" "$BASE64")" 2>&1)

# レスポンス解析
URL=$(echo "$RESPONSE" | python3 -c "import json,sys; d=json.load(sys.stdin); print(d.get('url',''))" 2>/dev/null || echo "")

if [ -n "$URL" ]; then
  echo "完了: $URL"
  open "$URL"
else
  echo "アップロード失敗: $RESPONSE"
  echo "ファイルを直接開きます..."
  open "$PPTX_PATH"
  exit 1
fi
