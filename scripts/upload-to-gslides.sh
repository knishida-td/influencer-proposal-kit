#!/usr/bin/env bash
# PPTXファイルをGoogle Drive APIで直接アップロード → Google Slides変換
# clasp認証トークンを流用（追加の認証設定不要）
#
# 使い方: ./upload-to-gslides.sh /tmp/proposal.pptx "提案資料名"
# 前提: clasp login 済み（~/.clasprc.json にトークンあり）
set -euo pipefail

PPTX_PATH="${1:?使い方: $0 <pptx-path> [name]}"
NAME="${2:-$(basename "$PPTX_PATH" .pptx)}"
CLASP_RC="$HOME/.clasprc.json"

if [ ! -f "$CLASP_RC" ]; then
  echo "clasp未認証です。先に clasp login を実行してください。"
  open "$PPTX_PATH"
  exit 0
fi

# トークン取得（期限切れの場合は自動リフレッシュ）
ACCESS_TOKEN=$(python3 -c "
import json, time, urllib.request, urllib.parse
with open('$CLASP_RC') as f:
    d = json.load(f)
tokens = d.get('tokens', d.get('token', {}))
tok = tokens.get('default', tokens) if isinstance(tokens, dict) and 'default' in tokens else tokens
access_token = tok.get('access_token', '')
expiry_date = tok.get('expiry_date', 0)
refresh_token = tok.get('refresh_token', '')
client_id = d.get('oauth2ClientSettings', {}).get('clientId', tok.get('client_id', ''))
client_secret = d.get('oauth2ClientSettings', {}).get('clientSecret', tok.get('client_secret', ''))

# Check if token is expired
if expiry_date and expiry_date < time.time() * 1000 and refresh_token and client_id and client_secret:
    data = urllib.parse.urlencode({
        'client_id': client_id,
        'client_secret': client_secret,
        'refresh_token': refresh_token,
        'grant_type': 'refresh_token'
    }).encode()
    req = urllib.request.Request('https://oauth2.googleapis.com/token', data=data)
    resp = json.loads(urllib.request.urlopen(req).read())
    access_token = resp.get('access_token', access_token)

print(access_token)
" 2>/dev/null)

if [ -z "$ACCESS_TOKEN" ]; then
  echo "トークン取得に失敗しました。clasp login を再実行してください。"
  open "$PPTX_PATH"
  exit 0
fi

echo "アップロード中: $NAME ..."

METADATA="{\"name\":\"$NAME\",\"mimeType\":\"application/vnd.google-apps.presentation\"}"

RESPONSE=$(curl -s -X POST \
  "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink" \
  -H "Authorization: Bearer $ACCESS_TOKEN" \
  -F "metadata=$METADATA;type=application/json" \
  -F "file=@$PPTX_PATH;type=application/vnd.openxmlformats-officedocument.presentationml.presentation" 2>&1)

URL=$(echo "$RESPONSE" | python3 -c "import json,sys; d=json.load(sys.stdin); print(d.get('webViewLink',''))" 2>/dev/null || echo "")

if [ -n "$URL" ]; then
  echo "完了: $URL"
  open "$URL"
else
  echo "アップロード失敗。ファイルを直接開きます..."
  open "$PPTX_PATH"
  exit 1
fi
