#!/usr/bin/env bash
# Xoá lịch sử chat / agent của Google Antigravity trên macOS.
# Mặc định chỉ liệt kê (dry-run). Thêm --yes để xoá thật.
# Giữ nguyên: đăng nhập (oauth_creds.json, google_accounts.json), cấu hình MCP,
#             settings/extension của IDE, knowledge (bộ nhớ dài hạn của agent).
#
# Dùng:  bash command_pc/clear-antigravity-history.sh          # xem trước
#        bash command_pc/clear-antigravity-history.sh --yes    # xoá thật
# Hãy thoát hẳn Antigravity (Cmd+Q) trước khi xoá.

set -euo pipefail

DO_DELETE=0
[[ "${1:-}" == "--yes" ]] && DO_DELETE=1

if pgrep -fi "antigravity" >/dev/null 2>&1 && [[ $DO_DELETE -eq 1 ]]; then
  echo "Antigravity đang chạy. Hãy thoát hẳn (Cmd+Q) rồi chạy lại."; exit 1
fi

G="$HOME/.gemini"
IDE_SUP="$HOME/Library/Application Support/Antigravity IDE/User"

candidates=(
  # Hội thoại, bản ghi agent, tóm tắt
  "$G/antigravity/conversations"
  "$G/antigravity/brain"
  "$G/antigravity/annotations"
  "$G/antigravity/conversation_summaries.db"
  "$G/antigravity/agyhub_summaries_proto.pb"
  "$G/antigravity/scratch"
  "$G/antigravity/log"
  "$G/antigravity/cache"
  "$G/antigravity-ide/conversations"
  "$G/antigravity-ide/brain"
  "$G/antigravity-ide/implicit"
  "$G/antigravity-ide/html_artifacts"
  "$G/antigravity-ide/browser_recordings"
  # Trạng thái chat theo workspace + lịch sử file cục bộ của IDE
  "$IDE_SUP/workspaceStorage"
  "$IDE_SUP/History"
)

targets=()
for p in "${candidates[@]}"; do
  [[ -e "$p" ]] && targets+=("$p")
done

if [[ ${#targets[@]} -eq 0 ]]; then
  echo "Không có gì để xoá."; exit 0
fi

echo "Sẽ xoá ${#targets[@]} mục:"
printf '  %s\n' "${targets[@]}"
du -ch "${targets[@]}" 2>/dev/null | tail -1 | sed 's/^/Tổng dung lượng: /'

if [[ $DO_DELETE -eq 0 ]]; then
  echo; echo "(dry-run) Chạy lại với --yes để xoá thật."; exit 0
fi

rm -rf -- "${targets[@]}"
echo "Đã xoá xong."
