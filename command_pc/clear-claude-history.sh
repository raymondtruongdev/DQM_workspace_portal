#!/usr/bin/env bash
# Xoá lịch sử chat Claude Code trên máy này.
# Mặc định chỉ liệt kê (dry-run). Thêm --yes để xoá thật.
# Giữ nguyên: settings, CLAUDE.md, skills, và thư mục memory của từng project.
#
# Dùng:  bash command_pc/clear-claude-history.sh          # xem trước
#        bash command_pc/clear-claude-history.sh --yes    # xoá thật
# Hãy đóng mọi phiên Claude Code trước khi xoá.

set -euo pipefail

CLAUDE_DIR="${CLAUDE_CONFIG_DIR:-$HOME/.claude}"
DO_DELETE=0
[[ "${1:-}" == "--yes" ]] && DO_DELETE=1

if [[ ! -d "$CLAUDE_DIR" ]]; then
  echo "Không thấy $CLAUDE_DIR"; exit 0
fi

targets=()

# 1) Transcript từng phiên: projects/<project>/*.jsonl và thư mục phiên (UUID);
#    KHÔNG đụng vào projects/<project>/memory
if [[ -d "$CLAUDE_DIR/projects" ]]; then
  while IFS= read -r -d '' p; do targets+=("$p"); done < <(
    find "$CLAUDE_DIR/projects" -mindepth 2 -maxdepth 2 \
      \( -name '*.jsonl' -o -type d -regex '.*/[0-9a-fA-F-]\{36\}' \) -print0
  )
fi

# 2) Lịch sử prompt (mũi tên lên) và các dữ liệu phiên khác
for p in \
  "$CLAUDE_DIR/history.jsonl" \
  "$CLAUDE_DIR/todos" \
  "$CLAUDE_DIR/shell-snapshots" \
  "$CLAUDE_DIR/file-history" \
  "$CLAUDE_DIR/session-env" \
  "$CLAUDE_DIR/paste-cache" \
  "$CLAUDE_DIR/debug" \
  "$CLAUDE_DIR/statsig"; do
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
