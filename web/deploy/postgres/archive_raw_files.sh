#!/bin/bash
set -eu

STATS_DIR="${STATS_BASE:-/mnt/data-disk/outage-stats/统计材料}"
SOURCE_DIR="$STATS_DIR/input/Newdata/daily"
ARCHIVE_ROOT="${RAW_ARCHIVE_DIR:-/mnt/data-disk/outage-stats/archive/raw}"
RETENTION_DAYS="${RAW_RETENTION_DAYS:-90}"

[ -d "$SOURCE_DIR" ] || exit 0
find "$SOURCE_DIR" -maxdepth 1 -type f \( -name '*.xlsx' -o -name '*.xls' \) -mtime "+$RETENTION_DAYS" -print0 |
while IFS= read -r -d '' source; do
  year="$(date -r "$source" +%Y)"
  month="$(date -r "$source" +%m)"
  target_dir="$ARCHIVE_ROOT/$year/$month"
  mkdir -p "$target_dir"
  target="$target_dir/$(basename "$source")"
  if [ -e "$target" ]; then
    target="$target_dir/$(date +%Y%m%d-%H%M%S)-$(basename "$source")"
  fi
  mv "$source" "$target"
done
