#!/bin/sh
set -eu

BASE_DIR="${STATS_BASE:-/mnt/data-disk/outage-stats/统计材料}"
ARCHIVE_DIR="${DB_ARCHIVE_DIR:-/mnt/data-disk/outage-stats/archive/db}"
RETENTION_DAYS="${DB_BACKUP_RETENTION_DAYS:-30}"

mkdir -p "$ARCHIVE_DIR"
timestamp="$(date +%Y%m%d-%H%M%S)"
target="$ARCHIVE_DIR/outage-stats-$timestamp.dump"

docker exec outage-stats-postgres pg_dump \
  --format=custom \
  --no-owner \
  --username="${POSTGRES_USER:-outage_stats}" \
  "${POSTGRES_DB:-outage_stats}" > "$target.tmp"
mv "$target.tmp" "$target"
find "$ARCHIVE_DIR" -type f -name 'outage-stats-*.dump' -mtime "+$RETENTION_DAYS" -delete

usage="$(df -P "$BASE_DIR" | awk 'NR==2 {gsub(/%/, "", $5); print $5}')"
case "$usage" in
  ''|*[!0-9]*) exit 0 ;;
esac
if [ "$usage" -ge 90 ]; then level="CRITICAL";
elif [ "$usage" -ge 80 ]; then level="WARNING";
elif [ "$usage" -ge 70 ]; then level="NOTICE";
else level="OK"; fi
logger -t outage-stats-disk "level=$level usage=${usage}% archive=$ARCHIVE_DIR"

