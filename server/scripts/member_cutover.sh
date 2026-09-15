#!/usr/bin/env bash
# Member cutover helper (2026-09-19). ASCII only: PowerShell mangles Japanese.
# Run on the server PC via Git Bash:  bash ~/member_cutover.sh <command> [file]
# Every command is safe to re-run. Nothing here touches the spreadsheet or GAS.
set -euo pipefail
cd ~/projects/mayumi-app/server
# Git Bash rewrites "/app/x" into "C:/Program Files/Git/app/x" before Docker sees it. Stop that.
export MSYS_NO_PATHCONV=1
export MSYS2_ARG_CONV_EXCL="*"

PY="docker compose exec -T web python manage.py"
PSQL="docker compose exec -T db psql -U postgres -d mayumi -t -A -c"

case "${1:-help}" in
  backup)
    # pg_dump -Fc into server/backups (same shape as scripts/backup ps1). Refuses tiny files.
    mkdir -p backups
    f="backups/mayumi-$(date +%Y%m%d-%H%M)-cutover.dump"
    docker compose exec -T db pg_dump -U postgres -Fc mayumi > "$f"
    sz=$(stat -c %s "$f" 2>/dev/null || stat -f %z "$f")
    if [ "$sz" -lt 1024 ]; then echo "BACKUP FAILED (too small): $f"; rm -f "$f"; exit 1; fi
    echo "backup ok: $f ($sz bytes)";;
  count)
    # Read-only. Member counts, push targets, and what the server created on its own
    # since <since> (default: today 00:00). Those are the rows to copy back by hand if we roll back.
    since="${2:-$(date +%Y-%m-%d) 00:00}"
    echo "members (not deleted): $($PSQL "select count(*) from members_member where deleted=false and deleted_at is null")"
    echo "members total        : $($PSQL "select count(*) from members_member")"
    echo "push targets (non-empty subscription): $($PSQL "select count(*) from members_member where coalesce(push_subscription,'')<>''")"
    echo "latest created_at    : $($PSQL "select member_id||' '||coalesce(created_at::text,'') from members_member order by created_at desc nulls last limit 1")"
    echo "--- created on the server since $since (JST) ---"
    $PSQL "select member_id||' '||to_char(created_at at time zone 'Asia/Tokyo','YYYY-MM-DD HH24:MI')||' '||coalesce(registration_source,'') from members_member where created_at >= timestamp '$since' at time zone 'Asia/Tokyo' order by created_at"
    echo "--- audit rows born on the server (sheet_row is null) since $since ---"
    $PSQL "select to_char(happened_at at time zone 'Asia/Tokyo','YYYY-MM-DD HH24:MI')||' '||kind||' '||coalesce(result,'')||' '||left(target,40) from records_auditlog where sheet_row is null and happened_at >= timestamp '$since' at time zone 'Asia/Tokyo' order by happened_at";;
  preview)   [ -n "${2:-}" ] || { echo "usage: preview <json in server/>"; exit 1; }
    $PY 会員を取り込む "/app/$2" --下見;;
  import)    [ -n "${2:-}" ] || { echo "usage: import <json in server/>"; exit 1; }
    $PY 会員を取り込む "/app/$2";;
  compare)   [ -n "${2:-}" ] || { echo "usage: compare <json in server/>"; exit 1; }
    $PY 会員を突き合わせる "/app/$2";;
  marks-preview) [ -n "${2:-}" ] || { echo "usage: marks-preview <json in server/>"; exit 1; }
    $PY 会員ごとの印を取り込む "/app/$2" --下見;;
  marks-import)  [ -n "${2:-}" ] || { echo "usage: marks-import <json in server/>"; exit 1; }
    $PY 会員ごとの印を取り込む "/app/$2";;
  shred)     [ -n "${2:-}" ] || { echo "usage: shred <json in server/>"; exit 1; }
    # The export JSON holds plaintext passcodes. Remove it once imported and compared.
    rm -f -- "$2" && echo "removed $2";;
  live-on)
    # Tell the Django manage screens that the server is now the source of truth for members.
    $PY 会員の正を切り替える server;;
  live-off)
    $PY 会員の正を切り替える sheet;;
  live)
    $PY 会員の正を切り替える;;
  check)
    # Read-only. Who could NOT get in after the switch, judged from the server ledger
    # (run again right after the import on the day). GAS "入れるか確かめる()" looks at the sheet; this is the server side.
    echo "no passcode at all (cannot log in):     $($PSQL "select count(*) from members_member where deleted=false and coalesce(passcode_hash,'')='' and coalesce(password_hash,'')=''")"
    echo "old password scheme only (hash+salt):    $($PSQL "select count(*) from members_member where deleted=false and coalesce(passcode_hash,'')='' and coalesce(password_hash,'')<>''")"
    echo "no birthday (cannot use recovery):       $($PSQL "select count(*) from members_member where deleted=false and birthday is null")"
    echo "no birthday and no phone (no reset):     $($PSQL "select count(*) from members_member where deleted=false and birthday is null and coalesce(phone,'')=''")"
    echo "empty name:                              $($PSQL "select count(*) from members_member where deleted=false and coalesce(name,'')=''")"
    echo "duplicate LINE user ids:                 $($PSQL "select count(*) from (select line_user_id from members_member where coalesce(line_user_id,'')<>'' group by line_user_id having count(*)>1) t")"
    echo "members with a role set:                 $($PSQL "select count(*)||' ('||string_agg(distinct role, ',')||')' from members_member where coalesce(role,'')<>''")";;
  probe)
    # After the switch: the customer-facing window must answer without an API key.
    for a in getNews getUserRewardStatus getRecoveryCandidates; do
      printf "%-22s " "$a"; curl -s -o /dev/null -w "%{http_code}\n" "https://mayumi-api.tail8efe0d.ts.net/api?action=$a&data=%7B%22memberId%22%3A%22MYM-0000%22%7D"
    done
    echo "web log (last 30 lines, errors only):"; docker compose logs --tail=200 web 2>&1 | grep -iE "error|traceback" | tail -30 || true;;
  *)
    cat <<'H'
usage: bash ~/member_cutover.sh <command> [file]
  backup                 pg_dump -> server/backups/mayumi-<date>-cutover.dump
  count [YYYY-MM-DD HH:MM]  counts, push targets, rows the server created since then (read only)
  preview <json>         dry run of the member import  (file placed in server/)
  import  <json>         member import
  compare <json>         sheet JSON vs server ledger, every field (read only)
  marks-preview <json>   dry run of the per-member marks import
  marks-import  <json>   per-member marks import
  shred   <json>         delete the export JSON (contains plaintext passcodes)
  live-on | live-off     raise / lower the "server is the source for members" flag (manage screens)
  live                   show the flag
  check                  who could not get in after the switch (read only)
  probe                  after the switch: public window answers + web errors
H
    ;;
esac
