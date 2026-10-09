#!/usr/bin/env bash
# 주간 부동산 보고서 생성 (Word + 자동수식 Excel)
# 사용법: bash scripts/run_weekly.sh [데이터.json]
# 출력: reports/weekly/<asOf>/  (asOf = 데이터 파일의 기준일)
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DATA="${1:-$ROOT/scripts/data/complexes.json}"
ASOF="$(python3 -c 'import json,sys; print(json.load(open(sys.argv[1], encoding="utf-8"))["asOf"])' "$DATA")"
OUT="$ROOT/reports/weekly/$ASOF"
mkdir -p "$OUT"

# 데이터·생성 코드가 지난 실행과 같으면 건너뜀 (docx/xlsx는 생성 시각이 달라 매번 새 파일이 되므로)
STAMP="$OUT/.source.sha256"
HASH="$(cat "$DATA" "$ROOT/scripts/build_report.js" "$ROOT/scripts/build_xlsx.py" | sha256sum | cut -d' ' -f1)"
if [ "${FORCE:-0}" != "1" ] && [ -f "$STAMP" ] && [ "$(cat "$STAMP")" = "$HASH" ]; then
  echo "변경 없음: $ASOF 데이터·코드가 동일하여 건너뜀 (강제 실행: FORCE=1)"
  exit 0
fi

[ -d "$ROOT/scripts/node_modules/docx" ] || npm ci --prefix "$ROOT/scripts" --silent
python3 -c 'import openpyxl' 2>/dev/null || pip install -q -r "$ROOT/scripts/requirements.txt"

node "$ROOT/scripts/build_report.js" "$OUT/부산서부권_34평_비교분석_보고서.docx" "$DATA"
python3 "$ROOT/scripts/build_xlsx.py" "$OUT/부산서부권_34평_갭투자_자동계산.xlsx" "$DATA"
echo "$HASH" > "$STAMP"
echo "완료: $OUT"
