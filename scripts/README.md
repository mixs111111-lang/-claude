# 주간 부동산 보고서 자동 생성

매주 **금요일 16:00 (KST)** 에 Word 보고서와 자동수식 Excel을 다시 만든다.

| 파일 | 역할 |
|---|---|
| `data/complexes.json` | 입력 데이터(기준일 `asOf`, 단지별 매매·전세·상승률·점수·근거). **매주 이 파일만 갱신** |
| `build_report.js` | Word 보고서 생성 (`node build_report.js <출력.docx> [데이터.json]`) |
| `build_xlsx.py` | 자동수식 Excel 생성 (`python3 build_xlsx.py <출력.xlsx> [데이터.json]`) |
| `run_weekly.sh` | 위 두 개를 실행해 `reports/weekly/<asOf>/` 에 저장 |
| `../.github/workflows/weekly-report.yml` | GitHub Actions 스케줄 (cron `0 7 * * 5` UTC = 금 16:00 KST), 수동 실행, 데이터 변경 시 자동 실행 |

## 로컬 실행

```bash
bash scripts/run_weekly.sh
```

## 주의
- 같은 `asOf`로 다시 돌리면 결과가 같아 커밋이 생기지 않는다. 새 주차 데이터는 `asOf`를 바꿔서 넣는다.
- GitHub 예약 실행은 트래픽에 따라 수 분~수십 분 늦을 수 있다.
- 공개 저장소는 60일간 커밋이 없으면 예약 워크플로가 자동 비활성화된다(Actions 탭에서 재활성화).
