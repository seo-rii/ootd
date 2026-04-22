# excel-oracle-win 프로토콜

## 목적

실제 Microsoft Excel 데스크톱을 정답 오라클로 사용하여
우리 라이브러리의 load/calc/save/render 결과를 검증한다.

## 전제

- Windows
- Microsoft Excel 설치
- 매크로 자동 실행 금지
- 테스트 전용 로컬 경로 사용
- 네트워크/외부 링크 차단

## 실행 순서

1. Excel.Application 생성
2. `Visible = false`
3. `DisplayAlerts = false`
4. `AutomationSecurity = msoAutomationSecurityForceDisable`
5. `Workbooks.Open(...)`
   - `UpdateLinks = 0`
   - `ReadOnly = true` (기본 oracle read 패스)
6. `Application.CalculateFullRebuild()`
7. workbook metadata 수집
8. probe range에 대해:
   - Formula
   - FormulaR1C1 (옵션)
   - Value2
   - NumberFormat
   - MergeArea
9. PDF export
   - workbook 또는 worksheet 단위
10. JSON 저장
11. Excel 종료 및 프로세스 정리

## 권장 probe 집합

- 각 시트의 UsedRange 경계 셀
- 모든 formula cell
- named range anchor
- table header/data/totals row
- merged area 좌상단과 내부 셀
- validation/CF 적용 셀 샘플
- dynamic array origin + spill area
- print area / frozen pane 인접 셀

## 산출물

- `oracle.json`
- `export.pdf`
- `run.log`

## 실패 분류

- open_failure
- calculate_failure
- export_failure
- semantic_mismatch
- format_mismatch
- unexpected_excel_repair
- crash

## 비교 단계

### A. Semantic compare
- Formula text
- Value2
- Error code
- UsedRange
- names/tables metadata

### B. Render compare
- PDF rasterize
- pixel diff with tolerance

### C. Save-then-open compare
- 우리 writer가 저장한 파일을 다시 Excel로 열어 동일 프로토콜 반복
