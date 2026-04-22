# Excel Compatibility Core 문서 번들 인덱스

이 ZIP은 Rust + WASM 기반 Excel Compatibility Core 프로젝트를 설계하기 위한 문서 묶음이다.

## 포함 문서

### 핵심 설계
- `excel_compatibility_architecture.md`
  - 전체 아키텍처, 스펙 우선순위, 구현 단계
- `excel_project_structure_and_interfaces.md`
  - Cargo workspace 구조, crate 경계, 핵심 인터페이스
- `core_interfaces.rs`
  - Rust trait/struct 스켈레톤
- `cargo-workspace-template.toml`
  - workspace 템플릿

### 스펙/메타
- `sources.toml`
  - 스펙 소스 레지스트리 초안
- `office-idl.schema.json`
  - canonical object model schema 초안

### 테스트
- `excel-oracle-win-protocol.md`
  - Windows Excel 오라클 테스트 절차
- `excel_test_corpus_and_validation_guide.md`
  - 코퍼스 수집, fixture 분류, validator/oracle/differential 테스트 가이드

### 추가 참고 자료
- `excel_engine_principles_and_pitfalls.md`
  - 계산/수식/날짜/Range/OOXML 보존 원리와 함정
- `excel_reference_map_and_reading_order.md`
  - 공식 자료 지도와 추천 읽기 순서

## 추천 읽기 순서

1. `excel_compatibility_architecture.md`
2. `excel_project_structure_and_interfaces.md`
3. `excel_engine_principles_and_pitfalls.md`
4. `excel_reference_map_and_reading_order.md`
5. `excel_test_corpus_and_validation_guide.md`
6. `excel-oracle-win-protocol.md`
7. `sources.toml`
8. `office-idl.schema.json`
9. `core_interfaces.rs`
10. `cargo-workspace-template.toml`

## 추천 다음 작업

1. repo 초기화 및 workspace 생성
2. `office-idl` / `office-codegen`부터 착수
3. `office-opc` + `excel-xlsx` lossless round-trip 구현
4. `excel-oracle-win` 테스트 러너 구축
5. fixture manifest와 synthetic generator 추가
