# OM Source Acquisition

이 문서는 `ootd`가 Excel Object Model의 machine-readable 계약을 어떤 순서로 수집하고,
무엇을 canonical input으로 삼는지 정리한다.

관련 pinned manifest는 [om_sources.toml](/home/seorii/dev/hancomac/ootd/specs/pinned/om_sources.toml)이다.

## Principle

- Excel OM의 정본 계약은 Learn 설명 문서가 아니라 Windows 환경에서 추출한 Excel COM type library와 PIA다.
- VBA / Interop 문서는 추출된 surface의 semantics, naming, examples, behavior를 보강하는 용도다.
- 따라서 `office-idl`은 문서 요약본이 아니라 type/member를 정규화한 canonical dataset을 받아야 한다.

## Stable Pins

현재 이 repo에서 확정한 값은 아래까지다.

- COM source family: `Excel` type library
- COM source version: `16.0`
- Platform: `windows`
- Managed projection: `Microsoft.Office.Interop.Excel`
- Assembly file: `Microsoft.Office.Interop.Excel.dll`

이 값들은 [om_sources.toml](/home/seorii/dev/hancomac/ootd/specs/pinned/om_sources.toml)에 고정했다.

## Environment-Specific Fields

아래 값은 이 Linux workspace에서 거짓으로 채우면 안 된다.
Windows + Excel 캡처 시점에만 확정한다.

- `product_family`
- `channel`
- `version`
- `build`
- `arch`
- `locale`
- 실제 type library / PIA 파일 경로

manifest에서는 이 상태를 `pending_windows_excel_capture`로 남겼다.

## Acquisition Procedure

1. Windows + Excel 설치 환경을 고정한다.
2. Office 제품군, 채널, build, arch, locale를 기록한다.
3. `oleview.exe`로 Excel type library identity와 interface surface를 확인한다.
4. vendor PIA가 있으면 우선 사용하고, 없으면 `tlbimp.exe`로 interop assembly를 생성한다.
5. reflection/metadata reader로 assembly surface를 읽어 `office-idl` schema에 맞게 정규화한다.
6. raw inspection artifact와 normalized dataset을 `specs/pinned/` 기준 경로에 저장한다.

## Required Capture Outputs

- `raw_typelib_identity.json`
- `excel_typelib_snapshot.idl`
- `excel_typelib_snapshot.odl`
- `excel_pia_identity.json`
- `excel_pia_public_surface.json`
- `office_idl_excel_om.json`

이 중 마지막 파일이 `ootd`의 canonical machine-readable OM contract가 된다.

## Tooling

- `oleview.exe`
  - 용도: 등록된 COM/type library 확인, interface/type info inspection
- `tlbimp.exe`
  - 용도: type library를 interop assembly로 변환
- `regasm.exe`
  - 용도: 필요 시 PIA 등록
- reflection/metadata reader
  - 용도: generated/vendor assembly에서 interface/member/enum/attribute metadata 추출

## Limitations

- Learn의 VBA/Interop 문서는 OM 전체 계약을 machine-readable하게 제공하지 않는다.
- PIA는 COM 원본의 projection이므로 완전한 1:1 복제가 아닐 수 있다.
- Office 채널과 build에 따라 실제 surface나 동작 차이가 날 수 있으므로 source pinning이 필수다.

## Next Step

- Windows capture runner를 만들고, [om_sources.toml](/home/seorii/dev/hancomac/ootd/specs/pinned/om_sources.toml)의 `pending` 필드를 실제 값으로 채운다.
- 이어서 `office-idl` 입력용 extractor를 구현해 `specs/pinned/office_idl_excel_om.json`을 생성한다.
