# Windows Capture Runner

`ootd`는 이 Linux workspace에서 Excel COM type library를 직접 추출할 수 없다.
따라서 실제 OM source capture는 별도의 Windows + Excel 환경에서 실행하는 runner가 필요하다.

## Purpose

- Excel type library와 PIA를 machine-readable artifact로 고정한다.
- Windows host에서 실행할 수 있는 invocation plan과 script emission boundary를 고정한다.
- 이후 `office-idl` extractor가 consume할 raw input을 일관된 위치와 파일명으로 남긴다.

## Inputs

- [om_sources.toml](/home/seorii/dev/hancomac/ootd/specs/pinned/om_sources.toml)
- [windows_capture.template.toml](/home/seorii/dev/hancomac/ootd/specs/pinned/windows_capture.template.toml)

## Expected Steps

1. Windows machine에서 `windows_capture.template.toml`을 실제 값으로 채운다.
2. runner는 template를 읽고 invocation plan을 계산한다.
3. 필요하면 script 또는 command batch를 emission-only 방식으로 생성한다.
4. Windows host에서 `oleview.exe`로 Excel type library identity를 확인한다.
5. vendor PIA가 있으면 그대로 사용하고, 없으면 `tlbimp.exe`로 interop assembly를 생성한다.
6. identity JSON, reconstructed IDL/ODL, public surface JSON을 지정된 output dir에 저장한다.
7. `office-codegen` generator가 capture bundle directory를 읽어 canonical `office-idl` JSON output을 생성한다.

## Required Outputs

- `capture_manifest.json`
- `capture.log`
- `raw_typelib_identity.json`
- `excel_typelib_snapshot.idl`
- `excel_typelib_snapshot.odl`
- `excel_pia_identity.json`
- `excel_pia_public_surface.json`
- `output_checksums.json`

## Output Layout

`output_dir` 아래에는 아래와 같은 하위 구조를 사용한다.

- `manifest/capture_manifest.json`
- `manifest/direct_exec_status.json`
- `manifest/direct_exec_status.template.json`
- `manifest/output_checksums.json`
- `manifest/execution_receipt.json`
- `manifest/execution_receipt.template.json`
- `logs/capture.log`
- `scripts/capture.ps1`
- `scripts/run_capture.cmd`
- `raw/raw_typelib_identity.json`
- `raw/excel_pia_identity.json`
- `snapshots/excel_typelib_snapshot.idl`
- `snapshots/excel_typelib_snapshot.odl`
- `snapshots/excel_pia_public_surface.json`

template에 plain filename만 들어 있더라도 runner는 artifact kind에 따라 `raw/` 또는 `snapshots/` 하위 경로로 정규화해야 한다.
PowerShell emission이 파일로 materialize되는 경우에는 `scripts/capture.ps1`을 사용한다.
direct-exec launcher는 `scripts/run_capture.cmd`를 사용한다.
Windows host launcher는 `manifest/direct_exec_status.json`을 launcher status result로 쓰고, template helper는 `manifest/direct_exec_status.template.json`을 제공한다.
capture script는 `manifest/execution_receipt.json`을 실제 capture receipt로 쓰고, template helper는 `manifest/execution_receipt.template.json`을 제공한다.

`office_idl_excel_om.json`은 Step 3의 canonical JSON writer가 생성하는 파생 산출물이다.
Step 2 runner는 이 파일을 생성하지 않고, 존재하더라도 덮어쓰지 않아야 한다.

### Direct-Exec Launcher Slice

launcher generation은 `scripts/run_capture.cmd`를 기준 경로로 삼는다.
launcher status template/result는 `manifest/direct_exec_status.template.json`과 `manifest/direct_exec_status.json`이다.
capture receipt template/result는 `manifest/execution_receipt.template.json`과 `manifest/execution_receipt.json`이다.
Step 2.3는 launcher generation과 template materialization만 닫는다.
launcher generation 이후에도 남는 작업은 Windows process spawn, exit status capture, host identity/tool result population, final manifest completion orchestration이었다.
이 실행 경계는 Step 2.4에서 library/CLI path로 구현한다.

`excel_pia_public_surface.json`의 synthetic shape 예시는
[excel_pia_public_surface.template.json](/home/seorii/dev/hancomac/ootd/specs/pinned/excel_pia_public_surface.template.json)에 고정했다.
이 capture는 getter/setter를 별도 member로 담고, extractor가 이를 `office-idl` property surface로 병합한다.

`raw_typelib_identity.json`의 synthetic shape 예시는
[raw_typelib_identity.template.json](/home/seorii/dev/hancomac/ootd/specs/pinned/raw_typelib_identity.template.json)에 고정했다.
이 capture는 type library GUID, interface IID, coclass CLSID를 담고, extractor는 PIA surface와 library/version을 대조한다.

## Current Boundary

현재 repo에는 [office-capture](/home/seorii/dev/hancomac/ootd/crates/office-capture/src/lib.rs) crate가 추가되어 있다.
이 crate는 pinned template load, Windows path normalization, fixed output layout resolution, raw bundle writer, manifest/checksum writing, invocation plan 계산, PowerShell script emission을 담당한다.
현재 script emission은 interop-derived `raw_typelib_identity.json`, `excel_pia_identity.json`, `excel_pia_public_surface.json` 자동 생성을 준비하고, `oleview` 기반 IDL/ODL snapshot export는 manual step으로 남겨 둔다.
이 문서가 다루는 Step 2 경계는 invocation plan, script emission, execution bundle materialization, receipt-driven bundle completion까지다.
현재 `office-capture`는 capture script를 `scripts/capture.ps1`로, direct-exec launcher를 `scripts/run_capture.cmd`로 materialize하고 `manifest/execution_plan.json`, `manifest/direct_exec_status.template.json`, `manifest/execution_receipt.template.json`을 함께 쓸 수 있으며, `manifest/execution_receipt.json`이 있으면 completed manifest/checksum을 다시 닫을 수 있다.
capture payload contract는 `raw_typelib_identity.json`, `excel_typelib_snapshot.idl`, `excel_typelib_snapshot.odl`, `excel_pia_identity.json`, `excel_pia_public_surface.json`의 5개 파일명으로 고정되어 있으며, plan summary, `execution_plan.json`, `execution_receipt.template.json`, completed `capture_manifest.json`에서 중복 없는 같은 목록을 노출한다.
completed `capture_manifest.json`의 `writableOutputs`는 이 5개 payload에 대응하는 canonical logical key set만 포함해야 하며, 각 값은 canonical relative output path(`raw/...` 또는 `snapshots/...`)로 끝나야 한다.
`office-codegen`은 예상 외 writable output key와 wrong-directory writable output path를 모두 거부한다.
completion 단계에서는 receipt가 `expectedCaptureOutputs`를 명시한 경우 plan의 payload contract와 중복 없이 정확히 일치해야 하며, legacy receipt처럼 해당 field가 없는 경우에는 호환을 위해 허용하되 manifest-level `expectedCaptureOutputs`는 계속 plan 기준으로 기록한다.
modern receipt는 `commandResults`와 `manualStepResults`의 각 status가 `completed`여야 한다.
또한 `expectedCaptureOutputs`가 있는 receipt에서는 조건 없는 command result와 manual step result 이름이 중복 없이 `execution_plan.json`의 plan과 일치해야 하며, 조건부 fallback command는 실행되지 않을 수 있어 optional로 취급한다.
알 수 없는 command result, 누락된 필수 command/manual step, pending/failed status는 completion error이며, 이 경우 completed manifest/checksum을 쓰지 않는다.
`output_checksums.json`은 manifest-level `expectedCaptureOutputs`의 5개 payload 파일명과 정확히 일치하는 bundle-relative output path checksum set을 제공해야 한다.
`office-codegen`은 각 expected payload filename이 checksum key에 정확히 하나만 나타나고 canonical relative path(`raw/...` 또는 `snapshots/...`)와 일치하는지도 재검증한다.
추가로 `run_execution_bundle` library path와 `--run-execution-bundle DIR` CLI mode를 통해 실제 Windows host에서 launcher를 spawn하고 `manifest/direct_exec_status.json`을 읽은 뒤 final manifest/checksum completion까지 이어지는 direct-exec orchestration을 제공한다.
non-Windows host에서는 이 path를 명시적으로 거부하고, materialized launcher/script 누락도 preflight error로 바로 반환한다.
`office-codegen`는 capture bundle directory를 입력으로 받아 canonical `office-idl` JSON output을 쓰는 generator path를 제공하고, synthetic bundle regression과 capture bundle validator로 Step 3 contract를 고정한다.
completed manifest가 있는 bundle에서는 `office-codegen`도 expected payload 목록, writable output basename, checksum coverage, checksum-listed file existence, SHA-256 digest match, embedded receipt status/result contract를 재검증한다.

Step 1의 확정 범위는 member-level `dispid`, getter/setter origin, type alias info, interface/class inheritance metadata까지다.
Step 2는 Windows machine에서 실제 artifact를 쓰는 contract와 layout을 고정하고, canonical `office-idl` 생성은 범위 밖으로 유지한다.
Step 2.1은 contract, failure mode, output layout, downstream boundary를 문서와 테스트로 고정한다.
Step 2.2는 emitted script file, execution bundle materialization, execution receipt file, receipt/manifest completion contract를 고정한다.
Step 2.3는 generated launcher path, launcher status template/result path, capture receipt path, 그리고 template materialization boundary를 고정한다.
Step 2.4는 direct-exec wrapper가 담당하는 process spawn, exit/status orchestration, host identity/tool result population boundary를 구현하고, 현재 남는 일은 실제 Windows host validation과 manual `oleview` workflow 정리다.
즉 `office_idl_excel_om.json`은 Step 3 generator output이며, 아직 실제 Windows-captured bundle을 pinning하는 작업은 남아 있다.

## Remaining Implementation

- 실제 Windows host에서 `--run-execution-bundle DIR`를 실행한 end-to-end validation fixture를 수집한다.
- 실제 Windows-captured bundle을 pinning해 canonical dataset을 대체한다.
- `oleview` manual step을 direct-exec wrapper와 어떻게 결합할지 정리한다.
