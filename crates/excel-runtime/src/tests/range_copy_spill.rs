use super::*;

fn assert_copy_failure_preserves_workbooks_and_session(
    runtime: &mut ExcelRuntime,
    source_workbook: WorkbookHandle,
    destination_workbook: WorkbookHandle,
    source: ObjectHandle,
    destination: ObjectHandle,
    protected_address: &str,
    label: &str,
) {
    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("copy session marker".to_string())],
        )
        .unwrap_or_else(|error| panic!("{label}: seed Find state: {error:?}"));
    runtime
        .dispatch_invoke(source, "Copy", &[])
        .unwrap_or_else(|error| panic!("{label}: seed copy state: {error:?}"));

    let source_before = runtime_workbook_persistence_snapshot(runtime, source_workbook);
    let destination_before = runtime_workbook_persistence_snapshot(runtime, destination_workbook);
    let source_dirty_before = runtime
        .workbook_dirty_domains(source_workbook)
        .unwrap_or_else(|error| panic!("{label}: source dirty domains: {error:?}"));
    let destination_dirty_before = runtime
        .workbook_dirty_domains(destination_workbook)
        .unwrap_or_else(|error| panic!("{label}: destination dirty domains: {error:?}"));
    let session_before = runtime_session_mutation_snapshot(runtime);

    let error = match runtime.dispatch_invoke(source, "Copy", &[OmValue::Object(destination)]) {
        Ok(value) => panic!("Range.Copy must reject {label}: {value:?}"),
        Err(error) => error,
    };

    assert_eq!(error.code, OmErrorCode::InvalidState, "{label}");
    assert!(
        error.message.contains(protected_address),
        "{label}: {error:?}",
    );
    assert_eq!(
        runtime_workbook_persistence_snapshot(runtime, source_workbook),
        source_before,
        "{label}: source workbook",
    );
    assert_eq!(
        runtime_workbook_persistence_snapshot(runtime, destination_workbook),
        destination_before,
        "{label}: destination workbook",
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(source_workbook)
            .unwrap_or_else(|error| {
                panic!("{label}: source dirty domains after failed copy: {error:?}")
            }),
        source_dirty_before,
        "{label}: source dirty domains",
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(destination_workbook)
            .unwrap_or_else(|error| {
                panic!("{label}: destination dirty domains after failed copy: {error:?}")
            }),
        destination_dirty_before,
        "{label}: destination dirty domains",
    );
    assert_eq!(
        runtime_session_mutation_snapshot(runtime),
        session_before,
        "{label}: session",
    );
}

#[test]
fn range_copy_destination_rejects_spill_after_normal_destination_atomically() {
    let (mut runtime, workbook, active_sheet, _) = runtime_with_sequence_spill();
    let source = expect_object_handle(
        runtime
            .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B1".to_string())])
            .expect("Range.Copy source"),
    );
    let destination = expect_object_handle(
        runtime
            .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("I10".to_string())])
            .expect("Range.Copy destination"),
    );

    assert_copy_failure_preserves_workbooks_and_session(
        &mut runtime,
        workbook,
        workbook,
        source,
        destination,
        "R10C10",
        "copy into spill anchor after normal destination",
    );
}

#[test]
fn range_copy_destination_rejects_spill_anchor_and_child_sources() {
    for (source_address, protected_address) in [("J10", "R10C10"), ("K10", "R10C11")] {
        let (mut runtime, workbook, active_sheet, _) = runtime_with_sequence_spill();
        let source = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Range",
                    &[OmValue::Text(source_address.to_string())],
                )
                .unwrap_or_else(|error| panic!("{source_address}: source range: {error:?}")),
        );
        let destination = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A20".to_string())])
                .unwrap_or_else(|error| panic!("{source_address}: destination range: {error:?}")),
        );

        assert_copy_failure_preserves_workbooks_and_session(
            &mut runtime,
            workbook,
            workbook,
            source,
            destination,
            protected_address,
            source_address,
        );
    }
}

#[test]
fn range_copy_cross_workbook_spill_failure_preserves_both_workbooks() {
    let (mut runtime, destination_workbook, destination_sheet, _) = runtime_with_sequence_spill();
    let source_workbook = runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open source workbook");
    let source_worksheets = expect_object_handle(
        runtime
            .dispatch_get(source_workbook.0, "Worksheets", &[])
            .expect("source Workbook.Worksheets"),
    );
    let source_sheet = expect_object_handle(
        runtime
            .dispatch_invoke(source_worksheets, "Item", &[OmValue::Number(1.0)])
            .expect("source Worksheets.Item(1)"),
    );
    let source = expect_object_handle(
        runtime
            .dispatch_invoke(source_sheet, "Range", &[OmValue::Text("A1:B1".to_string())])
            .expect("cross-workbook source range"),
    );
    let destination = expect_object_handle(
        runtime
            .dispatch_invoke(
                destination_sheet,
                "Range",
                &[OmValue::Text("I10".to_string())],
            )
            .expect("cross-workbook destination range"),
    );

    assert_copy_failure_preserves_workbooks_and_session(
        &mut runtime,
        source_workbook,
        destination_workbook,
        source,
        destination,
        "R10C10",
        "cross-workbook copy into spill anchor",
    );
}
