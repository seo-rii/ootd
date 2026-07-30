use super::*;

fn assert_cut_failure_preserves_workbooks_and_session(
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
            &[OmValue::Text("cut session marker".to_string())],
        )
        .unwrap_or_else(|error| panic!("{label}: seed Find state: {error:?}"));
    runtime
        .dispatch_invoke(source, "Cut", &[])
        .unwrap_or_else(|error| panic!("{label}: seed cut state: {error:?}"));

    let source_before = runtime_workbook_persistence_snapshot(runtime, source_workbook);
    let destination_before = runtime_workbook_persistence_snapshot(runtime, destination_workbook);
    let source_dirty_before = runtime
        .workbook_dirty_domains(source_workbook)
        .unwrap_or_else(|error| panic!("{label}: source dirty domains: {error:?}"));
    let destination_dirty_before = runtime
        .workbook_dirty_domains(destination_workbook)
        .unwrap_or_else(|error| panic!("{label}: destination dirty domains: {error:?}"));
    let session_before = runtime_session_mutation_snapshot(runtime);

    let error = match runtime.dispatch_invoke(source, "Cut", &[OmValue::Object(destination)]) {
        Ok(value) => panic!("Range.Cut must reject {label}: {value:?}"),
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
                panic!("{label}: source dirty domains after failed cut: {error:?}")
            }),
        source_dirty_before,
        "{label}: source dirty domains",
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(destination_workbook)
            .unwrap_or_else(|error| {
                panic!("{label}: destination dirty domains after failed cut: {error:?}")
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
fn range_cut_destination_rejects_spill_after_normal_destination_atomically() {
    let (mut runtime, workbook, active_sheet, _) = runtime_with_sequence_spill();
    let source = expect_object_handle(
        runtime
            .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B1".to_string())])
            .expect("Range.Cut source"),
    );
    let destination = expect_object_handle(
        runtime
            .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("I10".to_string())])
            .expect("Range.Cut destination"),
    );

    assert_cut_failure_preserves_workbooks_and_session(
        &mut runtime,
        workbook,
        workbook,
        source,
        destination,
        "R10C10",
        "cut into spill anchor after normal destination",
    );
}

#[test]
fn range_paste_special_cut_failure_preserves_workbook_and_session() {
    let (mut runtime, workbook, active_sheet, _) = runtime_with_sequence_spill();
    let source = expect_object_handle(
        runtime
            .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("A1:B1".to_string())])
            .expect("Range.Cut source"),
    );
    let destination = expect_object_handle(
        runtime
            .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("I10".to_string())])
            .expect("Range.PasteSpecial destination"),
    );
    runtime
        .dispatch_invoke(
            source,
            "Find",
            &[OmValue::Text("cut paste session marker".to_string())],
        )
        .expect("seed Find state");
    runtime
        .dispatch_invoke(source, "Cut", &[])
        .expect("seed cut clipboard");

    let workbook_before = runtime_workbook_persistence_snapshot(&runtime, workbook);
    let dirty_before = runtime
        .workbook_dirty_domains(workbook)
        .expect("dirty domains before failed cut paste");
    let session_before = runtime_session_mutation_snapshot(&runtime);

    let error = runtime
        .dispatch_invoke(destination, "PasteSpecial", &[])
        .expect_err("Range.PasteSpecial must reject cut into a spill anchor");

    assert_eq!(error.code, OmErrorCode::InvalidState);
    assert!(error.message.contains("R10C10"), "{error:?}");
    assert_eq!(
        runtime_workbook_persistence_snapshot(&runtime, workbook),
        workbook_before,
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(workbook)
            .expect("dirty domains after failed cut paste"),
        dirty_before,
    );
    assert_eq!(runtime_session_mutation_snapshot(&runtime), session_before);
}

#[test]
fn range_cut_destination_rejects_spill_anchor_and_child_sources() {
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

        assert_cut_failure_preserves_workbooks_and_session(
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
fn range_cut_cross_sheet_spill_failure_preserves_workbook() {
    let (mut runtime, workbook, destination_sheet, _) = runtime_with_sequence_spill();
    let worksheets = expect_object_handle(
        runtime
            .dispatch_get(workbook.0, "Worksheets", &[])
            .expect("Workbook.Worksheets"),
    );
    let source_sheet = expect_object_handle(
        runtime
            .dispatch_invoke(worksheets, "Add", &[])
            .expect("Worksheets.Add source sheet"),
    );
    let source = expect_object_handle(
        runtime
            .dispatch_invoke(source_sheet, "Range", &[OmValue::Text("A1:B1".to_string())])
            .expect("cross-sheet source range"),
    );
    runtime
        .dispatch_set(
            source,
            "Value2",
            OmValue::Array(
                OmArray::new(1, 2, vec![OmValue::Number(7.0), OmValue::Number(8.0)])
                    .expect("cross-sheet source values"),
            ),
            &[],
        )
        .expect("set cross-sheet source values");
    let destination = expect_object_handle(
        runtime
            .dispatch_invoke(
                destination_sheet,
                "Range",
                &[OmValue::Text("I10".to_string())],
            )
            .expect("cross-sheet destination range"),
    );

    assert_cut_failure_preserves_workbooks_and_session(
        &mut runtime,
        workbook,
        workbook,
        source,
        destination,
        "R10C10",
        "cross-sheet cut into spill anchor",
    );
}

#[test]
fn range_cut_cross_workbook_spill_failure_preserves_both_workbooks() {
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

    assert_cut_failure_preserves_workbooks_and_session(
        &mut runtime,
        source_workbook,
        destination_workbook,
        source,
        destination,
        "R10C10",
        "cross-workbook cut into spill anchor",
    );
}

#[test]
fn range_cut_destination_preserves_same_sheet_overlap_semantics() {
    let mut runtime = ExcelRuntime::new();
    runtime
        .open_workbook(OpenWorkbookSpec {
            bytes: synthetic_workbook_bytes(),
            format_hint: Some(FileFormat::Xlsx),
            profile: ExcelProfile::Excel365,
            read_only: false,
        })
        .expect("open workbook");
    let active_sheet = expect_object_handle(
        runtime
            .dispatch_get(runtime.root_application(), "ActiveSheet", &[])
            .expect("ActiveSheet"),
    );
    let source = expect_object_handle(
        runtime
            .dispatch_invoke(
                active_sheet,
                "Range",
                &[OmValue::Text("A20:B20".to_string())],
            )
            .expect("overlapping cut source"),
    );
    runtime
        .dispatch_set(
            source,
            "Value2",
            OmValue::Array(
                OmArray::new(1, 2, vec![OmValue::Number(1.0), OmValue::Number(2.0)])
                    .expect("overlapping cut source values"),
            ),
            &[],
        )
        .expect("set overlapping cut source values");
    let destination = expect_object_handle(
        runtime
            .dispatch_invoke(active_sheet, "Range", &[OmValue::Text("B20".to_string())])
            .expect("overlapping cut destination"),
    );

    runtime
        .dispatch_invoke(source, "Cut", &[OmValue::Object(destination)])
        .expect("overlapping Range.Cut Destination");

    for (address, expected) in [
        ("A20", OmValue::Empty),
        ("B20", OmValue::Number(1.0)),
        ("C20", OmValue::Number(2.0)),
    ] {
        let cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text(address.to_string())])
                .unwrap_or_else(|error| panic!("{address}: result range: {error:?}")),
        );
        assert_eq!(
            runtime
                .dispatch_get(cell, "Value2", &[])
                .unwrap_or_else(|error| panic!("{address}: result value: {error:?}")),
            expected,
            "{address}",
        );
    }
}
