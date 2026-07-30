use super::*;

fn assert_fill_failure_preserves_workbook_and_session(
    runtime: &mut ExcelRuntime,
    workbook: WorkbookHandle,
    target: ObjectHandle,
    member: &str,
    protected_address: &str,
    label: &str,
) {
    let before = runtime_workbook_persistence_snapshot(runtime, workbook);
    let dirty_before = runtime
        .workbook_dirty_domains(workbook)
        .unwrap_or_else(|error| panic!("{label}: workbook dirty domains: {error:?}"));
    let session_before = runtime_session_mutation_snapshot(runtime);

    let error = match runtime.dispatch_invoke(target, member, &[]) {
        Ok(value) => panic!("Range.{member} must reject {label}: {value:?}"),
        Err(error) => error,
    };

    assert_eq!(error.code, OmErrorCode::InvalidState, "{label}");
    assert!(
        error.message.contains(protected_address),
        "{label}: {error:?}",
    );
    assert_eq!(
        runtime_workbook_persistence_snapshot(runtime, workbook),
        before,
        "{label}",
    );
    assert_eq!(
        runtime
            .workbook_dirty_domains(workbook)
            .unwrap_or_else(|error| {
                panic!("{label}: dirty domains after failed fill: {error:?}")
            }),
        dirty_before,
        "{label}",
    );
    assert_eq!(
        runtime_session_mutation_snapshot(runtime),
        session_before,
        "{label}",
    );
}

#[test]
fn range_fill_spill_failure_preserves_workbook_state_and_dirty_domains() {
    for (member, source_address, source_rows, source_cols, target_address) in [
        ("FillDown", "J8:K8", 1, 2, "J8:K11"),
        ("FillRight", "H10:H11", 2, 1, "H10:K11"),
        ("FillUp", "J12:K12", 1, 2, "J9:K12"),
        ("FillLeft", "L10:L11", 2, 1, "I10:L11"),
    ] {
        let (mut runtime, workbook, active_sheet, _) = runtime_with_sequence_spill();
        let source = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Range",
                    &[OmValue::Text(source_address.to_string())],
                )
                .unwrap_or_else(|error| panic!("{member}: source range: {error:?}")),
        );
        runtime
            .dispatch_set(
                source,
                "Value2",
                OmValue::Array(
                    OmArray::new(
                        source_rows,
                        source_cols,
                        vec![
                            OmValue::Text(format!("{member} source"));
                            (source_rows * source_cols) as usize
                        ],
                    )
                    .unwrap_or_else(|error| panic!("{member}: source values: {error:?}")),
                ),
                &[],
            )
            .unwrap_or_else(|error| panic!("{member}: set source values: {error:?}"));
        let target = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Range",
                    &[OmValue::Text(target_address.to_string())],
                )
                .unwrap_or_else(|error| panic!("{member}: target range: {error:?}")),
        );
        runtime
            .dispatch_invoke(
                target,
                "Find",
                &[OmValue::Text("fill session marker".to_string())],
            )
            .unwrap_or_else(|error| panic!("{member}: seed Find state: {error:?}"));
        runtime
            .dispatch_invoke(source, "Copy", &[])
            .unwrap_or_else(|error| panic!("{member}: seed clipboard state: {error:?}"));
        assert_fill_failure_preserves_workbook_and_session(
            &mut runtime,
            workbook,
            target,
            member,
            "R10C10",
            member,
        );
    }
}

#[test]
fn range_fill_rejects_spill_sources_without_destination_overlap() {
    for (member, target_address, protected_address) in [
        ("FillDown", "J11:J13", "R11C10"),
        ("FillRight", "K10:M10", "R10C11"),
        ("FillUp", "J8:J10", "R10C10"),
        ("FillLeft", "H10:J10", "R10C10"),
    ] {
        let (mut runtime, workbook, active_sheet, _) = runtime_with_sequence_spill();
        let target = expect_object_handle(
            runtime
                .dispatch_invoke(
                    active_sheet,
                    "Range",
                    &[OmValue::Text(target_address.to_string())],
                )
                .unwrap_or_else(|error| panic!("{member}: target range: {error:?}")),
        );
        assert_fill_failure_preserves_workbook_and_session(
            &mut runtime,
            workbook,
            target,
            member,
            protected_address,
            member,
        );
    }
}

#[test]
fn range_multi_area_fill_spill_failure_rolls_back_earlier_areas() {
    let (mut runtime, workbook, active_sheet, _) = runtime_with_sequence_spill();
    for (address, value) in [("A1", 7.0), ("J9", 8.0)] {
        let cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text(address.to_string())])
                .unwrap_or_else(|error| panic!("{address}: source cell: {error:?}")),
        );
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(value), &[])
            .unwrap_or_else(|error| panic!("{address}: source value: {error:?}"));
    }
    let target = expect_object_handle(
        runtime
            .dispatch_invoke(
                active_sheet,
                "Range",
                &[OmValue::Text("A1:A2,J9:J10".to_string())],
            )
            .expect("multi-area FillDown target"),
    );
    assert_fill_failure_preserves_workbook_and_session(
        &mut runtime,
        workbook,
        target,
        "FillDown",
        "R10C10",
        "multi-area FillDown",
    );
}

#[test]
fn range_multi_area_fill_preserves_sequential_overlap_semantics() {
    let (mut runtime, _, active_sheet, _) = runtime_with_sequence_spill();
    for (address, value) in [("A20", 1.0), ("A21", 2.0)] {
        let cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text(address.to_string())])
                .unwrap_or_else(|error| panic!("{address}: seed range: {error:?}")),
        );
        runtime
            .dispatch_set(cell, "Value2", OmValue::Number(value), &[])
            .unwrap_or_else(|error| panic!("{address}: seed value: {error:?}"));
    }
    let target = expect_object_handle(
        runtime
            .dispatch_invoke(
                active_sheet,
                "Range",
                &[OmValue::Text("A20:A21,A21:A22".to_string())],
            )
            .expect("overlapping multi-area FillDown target"),
    );

    runtime
        .dispatch_invoke(target, "FillDown", &[])
        .expect("overlapping multi-area FillDown");

    for address in ["A21", "A22"] {
        let cell = expect_object_handle(
            runtime
                .dispatch_invoke(active_sheet, "Range", &[OmValue::Text(address.to_string())])
                .unwrap_or_else(|error| panic!("{address}: result range: {error:?}")),
        );
        assert_eq!(
            runtime
                .dispatch_get(cell, "Value2", &[])
                .unwrap_or_else(|error| panic!("{address}: result value: {error:?}")),
            OmValue::Number(1.0),
            "{address}",
        );
    }
}
