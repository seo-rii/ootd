//! A1/R1C1 reference parsing, formatting, conversion, and shifting.

use super::*;

pub(super) fn parse_formula_reference_text(
    input: &str,
    default_sheet_id: SheetId,
    state: &WorkbookState,
) -> Result<FormulaReference, FormulaEvalError> {
    let input = input.trim().strip_prefix('=').unwrap_or(input.trim());
    let mut areas = Vec::new();
    let mut explicit_area_count = 0usize;
    for part in split_reference_union_text(input).map_err(|_| FormulaEvalError::Ref)? {
        let reference = parse_formula_reference_area_text(part, default_sheet_id, state)?;
        explicit_area_count += reference.len();
        areas.extend(reference.areas().iter().copied());
    }
    FormulaReference::with_explicit_area_count(explicit_area_count, areas)
}

pub(super) fn parse_formula_reference_area_text(
    input: &str,
    default_sheet_id: SheetId,
    state: &WorkbookState,
) -> Result<FormulaReference, FormulaEvalError> {
    let input = input.trim();
    if input.is_empty() {
        return Err(FormulaEvalError::Ref);
    }

    let (sheet_ids, reference_text) = if input.starts_with('\'') {
        let mut cursor = 1usize;
        let mut sheet_name_or_span = String::new();
        loop {
            let Some(ch) = input[cursor..].chars().next() else {
                return Err(FormulaEvalError::Ref);
            };
            if ch == '\'' {
                let next_cursor = cursor + ch.len_utf8();
                if input[next_cursor..].starts_with('\'') {
                    sheet_name_or_span.push('\'');
                    cursor = next_cursor + 1;
                    continue;
                }
                if !input[next_cursor..].starts_with('!') {
                    return Err(FormulaEvalError::Ref);
                }
                let sheet_ids =
                    resolve_formula_sheet_name_or_span(state, sheet_name_or_span.as_str())?;
                break (sheet_ids, &input[next_cursor + 1..]);
            }
            sheet_name_or_span.push(ch);
            cursor += ch.len_utf8();
        }
    } else if let Some((sheet_name, reference_text)) = input.split_once('!') {
        let sheet_ids = resolve_formula_sheet_name_or_span(state, sheet_name.trim())?;
        (sheet_ids, reference_text)
    } else {
        (vec![default_sheet_id], input)
    };

    let rect = parse_rect_a1(reference_text.trim()).map_err(|_| FormulaEvalError::Ref)?;
    let areas = sheet_ids
        .into_iter()
        .map(|sheet_id| (sheet_id, rect))
        .collect::<Vec<_>>();
    FormulaReference::with_explicit_area_count(1, areas)
}

pub(super) fn resolve_formula_sheet_name_or_span(
    state: &WorkbookState,
    name_or_span: &str,
) -> Result<Vec<SheetId>, FormulaEvalError> {
    if let Some((start_sheet, end_sheet)) = name_or_span.split_once(':') {
        let start =
            resolve_formula_sheet_name(state, start_sheet.trim()).ok_or(FormulaEvalError::Ref)?;
        let end =
            resolve_formula_sheet_name(state, end_sheet.trim()).ok_or(FormulaEvalError::Ref)?;
        return formula_sheets_in_3d_span(state, start, end);
    }
    resolve_formula_sheet_name(state, name_or_span)
        .map(|sheet_id| vec![sheet_id])
        .ok_or(FormulaEvalError::Ref)
}

pub(super) fn resolve_formula_sheet_name(state: &WorkbookState, name: &str) -> Option<SheetId> {
    state
        .worksheets()
        .iter()
        .find(|worksheet| worksheet.name.eq_ignore_ascii_case(name))
        .map(|worksheet| worksheet.id)
}

pub(super) fn formula_sheets_in_3d_span(
    state: &WorkbookState,
    start: SheetId,
    end: SheetId,
) -> Result<Vec<SheetId>, FormulaEvalError> {
    let start_index = state
        .worksheets()
        .iter()
        .position(|worksheet| worksheet.id == start)
        .ok_or(FormulaEvalError::Ref)?;
    let end_index = state
        .worksheets()
        .iter()
        .position(|worksheet| worksheet.id == end)
        .ok_or(FormulaEvalError::Ref)?;
    let (first, last) = if start_index <= end_index {
        (start_index, end_index)
    } else {
        (end_index, start_index)
    };
    Ok(state.worksheets()[first..=last]
        .iter()
        .map(|worksheet| worksheet.id)
        .collect())
}

pub(crate) fn parse_rect_a1(input: &str) -> OmResult<Rect> {
    let input = input.trim();
    let mut parts = input.split(':');
    let first = parts
        .next()
        .ok_or_else(|| OmError::parse("empty A1 reference"))?;
    let second = parts.next();
    if parts.next().is_some() {
        return Err(OmError::parse("A1 range contains too many ':' separators"));
    }
    let first = parse_a1_endpoint(first)?;
    let Some(second) = second else {
        return match first {
            A1Endpoint::Cell(row, col) => Ok(Rect::single_cell(row, col)),
            A1Endpoint::Row(_) | A1Endpoint::Column(_) => Err(OmError::parse(format!(
                "A1 range {input:?} must use ':' for whole-row or whole-column selectors"
            ))),
        };
    };
    let second = parse_a1_endpoint(second)?;
    match (first, second) {
        (A1Endpoint::Cell(first_row, first_col), A1Endpoint::Cell(second_row, second_col)) => {
            Ok(Rect {
                row_first: first_row.min(second_row),
                row_last: first_row.max(second_row),
                col_first: first_col.min(second_col),
                col_last: first_col.max(second_col),
            })
        }
        (A1Endpoint::Row(first_row), A1Endpoint::Row(second_row)) => Ok(Rect {
            row_first: first_row.min(second_row),
            row_last: first_row.max(second_row),
            col_first: 1,
            col_last: EXCEL_MAX_COLUMN_INDEX,
        }),
        (A1Endpoint::Column(first_col), A1Endpoint::Column(second_col)) => Ok(Rect {
            row_first: 1,
            row_last: EXCEL_MAX_ROW_INDEX,
            col_first: first_col.min(second_col),
            col_last: first_col.max(second_col),
        }),
        _ => Err(OmError::parse(format!(
            "A1 range {input:?} cannot mix cell, row, and column endpoints"
        ))),
    }
}

pub(crate) fn split_reference_union_text(input: &str) -> OmResult<Vec<&str>> {
    let input = input.trim();
    if input.is_empty() {
        return Err(OmError::invalid_argument("range reference text is empty"));
    }

    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut in_quote = false;
    let mut in_brackets = false;
    let mut chars = input.char_indices().peekable();
    while let Some((index, ch)) = chars.next() {
        match ch {
            '\'' => {
                if in_quote && chars.peek().is_some_and(|(_, next)| *next == '\'') {
                    chars.next();
                } else {
                    in_quote = !in_quote;
                }
            }
            '[' if !in_quote => in_brackets = true,
            ']' if !in_quote => in_brackets = false,
            ',' if !in_quote && !in_brackets => {
                let part = input[start..index].trim();
                if part.is_empty() {
                    return Err(OmError::invalid_argument(
                        "multi-area range references cannot contain empty areas",
                    ));
                }
                parts.push(part);
                start = index + ch.len_utf8();
            }
            _ => {}
        }
    }

    if in_quote {
        return Err(OmError::invalid_argument(
            "range references use unmatched worksheet quoting",
        ));
    }
    if in_brackets {
        return Err(OmError::invalid_argument(
            "range references use invalid workbook qualification",
        ));
    }

    let part = input[start..].trim();
    if part.is_empty() {
        return Err(OmError::invalid_argument(
            "multi-area range references cannot contain empty areas",
        ));
    }
    parts.push(part);
    Ok(parts)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum A1Endpoint {
    Cell(u32, u32),
    Row(u32),
    Column(u32),
}

pub(super) fn parse_a1_endpoint(input: &str) -> OmResult<A1Endpoint> {
    let normalized = input.trim().replace('$', "");
    if normalized.is_empty() {
        return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
    }
    if normalized.chars().all(|ch| ch.is_ascii_digit()) {
        let row = normalized
            .parse::<u32>()
            .map_err(|_| OmError::parse(format!("invalid row index in {input:?}")))?;
        if row == 0 || row > EXCEL_MAX_ROW_INDEX {
            return Err(OmError::parse(format!(
                "A1 row reference {input:?} is outside worksheet bounds"
            )));
        }
        return Ok(A1Endpoint::Row(row));
    }
    if normalized.chars().all(|ch| ch.is_ascii_alphabetic()) {
        return Ok(A1Endpoint::Column(parse_column_label_a1(input)?));
    }

    let (row, col) = parse_cell_a1(input)?;
    Ok(A1Endpoint::Cell(row, col))
}

pub(super) fn parse_column_label_a1(input: &str) -> OmResult<u32> {
    let normalized = input.trim().replace('$', "").to_ascii_uppercase();
    if normalized.is_empty() || !normalized.chars().all(|ch| ch.is_ascii_alphabetic()) {
        return Err(OmError::parse(format!(
            "invalid A1 column reference {input:?}"
        )));
    }
    let mut col = 0u32;
    for ch in normalized.bytes() {
        col = col
            .checked_mul(26)
            .and_then(|value| value.checked_add((ch - b'A' + 1) as u32))
            .ok_or_else(|| OmError::parse("column index overflow"))?;
    }
    if col == 0 || col > EXCEL_MAX_COLUMN_INDEX {
        return Err(OmError::parse(format!(
            "A1 column reference {input:?} is outside worksheet bounds"
        )));
    }
    Ok(col)
}

pub(super) fn parse_cell_a1(input: &str) -> OmResult<(u32, u32)> {
    let trimmed = input.trim().trim_matches('$');
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in trimmed.chars() {
        if ch == '$' {
            continue;
        }
        if ch.is_ascii_alphabetic() && digits.is_empty() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        } else {
            return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
    }

    let mut col = 0u32;
    for ch in letters.bytes() {
        col = col
            .checked_mul(26)
            .and_then(|value| value.checked_add((ch - b'A' + 1) as u32))
            .ok_or_else(|| OmError::parse("column index overflow"))?;
    }
    let row = digits
        .parse::<u32>()
        .map_err(|_| OmError::parse(format!("invalid row index in {input:?}")))?;
    if row == 0 || col == 0 {
        return Err(OmError::parse(format!("invalid A1 reference {input:?}")));
    }
    if row > EXCEL_MAX_ROW_INDEX || col > EXCEL_MAX_COLUMN_INDEX {
        return Err(OmError::parse(format!(
            "A1 reference {input:?} is outside worksheet bounds"
        )));
    }
    Ok((row, col))
}

pub(crate) fn format_rect_address_with_flags(
    rect: Rect,
    row_absolute: bool,
    column_absolute: bool,
) -> String {
    let spans_all_rows = rect.row_first == 1 && rect.row_last == EXCEL_MAX_ROW_INDEX;
    let spans_all_columns = rect.col_first == 1 && rect.col_last == EXCEL_MAX_COLUMN_INDEX;
    if spans_all_rows && !spans_all_columns {
        let start = format_column_address(rect.col_first, column_absolute);
        if rect.col_first == rect.col_last {
            return format!("{start}:{start}");
        }
        let end = format_column_address(rect.col_last, column_absolute);
        return format!("{start}:{end}");
    }
    if spans_all_columns && !spans_all_rows {
        let start = format_row_address(rect.row_first, row_absolute);
        if rect.row_first == rect.row_last {
            return format!("{start}:{start}");
        }
        let end = format_row_address(rect.row_last, row_absolute);
        return format!("{start}:{end}");
    }

    let start = format_cell_address(
        rect.row_first,
        rect.col_first,
        row_absolute,
        column_absolute,
    );
    if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
        start
    } else {
        let end = format_cell_address(rect.row_last, rect.col_last, row_absolute, column_absolute);
        format!("{start}:{end}")
    }
}

pub(super) fn format_cell_address(
    row: u32,
    col: u32,
    row_absolute: bool,
    column_absolute: bool,
) -> String {
    format!(
        "{}{}",
        format_column_address(col, column_absolute),
        format_row_address(row, row_absolute)
    )
}

pub(super) fn format_column_address(col: u32, column_absolute: bool) -> String {
    let mut address = String::new();
    if column_absolute {
        address.push('$');
    }
    address.push_str(&column_to_letters(col));
    address
}

pub(super) fn format_row_address(row: u32, row_absolute: bool) -> String {
    let mut address = String::new();
    if row_absolute {
        address.push('$');
    }
    address.push_str(&row.to_string());
    address
}

pub(crate) fn format_external_address_qualifier(
    workbook_name: &str,
    worksheet_name: &str,
) -> String {
    let qualifier = format!("[{workbook_name}]{worksheet_name}");
    if excel_reference_qualifier_needs_quotes(workbook_name)
        || excel_reference_qualifier_needs_quotes(worksheet_name)
    {
        format!("'{}'!", qualifier.replace('\'', "''"))
    } else {
        format!("{qualifier}!")
    }
}

pub(super) fn excel_reference_qualifier_needs_quotes(value: &str) -> bool {
    value.is_empty()
        || value
            .chars()
            .any(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.')
}

pub(super) fn parse_a1_axis_reference_to_r1c1(
    formula: &str,
    start: usize,
    base_row: u32,
    base_col: u32,
) -> Option<(String, usize)> {
    let bytes = formula.as_bytes();
    let parse_column_endpoint = |mut cursor: usize| -> Option<(u32, bool, usize)> {
        let absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
            true
        } else {
            false
        };
        let letters_start = cursor;
        while cursor < bytes.len()
            && bytes[cursor].is_ascii_alphabetic()
            && cursor - letters_start < 3
        {
            cursor += 1;
        }
        if letters_start == cursor || (cursor < bytes.len() && bytes[cursor].is_ascii_alphabetic())
        {
            return None;
        }
        let col = parse_column_label_a1(&formula[letters_start..cursor]).ok()?;
        Some((col, absolute, cursor))
    };
    let parse_row_endpoint = |mut cursor: usize| -> Option<(u32, bool, usize)> {
        let absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
            true
        } else {
            false
        };
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if digits_start == cursor {
            return None;
        }
        let row = formula[digits_start..cursor].parse::<u32>().ok()?;
        if row == 0 || row > EXCEL_MAX_ROW_INDEX {
            return None;
        }
        Some((row, absolute, cursor))
    };

    if let Some((first_col, first_absolute, colon_index)) = parse_column_endpoint(start)
        && bytes.get(colon_index) == Some(&b':')
        && let Some((second_col, second_absolute, next_index)) =
            parse_column_endpoint(colon_index + 1)
        && formula[next_index..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
    {
        let start_ref = format_r1c1_column_reference(first_col, first_absolute, base_col);
        let end_ref = format_r1c1_column_reference(second_col, second_absolute, base_col);
        return Some((
            if start_ref == end_ref {
                start_ref
            } else {
                format!("{start_ref}:{end_ref}")
            },
            next_index,
        ));
    }

    if let Some((first_row, first_absolute, colon_index)) = parse_row_endpoint(start)
        && bytes.get(colon_index) == Some(&b':')
        && let Some((second_row, second_absolute, next_index)) = parse_row_endpoint(colon_index + 1)
        && formula[next_index..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
    {
        let start_ref = format_r1c1_row_reference(first_row, first_absolute, base_row);
        let end_ref = format_r1c1_row_reference(second_row, second_absolute, base_row);
        return Some((
            if start_ref == end_ref {
                start_ref
            } else {
                format!("{start_ref}:{end_ref}")
            },
            next_index,
        ));
    }

    None
}

/// Gives every shared-formula child the master formula translated to the child's position, so
/// evaluation, formula inspection, and dependency scans treat each child as a formula cell. The
/// XLSX rewriter keeps emitting the child's original formula-less `<f>` element.
pub(crate) fn expand_shared_formula_children(state: &mut WorkbookState) -> OmResult<()> {
    let sheet_ids = state.worksheet_data().keys().copied().collect::<Vec<_>>();
    for sheet_id in sheet_ids {
        let worksheet = state.worksheet_data_for_sheet_mut(sheet_id)?;
        let mut expansions = Vec::new();
        for (&anchor, group) in &worksheet.formula_groups {
            if group.kind != FormulaGroupKind::Shared {
                continue;
            }
            let Some(master) = worksheet
                .cells
                .get(&anchor)
                .and_then(|cell| cell.formula.as_ref())
            else {
                continue;
            };
            let relative = if master.is_r1c1 {
                master.text.clone()
            } else {
                convert_formula_a1_to_r1c1(&master.text, anchor.0, anchor.1)
            };
            for &child in &group.members {
                expansions.push((
                    child,
                    convert_formula_r1c1_to_a1(&relative, child.0, child.1),
                ));
            }
        }
        for (child, text) in expansions {
            let cell = worksheet.cells.entry(child).or_insert(CellData {
                value: CellValue::Blank,
                formula: None,
                style_id: None,
            });
            if cell.formula.is_none() {
                cell.formula = Some(FormulaSource {
                    text,
                    is_r1c1: false,
                });
            }
        }
    }
    Ok(())
}

/// The end of a quoted sheet name (`'…'` with doubled-quote escapes) or a bracketed structured or
/// external reference (`[…]`, nested) starting at `start`, which reference rewriting copies
/// verbatim.
fn formula_name_segment_end(formula: &str, start: usize) -> Option<usize> {
    let bytes = formula.as_bytes();
    match bytes.get(start)? {
        b'\'' => {
            let mut index = start + 1;
            while index < bytes.len() {
                if bytes[index] == b'\'' {
                    if bytes.get(index + 1) == Some(&b'\'') {
                        index += 2;
                        continue;
                    }
                    return Some(index + 1);
                }
                index += 1;
            }
            Some(bytes.len())
        }
        b'[' => {
            let mut depth = 0usize;
            for (index, byte) in bytes.iter().enumerate().skip(start) {
                match byte {
                    b'[' => depth += 1,
                    b']' => {
                        depth -= 1;
                        if depth == 0 {
                            return Some(index + 1);
                        }
                    }
                    _ => {}
                }
            }
            Some(bytes.len())
        }
        _ => None,
    }
}

pub(crate) fn convert_formula_a1_to_r1c1(formula: &str, base_row: u32, base_col: u32) -> String {
    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            let quoted_start = index;
            index += 1;
            while index < bytes.len() {
                if bytes[index] == b'"' {
                    index += 1;
                    if index < bytes.len() && bytes[index] == b'"' {
                        index += 1;
                        continue;
                    }
                    break;
                }
                let ch = formula[index..]
                    .chars()
                    .next()
                    .expect("valid formula char boundary");
                index += ch.len_utf8();
            }
            output.push_str(&formula[quoted_start..index]);
            continue;
        }

        // Quoted sheet names and bracketed structured or external references are names, not
        // cell references, even when their text looks like one (`'Q1'!A1`, `Table1[Col1]`).
        if let Some(end) = formula_name_segment_end(formula, index) {
            output.push_str(&formula[index..end]);
            index = end;
            continue;
        }

        let previous_is_boundary = formula[..index]
            .chars()
            .next_back()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if previous_is_boundary {
            if let Some((reference, next_index)) =
                parse_a1_axis_reference_to_r1c1(formula, index, base_row, base_col)
            {
                output.push_str(&reference);
                index = next_index;
                continue;
            }

            let reference_start = index;
            let mut cursor = index;
            let column_absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
                cursor += 1;
                true
            } else {
                false
            };
            let letters_start = cursor;
            while cursor < bytes.len()
                && bytes[cursor].is_ascii_alphabetic()
                && cursor - letters_start < 3
            {
                cursor += 1;
            }
            let letters_end = cursor;
            if letters_end > letters_start
                && (cursor >= bytes.len() || !bytes[cursor].is_ascii_alphabetic())
            {
                let row_absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
                    cursor += 1;
                    true
                } else {
                    false
                };
                let digits_start = cursor;
                while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
                    cursor += 1;
                }
                if digits_start < cursor {
                    let next_char = formula[cursor..].chars().next();
                    let next_is_boundary = next_char.is_none_or(|ch| {
                        !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '('
                    });
                    if next_is_boundary {
                        let mut col = 0u32;
                        for byte in &bytes[letters_start..letters_end] {
                            col = col * 26 + (byte.to_ascii_uppercase() - b'A' + 1) as u32;
                        }
                        let row = formula[digits_start..cursor].parse::<u32>().ok();
                        if let Some(row) = row {
                            if row > 0
                                && col > 0
                                && row <= EXCEL_MAX_ROW_INDEX
                                && col <= EXCEL_MAX_COLUMN_INDEX
                            {
                                output.push_str(&format_r1c1_reference(
                                    row,
                                    col,
                                    row_absolute,
                                    column_absolute,
                                    base_row,
                                    base_col,
                                ));
                                index = cursor;
                                continue;
                            }
                        }
                    }
                }
            }
            index = reference_start;
        }

        let ch = formula[index..]
            .chars()
            .next()
            .expect("valid formula char boundary");
        output.push(ch);
        index += ch.len_utf8();
    }
    output
}

pub(super) fn format_r1c1_reference(
    row: u32,
    col: u32,
    row_absolute: bool,
    column_absolute: bool,
    base_row: u32,
    base_col: u32,
) -> String {
    format!(
        "{}{}",
        format_r1c1_row_reference(row, row_absolute, base_row),
        format_r1c1_column_reference(col, column_absolute, base_col)
    )
}

pub(super) fn format_r1c1_row_reference(row: u32, row_absolute: bool, base_row: u32) -> String {
    if row_absolute {
        format!("R{row}")
    } else {
        let delta = i64::from(row) - i64::from(base_row);
        if delta == 0 {
            "R".to_string()
        } else {
            format!("R[{delta}]")
        }
    }
}

pub(super) fn format_r1c1_column_reference(
    col: u32,
    column_absolute: bool,
    base_col: u32,
) -> String {
    if column_absolute {
        format!("C{col}")
    } else {
        let delta = i64::from(col) - i64::from(base_col);
        if delta == 0 {
            "C".to_string()
        } else {
            format!("C[{delta}]")
        }
    }
}

pub(crate) fn format_rect_r1c1_address_with_flags(
    rect: Rect,
    row_absolute: bool,
    column_absolute: bool,
    base_row: u32,
    base_col: u32,
) -> String {
    let spans_all_rows = rect.row_first == 1 && rect.row_last == EXCEL_MAX_ROW_INDEX;
    let spans_all_columns = rect.col_first == 1 && rect.col_last == EXCEL_MAX_COLUMN_INDEX;
    if spans_all_rows && !spans_all_columns {
        let start = format_r1c1_column_reference(rect.col_first, column_absolute, base_col);
        if rect.col_first == rect.col_last {
            return start;
        }
        let end = format_r1c1_column_reference(rect.col_last, column_absolute, base_col);
        return format!("{start}:{end}");
    }
    if spans_all_columns && !spans_all_rows {
        let start = format_r1c1_row_reference(rect.row_first, row_absolute, base_row);
        if rect.row_first == rect.row_last {
            return start;
        }
        let end = format_r1c1_row_reference(rect.row_last, row_absolute, base_row);
        return format!("{start}:{end}");
    }

    let start = format_r1c1_reference(
        rect.row_first,
        rect.col_first,
        row_absolute,
        column_absolute,
        base_row,
        base_col,
    );
    if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
        start
    } else {
        let end = format_r1c1_reference(
            rect.row_last,
            rect.col_last,
            row_absolute,
            column_absolute,
            base_row,
            base_col,
        );
        format!("{start}:{end}")
    }
}

pub(crate) fn convert_formula_r1c1_to_a1(formula: &str, base_row: u32, base_col: u32) -> String {
    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            let quoted_start = index;
            index += 1;
            while index < bytes.len() {
                if bytes[index] == b'"' {
                    index += 1;
                    if index < bytes.len() && bytes[index] == b'"' {
                        index += 1;
                        continue;
                    }
                    break;
                }
                let ch = formula[index..]
                    .chars()
                    .next()
                    .expect("valid formula char boundary");
                index += ch.len_utf8();
            }
            output.push_str(&formula[quoted_start..index]);
            continue;
        }

        // Quoted sheet names and bracketed structured or external references are names, not
        // cell references, even when their text looks like one (`'Q1'!A1`, `Table1[Col1]`).
        if let Some(end) = formula_name_segment_end(formula, index) {
            output.push_str(&formula[index..end]);
            index = end;
            continue;
        }

        let previous_is_boundary = formula[..index]
            .chars()
            .next_back()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if previous_is_boundary && matches!(bytes[index], b'R' | b'r') {
            if let Some((row, row_absolute, col, column_absolute, next_index)) =
                parse_r1c1_reference(formula, index, base_row, base_col)
            {
                let next_char = formula[next_index..].chars().next();
                let next_is_boundary = next_char.is_none_or(|ch| {
                    !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '('
                });
                if next_is_boundary {
                    if row < 1
                        || row > i64::from(EXCEL_MAX_ROW_INDEX)
                        || col < 1
                        || col > i64::from(EXCEL_MAX_COLUMN_INDEX)
                    {
                        output.push_str("#REF!");
                    } else {
                        output.push_str(&format_cell_address(
                            row as u32,
                            col as u32,
                            row_absolute,
                            column_absolute,
                        ));
                    }
                    index = next_index;
                    continue;
                }
            }
        }
        if previous_is_boundary
            && matches!(bytes[index], b'R' | b'r' | b'C' | b'c')
            && let Some((reference, next_index)) =
                parse_r1c1_axis_reference_to_a1(formula, index, base_row, base_col)
        {
            output.push_str(&reference);
            index = next_index;
            continue;
        }

        let ch = formula[index..]
            .chars()
            .next()
            .expect("valid formula char boundary");
        output.push(ch);
        index += ch.len_utf8();
    }
    output
}

pub(super) fn parse_r1c1_axis_reference_to_a1(
    formula: &str,
    start: usize,
    base_row: u32,
    base_col: u32,
) -> Option<(String, usize)> {
    let bytes = formula.as_bytes();
    if matches!(bytes.get(start), Some(b'R' | b'r')) {
        let (first_row, first_absolute, cursor) =
            parse_r1c1_axis(bytes, start + 1, i64::from(base_row))?;
        if matches!(bytes.get(cursor), Some(b'C' | b'c')) {
            return None;
        }
        if bytes.get(cursor) == Some(&b':') {
            if !matches!(bytes.get(cursor + 1), Some(b'R' | b'r')) {
                return None;
            }
            let (second_row, second_absolute, next_index) =
                parse_r1c1_axis(bytes, cursor + 2, i64::from(base_row))?;
            if !formula[next_index..]
                .chars()
                .next()
                .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
            {
                return None;
            }
            let first = format_a1_row_axis(first_row, first_absolute)?;
            let second = format_a1_row_axis(second_row, second_absolute)?;
            return Some((format!("{first}:{second}"), next_index));
        }
        if formula[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
        {
            let first = format_a1_row_axis(first_row, first_absolute)?;
            return Some((format!("{first}:{first}"), cursor));
        }
    }

    if matches!(bytes.get(start), Some(b'C' | b'c')) {
        let (first_col, first_absolute, cursor) =
            parse_r1c1_axis(bytes, start + 1, i64::from(base_col))?;
        if bytes.get(cursor) == Some(&b':') {
            if !matches!(bytes.get(cursor + 1), Some(b'C' | b'c')) {
                return None;
            }
            let (second_col, second_absolute, next_index) =
                parse_r1c1_axis(bytes, cursor + 2, i64::from(base_col))?;
            if !formula[next_index..]
                .chars()
                .next()
                .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
            {
                return None;
            }
            let first = format_a1_column_axis(first_col, first_absolute)?;
            let second = format_a1_column_axis(second_col, second_absolute)?;
            return Some((format!("{first}:{second}"), next_index));
        }
        if formula[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
        {
            let first = format_a1_column_axis(first_col, first_absolute)?;
            return Some((format!("{first}:{first}"), cursor));
        }
    }

    None
}

pub(super) fn format_a1_row_axis(row: i64, absolute: bool) -> Option<String> {
    if row < 1 || row > i64::from(EXCEL_MAX_ROW_INDEX) {
        None
    } else {
        Some(format_row_address(row as u32, absolute))
    }
}

pub(super) fn format_a1_column_axis(col: i64, absolute: bool) -> Option<String> {
    if col < 1 || col > i64::from(EXCEL_MAX_COLUMN_INDEX) {
        None
    } else {
        Some(format_column_address(col as u32, absolute))
    }
}

pub(super) fn parse_r1c1_reference(
    formula: &str,
    start: usize,
    base_row: u32,
    base_col: u32,
) -> Option<(i64, bool, i64, bool, usize)> {
    let bytes = formula.as_bytes();
    let mut cursor = start;
    if !matches!(bytes.get(cursor), Some(b'R' | b'r')) {
        return None;
    }
    cursor += 1;
    let (row, row_absolute, next_cursor) = parse_r1c1_axis(bytes, cursor, i64::from(base_row))?;
    cursor = next_cursor;
    if !matches!(bytes.get(cursor), Some(b'C' | b'c')) {
        return None;
    }
    cursor += 1;
    let (col, column_absolute, next_cursor) = parse_r1c1_axis(bytes, cursor, i64::from(base_col))?;
    Some((row, row_absolute, col, column_absolute, next_cursor))
}

pub(super) fn parse_r1c1_axis(bytes: &[u8], start: usize, base: i64) -> Option<(i64, bool, usize)> {
    if bytes.get(start) == Some(&b'[') {
        let mut cursor = start + 1;
        let sign = if bytes.get(cursor) == Some(&b'-') {
            cursor += 1;
            -1
        } else {
            if bytes.get(cursor) == Some(&b'+') {
                cursor += 1;
            }
            1
        };
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if digits_start == cursor || bytes.get(cursor) != Some(&b']') {
            return None;
        }
        let value = std::str::from_utf8(&bytes[digits_start..cursor])
            .ok()?
            .parse::<i64>()
            .ok()?;
        Some((base + sign * value, false, cursor + 1))
    } else {
        let mut cursor = start;
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if digits_start == cursor {
            Some((base, false, cursor))
        } else {
            let value = std::str::from_utf8(&bytes[digits_start..cursor])
                .ok()?
                .parse::<i64>()
                .ok()?;
            Some((value, true, cursor))
        }
    }
}

pub(crate) fn shift_formula_a1_references(formula: &str, row_delta: i64, col_delta: i64) -> String {
    if row_delta == 0 && col_delta == 0 {
        return formula.to_string();
    }

    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            let quoted_start = index;
            index += 1;
            while index < bytes.len() {
                if bytes[index] == b'"' {
                    index += 1;
                    if index < bytes.len() && bytes[index] == b'"' {
                        index += 1;
                        continue;
                    }
                    break;
                }
                let ch = formula[index..]
                    .chars()
                    .next()
                    .expect("valid formula char boundary");
                index += ch.len_utf8();
            }
            output.push_str(&formula[quoted_start..index]);
            continue;
        }

        // Quoted sheet names and bracketed structured or external references are names, not
        // cell references, even when their text looks like one (`'Q1'!A1`, `Table1[Col1]`).
        if let Some(end) = formula_name_segment_end(formula, index) {
            output.push_str(&formula[index..end]);
            index = end;
            continue;
        }

        let previous_is_boundary = formula[..index]
            .chars()
            .next_back()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if previous_is_boundary {
            if let Some((reference, next_index)) =
                shift_formula_a1_axis_reference(formula, index, row_delta, col_delta)
            {
                output.push_str(&reference);
                index = next_index;
                continue;
            }

            let reference_start = index;
            let mut cursor = index;
            let column_absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
                cursor += 1;
                true
            } else {
                false
            };
            let letters_start = cursor;
            while cursor < bytes.len()
                && bytes[cursor].is_ascii_alphabetic()
                && cursor - letters_start < 3
            {
                cursor += 1;
            }
            let letters_end = cursor;
            if letters_end > letters_start
                && (cursor >= bytes.len() || !bytes[cursor].is_ascii_alphabetic())
            {
                let row_absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
                    cursor += 1;
                    true
                } else {
                    false
                };
                let digits_start = cursor;
                while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
                    cursor += 1;
                }
                if digits_start < cursor {
                    let next_char = formula[cursor..].chars().next();
                    let next_is_boundary = next_char.is_none_or(|ch| {
                        !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '('
                    });
                    if next_is_boundary {
                        let mut col = 0u32;
                        for byte in &bytes[letters_start..letters_end] {
                            col = col * 26 + (byte.to_ascii_uppercase() - b'A' + 1) as u32;
                        }
                        let row = formula[digits_start..cursor].parse::<u32>().ok();
                        if let Some(row) = row {
                            if row > 0
                                && col > 0
                                && row <= EXCEL_MAX_ROW_INDEX
                                && col <= EXCEL_MAX_COLUMN_INDEX
                            {
                                let shifted_row = if row_absolute {
                                    i64::from(row)
                                } else {
                                    i64::from(row) + row_delta
                                };
                                let shifted_col = if column_absolute {
                                    i64::from(col)
                                } else {
                                    i64::from(col) + col_delta
                                };
                                if shifted_row < 1
                                    || shifted_row > i64::from(EXCEL_MAX_ROW_INDEX)
                                    || shifted_col < 1
                                    || shifted_col > i64::from(EXCEL_MAX_COLUMN_INDEX)
                                {
                                    output.push_str("#REF!");
                                } else {
                                    output.push_str(&format_cell_address(
                                        shifted_row as u32,
                                        shifted_col as u32,
                                        row_absolute,
                                        column_absolute,
                                    ));
                                }
                                index = cursor;
                                continue;
                            }
                        }
                    }
                }
            }
            index = reference_start;
        }

        let ch = formula[index..]
            .chars()
            .next()
            .expect("valid formula char boundary");
        output.push(ch);
        index += ch.len_utf8();
    }
    output
}

pub(super) fn shift_formula_a1_axis_reference(
    formula: &str,
    start: usize,
    row_delta: i64,
    col_delta: i64,
) -> Option<(String, usize)> {
    let bytes = formula.as_bytes();
    let parse_column_endpoint = |mut cursor: usize| -> Option<(u32, bool, usize)> {
        let absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
            true
        } else {
            false
        };
        let letters_start = cursor;
        while cursor < bytes.len()
            && bytes[cursor].is_ascii_alphabetic()
            && cursor - letters_start < 3
        {
            cursor += 1;
        }
        if letters_start == cursor || (cursor < bytes.len() && bytes[cursor].is_ascii_alphabetic())
        {
            return None;
        }
        let col = parse_column_label_a1(&formula[letters_start..cursor]).ok()?;
        Some((col, absolute, cursor))
    };
    let parse_row_endpoint = |mut cursor: usize| -> Option<(u32, bool, usize)> {
        let absolute = if cursor < bytes.len() && bytes[cursor] == b'$' {
            cursor += 1;
            true
        } else {
            false
        };
        let digits_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
            cursor += 1;
        }
        if digits_start == cursor {
            return None;
        }
        let row = formula[digits_start..cursor].parse::<u32>().ok()?;
        if row == 0 || row > EXCEL_MAX_ROW_INDEX {
            return None;
        }
        Some((row, absolute, cursor))
    };
    let next_is_boundary = |cursor: usize| {
        formula[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
    };

    if let Some((first_col, first_absolute, colon_index)) = parse_column_endpoint(start)
        && bytes.get(colon_index) == Some(&b':')
        && let Some((second_col, second_absolute, next_index)) =
            parse_column_endpoint(colon_index + 1)
        && next_is_boundary(next_index)
    {
        let first = if first_absolute {
            i64::from(first_col)
        } else {
            i64::from(first_col) + col_delta
        };
        let second = if second_absolute {
            i64::from(second_col)
        } else {
            i64::from(second_col) + col_delta
        };
        if first < 1
            || first > i64::from(EXCEL_MAX_COLUMN_INDEX)
            || second < 1
            || second > i64::from(EXCEL_MAX_COLUMN_INDEX)
        {
            return Some(("#REF!".to_string(), next_index));
        }
        return Some((
            format!(
                "{}:{}",
                format_column_address(first as u32, first_absolute),
                format_column_address(second as u32, second_absolute)
            ),
            next_index,
        ));
    }

    if let Some((first_row, first_absolute, colon_index)) = parse_row_endpoint(start)
        && bytes.get(colon_index) == Some(&b':')
        && let Some((second_row, second_absolute, next_index)) = parse_row_endpoint(colon_index + 1)
        && next_is_boundary(next_index)
    {
        let first = if first_absolute {
            i64::from(first_row)
        } else {
            i64::from(first_row) + row_delta
        };
        let second = if second_absolute {
            i64::from(second_row)
        } else {
            i64::from(second_row) + row_delta
        };
        if first < 1
            || first > i64::from(EXCEL_MAX_ROW_INDEX)
            || second < 1
            || second > i64::from(EXCEL_MAX_ROW_INDEX)
        {
            return Some(("#REF!".to_string(), next_index));
        }
        return Some((
            format!(
                "{}:{}",
                format_row_address(first as u32, first_absolute),
                format_row_address(second as u32, second_absolute)
            ),
            next_index,
        ));
    }

    None
}

pub(super) fn column_to_letters(mut col: u32) -> String {
    let mut letters = Vec::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        letters.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    letters.iter().rev().collect()
}

#[cfg(test)]
mod tests {
    use super::{
        convert_formula_a1_to_r1c1, convert_formula_r1c1_to_a1, shift_formula_a1_references,
    };

    #[test]
    fn r1c1_conversion_keeps_names_that_look_like_references() {
        for (a1, r1c1) in [
            ("'Q1'!A1+B2", "'Q1'!R[-1]C[-1]+RC"),
            ("SUM(Table1[Col1])+B2", "SUM(Table1[Col1])+RC"),
            ("\"A1\"&B2", "\"A1\"&RC"),
        ] {
            assert_eq!(convert_formula_a1_to_r1c1(a1, 2, 2), r1c1, "to R1C1 {a1}");
            assert_eq!(convert_formula_r1c1_to_a1(r1c1, 2, 2), a1, "to A1 {r1c1}");
        }
    }

    #[test]
    fn shifting_moves_references_but_not_names_that_look_like_them() {
        assert_eq!(shift_formula_a1_references("'Q1'!A1+1", 1, 0), "'Q1'!A2+1");
        assert_eq!(
            shift_formula_a1_references("'Q1 Sales'!B2", 0, 1),
            "'Q1 Sales'!C2"
        );
        assert_eq!(
            shift_formula_a1_references("'It''s A1'!A1", 1, 1),
            "'It''s A1'!B2"
        );
        assert_eq!(
            shift_formula_a1_references("SUM(Table1[Col1])+B1", 1, 0),
            "SUM(Table1[Col1])+B2"
        );
        assert_eq!(
            shift_formula_a1_references("[1]Sheet1!A1+\"A1\"", 1, 0),
            "[1]Sheet1!A2+\"A1\""
        );
    }
}
