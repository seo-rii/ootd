//! Rewriting A1 references when whole rows or columns are inserted or deleted.

use crate::ExcelLimits;

/// The axis a structural edit inserts or deletes along.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StructuralAxis {
    Rows,
    Columns,
}

/// Whole rows or columns inserted before `first`, or deleted from `first` onward.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct StructuralShift {
    pub axis: StructuralAxis,
    pub first: u32,
    pub count: u32,
    pub insert: bool,
}

/// Why a formula's references cannot be retargeted automatically.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RetargetBlocker {
    /// A sheet qualifier that does not name a worksheet of the workbook.
    UnknownSheet(String),
    /// A 3D reference (`Sheet1:Sheet3!A1`), whose span may include the edited sheet.
    ThreeDimensionalReference,
    /// A relative endpoint on the edited axis that would move, where only absolute references
    /// can be retargeted (defined names, whose relative parts resolve at each caller).
    RelativeReference,
}

/// How a reference's sheet relates to the edited sheet.
pub enum SheetMatch {
    Edited,
    Other,
    Unknown,
}

/// Rewrites every A1 reference to the edited sheet in `formula`.
///
/// `sheet_of(None)` classifies an unqualified reference and `sheet_of(Some(name))` a reference
/// qualified with an (unescaped) sheet name. External-workbook references (`[1]Sheet1!A1`),
/// string literals, and structured references are left unchanged. On an insert, endpoints at or
/// after `first` move by `count`; a range end pushed past the grid stays at the last row or column
/// and any other endpoint pushed past it becomes `#REF!`. On a delete, endpoints after the
/// deleted span move back by `count`, ranges that lose only part of their span shrink, and
/// references entirely inside the span become `#REF!`.
pub fn retarget_formula_references(
    formula: &str,
    shift: StructuralShift,
    require_absolute: bool,
    sheet_of: impl Fn(Option<&str>) -> SheetMatch,
) -> Result<String, RetargetBlocker> {
    let bytes = formula.as_bytes();
    let mut output = String::with_capacity(formula.len());
    let mut index = 0usize;
    let mut external_pending = false;
    while index < bytes.len() {
        match bytes[index] {
            b'"' => {
                let end = quoted_end(formula, index, b'"');
                output.push_str(&formula[index..end]);
                index = end;
                continue;
            }
            b'[' => {
                let end = bracket_end(formula, index);
                // A bracket at a token boundary is an external-workbook index; one after an
                // identifier is a structured reference.
                external_pending = previous_is_boundary(formula, index);
                output.push_str(&formula[index..end]);
                index = end;
                continue;
            }
            _ => {}
        }
        if !previous_is_boundary(formula, index) {
            let ch = next_char(formula, index);
            output.push(ch);
            index += ch.len_utf8();
            continue;
        }

        let token_start = index;
        let (qualifier, reference_start) = match sheet_qualifier(formula, index) {
            Some((qualifier, after)) => (Some(qualifier), after),
            None => (None, index),
        };
        let Some((reference, reference_end)) = parse_reference(formula, reference_start) else {
            if qualifier.is_some() {
                // A qualified name such as `Sheet1!Total`: copy the qualifier and the name.
                let name_end = formula[reference_start..]
                    .char_indices()
                    .find(|(_, ch)| !(ch.is_alphanumeric() || matches!(ch, '_' | '.')))
                    .map_or(formula.len(), |(offset, _)| reference_start + offset);
                output.push_str(&formula[token_start..name_end]);
                index = name_end;
                external_pending = false;
            } else {
                let ch = next_char(formula, index);
                output.push(ch);
                index += ch.len_utf8();
            }
            continue;
        };
        let original = &formula[token_start..reference_end];
        let qualifier_text = &formula[token_start..reference_start];
        index = reference_end;
        if std::mem::take(&mut external_pending) {
            output.push_str(original);
            continue;
        }
        let sheet = match &qualifier {
            Some(Qualifier::Span) => return Err(RetargetBlocker::ThreeDimensionalReference),
            Some(Qualifier::Sheet(name)) => sheet_of(Some(name)),
            None => sheet_of(None),
        };
        match sheet {
            SheetMatch::Edited => {}
            SheetMatch::Other => {
                output.push_str(original);
                continue;
            }
            SheetMatch::Unknown => {
                return Err(RetargetBlocker::UnknownSheet(match qualifier {
                    Some(Qualifier::Sheet(name)) => name,
                    _ => String::new(),
                }));
            }
        }
        let retargeted = reference.retarget(shift);
        if !matches!(retargeted, Retargeted::Unchanged)
            && require_absolute
            && reference.has_relative_on(shift.axis)
        {
            return Err(RetargetBlocker::RelativeReference);
        }
        match retargeted {
            Retargeted::Unchanged => output.push_str(original),
            Retargeted::Moved(moved) => {
                output.push_str(qualifier_text);
                output.push_str(&moved.render());
            }
            Retargeted::Deleted => output.push_str("#REF!"),
        }
    }
    Ok(output)
}

enum Qualifier {
    Sheet(String),
    Span,
}

fn next_char(formula: &str, index: usize) -> char {
    formula[index..]
        .chars()
        .next()
        .expect("index is a char boundary")
}

fn previous_is_boundary(formula: &str, index: usize) -> bool {
    formula[..index]
        .chars()
        .next_back()
        .is_none_or(|ch| !(ch.is_alphanumeric() || matches!(ch, '_' | '.' | '$' | '\\')))
}

fn quoted_end(formula: &str, start: usize, quote: u8) -> usize {
    let bytes = formula.as_bytes();
    let mut index = start + 1;
    while index < bytes.len() {
        if bytes[index] == quote {
            if bytes.get(index + 1) == Some(&quote) {
                index += 2;
                continue;
            }
            return index + 1;
        }
        index += 1;
    }
    bytes.len()
}

fn bracket_end(formula: &str, start: usize) -> usize {
    let bytes = formula.as_bytes();
    let mut depth = 0usize;
    for (index, byte) in bytes.iter().enumerate().skip(start) {
        match byte {
            b'[' => depth += 1,
            b']' => {
                depth -= 1;
                if depth == 0 {
                    return index + 1;
                }
            }
            _ => {}
        }
    }
    bytes.len()
}

/// A `'Sheet Name'!` or `Sheet1!` qualifier starting at `start`, with the index after `!`.
fn sheet_qualifier(formula: &str, start: usize) -> Option<(Qualifier, usize)> {
    let bytes = formula.as_bytes();
    let (raw, after_name) = if bytes[start] == b'\'' {
        let end = quoted_end(formula, start, b'\'');
        if end > bytes.len() || end <= start + 1 {
            return None;
        }
        (formula[start + 1..end - 1].replace("''", "'"), end)
    } else {
        let end = formula[start..]
            .char_indices()
            .find(|(_, ch)| !(ch.is_alphanumeric() || matches!(ch, '_' | '.' | ':')))
            .map_or(formula.len(), |(offset, _)| start + offset);
        if end == start {
            return None;
        }
        (formula[start..end].to_string(), end)
    };
    if bytes.get(after_name) != Some(&b'!') {
        return None;
    }
    let qualifier = if raw.contains(':') {
        Qualifier::Span
    } else {
        Qualifier::Sheet(raw)
    };
    Some((qualifier, after_name + 1))
}

#[derive(Debug, Clone, Copy)]
struct Endpoint {
    value: u32,
    absolute: bool,
}

#[derive(Debug, Clone, Copy)]
enum Reference {
    Cells {
        first: (Endpoint, Endpoint),
        last: Option<(Endpoint, Endpoint)>,
    },
    Rows(Endpoint, Endpoint),
    Columns(Endpoint, Endpoint),
}

enum Retargeted {
    Unchanged,
    Moved(Reference),
    Deleted,
}

fn parse_column(formula: &str, start: usize) -> Option<(Endpoint, usize)> {
    let bytes = formula.as_bytes();
    let mut index = start;
    let absolute = bytes.get(index) == Some(&b'$');
    if absolute {
        index += 1;
    }
    let letters = index;
    while index < bytes.len() && bytes[index].is_ascii_alphabetic() && index - letters < 3 {
        index += 1;
    }
    if index == letters || bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        return None;
    }
    let mut value = 0u32;
    for byte in &bytes[letters..index] {
        value = value * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    (value <= ExcelLimits::MAX_COLUMN_INDEX).then_some((Endpoint { value, absolute }, index))
}

fn parse_row(formula: &str, start: usize) -> Option<(Endpoint, usize)> {
    let bytes = formula.as_bytes();
    let mut index = start;
    let absolute = bytes.get(index) == Some(&b'$');
    if absolute {
        index += 1;
    }
    let digits = index;
    while index < bytes.len() && bytes[index].is_ascii_digit() {
        index += 1;
    }
    if index == digits {
        return None;
    }
    let value = formula[digits..index].parse::<u32>().ok()?;
    (1..=ExcelLimits::MAX_ROW_INDEX)
        .contains(&value)
        .then_some((Endpoint { value, absolute }, index))
}

fn parse_cell(formula: &str, start: usize) -> Option<((Endpoint, Endpoint), usize)> {
    let (column, after_column) = parse_column(formula, start)?;
    let (row, after_row) = parse_row(formula, after_column)?;
    Some(((column, row), after_row))
}

fn reference_ends_here(formula: &str, index: usize) -> bool {
    formula[index..]
        .chars()
        .next()
        .is_none_or(|ch| !(ch.is_alphanumeric() || matches!(ch, '_' | '.' | '(' | '!' | '$')))
}

fn parse_reference(formula: &str, start: usize) -> Option<(Reference, usize)> {
    let bytes = formula.as_bytes();
    if let Some((first, after_first)) = parse_cell(formula, start) {
        if bytes.get(after_first) == Some(&b':')
            && let Some((last, after_last)) = parse_cell(formula, after_first + 1)
            && reference_ends_here(formula, after_last)
        {
            return Some((
                Reference::Cells {
                    first,
                    last: Some(last),
                },
                after_last,
            ));
        }
        if reference_ends_here(formula, after_first) {
            return Some((Reference::Cells { first, last: None }, after_first));
        }
        return None;
    }
    if let Some((first, after_first)) = parse_column(formula, start)
        && bytes.get(after_first) == Some(&b':')
        && let Some((last, after_last)) = parse_column(formula, after_first + 1)
        && reference_ends_here(formula, after_last)
    {
        return Some((Reference::Columns(first, last), after_last));
    }
    if let Some((first, after_first)) = parse_row(formula, start)
        && bytes.get(after_first) == Some(&b':')
        && let Some((last, after_last)) = parse_row(formula, after_first + 1)
        && reference_ends_here(formula, after_last)
    {
        return Some((Reference::Rows(first, last), after_last));
    }
    None
}

fn column_label(mut value: u32) -> String {
    let mut letters = Vec::new();
    while value > 0 {
        letters.push(char::from(b'A' + ((value - 1) % 26) as u8));
        value = (value - 1) / 26;
    }
    letters.iter().rev().collect()
}

fn render_endpoint(endpoint: Endpoint, column: bool) -> String {
    let dollar = if endpoint.absolute { "$" } else { "" };
    if column {
        format!("{dollar}{}", column_label(endpoint.value))
    } else {
        format!("{dollar}{}", endpoint.value)
    }
}

/// Moves a `first..=last` span on the edited axis.
fn retarget_span(
    first: u32,
    last: u32,
    shift: StructuralShift,
    max: u32,
    is_range: bool,
) -> Option<(u32, u32)> {
    let (low, high) = (first.min(last), first.max(last));
    if shift.insert {
        let moved = |value: u32| {
            if value >= shift.first {
                value.checked_add(shift.count)
            } else {
                Some(value)
            }
        };
        let new_low = moved(low).filter(|value| *value <= max)?;
        let new_high = match moved(high) {
            Some(value) if value <= max => value,
            _ if is_range => max,
            _ => return None,
        };
        return Some((new_low, new_high));
    }
    let deleted_last = shift.first + shift.count - 1;
    if low >= shift.first && high <= deleted_last {
        return None;
    }
    let new_low = if low < shift.first {
        low
    } else if low <= deleted_last {
        shift.first
    } else {
        low - shift.count
    };
    let new_high = if high < shift.first {
        high
    } else if high <= deleted_last {
        shift.first - 1
    } else {
        high - shift.count
    };
    Some((new_low, new_high))
}

impl Reference {
    fn has_relative_on(&self, axis: StructuralAxis) -> bool {
        let relative = |endpoint: &Endpoint| !endpoint.absolute;
        match (self, axis) {
            (Reference::Cells { first, last }, StructuralAxis::Rows) => {
                relative(&first.1) || last.is_some_and(|last| relative(&last.1))
            }
            (Reference::Cells { first, last }, StructuralAxis::Columns) => {
                relative(&first.0) || last.is_some_and(|last| relative(&last.0))
            }
            (Reference::Rows(first, last), StructuralAxis::Rows)
            | (Reference::Columns(first, last), StructuralAxis::Columns) => {
                relative(first) || relative(last)
            }
            _ => false,
        }
    }

    fn retarget(&self, shift: StructuralShift) -> Retargeted {
        let max = match shift.axis {
            StructuralAxis::Rows => ExcelLimits::MAX_ROW_INDEX,
            StructuralAxis::Columns => ExcelLimits::MAX_COLUMN_INDEX,
        };
        let pick = |cell: &(Endpoint, Endpoint)| match shift.axis {
            StructuralAxis::Rows => cell.1,
            StructuralAxis::Columns => cell.0,
        };
        let (first, last, is_range) = match self {
            Reference::Cells { first, last } => (
                pick(first),
                last.as_ref().map_or(pick(first), pick),
                last.is_some(),
            ),
            Reference::Rows(first, last) if shift.axis == StructuralAxis::Rows => {
                (*first, *last, true)
            }
            Reference::Columns(first, last) if shift.axis == StructuralAxis::Columns => {
                (*first, *last, true)
            }
            _ => return Retargeted::Unchanged,
        };
        let Some((low, high)) = retarget_span(first.value, last.value, shift, max, is_range) else {
            return Retargeted::Deleted;
        };
        let (new_first, new_last) = if first.value <= last.value {
            (low, high)
        } else {
            (high, low)
        };
        if new_first == first.value && new_last == last.value {
            return Retargeted::Unchanged;
        }
        let with = |endpoint: Endpoint, value: u32| Endpoint { value, ..endpoint };
        let place = |cell: (Endpoint, Endpoint), value: u32| match shift.axis {
            StructuralAxis::Rows => (cell.0, with(cell.1, value)),
            StructuralAxis::Columns => (with(cell.0, value), cell.1),
        };
        Retargeted::Moved(match *self {
            Reference::Cells {
                first: cell,
                last: None,
            } => Reference::Cells {
                first: place(cell, new_first),
                last: None,
            },
            Reference::Cells {
                first: cell,
                last: Some(last_cell),
            } => Reference::Cells {
                first: place(cell, new_first),
                last: Some(place(last_cell, new_last)),
            },
            Reference::Rows(a, b) => Reference::Rows(with(a, new_first), with(b, new_last)),
            Reference::Columns(a, b) => Reference::Columns(with(a, new_first), with(b, new_last)),
        })
    }

    fn render(&self) -> String {
        let cell = |cell: &(Endpoint, Endpoint)| {
            format!(
                "{}{}",
                render_endpoint(cell.0, true),
                render_endpoint(cell.1, false)
            )
        };
        match self {
            Reference::Cells { first, last: None } => cell(first),
            Reference::Cells {
                first,
                last: Some(last),
            } => format!("{}:{}", cell(first), cell(last)),
            Reference::Rows(a, b) => {
                format!(
                    "{}:{}",
                    render_endpoint(*a, false),
                    render_endpoint(*b, false)
                )
            }
            Reference::Columns(a, b) => {
                format!(
                    "{}:{}",
                    render_endpoint(*a, true),
                    render_endpoint(*b, true)
                )
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rows(first: u32, count: u32, insert: bool) -> StructuralShift {
        StructuralShift {
            axis: StructuralAxis::Rows,
            first,
            count,
            insert,
        }
    }

    fn on_data(qualifier: Option<&str>) -> SheetMatch {
        match qualifier {
            None => SheetMatch::Edited,
            Some(name) if name.eq_ignore_ascii_case("Data") => SheetMatch::Edited,
            Some("Other" | "Q1") => SheetMatch::Other,
            Some(_) => SheetMatch::Unknown,
        }
    }

    fn retarget(formula: &str, shift: StructuralShift) -> String {
        retarget_formula_references(formula, shift, false, on_data).expect(formula)
    }

    #[test]
    fn row_inserts_move_and_expand_references() {
        let insert = rows(5, 2, true);
        for (before, after) in [
            ("A4+A5", "A4+A7"),
            ("SUM($B$3:$B$9)", "SUM($B$3:$B$11)"),
            ("SUM(B5:B9)", "SUM(B7:B11)"),
            ("Data!C10*'data'!C1", "Data!C12*'data'!C1"),
            ("Other!A9+'Q1'!A9", "Other!A9+'Q1'!A9"),
            ("SUM(5:6)+SUM(A:A)", "SUM(7:8)+SUM(A:A)"),
            ("\"A9\"&LOG10(A9)", "\"A9\"&LOG10(A11)"),
            ("[1]Data!A9+Table1[A9]", "[1]Data!A9+Table1[A9]"),
            ("SUM(A1:A1048576)", "SUM(A1:A1048576)"),
            ("A1048575", "#REF!"),
            ("Data!Total+A9", "Data!Total+A11"),
        ] {
            assert_eq!(retarget(before, insert), after, "{before}");
        }
    }

    #[test]
    fn row_deletes_shrink_and_invalidate_references() {
        let delete = rows(5, 3, false);
        for (before, after) in [
            ("A4+A8", "A4+A5"),
            ("A6", "#REF!"),
            ("SUM(A5:A7)", "SUM(#REF!)"),
            ("SUM(A3:A10)", "SUM(A3:A7)"),
            ("SUM(A6:A10)", "SUM(A5:A7)"),
            ("SUM(A2:A6)", "SUM(A2:A4)"),
            ("SUM(9:12)", "SUM(6:9)"),
        ] {
            assert_eq!(retarget(before, delete), after, "{before}");
        }
    }

    #[test]
    fn column_shifts_move_column_endpoints_only() {
        let insert = StructuralShift {
            axis: StructuralAxis::Columns,
            first: 2,
            count: 1,
            insert: true,
        };
        assert_eq!(
            retarget("A1+B1+SUM(B:C)+SUM(1:2)", insert),
            "A1+C1+SUM(C:D)+SUM(1:2)"
        );
        assert_eq!(retarget("$XFD$1", insert), "#REF!");
    }

    #[test]
    fn blocked_formulas_are_reported() {
        assert_eq!(
            retarget_formula_references("Missing!A9", rows(5, 1, true), false, on_data),
            Err(RetargetBlocker::UnknownSheet("Missing".to_string()))
        );
        assert_eq!(
            retarget_formula_references("SUM(Data:Other!A1)", rows(5, 1, true), false, on_data),
            Err(RetargetBlocker::ThreeDimensionalReference)
        );
        assert_eq!(
            retarget_formula_references("Data!A9", rows(5, 1, true), true, on_data),
            Err(RetargetBlocker::RelativeReference)
        );
        assert_eq!(
            retarget_formula_references("Data!$A$9+Data!A1", rows(5, 1, true), true, on_data),
            Ok("Data!$A$10+Data!A1".to_string())
        );
    }
}
