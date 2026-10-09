//! Converting formula text between A1 and R1C1 reference styles.
//!
//! R1C1 references are relative to a base cell: `R[-1]C` is the cell above it. Conversion keeps
//! string literals, quoted sheet names, and bracketed structured or external references unchanged.

use crate::ExcelLimits;

fn column_letters(mut col: u32) -> String {
    let mut letters = Vec::new();
    while col > 0 {
        letters.push(char::from(b'A' + ((col - 1) % 26) as u8));
        col = (col - 1) / 26;
    }
    letters.iter().rev().collect()
}

/// The column index of A1 column letters, or `None` outside the worksheet grid.
fn column_index(letters: &str) -> Option<u32> {
    if letters.is_empty() || !letters.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        return None;
    }
    let mut col = 0u32;
    for byte in letters.bytes() {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))?;
    }
    (1..=ExcelLimits::MAX_COLUMN_INDEX)
        .contains(&col)
        .then_some(col)
}

fn format_column_address(col: u32, column_absolute: bool) -> String {
    format!(
        "{}{}",
        if column_absolute { "$" } else { "" },
        column_letters(col)
    )
}

fn format_row_address(row: u32, row_absolute: bool) -> String {
    format!("{}{row}", if row_absolute { "$" } else { "" })
}

fn format_cell_address(row: u32, col: u32, row_absolute: bool, column_absolute: bool) -> String {
    format!(
        "{}{}",
        format_column_address(col, column_absolute),
        format_row_address(row, row_absolute)
    )
}

pub fn parse_a1_axis_reference_to_r1c1(
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
        let col = column_index(&formula[letters_start..cursor])?;
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
        if row == 0 || row > ExcelLimits::MAX_ROW_INDEX {
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

/// The end of a quoted sheet name (`'…'` with doubled-quote escapes) or a bracketed structured or
/// external reference (`[…]`, nested) starting at `start`, which reference rewriting copies
/// verbatim.
pub fn formula_name_segment_end(formula: &str, start: usize) -> Option<usize> {
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

pub fn convert_formula_a1_to_r1c1(formula: &str, base_row: u32, base_col: u32) -> String {
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
                        if let Some(row) = row
                            && row > 0
                            && col > 0
                            && row <= ExcelLimits::MAX_ROW_INDEX
                            && col <= ExcelLimits::MAX_COLUMN_INDEX
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

pub fn format_r1c1_reference(
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

pub fn format_r1c1_row_reference(row: u32, row_absolute: bool, base_row: u32) -> String {
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

pub fn format_r1c1_column_reference(col: u32, column_absolute: bool, base_col: u32) -> String {
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

pub fn convert_formula_r1c1_to_a1(formula: &str, base_row: u32, base_col: u32) -> String {
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
        if previous_is_boundary
            && matches!(bytes[index], b'R' | b'r')
            && let Some((row, row_absolute, col, column_absolute, next_index)) =
                parse_r1c1_reference(formula, index, base_row, base_col)
        {
            let next_char = formula[next_index..].chars().next();
            let next_is_boundary = next_char.is_none_or(|ch| {
                !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '('
            });
            if next_is_boundary {
                if row < 1
                    || row > i64::from(ExcelLimits::MAX_ROW_INDEX)
                    || col < 1
                    || col > i64::from(ExcelLimits::MAX_COLUMN_INDEX)
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

pub fn parse_r1c1_axis_reference_to_a1(
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

pub fn format_a1_row_axis(row: i64, absolute: bool) -> Option<String> {
    if row < 1 || row > i64::from(ExcelLimits::MAX_ROW_INDEX) {
        None
    } else {
        Some(format_row_address(row as u32, absolute))
    }
}

pub fn format_a1_column_axis(col: i64, absolute: bool) -> Option<String> {
    if col < 1 || col > i64::from(ExcelLimits::MAX_COLUMN_INDEX) {
        None
    } else {
        Some(format_column_address(col as u32, absolute))
    }
}

pub fn parse_r1c1_reference(
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

pub fn parse_r1c1_axis(bytes: &[u8], start: usize, base: i64) -> Option<(i64, bool, usize)> {
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
