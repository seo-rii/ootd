use serde::{Deserialize, Serialize};

use crate::{
    CellError, CellValue, ExcelLimits, FormulaSource, OmArray, OmError, OmResult, RangeRef, Rect,
    SheetScope, WorkbookId,
};

pub fn formula_contains_a1_reference(formula: &str) -> bool {
    let bytes = formula.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'"' {
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
                    .expect("valid formula character boundary");
                index += ch.len_utf8();
            }
            continue;
        }

        let previous_is_boundary = formula[..index]
            .chars()
            .next_back()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if previous_is_boundary
            && (a1_cell_reference_end(formula, index).is_some()
                || a1_axis_range_end(formula, index).is_some())
        {
            return true;
        }

        let ch = formula[index..]
            .chars()
            .next()
            .expect("valid formula character boundary");
        index += ch.len_utf8();
    }
    false
}

fn a1_cell_reference_end(formula: &str, start: usize) -> Option<usize> {
    let bytes = formula.as_bytes();
    let mut cursor = start;
    if bytes.get(cursor) == Some(&b'$') {
        cursor += 1;
    }
    let letters_start = cursor;
    while cursor < bytes.len() && bytes[cursor].is_ascii_alphabetic() && cursor - letters_start < 3
    {
        cursor += 1;
    }
    if letters_start == cursor || bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) {
        return None;
    }
    let column = formula[letters_start..cursor]
        .bytes()
        .try_fold(0u32, |column, byte| {
            column
                .checked_mul(26)?
                .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
        })?;
    if column == 0 || column > ExcelLimits::MAX_COLUMN_INDEX {
        return None;
    }
    if bytes.get(cursor) == Some(&b'$') {
        cursor += 1;
    }
    let digits_start = cursor;
    while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
        cursor += 1;
    }
    if digits_start == cursor {
        return None;
    }
    let row = formula[digits_start..cursor].parse::<u32>().ok()?;
    if row == 0 || row > ExcelLimits::MAX_ROW_INDEX || !a1_reference_boundary(formula, cursor) {
        return None;
    }
    Some(cursor)
}

fn a1_axis_range_end(formula: &str, start: usize) -> Option<usize> {
    let bytes = formula.as_bytes();
    for column_axis in [true, false] {
        let parse_axis = |mut cursor: usize| -> Option<usize> {
            if bytes.get(cursor) == Some(&b'$') {
                cursor += 1;
            }
            let token_start = cursor;
            if column_axis {
                while cursor < bytes.len()
                    && bytes[cursor].is_ascii_alphabetic()
                    && cursor - token_start < 3
                {
                    cursor += 1;
                }
                if token_start == cursor || bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) {
                    return None;
                }
                let column =
                    formula[token_start..cursor]
                        .bytes()
                        .try_fold(0u32, |column, byte| {
                            column
                                .checked_mul(26)?
                                .checked_add(u32::from(byte.to_ascii_uppercase() - b'A' + 1))
                        })?;
                (column > 0 && column <= ExcelLimits::MAX_COLUMN_INDEX).then_some(cursor)
            } else {
                while cursor < bytes.len() && bytes[cursor].is_ascii_digit() {
                    cursor += 1;
                }
                if token_start == cursor {
                    return None;
                }
                let row = formula[token_start..cursor].parse::<u32>().ok()?;
                (row > 0 && row <= ExcelLimits::MAX_ROW_INDEX).then_some(cursor)
            }
        };
        let Some(first_end) = parse_axis(start) else {
            continue;
        };
        if bytes.get(first_end) != Some(&b':') {
            continue;
        }
        let Some(second_end) = parse_axis(first_end + 1) else {
            continue;
        };
        if a1_reference_boundary(formula, second_end) {
            return Some(second_end);
        }
    }
    None
}

fn a1_reference_boundary(formula: &str, index: usize) -> bool {
    formula[index..]
        .chars()
        .next()
        .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.' && ch != '(')
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct RangeArea {
    pub scope: SheetScope,
    pub rect: Rect,
}

impl RangeArea {
    pub fn new(scope: SheetScope, rect: Rect) -> OmResult<Self> {
        ExcelLimits::validate_rect(rect)?;

        Ok(Self { scope, rect })
    }

    pub fn checked_cell_count(&self) -> OmResult<u64> {
        self.rect.checked_cell_count()
    }

    pub fn checked_cell_count_usize(&self) -> OmResult<usize> {
        self.rect.checked_cell_count_usize()
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct RangeSet {
    workbook_id: WorkbookId,
    areas: Vec<RangeArea>,
}

impl RangeSet {
    pub fn new(workbook_id: WorkbookId, areas: Vec<RangeArea>) -> OmResult<Self> {
        if areas.is_empty() {
            return Err(OmError::invalid_argument(
                "range set must contain at least one area",
            ));
        }

        for area in &areas {
            RangeArea::new(area.scope, area.rect)?;
        }

        Ok(Self { workbook_id, areas })
    }

    pub fn single_area(workbook_id: WorkbookId, scope: SheetScope, rect: Rect) -> OmResult<Self> {
        Self::new(workbook_id, vec![RangeArea::new(scope, rect)?])
    }

    pub fn single_rect(
        workbook_id: WorkbookId,
        sheet_id: crate::SheetId,
        rect: Rect,
    ) -> OmResult<Self> {
        Self::single_area(workbook_id, SheetScope::Single(sheet_id), rect)
    }

    pub fn workbook_id(&self) -> WorkbookId {
        self.workbook_id
    }

    pub fn areas(&self) -> &[RangeArea] {
        &self.areas
    }

    pub fn explicit_areas(&self) -> &[RangeArea] {
        self.areas()
    }

    pub fn len(&self) -> usize {
        self.areas.len()
    }

    pub fn is_empty(&self) -> bool {
        self.areas.is_empty()
    }

    pub fn is_multi_area(&self) -> bool {
        self.areas.len() > 1
    }

    pub fn into_areas(self) -> Vec<RangeArea> {
        self.areas
    }
}

impl TryFrom<RangeRef> for RangeSet {
    type Error = OmError;

    fn try_from(range: RangeRef) -> OmResult<Self> {
        let areas = range
            .areas
            .into_iter()
            .map(|rect| RangeArea::new(range.scope, rect))
            .collect::<OmResult<Vec<_>>>()?;
        Self::new(range.workbook_id, areas)
    }
}

impl TryFrom<&RangeRef> for RangeSet {
    type Error = OmError;

    fn try_from(range: &RangeRef) -> OmResult<Self> {
        Self::try_from(range.clone())
    }
}

impl TryFrom<RangeSet> for RangeRef {
    type Error = OmError;

    fn try_from(range: RangeSet) -> OmResult<Self> {
        let Some(first_area) = range.areas.first() else {
            return Err(OmError::invalid_argument(
                "range set must contain at least one area",
            ));
        };
        let scope = first_area.scope;

        if range.areas.iter().any(|area| area.scope != scope) {
            return Err(OmError::invalid_argument(
                "range set areas do not share a common sheet scope",
            ));
        }

        Ok(Self {
            workbook_id: range.workbook_id,
            scope,
            areas: range.areas.into_iter().map(|area| area.rect).collect(),
        })
    }
}

impl TryFrom<&RangeSet> for RangeRef {
    type Error = OmError;

    fn try_from(range: &RangeSet) -> OmResult<Self> {
        Self::try_from(range.clone())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ExternalReference {
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum ReferenceTarget {
    Range(RangeSet),
    Value(CellValue),
    Array(OmArray),
    Formula(FormulaSource),
    External(ExternalReference),
    Invalid(CellError),
    UnsupportedRaw(String),
}

#[cfg(test)]
mod tests {
    use super::{RangeArea, RangeSet, formula_contains_a1_reference};
    use crate::{ExcelLimits, OmErrorCode, RangeRef, Rect, SheetId, SheetScope, WorkbookId};

    #[test]
    fn range_set_preserves_input_area_order() {
        let first = RangeArea::new(SheetScope::Single(SheetId(2)), Rect::single_cell(1, 3))
            .expect("first area");
        let second = RangeArea::new(SheetScope::Single(SheetId(2)), Rect::single_cell(1, 1))
            .expect("second area");

        let range =
            RangeSet::new(WorkbookId(1), vec![first, second]).expect("multi-area range set");

        assert_eq!(range.areas(), &[first, second]);
        assert!(range.is_multi_area());
    }

    #[test]
    fn formula_a1_reference_detection_skips_text_and_function_names() {
        for formula in [
            "SUM(A1)",
            "$A$1",
            "'Sheet 1'!B2",
            "A:A",
            "$1:$3",
            "[Book.xlsx]Sheet1!XFD1048576",
            "A1#",
            "A1:B2",
        ] {
            assert!(formula_contains_a1_reference(formula), "{formula}");
        }
        for formula in [
            r#"UPPER("A1")"#,
            "LOG10(100)",
            "Named_Value",
            "XFE1",
            "A1048577",
            "R[-1]C",
        ] {
            assert!(!formula_contains_a1_reference(formula), "{formula}");
        }
    }

    #[test]
    fn range_set_rejects_empty_construction() {
        let error = RangeSet::new(WorkbookId(1), Vec::new())
            .expect_err("empty range sets should be rejected");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn range_set_rejects_zero_based_rect() {
        let error = RangeArea::new(SheetScope::Single(SheetId(2)), Rect::single_cell(0, 1))
            .expect_err("zero-based coordinates should be rejected");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn range_area_rejects_inverted_rectangles() {
        let error = RangeArea::new(
            SheetScope::Single(SheetId(2)),
            Rect {
                row_first: 3,
                row_last: 2,
                col_first: 1,
                col_last: 1,
            },
        )
        .expect_err("inverted rectangles should be rejected");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn range_area_rejects_coordinates_beyond_the_excel_grid() {
        for rect in [
            Rect::single_cell(1_048_577, 1),
            Rect::single_cell(1, 16_385),
            Rect {
                row_first: 1,
                row_last: 1_048_577,
                col_first: 1,
                col_last: 1,
            },
            Rect {
                row_first: 1,
                row_last: 1,
                col_first: 1,
                col_last: 16_385,
            },
        ] {
            let error = RangeArea::new(SheetScope::Single(SheetId(2)), rect)
                .expect_err("out-of-grid range coordinates should be rejected");

            assert_eq!(error.code, OmErrorCode::InvalidArgument);
        }
    }

    #[test]
    fn range_area_reports_checked_excel_grid_cell_counts() {
        let area = RangeArea::new(
            SheetScope::Single(SheetId(2)),
            Rect {
                row_first: 1,
                row_last: 1_048_576,
                col_first: 1,
                col_last: 16_384,
            },
        )
        .expect("full Excel grid");

        assert_eq!(
            area.checked_cell_count().expect("u64 count"),
            ExcelLimits::MAX_CELL_COUNT
        );
        if usize::BITS >= 64 {
            let expected = usize::try_from(ExcelLimits::MAX_CELL_COUNT)
                .expect("64-bit usize represents the full Excel grid");
            assert_eq!(
                area.checked_cell_count_usize().expect("usize count"),
                expected
            );
        } else {
            assert_eq!(
                area.checked_cell_count_usize()
                    .expect_err("32-bit usize cannot represent the full grid")
                    .code,
                OmErrorCode::ResourceLimit
            );
        }
    }

    #[test]
    fn range_set_does_not_merge_adjacent_areas_by_default() {
        let first = RangeArea::new(
            SheetScope::Single(SheetId(2)),
            Rect {
                row_first: 1,
                row_last: 2,
                col_first: 1,
                col_last: 1,
            },
        )
        .expect("first area");
        let second = RangeArea::new(
            SheetScope::Single(SheetId(2)),
            Rect {
                row_first: 3,
                row_last: 4,
                col_first: 1,
                col_last: 1,
            },
        )
        .expect("second area");

        let range = RangeSet::new(WorkbookId(1), vec![first, second]).expect("range set");

        assert_eq!(range.len(), 2);
        assert_eq!(range.areas(), &[first, second]);
    }

    #[test]
    fn range_ref_converts_to_range_set() {
        let range_ref = RangeRef::single_rect(
            WorkbookId(1),
            SheetId(2),
            Rect {
                row_first: 1,
                row_last: 3,
                col_first: 4,
                col_last: 6,
            },
        );

        let range_set = RangeSet::try_from(range_ref).expect("range ref conversion");

        assert_eq!(range_set.workbook_id(), WorkbookId(1));
        assert_eq!(
            range_set.areas(),
            &[RangeArea {
                scope: SheetScope::Single(SheetId(2)),
                rect: Rect {
                    row_first: 1,
                    row_last: 3,
                    col_first: 4,
                    col_last: 6,
                },
            }]
        );
    }

    #[test]
    fn range_set_try_into_range_ref_requires_common_scope() {
        let range_set = RangeSet::new(
            WorkbookId(1),
            vec![
                RangeArea::new(SheetScope::Single(SheetId(2)), Rect::single_cell(1, 1))
                    .expect("first area"),
                RangeArea::new(SheetScope::Single(SheetId(3)), Rect::single_cell(1, 2))
                    .expect("second area"),
            ],
        )
        .expect("mixed sheet range set");

        let error = RangeRef::try_from(range_set)
            .expect_err("mixed sheet range sets should not convert to RangeRef");

        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }
}
