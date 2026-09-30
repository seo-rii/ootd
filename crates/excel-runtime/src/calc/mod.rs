use super::{
    APPLICATION_VERSION, EXCEL_MAX_COLUMN_INDEX, EXCEL_MAX_ROW_INDEX, RuntimeDateOrder,
    RuntimeEnvironment, RuntimeLocale, xml_local_name,
};
use excel_model::{CellData, FormulaGroupKind, WorkbookState};
use office_common::{
    CellError, CellValue, DefinedNameId, FormulaSource, NameScope, OmError, OmErrorCode, OmResult,
    OmValue, Rect, SheetId,
};
use quick_xml::Reader;
use quick_xml::events::Event;
use regex::{Regex, RegexBuilder};
use std::cmp::Ordering;
use std::collections::BTreeSet;
use std::io::Cursor;

mod aggregate_functions;
mod context;
mod date_time;
mod engineering;
mod evaluator;
mod financial;
mod lookup;
mod matrix;
mod numeric;
mod parser_date_time;
mod parser_financial;
mod parser_logic;
mod parser_lookup;
mod parser_math;
mod parser_statistics;
mod parser_text;
mod reference_text;
mod scalar_functions;
mod text;

use aggregate_functions::*;
pub(super) use context::CalcContext;
#[cfg(test)]
pub(super) use context::formula_current_excel_serial;
#[cfg(test)]
pub(super) use date_time::formula_date_serial_from_args;
use date_time::*;
pub(super) use engineering::*;
pub(super) use evaluator::*;
use financial::*;
use lookup::*;
use matrix::*;
use numeric::*;
pub(super) use reference_text::*;
use scalar_functions::*;
pub(super) use text::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FormulaEvalError {
    Unsupported,
    Circular,
    Null,
    Div0,
    Value,
    Ref,
    Name,
    NA,
    Num,
    GettingData,
    Spill,
    Calc,
    Field,
    Blocked,
    Busy,
    Connect,
    Python,
    Timeout,
    Unknown,
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct FormulaArrayResult {
    pub(super) rows: usize,
    pub(super) cols: usize,
    pub(super) values: Vec<CellValue>,
}

impl FormulaArrayResult {
    pub(super) fn single(value: CellValue) -> Self {
        Self {
            rows: 1,
            cols: 1,
            values: vec![value],
        }
    }
}

impl FormulaEvalError {
    pub(super) fn into_cell_value(self) -> Option<CellValue> {
        match self {
            FormulaEvalError::Unsupported => None,
            FormulaEvalError::Circular => Some(CellValue::Error(CellError::Calc)),
            FormulaEvalError::Null => Some(CellValue::Error(CellError::Null)),
            FormulaEvalError::Div0 => Some(CellValue::Error(CellError::Div0)),
            FormulaEvalError::Value => Some(CellValue::Error(CellError::Value)),
            FormulaEvalError::Ref => Some(CellValue::Error(CellError::Ref)),
            FormulaEvalError::Name => Some(CellValue::Error(CellError::Name)),
            FormulaEvalError::NA => Some(CellValue::Error(CellError::NA)),
            FormulaEvalError::Num => Some(CellValue::Error(CellError::Num)),
            FormulaEvalError::GettingData => Some(CellValue::Error(CellError::GettingData)),
            FormulaEvalError::Spill => Some(CellValue::Error(CellError::Spill)),
            FormulaEvalError::Calc => Some(CellValue::Error(CellError::Calc)),
            FormulaEvalError::Field => Some(CellValue::Error(CellError::Field)),
            FormulaEvalError::Blocked => Some(CellValue::Error(CellError::Blocked)),
            FormulaEvalError::Busy => Some(CellValue::Error(CellError::Busy)),
            FormulaEvalError::Connect => Some(CellValue::Error(CellError::Connect)),
            FormulaEvalError::Python => Some(CellValue::Error(CellError::Python)),
            FormulaEvalError::Timeout => Some(CellValue::Error(CellError::Timeout)),
            FormulaEvalError::Unknown => Some(CellValue::Error(CellError::Unknown)),
        }
    }
}

fn formula_eval_error_from_cell_error(error: CellError) -> FormulaEvalError {
    match error {
        CellError::Null => FormulaEvalError::Null,
        CellError::Div0 => FormulaEvalError::Div0,
        CellError::Ref => FormulaEvalError::Ref,
        CellError::Name => FormulaEvalError::Name,
        CellError::NA => FormulaEvalError::NA,
        CellError::Num => FormulaEvalError::Num,
        CellError::GettingData => FormulaEvalError::GettingData,
        CellError::Spill => FormulaEvalError::Spill,
        CellError::Calc => FormulaEvalError::Calc,
        CellError::Field => FormulaEvalError::Field,
        CellError::Blocked => FormulaEvalError::Blocked,
        CellError::Busy => FormulaEvalError::Busy,
        CellError::Connect => FormulaEvalError::Connect,
        CellError::Python => FormulaEvalError::Python,
        CellError::Timeout => FormulaEvalError::Timeout,
        CellError::Unknown | CellError::UnknownLexical(_) => FormulaEvalError::Unknown,
        CellError::Value => FormulaEvalError::Value,
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaComparisonOperator {
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
}

impl FormulaComparisonOperator {
    fn evaluate(self, left: f64, right: f64) -> bool {
        match self {
            FormulaComparisonOperator::Equal => left == right,
            FormulaComparisonOperator::NotEqual => left != right,
            FormulaComparisonOperator::LessThan => left < right,
            FormulaComparisonOperator::LessThanOrEqual => left <= right,
            FormulaComparisonOperator::GreaterThan => left > right,
            FormulaComparisonOperator::GreaterThanOrEqual => left >= right,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
enum FormulaCriteria {
    Blank,
    NonBlank,
    Number {
        operator: FormulaComparisonOperator,
        value: f64,
    },
    Text {
        operator: FormulaComparisonOperator,
        pattern: String,
    },
}

impl FormulaCriteria {
    fn matches(&self, cell_value: &CellValue) -> bool {
        match self {
            FormulaCriteria::Blank => matches!(cell_value, CellValue::Blank),
            FormulaCriteria::NonBlank => !matches!(cell_value, CellValue::Blank),
            FormulaCriteria::Number {
                operator,
                value: expected,
            } => match cell_value {
                CellValue::Number(actual) => operator.evaluate(*actual, *expected),
                CellValue::Bool(actual) => {
                    operator.evaluate(if *actual { 1.0 } else { 0.0 }, *expected)
                }
                CellValue::Blank
                | CellValue::Text(_)
                | CellValue::Error(_)
                | CellValue::IsoDateTime(_)
                | CellValue::RichText(_) => false,
            },
            FormulaCriteria::Text { operator, pattern } => {
                let Some(actual) = cell_value.as_text() else {
                    return false;
                };
                match operator {
                    FormulaComparisonOperator::Equal => {
                        formula_wildcard_matches(pattern, actual, true)
                    }
                    FormulaComparisonOperator::NotEqual => {
                        !formula_wildcard_matches(pattern, actual, true)
                    }
                    _ => false,
                }
            }
        }
    }

    fn from_numeric_value(value: f64) -> Self {
        Self::Number {
            operator: FormulaComparisonOperator::Equal,
            value,
        }
    }

    fn from_string_literal(literal: String) -> Self {
        if literal.is_empty() {
            return Self::Blank;
        }
        if let Some((operator, value)) = parse_formula_criteria_numeric_literal(literal.as_str()) {
            return Self::Number { operator, value };
        }
        if literal == "<>" {
            return Self::NonBlank;
        }
        if let Some(pattern) = literal.strip_prefix("<>") {
            return Self::Text {
                operator: FormulaComparisonOperator::NotEqual,
                pattern: pattern.to_string(),
            };
        }
        if let Some(pattern) = literal.strip_prefix('=') {
            if pattern.is_empty() {
                return Self::Blank;
            }
            return Self::Text {
                operator: FormulaComparisonOperator::Equal,
                pattern: pattern.to_string(),
            };
        }
        Self::Text {
            operator: FormulaComparisonOperator::Equal,
            pattern: literal,
        }
    }
}

#[derive(Debug, Clone)]
struct FormulaCriteriaRange {
    sheet_id: SheetId,
    rect: Rect,
    criteria: FormulaCriteria,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct FormulaReference {
    areas: Vec<(SheetId, Rect)>,
    explicit_area_count: usize,
}

impl FormulaReference {
    fn with_explicit_area_count(
        explicit_area_count: usize,
        areas: Vec<(SheetId, Rect)>,
    ) -> Result<Self, FormulaEvalError> {
        if areas.is_empty() {
            return Err(FormulaEvalError::Ref);
        }
        if explicit_area_count == 0 {
            return Err(FormulaEvalError::Ref);
        }
        Ok(Self {
            areas,
            explicit_area_count,
        })
    }

    fn single(sheet_id: SheetId, rect: Rect) -> Self {
        Self {
            areas: vec![(sheet_id, rect)],
            explicit_area_count: 1,
        }
    }

    fn single_area(&self) -> Result<(SheetId, Rect), FormulaEvalError> {
        if self.explicit_area_count != 1 {
            return Err(FormulaEvalError::Value);
        }
        self.areas.first().copied().ok_or(FormulaEvalError::Ref)
    }

    fn len(&self) -> usize {
        self.explicit_area_count
    }

    fn areas(&self) -> &[(SheetId, Rect)] {
        &self.areas
    }
}

#[derive(Debug, Clone, PartialEq)]
enum FormulaValueProbe {
    Blank,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(FormulaEvalError),
    Omitted,
    Lambda {
        parameters: Vec<String>,
        body: String,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaLogicalFunction {
    And,
    Or,
    Xor,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaGroupByAggregation {
    Sum,
    Average,
    Count,
    CountA,
    Max,
    Min,
    Product,
}

impl FormulaGroupByAggregation {
    fn from_name(name: &str) -> Option<Self> {
        if name.eq_ignore_ascii_case("SUM") {
            Some(Self::Sum)
        } else if name.eq_ignore_ascii_case("AVERAGE") {
            Some(Self::Average)
        } else if name.eq_ignore_ascii_case("COUNT") {
            Some(Self::Count)
        } else if name.eq_ignore_ascii_case("COUNTA") {
            Some(Self::CountA)
        } else if name.eq_ignore_ascii_case("MAX") {
            Some(Self::Max)
        } else if name.eq_ignore_ascii_case("MIN") {
            Some(Self::Min)
        } else if name.eq_ignore_ascii_case("PRODUCT") {
            Some(Self::Product)
        } else {
            None
        }
    }

    fn evaluate(self, values: &[FormulaValueProbe]) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut numbers = Vec::new();
        let mut counta = 0_u64;
        for value in values {
            match value {
                FormulaValueProbe::Blank => {}
                FormulaValueProbe::Bool(_) | FormulaValueProbe::Text(_) => counta += 1,
                FormulaValueProbe::Number(number) => {
                    counta += 1;
                    numbers.push(*number);
                }
                FormulaValueProbe::Error(error) => return Err(*error),
                FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {}
            }
        }
        match self {
            Self::Sum => Ok(FormulaValueProbe::Number(numbers.iter().sum())),
            Self::Average => {
                if numbers.is_empty() {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok(FormulaValueProbe::Number(
                        numbers.iter().sum::<f64>() / numbers.len() as f64,
                    ))
                }
            }
            Self::Count => Ok(FormulaValueProbe::Number(numbers.len() as f64)),
            Self::CountA => Ok(FormulaValueProbe::Number(counta as f64)),
            Self::Max => Ok(FormulaValueProbe::Number(
                numbers.iter().copied().reduce(f64::max).unwrap_or(0.0),
            )),
            Self::Min => Ok(FormulaValueProbe::Number(
                numbers.iter().copied().reduce(f64::min).unwrap_or(0.0),
            )),
            Self::Product => Ok(FormulaValueProbe::Number(numbers.iter().product())),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaLookupMode {
    Exact,
    ApproxAscending,
    ApproxDescending,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaXLookupMatchMode {
    Exact,
    ExactOrNextSmaller,
    ExactOrNextLarger,
    Wildcard,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaXLookupSearchMode {
    Forward,
    Reverse,
    BinaryAscending,
    BinaryDescending,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaLookupOrientation {
    FirstColumn,
    FirstRow,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FormulaWildcardToken {
    Literal(char),
    AnyChar,
    AnySequence,
}

fn formula_value_probe_from_cell_value(value: CellValue) -> FormulaValueProbe {
    match value {
        CellValue::Blank => FormulaValueProbe::Blank,
        CellValue::Bool(value) => FormulaValueProbe::Bool(value),
        CellValue::Number(value) => FormulaValueProbe::Number(value),
        CellValue::Text(value) => FormulaValueProbe::Text(value),
        CellValue::Error(error) => {
            FormulaValueProbe::Error(formula_eval_error_from_cell_error(error))
        }
        CellValue::IsoDateTime(value) => FormulaValueProbe::Text(value.into_string()),
        CellValue::RichText(value) => FormulaValueProbe::Text(value.into_string()),
    }
}

fn formula_cell_value_from_probe(value: FormulaValueProbe) -> Option<CellValue> {
    match value {
        FormulaValueProbe::Blank => Some(CellValue::Blank),
        FormulaValueProbe::Bool(value) => Some(CellValue::Bool(value)),
        FormulaValueProbe::Number(value) => Some(CellValue::Number(value)),
        FormulaValueProbe::Text(value) => Some(CellValue::Text(value)),
        FormulaValueProbe::Error(error) => error.into_cell_value(),
        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => None,
    }
}

fn formula_number_from_value_probe(value: FormulaValueProbe) -> Result<f64, FormulaEvalError> {
    match value {
        FormulaValueProbe::Blank => Ok(0.0),
        FormulaValueProbe::Bool(value) => Ok(if value { 1.0 } else { 0.0 }),
        FormulaValueProbe::Number(value) => Ok(value),
        FormulaValueProbe::Text(_) => Err(FormulaEvalError::Value),
        FormulaValueProbe::Error(error) => Err(error),
        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
            Err(FormulaEvalError::Value)
        }
    }
}

fn formula_text_from_value_probe(value: FormulaValueProbe) -> Result<String, FormulaEvalError> {
    match value {
        FormulaValueProbe::Blank => Ok(String::new()),
        FormulaValueProbe::Bool(value) => Ok(if value { "TRUE" } else { "FALSE" }.into()),
        FormulaValueProbe::Number(value) => formula_text_from_number(value),
        FormulaValueProbe::Text(value) => Ok(value),
        FormulaValueProbe::Error(error) => Err(error),
        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
            Err(FormulaEvalError::Value)
        }
    }
}

fn formula_integer_argument(value: f64) -> Result<i64, FormulaEvalError> {
    if !value.is_finite() || value.fract() != 0.0 {
        return Err(FormulaEvalError::Value);
    }
    if value < i64::MIN as f64 || value > i64::MAX as f64 {
        return Err(FormulaEvalError::Num);
    }
    Ok(value as i64)
}

fn formula_source_has_top_level_function(formula: &FormulaSource, name: &str) -> bool {
    let text = formula.text.trim_start();
    let Some(prefix) = text.get(..name.len()) else {
        return false;
    };
    if !prefix.eq_ignore_ascii_case(name) {
        return false;
    }
    text[name.len()..].trim_start().starts_with('(')
}

fn formula_radix_argument(value: f64) -> Result<u32, FormulaEvalError> {
    let value = formula_integer_argument(value)?;
    if !(2..=36).contains(&value) {
        return Err(FormulaEvalError::Num);
    }
    u32::try_from(value).map_err(|_| FormulaEvalError::Num)
}

fn formula_non_negative_count_argument(value: f64) -> Result<usize, FormulaEvalError> {
    let count = formula_integer_argument(value)?;
    if count < 0 {
        return Err(FormulaEvalError::Value);
    }
    usize::try_from(count).map_err(|_| FormulaEvalError::Num)
}

fn formula_positive_position_argument(value: f64) -> Result<usize, FormulaEvalError> {
    let position = formula_integer_argument(value)?;
    if position < 1 {
        return Err(FormulaEvalError::Value);
    }
    usize::try_from(position).map_err(|_| FormulaEvalError::Num)
}

fn formula_xlookup_match_mode_argument(
    value: f64,
) -> Result<FormulaXLookupMatchMode, FormulaEvalError> {
    match formula_integer_argument(value)? {
        0 => Ok(FormulaXLookupMatchMode::Exact),
        -1 => Ok(FormulaXLookupMatchMode::ExactOrNextSmaller),
        1 => Ok(FormulaXLookupMatchMode::ExactOrNextLarger),
        2 => Ok(FormulaXLookupMatchMode::Wildcard),
        _ => Err(FormulaEvalError::Value),
    }
}

fn formula_xlookup_search_mode_argument(
    value: f64,
) -> Result<FormulaXLookupSearchMode, FormulaEvalError> {
    match formula_integer_argument(value)? {
        1 => Ok(FormulaXLookupSearchMode::Forward),
        -1 => Ok(FormulaXLookupSearchMode::Reverse),
        2 => Ok(FormulaXLookupSearchMode::BinaryAscending),
        -2 => Ok(FormulaXLookupSearchMode::BinaryDescending),
        _ => Err(FormulaEvalError::Value),
    }
}

fn div_floor(value: i64, divisor: i64) -> i64 {
    let quotient = value / divisor;
    let remainder = value % divisor;
    if remainder != 0 && ((remainder > 0) != (divisor > 0)) {
        quotient - 1
    } else {
        quotient
    }
}

struct FormulaParser<'a, 'b, 'state> {
    input: &'a str,
    index: usize,
    evaluator: &'b mut FormulaEvaluator<'state>,
    sheet_id: SheetId,
    current_position: Option<(u32, u32)>,
    bindings: Vec<(String, FormulaValueProbe)>,
}

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    fn new(
        input: &'a str,
        evaluator: &'b mut FormulaEvaluator<'state>,
        sheet_id: SheetId,
        current_position: Option<(u32, u32)>,
    ) -> Self {
        Self {
            input: input.trim().strip_prefix('=').unwrap_or(input.trim()),
            index: 0,
            evaluator,
            sheet_id,
            current_position,
            bindings: Vec::new(),
        }
    }

    fn parse_dynamic_array_formula(&mut self) -> Result<FormulaArrayResult, FormulaEvalError> {
        self.skip_whitespace();
        let Some(identifier) = self.parse_identifier() else {
            return Err(FormulaEvalError::Unsupported);
        };
        if ![
            "FILTER",
            "SORT",
            "SORTBY",
            "UNIQUE",
            "TAKE",
            "DROP",
            "CHOOSECOLS",
            "CHOOSEROWS",
            "TRANSPOSE",
            "SEQUENCE",
            "EXPAND",
            "HSTACK",
            "VSTACK",
            "TOCOL",
            "TOROW",
            "WRAPROWS",
            "WRAPCOLS",
        ]
        .iter()
        .any(|name| identifier.eq_ignore_ascii_case(name))
        {
            return Err(FormulaEvalError::Unsupported);
        }
        self.skip_whitespace();
        if !self.consume_char('(') {
            return Err(FormulaEvalError::Unsupported);
        }

        if identifier.eq_ignore_ascii_case("SEQUENCE") {
            let rows = formula_integer_argument(self.parse_comparison()?)?;
            if rows < 1 {
                return Err(FormulaEvalError::Value);
            }
            let mut cols = 1_i64;
            let mut start = 1.0_f64;
            let mut step = 1.0_f64;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    cols = formula_integer_argument(self.parse_comparison()?)?;
                    if cols < 1 {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        start = self.parse_comparison()?;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        step = self.parse_comparison()?;
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() || !start.is_finite() || !step.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            let rows = usize::try_from(rows).map_err(|_| FormulaEvalError::Value)?;
            let cols = usize::try_from(cols).map_err(|_| FormulaEvalError::Value)?;
            let len = rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?;
            let values = (0..len)
                .map(|index| {
                    formula_checked_numeric_result(start + step * index as f64)
                        .map(CellValue::Number)
                })
                .collect::<Result<Vec<_>, _>>()?;
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        let (source_sheet_id, source_rect) = self.parse_reference_argument()?;
        let source_rows = source_rect.height() as usize;
        let source_cols = source_rect.width() as usize;
        let mut source_values = Vec::with_capacity(
            source_rect
                .checked_cell_count_usize()
                .map_err(|_| FormulaEvalError::Num)?,
        );
        for row in source_rect.row_first..=source_rect.row_last {
            for col in source_rect.col_first..=source_rect.col_last {
                source_values.push(self.evaluator.cell_value_or_blank(
                    source_sheet_id,
                    row,
                    col,
                )?);
            }
        }

        if identifier.eq_ignore_ascii_case("FILTER") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (include_sheet_id, include_rect) = self.parse_reference_argument()?;
            self.skip_whitespace();
            let if_empty = if self.consume_char(')') {
                None
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let value = self.parse_value_probe_argument()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
                Some(value)
            };
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }

            let include_value = |value: CellValue| -> Result<bool, FormulaEvalError> {
                match value {
                    CellValue::Blank => Ok(false),
                    CellValue::Bool(value) => Ok(value),
                    CellValue::Number(value) => Ok(value != 0.0),
                    CellValue::Text(_) | CellValue::IsoDateTime(_) | CellValue::RichText(_) => {
                        Err(FormulaEvalError::Value)
                    }
                    CellValue::Error(error) => Err(formula_eval_error_from_cell_error(error)),
                }
            };
            let mut values = Vec::new();
            let (rows, cols) = if include_rect.height() == source_rect.height()
                && include_rect.width() == 1
            {
                let mut selected_rows = 0_usize;
                for row_offset in 0..source_rect.height() {
                    let include = self.evaluator.cell_value_or_blank(
                        include_sheet_id,
                        include_rect.row_first + row_offset,
                        include_rect.col_first,
                    )?;
                    if include_value(include)? {
                        selected_rows += 1;
                        let source_start = row_offset as usize * source_cols;
                        values.extend_from_slice(
                            &source_values[source_start..source_start + source_cols],
                        );
                    }
                }
                (selected_rows, source_cols)
            } else if include_rect.width() == source_rect.width() && include_rect.height() == 1 {
                let selected_columns = (0..source_rect.width())
                    .map(|col_offset| {
                        self.evaluator
                            .cell_value_or_blank(
                                include_sheet_id,
                                include_rect.row_first,
                                include_rect.col_first + col_offset,
                            )
                            .and_then(&include_value)
                            .map(|include| (col_offset as usize, include))
                    })
                    .collect::<Result<Vec<_>, _>>()?
                    .into_iter()
                    .filter_map(|(col_offset, include)| include.then_some(col_offset))
                    .collect::<Vec<_>>();
                for row_offset in 0..source_rows {
                    for &col_offset in &selected_columns {
                        values.push(source_values[row_offset * source_cols + col_offset].clone());
                    }
                }
                (source_rows, selected_columns.len())
            } else {
                return Err(FormulaEvalError::Value);
            };

            if rows == 0 || cols == 0 {
                return if_empty
                    .and_then(formula_cell_value_from_probe)
                    .map(FormulaArrayResult::single)
                    .ok_or(FormulaEvalError::Calc);
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("EXPAND") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let rows = formula_integer_argument(self.parse_comparison()?)?;
            if rows < source_rows as i64 {
                return Err(FormulaEvalError::Value);
            }
            let mut cols = source_cols as i64;
            let mut pad = CellValue::Error(CellError::NA);
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    cols = formula_integer_argument(self.parse_comparison()?)?;
                }
                if cols < source_cols as i64 {
                    return Err(FormulaEvalError::Value);
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    pad = formula_cell_value_from_probe(self.parse_value_probe_argument()?)
                        .ok_or(FormulaEvalError::Value)?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() || rows < 1 || cols < 1 {
                return Err(FormulaEvalError::Value);
            }
            let rows = usize::try_from(rows).map_err(|_| FormulaEvalError::Value)?;
            let cols = usize::try_from(cols).map_err(|_| FormulaEvalError::Value)?;
            let mut values = vec![pad; rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?];
            for row in 0..source_rows {
                for col in 0..source_cols {
                    values[row * cols + col] = source_values[row * source_cols + col].clone();
                }
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("HSTACK") || identifier.eq_ignore_ascii_case("VSTACK") {
            let mut matrices = vec![(source_rows, source_cols, source_values)];
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    break;
                }
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let (sheet_id, rect) = self.parse_reference_argument()?;
                let rows = rect.height() as usize;
                let cols = rect.width() as usize;
                let mut values = Vec::with_capacity(
                    rect.checked_cell_count_usize()
                        .map_err(|_| FormulaEvalError::Num)?,
                );
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        values.push(self.evaluator.cell_value_or_blank(sheet_id, row, col)?);
                    }
                }
                matrices.push((rows, cols, values));
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            if identifier.eq_ignore_ascii_case("HSTACK") {
                let rows = matrices
                    .iter()
                    .map(|(rows, _, _)| *rows)
                    .max()
                    .ok_or(FormulaEvalError::Calc)?;
                let cols = matrices.iter().try_fold(0_usize, |total, (_, cols, _)| {
                    total.checked_add(*cols).ok_or(FormulaEvalError::Num)
                })?;
                let mut values = vec![
                    CellValue::Error(CellError::NA);
                    rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?
                ];
                let mut col_start = 0_usize;
                for (matrix_rows, matrix_cols, matrix_values) in matrices {
                    for row in 0..matrix_rows {
                        for col in 0..matrix_cols {
                            values[row * cols + col_start + col] =
                                matrix_values[row * matrix_cols + col].clone();
                        }
                    }
                    col_start += matrix_cols;
                }
                return Ok(FormulaArrayResult { rows, cols, values });
            }
            let rows = matrices.iter().try_fold(0_usize, |total, (rows, _, _)| {
                total.checked_add(*rows).ok_or(FormulaEvalError::Num)
            })?;
            let cols = matrices
                .iter()
                .map(|(_, cols, _)| *cols)
                .max()
                .ok_or(FormulaEvalError::Calc)?;
            let mut values = vec![
                CellValue::Error(CellError::NA);
                rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?
            ];
            let mut row_start = 0_usize;
            for (matrix_rows, matrix_cols, matrix_values) in matrices {
                for row in 0..matrix_rows {
                    for col in 0..matrix_cols {
                        values[(row_start + row) * cols + col] =
                            matrix_values[row * matrix_cols + col].clone();
                    }
                }
                row_start += matrix_rows;
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("TOCOL") || identifier.eq_ignore_ascii_case("TOROW") {
            let mut ignore = 0_i64;
            let mut scan_by_column = false;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    ignore = formula_integer_argument(self.parse_comparison()?)?;
                }
                if !(0..=3).contains(&ignore) {
                    return Err(FormulaEvalError::Value);
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    scan_by_column = self.parse_comparison()? != 0.0;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let include_value = |value: &CellValue| {
                let ignore_blank = matches!(ignore, 1 | 3);
                let ignore_error = matches!(ignore, 2 | 3);
                !(ignore_blank && matches!(value, CellValue::Blank))
                    && !(ignore_error && matches!(value, CellValue::Error(_)))
            };
            let mut values = Vec::with_capacity(source_values.len());
            if scan_by_column {
                for col in 0..source_cols {
                    for row in 0..source_rows {
                        let value = &source_values[row * source_cols + col];
                        if include_value(value) {
                            values.push(value.clone());
                        }
                    }
                }
            } else {
                values.extend(source_values.into_iter().filter(include_value));
            }
            if values.is_empty() {
                return Err(FormulaEvalError::Calc);
            }
            let (rows, cols) = if identifier.eq_ignore_ascii_case("TOCOL") {
                (values.len(), 1)
            } else {
                (1, values.len())
            };
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("WRAPROWS")
            || identifier.eq_ignore_ascii_case("WRAPCOLS")
        {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let wrap_count = formula_integer_argument(self.parse_comparison()?)?;
            if wrap_count < 1 {
                return Err(FormulaEvalError::Value);
            }
            let mut pad = CellValue::Error(CellError::NA);
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                pad = formula_cell_value_from_probe(self.parse_value_probe_argument()?)
                    .ok_or(FormulaEvalError::Value)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let wrap_count = usize::try_from(wrap_count).map_err(|_| FormulaEvalError::Value)?;
            let chunk_count = source_values.len().div_ceil(wrap_count);
            let (rows, cols) = if identifier.eq_ignore_ascii_case("WRAPROWS") {
                (chunk_count, wrap_count)
            } else {
                (wrap_count, chunk_count)
            };
            let mut values = vec![pad; rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?];
            if identifier.eq_ignore_ascii_case("WRAPROWS") {
                values[..source_values.len()].clone_from_slice(&source_values);
            } else {
                for (index, value) in source_values.into_iter().enumerate() {
                    let row = index % wrap_count;
                    let col = index / wrap_count;
                    values[row * cols + col] = value;
                }
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("SORT") || identifier.eq_ignore_ascii_case("SORTBY") {
            let compare_values =
                |left: &CellValue, right: &CellValue| -> Result<Ordering, FormulaEvalError> {
                    let left = formula_value_probe_from_cell_value(left.clone());
                    let right = formula_value_probe_from_cell_value(right.clone());
                    if let FormulaValueProbe::Error(error) = left {
                        return Err(error);
                    }
                    if let FormulaValueProbe::Error(error) = right {
                        return Err(error);
                    }
                    if let Some(ordering) = formula_value_probe_ordering(&left, &right)? {
                        return Ok(ordering);
                    }
                    let rank = |value: &FormulaValueProbe| match value {
                        FormulaValueProbe::Blank => 0,
                        FormulaValueProbe::Number(_) => 1,
                        FormulaValueProbe::Text(_) => 2,
                        FormulaValueProbe::Bool(_) => 3,
                        FormulaValueProbe::Error(_) => 4,
                        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => 5,
                    };
                    Ok(rank(&left).cmp(&rank(&right)))
                };
            let mut key_sets = Vec::<(Vec<CellValue>, bool)>::new();
            let by_column;
            if identifier.eq_ignore_ascii_case("SORT") {
                let mut sort_index = 1_i64;
                let mut sort_order = 1_i64;
                let mut sort_by_column = false;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        sort_index = formula_integer_argument(self.parse_comparison()?)?;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                            sort_order = formula_integer_argument(self.parse_comparison()?)?;
                        }
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            if !self.consume_char(',') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                            self.skip_whitespace();
                            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                                sort_by_column = self.parse_comparison()? != 0.0;
                            }
                            self.skip_whitespace();
                            if !self.consume_char(')') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                        }
                    }
                }
                if sort_index < 1 || !matches!(sort_order, -1 | 1) {
                    return Err(FormulaEvalError::Value);
                }
                by_column = sort_by_column;
                let sort_index =
                    usize::try_from(sort_index - 1).map_err(|_| FormulaEvalError::Value)?;
                let keys = if by_column {
                    if sort_index >= source_rows {
                        return Err(FormulaEvalError::Value);
                    }
                    (0..source_cols)
                        .map(|col| source_values[sort_index * source_cols + col].clone())
                        .collect()
                } else {
                    if sort_index >= source_cols {
                        return Err(FormulaEvalError::Value);
                    }
                    (0..source_rows)
                        .map(|row| source_values[row * source_cols + sort_index].clone())
                        .collect()
                };
                key_sets.push((keys, sort_order == -1));
            } else {
                self.skip_whitespace();
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let mut orientation = None;
                loop {
                    let (key_sheet_id, key_rect) = self.parse_reference_argument()?;
                    let current_by_column = if key_rect.height() as usize == source_rows
                        && key_rect.width() == 1
                    {
                        false
                    } else if key_rect.width() as usize == source_cols && key_rect.height() == 1 {
                        true
                    } else {
                        return Err(FormulaEvalError::Value);
                    };
                    if orientation.is_some_and(|value| value != current_by_column) {
                        return Err(FormulaEvalError::Value);
                    }
                    orientation = Some(current_by_column);
                    let keys = if current_by_column {
                        (key_rect.col_first..=key_rect.col_last)
                            .map(|col| {
                                self.evaluator.cell_value_or_blank(
                                    key_sheet_id,
                                    key_rect.row_first,
                                    col,
                                )
                            })
                            .collect::<Result<Vec<_>, _>>()?
                    } else {
                        (key_rect.row_first..=key_rect.row_last)
                            .map(|row| {
                                self.evaluator.cell_value_or_blank(
                                    key_sheet_id,
                                    row,
                                    key_rect.col_first,
                                )
                            })
                            .collect::<Result<Vec<_>, _>>()?
                    };
                    self.skip_whitespace();
                    let sort_order = if self.consume_char(')') {
                        key_sets.push((keys, false));
                        break;
                    } else {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                            1
                        } else {
                            formula_integer_argument(self.parse_comparison()?)?
                        }
                    };
                    if !matches!(sort_order, -1 | 1) {
                        return Err(FormulaEvalError::Value);
                    }
                    key_sets.push((keys, sort_order == -1));
                    self.skip_whitespace();
                    if self.consume_char(')') {
                        break;
                    }
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                }
                by_column = orientation.unwrap_or(false);
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }

            let mut indexes = if by_column {
                (0..source_cols).collect::<Vec<_>>()
            } else {
                (0..source_rows).collect::<Vec<_>>()
            };
            for index in 1..indexes.len() {
                let current = indexes[index];
                let mut position = index;
                while position > 0 {
                    let previous = indexes[position - 1];
                    let mut ordering = Ordering::Equal;
                    for (keys, descending) in &key_sets {
                        ordering = compare_values(&keys[previous], &keys[current])?;
                        if *descending {
                            ordering = ordering.reverse();
                        }
                        if ordering != Ordering::Equal {
                            break;
                        }
                    }
                    if ordering != Ordering::Greater {
                        break;
                    }
                    indexes[position] = previous;
                    position -= 1;
                }
                indexes[position] = current;
            }
            let mut values = Vec::with_capacity(source_values.len());
            if by_column {
                for row in 0..source_rows {
                    for &col in &indexes {
                        values.push(source_values[row * source_cols + col].clone());
                    }
                }
            } else {
                for &row in &indexes {
                    values.extend_from_slice(
                        &source_values[row * source_cols..(row + 1) * source_cols],
                    );
                }
            }
            return Ok(FormulaArrayResult {
                rows: source_rows,
                cols: source_cols,
                values,
            });
        }

        if identifier.eq_ignore_ascii_case("UNIQUE") {
            let mut by_column = false;
            let mut exactly_once = false;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    by_column = self.parse_comparison()? != 0.0;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        exactly_once = self.parse_comparison()? != 0.0;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let record_count = if by_column { source_cols } else { source_rows };
            let record_len = if by_column { source_rows } else { source_cols };
            let records_equal = |left: usize, right: usize| -> Result<bool, FormulaEvalError> {
                for offset in 0..record_len {
                    let left_value = if by_column {
                        &source_values[offset * source_cols + left]
                    } else {
                        &source_values[left * source_cols + offset]
                    };
                    let right_value = if by_column {
                        &source_values[offset * source_cols + right]
                    } else {
                        &source_values[right * source_cols + offset]
                    };
                    if !formula_value_probe_exact_match(
                        &formula_value_probe_from_cell_value(left_value.clone()),
                        &formula_value_probe_from_cell_value(right_value.clone()),
                    )? {
                        return Ok(false);
                    }
                }
                Ok(true)
            };
            let mut selected = Vec::new();
            for candidate in 0..record_count {
                let mut prior_match = false;
                let mut count = 0_usize;
                for other in 0..record_count {
                    if records_equal(candidate, other)? {
                        count += 1;
                        if other < candidate {
                            prior_match = true;
                        }
                    }
                }
                if (exactly_once && count == 1) || (!exactly_once && !prior_match) {
                    selected.push(candidate);
                }
            }
            if selected.is_empty() {
                return Err(FormulaEvalError::Calc);
            }
            let mut values = Vec::with_capacity(
                selected
                    .len()
                    .checked_mul(record_len)
                    .ok_or(FormulaEvalError::Num)?,
            );
            if by_column {
                for row in 0..source_rows {
                    for &col in &selected {
                        values.push(source_values[row * source_cols + col].clone());
                    }
                }
                return Ok(FormulaArrayResult {
                    rows: source_rows,
                    cols: selected.len(),
                    values,
                });
            }
            for &row in &selected {
                values
                    .extend_from_slice(&source_values[row * source_cols..(row + 1) * source_cols]);
            }
            return Ok(FormulaArrayResult {
                rows: selected.len(),
                cols: source_cols,
                values,
            });
        }

        if identifier.eq_ignore_ascii_case("TAKE") || identifier.eq_ignore_ascii_case("DROP") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let row_count = formula_integer_argument(self.parse_comparison()?)?;
            let mut col_count = if identifier.eq_ignore_ascii_case("TAKE") {
                source_cols as i64
            } else {
                0
            };
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                col_count = formula_integer_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let (row_start, rows, col_start, cols) = if identifier.eq_ignore_ascii_case("TAKE") {
                if row_count == 0
                    || row_count.unsigned_abs() > source_rows as u64
                    || col_count == 0
                    || col_count.unsigned_abs() > source_cols as u64
                {
                    return Err(FormulaEvalError::Calc);
                }
                let rows = row_count.unsigned_abs() as usize;
                let cols = col_count.unsigned_abs() as usize;
                (
                    if row_count < 0 { source_rows - rows } else { 0 },
                    rows,
                    if col_count < 0 { source_cols - cols } else { 0 },
                    cols,
                )
            } else {
                if row_count.unsigned_abs() >= source_rows as u64
                    || col_count.unsigned_abs() >= source_cols as u64
                {
                    return Err(FormulaEvalError::Calc);
                }
                let dropped_rows = row_count.unsigned_abs() as usize;
                let dropped_cols = col_count.unsigned_abs() as usize;
                (
                    if row_count > 0 { dropped_rows } else { 0 },
                    source_rows - dropped_rows,
                    if col_count > 0 { dropped_cols } else { 0 },
                    source_cols - dropped_cols,
                )
            };
            let mut values =
                Vec::with_capacity(rows.checked_mul(cols).ok_or(FormulaEvalError::Num)?);
            for row in row_start..row_start + rows {
                values.extend_from_slice(
                    &source_values
                        [row * source_cols + col_start..row * source_cols + col_start + cols],
                );
            }
            return Ok(FormulaArrayResult { rows, cols, values });
        }

        if identifier.eq_ignore_ascii_case("CHOOSECOLS")
            || identifier.eq_ignore_ascii_case("CHOOSEROWS")
        {
            let selecting_columns = identifier.eq_ignore_ascii_case("CHOOSECOLS");
            let size = if selecting_columns {
                source_cols
            } else {
                source_rows
            };
            let mut selected = Vec::new();
            loop {
                self.skip_whitespace();
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let index = formula_integer_argument(self.parse_comparison()?)?;
                if index == 0 || index.unsigned_abs() > size as u64 {
                    return Err(FormulaEvalError::Value);
                }
                selected.push(if index > 0 {
                    usize::try_from(index - 1).map_err(|_| FormulaEvalError::Value)?
                } else {
                    size - index.unsigned_abs() as usize
                });
                self.skip_whitespace();
                if self.consume_char(')') {
                    break;
                }
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let mut values = Vec::new();
            if selecting_columns {
                values.reserve(
                    source_rows
                        .checked_mul(selected.len())
                        .ok_or(FormulaEvalError::Num)?,
                );
                for row in 0..source_rows {
                    for &col in &selected {
                        values.push(source_values[row * source_cols + col].clone());
                    }
                }
                return Ok(FormulaArrayResult {
                    rows: source_rows,
                    cols: selected.len(),
                    values,
                });
            }
            values.reserve(
                selected
                    .len()
                    .checked_mul(source_cols)
                    .ok_or(FormulaEvalError::Num)?,
            );
            for &row in &selected {
                values
                    .extend_from_slice(&source_values[row * source_cols..(row + 1) * source_cols]);
            }
            return Ok(FormulaArrayResult {
                rows: selected.len(),
                cols: source_cols,
                values,
            });
        }

        if identifier.eq_ignore_ascii_case("TRANSPOSE") {
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if self.index != self.input.len() {
                return Err(FormulaEvalError::Unsupported);
            }
            let mut values = Vec::with_capacity(source_values.len());
            for col in 0..source_cols {
                for row in 0..source_rows {
                    values.push(source_values[row * source_cols + col].clone());
                }
            }
            return Ok(FormulaArrayResult {
                rows: source_cols,
                cols: source_rows,
                values,
            });
        }

        Err(FormulaEvalError::Unsupported)
    }

    fn binding_value(&self, name: &str) -> Option<FormulaValueProbe> {
        self.bindings
            .iter()
            .rev()
            .find(|(binding_name, _)| binding_name.eq_ignore_ascii_case(name))
            .map(|(_, value)| value.clone())
    }

    fn parse_text_formula(&mut self) -> Result<String, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(text) = self.parse_string_literal()? {
            self.skip_whitespace();
            return if self.index == self.input.len() {
                Ok(text)
            } else {
                Err(FormulaEvalError::Unsupported)
            };
        }
        let Some(identifier) = self.parse_identifier() else {
            return Err(FormulaEvalError::Unsupported);
        };
        if !formula_text_function_name(identifier.as_str()) {
            self.skip_whitespace();
            if !self.consume_char('(') {
                return Err(FormulaEvalError::Unsupported);
            }
            let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? else {
                return Err(FormulaEvalError::Unsupported);
            };
            self.skip_whitespace();
            return if self.index == self.input.len() {
                match value {
                    FormulaValueProbe::Text(text) => Ok(text),
                    FormulaValueProbe::Error(error) => Err(error),
                    _ => Err(FormulaEvalError::Unsupported),
                }
            } else {
                Err(FormulaEvalError::Unsupported)
            };
        }
        self.skip_whitespace();
        if !self.consume_char('(') {
            return Err(FormulaEvalError::Unsupported);
        }
        let value = self.parse_text_function(identifier.as_str())?;
        self.skip_whitespace();
        if self.index == self.input.len() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Unsupported)
        }
    }

    fn parse_formula(&mut self) -> Result<f64, FormulaEvalError> {
        let value = self.parse_comparison()?;
        self.skip_whitespace();
        if self.index == self.input.len() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Unsupported)
        }
    }

    fn parse_value_probe_formula(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if self.index == self.input.len() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Unsupported)
        }
    }

    fn parse_comparison(&mut self) -> Result<f64, FormulaEvalError> {
        let mut value = self.parse_expression()?;
        loop {
            self.skip_whitespace();
            let Some(operator) = self.consume_comparison_operator() else {
                return Ok(value);
            };
            let right = self.parse_expression()?;
            value = if operator.evaluate(value, right) {
                1.0
            } else {
                0.0
            };
        }
    }

    fn parse_expression(&mut self) -> Result<f64, FormulaEvalError> {
        let mut value = self.parse_term()?;
        loop {
            self.skip_whitespace();
            if self.consume_char('+') {
                value += self.parse_term()?;
            } else if self.consume_char('-') {
                value -= self.parse_term()?;
            } else {
                return Ok(value);
            }
        }
    }

    fn parse_term(&mut self) -> Result<f64, FormulaEvalError> {
        let mut value = self.parse_factor()?;
        loop {
            self.skip_whitespace();
            if self.consume_char('*') {
                value *= self.parse_factor()?;
            } else if self.consume_char('/') {
                let divisor = self.parse_factor()?;
                if divisor == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                value /= divisor;
            } else {
                return Ok(value);
            }
        }
    }

    fn parse_factor(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char('+') {
            return self.parse_factor();
        }
        if self.consume_char('-') {
            return Ok(-self.parse_factor()?);
        }
        if self.consume_char('(') {
            let value = self.parse_expression()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(value);
        }
        if let Some(number) = self.parse_number()? {
            return Ok(number);
        }
        if let Some(error) = self.parse_error_literal() {
            return Err(error);
        }
        let checkpoint = self.index;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                if let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? {
                    return formula_number_from_value_probe(value);
                }
                return self.parse_function(identifier.as_str());
            }
        }
        self.index = checkpoint;
        if let Some((reference, next_index)) = self.try_parse_reference_set()? {
            self.index = next_index;
            let (target_sheet_id, rect) = reference.single_area()?;
            if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
                return self.evaluator.numeric_cell_value(
                    target_sheet_id,
                    rect.row_first,
                    rect.col_first,
                );
            }
            return Err(FormulaEvalError::Value);
        }
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                if let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? {
                    return formula_number_from_value_probe(value);
                }
                return self.parse_function(identifier.as_str());
            }
            if identifier.eq_ignore_ascii_case("TRUE") {
                return Ok(1.0);
            }
            if identifier.eq_ignore_ascii_case("FALSE") {
                return Ok(0.0);
            }
            if let Some(value) = self.binding_value(identifier.as_str()) {
                return formula_number_from_value_probe(value);
            }
            if let Some(value) = self.defined_name_value_probe(identifier.as_str())? {
                return formula_number_from_value_probe(value);
            }
            return Err(FormulaEvalError::Name);
        }
        Err(FormulaEvalError::Unsupported)
    }

    /// Consumes an error literal such as `#N/A` or `#VALUE!` and returns the error it denotes.
    fn parse_error_literal(&mut self) -> Option<FormulaEvalError> {
        const LITERALS: [(&str, FormulaEvalError); 8] = [
            ("#NULL!", FormulaEvalError::Null),
            ("#DIV/0!", FormulaEvalError::Div0),
            ("#VALUE!", FormulaEvalError::Value),
            ("#REF!", FormulaEvalError::Ref),
            ("#NAME?", FormulaEvalError::Name),
            ("#NUM!", FormulaEvalError::Num),
            ("#N/A", FormulaEvalError::NA),
            ("#CALC!", FormulaEvalError::Calc),
        ];
        let rest = &self.input[self.index..];
        let (literal, error) = LITERALS.into_iter().find(|(literal, _)| {
            rest.get(..literal.len())
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case(literal))
        })?;
        self.index += literal.len();
        Some(error)
    }

    fn parse_function(&mut self, name: &str) -> Result<f64, FormulaEvalError> {
        let percentile_value =
            |mut values: Vec<f64>, k: f64, exclusive: bool| -> Result<f64, FormulaEvalError> {
                if values.is_empty() || !k.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
                values.sort_by(|left, right| left.total_cmp(right));
                if exclusive {
                    if k <= 0.0 || k >= 1.0 {
                        return Err(FormulaEvalError::Num);
                    }
                    let rank = k * (values.len() as f64 + 1.0);
                    if rank < 1.0 || rank > values.len() as f64 {
                        return Err(FormulaEvalError::Num);
                    }
                    let lower_rank = rank.floor();
                    let upper_rank = rank.ceil();
                    if lower_rank == upper_rank {
                        return Ok(values[lower_rank as usize - 1]);
                    }
                    let lower_index = lower_rank as usize - 1;
                    let upper_index = upper_rank as usize - 1;
                    let fraction = rank - lower_rank;
                    return Ok(values[lower_index]
                        + (values[upper_index] - values[lower_index]) * fraction);
                }
                if !(0.0..=1.0).contains(&k) {
                    return Err(FormulaEvalError::Num);
                }
                let rank = k * (values.len() as f64 - 1.0);
                let lower_index = rank.floor() as usize;
                let upper_index = rank.ceil() as usize;
                if lower_index == upper_index {
                    return Ok(values[lower_index]);
                }
                let fraction = rank - lower_index as f64;
                Ok(values[lower_index] + (values[upper_index] - values[lower_index]) * fraction)
            };
        let percent_rank_value = |mut values: Vec<f64>,
                                  x: f64,
                                  significance: i64,
                                  exclusive: bool|
         -> Result<f64, FormulaEvalError> {
            if values.is_empty() || !x.is_finite() {
                return Err(FormulaEvalError::Num);
            }
            if significance < 1 {
                return Err(FormulaEvalError::Num);
            }
            let significance = i32::try_from(significance).map_err(|_| FormulaEvalError::Num)?;
            let factor = 10_f64.powi(significance);
            if !factor.is_finite() {
                return Err(FormulaEvalError::Num);
            }
            values.sort_by(|left, right| left.total_cmp(right));
            let Some(minimum) = values.first().copied() else {
                return Err(FormulaEvalError::Num);
            };
            let Some(maximum) = values.last().copied() else {
                return Err(FormulaEvalError::Num);
            };
            if x < minimum || x > maximum {
                return Err(FormulaEvalError::NA);
            }
            if !exclusive && values.len() == 1 {
                return Ok(0.0);
            }
            if exclusive && values.len() == 1 {
                return Ok(0.5);
            }
            let rank_for_index = |index: usize| -> f64 {
                if exclusive {
                    (index as f64 + 1.0) / (values.len() as f64 + 1.0)
                } else {
                    index as f64 / (values.len() as f64 - 1.0)
                }
            };
            let rank = if let Some(index) = values.iter().position(|value| *value == x) {
                rank_for_index(index)
            } else {
                let Some(upper_index) = values.iter().position(|value| *value > x) else {
                    return Err(FormulaEvalError::NA);
                };
                if upper_index == 0 {
                    return Err(FormulaEvalError::NA);
                }
                let lower_index = upper_index - 1;
                let lower_rank = rank_for_index(lower_index);
                let upper_rank = rank_for_index(upper_index);
                lower_rank
                    + (upper_rank - lower_rank) * (x - values[lower_index])
                        / (values[upper_index] - values[lower_index])
            };
            Ok((rank * factor).trunc() / factor)
        };
        if name.eq_ignore_ascii_case("IF") {
            return self.parse_if_function();
        }
        if name.eq_ignore_ascii_case("IFS") {
            return formula_number_from_value_probe(self.parse_ifs_value_function()?);
        }
        if name.eq_ignore_ascii_case("SWITCH") {
            return formula_number_from_value_probe(self.parse_switch_value_function()?);
        }
        if name.eq_ignore_ascii_case("AND") {
            return self.parse_logical_function(FormulaLogicalFunction::And);
        }
        if name.eq_ignore_ascii_case("OR") {
            return self.parse_logical_function(FormulaLogicalFunction::Or);
        }
        if name.eq_ignore_ascii_case("XOR") {
            return self.parse_logical_function(FormulaLogicalFunction::Xor);
        }
        if name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE") {
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(if name.eq_ignore_ascii_case("TRUE") {
                1.0
            } else {
                0.0
            });
        }
        if name.eq_ignore_ascii_case("DATEDIF") {
            return self.parse_datedif_function();
        }
        if name.eq_ignore_ascii_case("WORKDAY") {
            return self.parse_workday_function();
        }
        if name.eq_ignore_ascii_case("WORKDAY.INTL") {
            return self.parse_workday_intl_function();
        }
        if name.eq_ignore_ascii_case("NETWORKDAYS") {
            return self.parse_networkdays_function();
        }
        if name.eq_ignore_ascii_case("NETWORKDAYS.INTL") {
            return self.parse_networkdays_intl_function();
        }
        if name.eq_ignore_ascii_case("SERIESSUM") {
            return self.parse_series_sum_function();
        }
        if name.eq_ignore_ascii_case("AGGREGATE") {
            return self.parse_aggregate_function();
        }
        if name.eq_ignore_ascii_case("SUBTOTAL") {
            return self.parse_subtotal_function();
        }
        if let Some(function) = FormulaScalarFunction::from_name(name) {
            return self.parse_scalar_function(function);
        }
        if name.eq_ignore_ascii_case("COUNTA") {
            return self.parse_counta_function();
        }
        if name.eq_ignore_ascii_case("COUNTBLANK") {
            return self.parse_countblank_function();
        }
        if name.eq_ignore_ascii_case("CHOOSE") {
            return self.parse_choose_function();
        }
        if name.eq_ignore_ascii_case("COLUMN") {
            return self.parse_column_function();
        }
        if name.eq_ignore_ascii_case("COLUMNS") || name.eq_ignore_ascii_case("COLS") {
            return self.parse_columns_function();
        }
        if name.eq_ignore_ascii_case("COUNTIF") {
            return self.parse_countif_function();
        }
        if name.eq_ignore_ascii_case("IFERROR") {
            return self.parse_iferror_function();
        }
        if name.eq_ignore_ascii_case("IFNA") {
            return self.parse_ifna_function();
        }
        if name.eq_ignore_ascii_case("ISERROR") {
            return self.parse_error_test_function(true, false);
        }
        if name.eq_ignore_ascii_case("ISERR") {
            return self.parse_error_test_function(false, false);
        }
        if name.eq_ignore_ascii_case("ISNA") {
            return self.parse_error_test_function(false, true);
        }
        if name.eq_ignore_ascii_case("ISBLANK") {
            return self.parse_value_probe_test_function(|value| {
                matches!(value, FormulaValueProbe::Blank)
            });
        }
        if name.eq_ignore_ascii_case("ISNUMBER") {
            return self.parse_value_probe_test_function(|value| {
                matches!(value, FormulaValueProbe::Number(_))
            });
        }
        if name.eq_ignore_ascii_case("ISLOGICAL") {
            return self.parse_value_probe_test_function(|value| {
                matches!(value, FormulaValueProbe::Bool(_))
            });
        }
        if name.eq_ignore_ascii_case("ISNONTEXT") {
            return self.parse_value_probe_test_function(|value| {
                !matches!(value, FormulaValueProbe::Text(_))
            });
        }
        if name.eq_ignore_ascii_case("ISTEXT") {
            return self.parse_value_probe_test_function(|value| {
                matches!(value, FormulaValueProbe::Text(_))
            });
        }
        if name.eq_ignore_ascii_case("ISREF") {
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((_, _, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.consume_char(')') {
                    return Ok(1.0);
                }
            }
            self.index = checkpoint;
            let _ = self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(0.0);
        }
        if name.eq_ignore_ascii_case("ISFORMULA") {
            return self.parse_isformula_function();
        }
        if name.eq_ignore_ascii_case("TYPE") {
            return self.parse_type_function();
        }
        if name.eq_ignore_ascii_case("ERROR.TYPE") {
            return self.parse_error_type_function();
        }
        if name.eq_ignore_ascii_case("N") {
            let value = self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return match value {
                FormulaValueProbe::Blank | FormulaValueProbe::Text(_) => Ok(0.0),
                FormulaValueProbe::Bool(value) => Ok(if value { 1.0 } else { 0.0 }),
                FormulaValueProbe::Number(value) => Ok(value),
                FormulaValueProbe::Error(error) => Err(error),
                FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
                    Err(FormulaEvalError::Value)
                }
            };
        }
        if name.eq_ignore_ascii_case("DECIMAL") {
            return self.parse_decimal_function();
        }
        if name.eq_ignore_ascii_case("CONVERT") {
            return self.parse_convert_function();
        }
        if name.eq_ignore_ascii_case("EUROCONVERT") {
            return self.parse_euroconvert_function();
        }
        if name.eq_ignore_ascii_case("MDETERM") {
            return self.parse_mdeterm_function();
        }
        if name.eq_ignore_ascii_case("IMABS")
            || name.eq_ignore_ascii_case("IMAGINARY")
            || name.eq_ignore_ascii_case("IMARGUMENT")
            || name.eq_ignore_ascii_case("IMREAL")
        {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if name.eq_ignore_ascii_case("IMABS") {
                let result = value.real.hypot(value.imaginary);
                return if result.is_finite() {
                    Ok(result)
                } else {
                    Err(FormulaEvalError::Num)
                };
            }
            if name.eq_ignore_ascii_case("IMAGINARY") {
                return Ok(value.imaginary);
            }
            if name.eq_ignore_ascii_case("IMARGUMENT") {
                if value.real == 0.0 && value.imaginary == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                return Ok(value.imaginary.atan2(value.real));
            }
            return Ok(value.real);
        }
        if name.eq_ignore_ascii_case("BIN2DEC") {
            return self.parse_engineering_decimal_function(2, 10, 10);
        }
        if name.eq_ignore_ascii_case("OCT2DEC") {
            return self.parse_engineering_decimal_function(8, 30, 10);
        }
        if name.eq_ignore_ascii_case("HEX2DEC") {
            return self.parse_engineering_decimal_function(16, 40, 10);
        }
        if name.eq_ignore_ascii_case("DOLLARDE") {
            return self.parse_dollarde_function();
        }
        if name.eq_ignore_ascii_case("DOLLARFR") {
            return self.parse_dollarfr_function();
        }
        if name.eq_ignore_ascii_case("FVSCHEDULE") {
            self.skip_whitespace();
            if self.parse_string_literal()?.is_some() {
                return Err(FormulaEvalError::Value);
            }
            let mut value = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let mut schedule = Vec::new();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_some_and(|ch| ch == ')') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Blank => schedule.push(0.0),
                                CellValue::Number(number) => schedule.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Bool(_)
                                | CellValue::Text(_)
                                | CellValue::IsoDateTime(_)
                                | CellValue::RichText(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    schedule.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                schedule.push(self.parse_comparison()?);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if !value.is_finite() || schedule.iter().any(|rate| !rate.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            for rate in schedule {
                value *= 1.0 + rate;
            }
            return if value.is_finite() {
                Ok(value)
            } else {
                Err(FormulaEvalError::Num)
            };
        }
        if name.eq_ignore_ascii_case("NPV") {
            let rate = self.parse_comparison()?;
            if !rate.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let discount = 1.0 + rate;
            let mut discount_factor = 1.0;
            let mut total = 0.0;
            let mut saw_value_argument = false;
            macro_rules! record_cash_flow {
                ($cash_flow:expr) => {{
                    let cash_flow = $cash_flow;
                    if !cash_flow.is_finite() {
                        return Err(FormulaEvalError::Value);
                    }
                    discount_factor *= discount;
                    if discount_factor == 0.0 {
                        return Err(FormulaEvalError::Div0);
                    }
                    if !discount_factor.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    total += cash_flow / discount_factor;
                    if !total.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }};
            }
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    return if saw_value_argument {
                        Ok(total)
                    } else {
                        Err(FormulaEvalError::Value)
                    };
                }
                saw_value_argument = true;
                let checkpoint = self.index;
                if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                    self.index = next_index;
                    self.skip_whitespace();
                    if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')')) {
                        for row in rect.row_first..=rect.row_last {
                            for col in rect.col_first..=rect.col_last {
                                match self
                                    .evaluator
                                    .cell_value_or_blank(target_sheet_id, row, col)
                                {
                                    Ok(CellValue::Number(number)) => record_cash_flow!(number),
                                    Ok(_) => {}
                                    Err(FormulaEvalError::Unsupported) => {
                                        return Err(FormulaEvalError::Unsupported);
                                    }
                                    Err(_) => {}
                                }
                            }
                        }
                    } else {
                        self.index = checkpoint;
                        if self.parse_string_literal()?.is_none() {
                            let identifier_checkpoint = self.index;
                            if let Some(identifier) = self.parse_identifier() {
                                self.skip_whitespace();
                                if !(identifier.eq_ignore_ascii_case("TRUE")
                                    || identifier.eq_ignore_ascii_case("FALSE"))
                                    || self.peek_char() == Some('(')
                                {
                                    self.index = identifier_checkpoint;
                                    if let Ok(value) = self.parse_catchable_argument()? {
                                        record_cash_flow!(value);
                                    }
                                }
                            } else {
                                self.index = identifier_checkpoint;
                                if let Ok(value) = self.parse_catchable_argument()? {
                                    record_cash_flow!(value);
                                }
                            }
                        }
                    }
                } else if self.parse_string_literal()?.is_none() {
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if !(identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            || self.peek_char() == Some('(')
                        {
                            self.index = identifier_checkpoint;
                            if let Ok(value) = self.parse_catchable_argument()? {
                                record_cash_flow!(value);
                            }
                        }
                    } else {
                        self.index = identifier_checkpoint;
                        if let Ok(value) = self.parse_catchable_argument()? {
                            record_cash_flow!(value);
                        }
                    }
                }
                self.skip_whitespace();
                if self.consume_char(',') {
                    continue;
                }
                if self.consume_char(')') {
                    return Ok(total);
                }
                return Err(FormulaEvalError::Unsupported);
            }
        }
        if name.eq_ignore_ascii_case("XNPV") {
            let rate = self.parse_comparison()?;
            if !rate.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }

            let mut values = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| ch == ',') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Number(number) => values.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank
                                | CellValue::Bool(_)
                                | CellValue::Text(_)
                                | CellValue::IsoDateTime(_)
                                | CellValue::RichText(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if (identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            && self.peek_char() != Some('(')
                        {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.index = identifier_checkpoint;
                    values.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                let identifier_checkpoint = self.index;
                if let Some(identifier) = self.parse_identifier() {
                    self.skip_whitespace();
                    if (identifier.eq_ignore_ascii_case("TRUE")
                        || identifier.eq_ignore_ascii_case("FALSE"))
                        && self.peek_char() != Some('(')
                    {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.index = identifier_checkpoint;
                values.push(self.parse_comparison()?);
            }

            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }

            let mut dates = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_some_and(|ch| ch == ')') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            let date = match self.evaluator.cell_value_or_blank(
                                target_sheet_id,
                                row,
                                col,
                            )? {
                                CellValue::Number(number) => number,
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank
                                | CellValue::Bool(_)
                                | CellValue::Text(_)
                                | CellValue::IsoDateTime(_)
                                | CellValue::RichText(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            };
                            let serial = formula_serial_integer(date)
                                .and_then(|serial| {
                                    self.evaluator
                                        .context
                                        .date_system()
                                        .ymd(serial as f64)
                                        .map(|_| serial)
                                })
                                .map_err(|_| FormulaEvalError::Value)?;
                            dates.push(serial);
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if (identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            && self.peek_char() != Some('(')
                        {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.index = identifier_checkpoint;
                    let serial = formula_serial_integer(self.parse_comparison()?)
                        .and_then(|serial| {
                            self.evaluator
                                .context
                                .date_system()
                                .ymd(serial as f64)
                                .map(|_| serial)
                        })
                        .map_err(|_| FormulaEvalError::Value)?;
                    dates.push(serial);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                let identifier_checkpoint = self.index;
                if let Some(identifier) = self.parse_identifier() {
                    self.skip_whitespace();
                    if (identifier.eq_ignore_ascii_case("TRUE")
                        || identifier.eq_ignore_ascii_case("FALSE"))
                        && self.peek_char() != Some('(')
                    {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.index = identifier_checkpoint;
                let serial = formula_serial_integer(self.parse_comparison()?)
                    .and_then(|serial| {
                        self.evaluator
                            .context
                            .date_system()
                            .ymd(serial as f64)
                            .map(|_| serial)
                    })
                    .map_err(|_| FormulaEvalError::Value)?;
                dates.push(serial);
            }

            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if values.len() != dates.len()
                || !values.iter().any(|value| *value > 0.0)
                || !values.iter().any(|value| *value < 0.0)
            {
                return Err(FormulaEvalError::Num);
            }
            if values.iter().any(|value| !value.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            let start_date = dates[0];
            let discount = 1.0 + rate;
            let mut total = 0.0;
            for (value, date) in values.iter().zip(dates.iter()) {
                if *date < start_date {
                    return Err(FormulaEvalError::Num);
                }
                let years = (*date - start_date) as f64 / 365.0;
                let denominator = discount.powf(years);
                if denominator == 0.0 || !denominator.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
                total += value / denominator;
                if !total.is_finite() {
                    return Err(FormulaEvalError::Num);
                }
            }
            return Ok(total);
        }
        if name.eq_ignore_ascii_case("XIRR") {
            let mut values = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| ch == ',') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Number(number) => values.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank
                                | CellValue::Bool(_)
                                | CellValue::Text(_)
                                | CellValue::IsoDateTime(_)
                                | CellValue::RichText(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if (identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            && self.peek_char() != Some('(')
                        {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.index = identifier_checkpoint;
                    values.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                let identifier_checkpoint = self.index;
                if let Some(identifier) = self.parse_identifier() {
                    self.skip_whitespace();
                    if (identifier.eq_ignore_ascii_case("TRUE")
                        || identifier.eq_ignore_ascii_case("FALSE"))
                        && self.peek_char() != Some('(')
                    {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.index = identifier_checkpoint;
                values.push(self.parse_comparison()?);
            }

            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }

            let mut dates = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            let date = match self.evaluator.cell_value_or_blank(
                                target_sheet_id,
                                row,
                                col,
                            )? {
                                CellValue::Number(number) => number,
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank
                                | CellValue::Bool(_)
                                | CellValue::Text(_)
                                | CellValue::IsoDateTime(_)
                                | CellValue::RichText(_) => {
                                    return Err(FormulaEvalError::Value);
                                }
                            };
                            let serial = formula_serial_integer(date)
                                .and_then(|serial| {
                                    self.evaluator
                                        .context
                                        .date_system()
                                        .ymd(serial as f64)
                                        .map(|_| serial)
                                })
                                .map_err(|_| FormulaEvalError::Value)?;
                            dates.push(serial);
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    let identifier_checkpoint = self.index;
                    if let Some(identifier) = self.parse_identifier() {
                        self.skip_whitespace();
                        if (identifier.eq_ignore_ascii_case("TRUE")
                            || identifier.eq_ignore_ascii_case("FALSE"))
                            && self.peek_char() != Some('(')
                        {
                            return Err(FormulaEvalError::Value);
                        }
                    }
                    self.index = identifier_checkpoint;
                    let serial = formula_serial_integer(self.parse_comparison()?)
                        .and_then(|serial| {
                            self.evaluator
                                .context
                                .date_system()
                                .ymd(serial as f64)
                                .map(|_| serial)
                        })
                        .map_err(|_| FormulaEvalError::Value)?;
                    dates.push(serial);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                let identifier_checkpoint = self.index;
                if let Some(identifier) = self.parse_identifier() {
                    self.skip_whitespace();
                    if (identifier.eq_ignore_ascii_case("TRUE")
                        || identifier.eq_ignore_ascii_case("FALSE"))
                        && self.peek_char() != Some('(')
                    {
                        return Err(FormulaEvalError::Value);
                    }
                }
                self.index = identifier_checkpoint;
                let serial = formula_serial_integer(self.parse_comparison()?)
                    .and_then(|serial| {
                        self.evaluator
                            .context
                            .date_system()
                            .ymd(serial as f64)
                            .map(|_| serial)
                    })
                    .map_err(|_| FormulaEvalError::Value)?;
                dates.push(serial);
            }

            self.skip_whitespace();
            let mut guess = 0.1;
            if self.consume_char(',') {
                guess = self.parse_comparison()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            } else if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if !guess.is_finite() || values.iter().any(|value| !value.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            if guess <= -1.0
                || values.len() != dates.len()
                || !values.iter().any(|value| *value > 0.0)
                || !values.iter().any(|value| *value < 0.0)
            {
                return Err(FormulaEvalError::Num);
            }
            let start_date = dates[0];
            let xirr_value = |rate: f64| -> Result<(f64, f64), FormulaEvalError> {
                if !rate.is_finite() || rate <= -1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let discount = 1.0 + rate;
                let mut value = 0.0;
                let mut derivative = 0.0;
                for (cash_flow, date) in values.iter().zip(dates.iter()) {
                    if *date < start_date {
                        return Err(FormulaEvalError::Num);
                    }
                    let years = (*date - start_date) as f64 / 365.0;
                    let denominator = discount.powf(years);
                    if denominator == 0.0 || !denominator.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    value += cash_flow / denominator;
                    derivative -= years * cash_flow / (denominator * discount);
                }
                if value.is_finite() && derivative.is_finite() {
                    Ok((value, derivative))
                } else {
                    Err(FormulaEvalError::Num)
                }
            };

            const XIRR_MAX_ITERATIONS: usize = 100;
            const XIRR_TOLERANCE: f64 = 1e-8;
            let mut rate = guess;
            for _ in 0..XIRR_MAX_ITERATIONS {
                let (value, derivative) = xirr_value(rate)?;
                if value.abs() <= XIRR_TOLERANCE {
                    return Ok(rate);
                }
                if derivative == 0.0 {
                    break;
                }
                let next_rate = rate - value / derivative;
                if !next_rate.is_finite() || next_rate <= -1.0 {
                    break;
                }
                if (next_rate - rate).abs() <= XIRR_TOLERANCE {
                    return Ok(next_rate);
                }
                rate = next_rate;
            }
            return Err(FormulaEvalError::Num);
        }
        if name.eq_ignore_ascii_case("IRR") {
            let mut values = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')')) {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Number(number) => values.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank
                                | CellValue::Bool(_)
                                | CellValue::Text(_)
                                | CellValue::IsoDateTime(_)
                                | CellValue::RichText(_) => {}
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    values.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                values.push(self.parse_comparison()?);
            }
            self.skip_whitespace();
            let mut guess = 0.1;
            if self.consume_char(',') {
                guess = self.parse_comparison()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            } else if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if !guess.is_finite() || values.iter().any(|value| !value.is_finite()) {
                return Err(FormulaEvalError::Value);
            }
            if guess <= -1.0
                || values.len() < 2
                || !values.iter().any(|value| *value > 0.0)
                || !values.iter().any(|value| *value < 0.0)
            {
                return Err(FormulaEvalError::Num);
            }

            let irr_value = |rate: f64| -> Result<(f64, f64), FormulaEvalError> {
                if !rate.is_finite() || rate <= -1.0 {
                    return Err(FormulaEvalError::Num);
                }
                let factor = 1.0 + rate;
                let mut denominator = 1.0;
                let mut value = 0.0;
                let mut derivative = 0.0;
                for (index, cash_flow) in values.iter().enumerate() {
                    if index > 0 {
                        denominator *= factor;
                        if denominator == 0.0 || !denominator.is_finite() {
                            return Err(FormulaEvalError::Num);
                        }
                    }
                    value += cash_flow / denominator;
                    if index > 0 {
                        derivative -= index as f64 * cash_flow / (denominator * factor);
                    }
                }
                if value.is_finite() && derivative.is_finite() {
                    Ok((value, derivative))
                } else {
                    Err(FormulaEvalError::Num)
                }
            };

            const IRR_MAX_ITERATIONS: usize = 20;
            const IRR_TOLERANCE: f64 = 1e-7;
            let mut rate = guess;
            for _ in 0..IRR_MAX_ITERATIONS {
                let (value, derivative) = irr_value(rate)?;
                if value.abs() <= IRR_TOLERANCE {
                    return Ok(rate);
                }
                if derivative == 0.0 {
                    break;
                }
                let next_rate = rate - value / derivative;
                if !next_rate.is_finite() || next_rate <= -1.0 {
                    break;
                }
                if (next_rate - rate).abs() <= IRR_TOLERANCE {
                    return Ok(next_rate);
                }
                rate = next_rate;
            }
            return Err(FormulaEvalError::Num);
        }
        if name.eq_ignore_ascii_case("MIRR") {
            let mut values = Vec::new();
            self.skip_whitespace();
            let checkpoint = self.index;
            if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| ch == ',') {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            match self
                                .evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?
                            {
                                CellValue::Number(number) => values.push(number),
                                CellValue::Error(error) => {
                                    return Err(formula_eval_error_from_cell_error(error));
                                }
                                CellValue::Blank
                                | CellValue::Bool(_)
                                | CellValue::Text(_)
                                | CellValue::IsoDateTime(_)
                                | CellValue::RichText(_) => {}
                            }
                        }
                    }
                } else {
                    self.index = checkpoint;
                    if self.parse_string_literal()?.is_some() {
                        return Err(FormulaEvalError::Value);
                    }
                    values.push(self.parse_comparison()?);
                }
            } else {
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                values.push(self.parse_comparison()?);
            }
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let finance_rate = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let reinvest_rate = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if !finance_rate.is_finite()
                || !reinvest_rate.is_finite()
                || values.iter().any(|value| !value.is_finite())
            {
                return Err(FormulaEvalError::Value);
            }
            if values.len() < 2
                || !values.iter().any(|value| *value > 0.0)
                || !values.iter().any(|value| *value < 0.0)
            {
                return Err(FormulaEvalError::Div0);
            }
            let finance_factor = 1.0 + finance_rate;
            let reinvest_factor = 1.0 + reinvest_rate;
            let periods = values.len() - 1;
            let mut future_positive = 0.0;
            let mut present_negative = 0.0;
            for (index, value) in values.iter().enumerate() {
                if *value > 0.0 {
                    let exponent =
                        i32::try_from(periods - index).map_err(|_| FormulaEvalError::Num)?;
                    future_positive += value * reinvest_factor.powi(exponent);
                    if !future_positive.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                } else if *value < 0.0 {
                    let exponent = i32::try_from(index).map_err(|_| FormulaEvalError::Num)?;
                    let denominator = finance_factor.powi(exponent);
                    if denominator == 0.0 {
                        return Err(FormulaEvalError::Div0);
                    }
                    if !denominator.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                    present_negative += value / denominator;
                    if !present_negative.is_finite() {
                        return Err(FormulaEvalError::Num);
                    }
                }
            }
            if future_positive <= 0.0 || present_negative >= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            let result = (future_positive / -present_negative).powf(1.0 / periods as f64) - 1.0;
            return if result.is_finite() {
                Ok(result)
            } else {
                Err(FormulaEvalError::Num)
            };
        }
        if name.eq_ignore_ascii_case("FV") {
            return self.parse_fv_function();
        }
        if name.eq_ignore_ascii_case("PV") {
            return self.parse_pv_function();
        }
        if name.eq_ignore_ascii_case("PMT") {
            return self.parse_pmt_function();
        }
        if name.eq_ignore_ascii_case("IPMT") {
            return self.parse_ipmt_function();
        }
        if name.eq_ignore_ascii_case("PPMT") {
            return self.parse_ppmt_function();
        }
        if name.eq_ignore_ascii_case("CUMIPMT") {
            return self.parse_cumulative_payment_function(false);
        }
        if name.eq_ignore_ascii_case("CUMPRINC") {
            return self.parse_cumulative_payment_function(true);
        }
        if name.eq_ignore_ascii_case("NPER") {
            return self.parse_nper_function();
        }
        if name.eq_ignore_ascii_case("RATE") {
            return self.parse_rate_function();
        }
        if name.eq_ignore_ascii_case("ISPMT") {
            return self.parse_ispmt_function();
        }
        if name.eq_ignore_ascii_case("ARABIC") {
            return self.parse_arabic_function();
        }
        if name.eq_ignore_ascii_case("CODE") || name.eq_ignore_ascii_case("UNICODE") {
            return self.parse_character_code_function();
        }
        if name.eq_ignore_ascii_case("LEN") {
            return self.parse_len_function(false);
        }
        if name.eq_ignore_ascii_case("LENB") {
            return self.parse_len_function(true);
        }
        if name.eq_ignore_ascii_case("FIND") {
            return self.parse_find_function(false, false);
        }
        if name.eq_ignore_ascii_case("FINDB") {
            return self.parse_find_function(false, true);
        }
        if name.eq_ignore_ascii_case("SEARCH") {
            return self.parse_find_function(true, false);
        }
        if name.eq_ignore_ascii_case("SEARCHB") {
            return self.parse_find_function(true, true);
        }
        if name.eq_ignore_ascii_case("REGEXTEST") {
            return self.parse_regex_test_function();
        }
        if name.eq_ignore_ascii_case("EXACT") {
            return self.parse_exact_function();
        }
        if name.eq_ignore_ascii_case("VALUE") {
            return self.parse_value_function();
        }
        if name.eq_ignore_ascii_case("NUMBERVALUE") {
            return self.parse_numbervalue_function();
        }
        if name.eq_ignore_ascii_case("DATEVALUE") {
            return self.parse_datevalue_function();
        }
        if name.eq_ignore_ascii_case("TIMEVALUE") {
            return self.parse_timevalue_function();
        }
        if name.eq_ignore_ascii_case("NA") {
            return self.parse_na_function();
        }
        if name.eq_ignore_ascii_case("CELL") {
            return formula_number_from_value_probe(self.parse_cell_value_function()?);
        }
        if name.eq_ignore_ascii_case("INFO") {
            return formula_number_from_value_probe(self.parse_info_value_function()?);
        }
        if name.eq_ignore_ascii_case("ROW") {
            return self.parse_row_function();
        }
        if name.eq_ignore_ascii_case("ROWS") {
            return self.parse_rows_function();
        }
        if name.eq_ignore_ascii_case("AREAS") {
            return self.parse_areas_function();
        }
        if name.eq_ignore_ascii_case("SHEET") {
            return self.parse_sheet_function();
        }
        if name.eq_ignore_ascii_case("SHEETS") {
            return self.parse_sheets_function();
        }
        if name.eq_ignore_ascii_case("INDEX") {
            return self.parse_index_function();
        }
        if name.eq_ignore_ascii_case("MATCH") {
            return self.parse_match_function();
        }
        if name.eq_ignore_ascii_case("XMATCH") {
            return self.parse_xmatch_function();
        }
        if name.eq_ignore_ascii_case("LOOKUP") {
            return self.parse_lookup_function();
        }
        if name.eq_ignore_ascii_case("VLOOKUP") {
            return self.parse_vlookup_function();
        }
        if name.eq_ignore_ascii_case("HLOOKUP") {
            return self.parse_hlookup_function();
        }
        if name.eq_ignore_ascii_case("XLOOKUP") {
            return formula_number_from_value_probe(self.parse_xlookup_value_function()?);
        }
        if name.eq_ignore_ascii_case("LET") {
            return formula_number_from_value_probe(self.parse_let_value_function()?);
        }
        if name.eq_ignore_ascii_case("LAMBDA") {
            let lambda = self.parse_lambda_value_function()?;
            self.skip_whitespace();
            if !self.consume_char('(') {
                return Err(FormulaEvalError::Calc);
            }
            return formula_number_from_value_probe(self.parse_lambda_call_arguments(lambda)?);
        }
        if name.eq_ignore_ascii_case("ISOMITTED") {
            return self.parse_isomitted_function();
        }
        if name.eq_ignore_ascii_case("MAKEARRAY") {
            return formula_number_from_value_probe(self.parse_makearray_value_function()?);
        }
        if name.eq_ignore_ascii_case("REDUCE") || name.eq_ignore_ascii_case("SCAN") {
            return formula_number_from_value_probe(self.parse_reduce_scan_value_function(name)?);
        }
        if name.eq_ignore_ascii_case("GETPIVOTDATA") {
            return formula_number_from_value_probe(self.parse_getpivotdata_value_function()?);
        }
        if name.eq_ignore_ascii_case("CUBESETCOUNT") {
            return self.parse_cubesetcount_function();
        }
        if name.eq_ignore_ascii_case("CUBEVALUE")
            || name.eq_ignore_ascii_case("RTD")
            || name.eq_ignore_ascii_case("STOCKHISTORY")
            || name.eq_ignore_ascii_case("COPILOT")
        {
            return self.parse_external_data_unavailable_function();
        }
        if name.eq_ignore_ascii_case("FIELDVALUE") {
            return self.parse_external_field_unavailable_function();
        }
        if name.eq_ignore_ascii_case("PY") {
            return self.parse_external_python_unavailable_function();
        }
        if name.eq_ignore_ascii_case("CALL") || name.eq_ignore_ascii_case("REGISTER.ID") {
            return self.parse_external_platform_unavailable_function();
        }
        if name.eq_ignore_ascii_case("INDIRECT")
            || name.eq_ignore_ascii_case("OFFSET")
            || name.eq_ignore_ascii_case("TRIMRANGE")
        {
            return formula_number_from_value_probe(
                self.parse_reference_projection_value_function(name)?,
            );
        }
        if formula_array_projection_function_name(name) {
            return formula_number_from_value_probe(
                self.parse_array_projection_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("FREQUENCY") {
            return self.parse_frequency_function();
        }
        if name.eq_ignore_ascii_case("MMULT") {
            return self.parse_mmult_function();
        }
        if name.eq_ignore_ascii_case("MINVERSE") {
            return self.parse_minverse_function();
        }
        if name.eq_ignore_ascii_case("MUNIT") {
            return self.parse_munit_function();
        }
        if name.eq_ignore_ascii_case("SEQUENCE") {
            return self.parse_sequence_function();
        }
        if name.eq_ignore_ascii_case("RANDARRAY") {
            return self.parse_randarray_function();
        }
        if name.eq_ignore_ascii_case("CHISQ.TEST") || name.eq_ignore_ascii_case("CHITEST") {
            return self.parse_chisq_test_function();
        }
        if name.eq_ignore_ascii_case("F.TEST") || name.eq_ignore_ascii_case("FTEST") {
            return self.parse_f_test_function();
        }
        if name.eq_ignore_ascii_case("T.TEST") || name.eq_ignore_ascii_case("TTEST") {
            return self.parse_t_test_function();
        }
        if name.eq_ignore_ascii_case("Z.TEST") || name.eq_ignore_ascii_case("ZTEST") {
            return self.parse_z_test_function();
        }
        if name.eq_ignore_ascii_case("DAVERAGE")
            || name.eq_ignore_ascii_case("DCOUNT")
            || name.eq_ignore_ascii_case("DCOUNTA")
            || name.eq_ignore_ascii_case("DGET")
            || name.eq_ignore_ascii_case("DMAX")
            || name.eq_ignore_ascii_case("DMIN")
            || name.eq_ignore_ascii_case("DPRODUCT")
            || name.eq_ignore_ascii_case("DSTDEV")
            || name.eq_ignore_ascii_case("DSTDEVP")
            || name.eq_ignore_ascii_case("DSUM")
            || name.eq_ignore_ascii_case("DVAR")
            || name.eq_ignore_ascii_case("DVARP")
        {
            return formula_number_from_value_probe(self.parse_database_value_function(name)?);
        }
        if name.eq_ignore_ascii_case("SUMIF") {
            return self.parse_sumif_function();
        }
        if name.eq_ignore_ascii_case("AVERAGEIF") {
            return self.parse_averageif_function();
        }
        if name.eq_ignore_ascii_case("COUNTIFS") {
            return self.parse_countifs_function();
        }
        if name.eq_ignore_ascii_case("SUMIFS") {
            return self.parse_sumifs_function();
        }
        if name.eq_ignore_ascii_case("AVERAGEIFS") {
            return self.parse_averageifs_function();
        }
        if name.eq_ignore_ascii_case("MINIFS") {
            return self.parse_minifs_function();
        }
        if name.eq_ignore_ascii_case("MAXIFS") {
            return self.parse_maxifs_function();
        }
        if name.eq_ignore_ascii_case("AVERAGEA")
            || name.eq_ignore_ascii_case("MINA")
            || name.eq_ignore_ascii_case("MAXA")
            || name.eq_ignore_ascii_case("VARA")
            || name.eq_ignore_ascii_case("VARPA")
            || name.eq_ignore_ascii_case("STDEVA")
            || name.eq_ignore_ascii_case("STDEVPA")
        {
            return self.parse_aggregate_a_function(name);
        }
        if name.eq_ignore_ascii_case("SUMPRODUCT") {
            let mut arguments: Vec<(u32, u32, Vec<f64>)> = Vec::new();
            let finish = |arguments: &[(u32, u32, Vec<f64>)]| -> Result<f64, FormulaEvalError> {
                let Some((base_height, base_width, first_values)) = arguments.first() else {
                    return Err(FormulaEvalError::Value);
                };
                if arguments
                    .iter()
                    .any(|(height, width, _)| height != base_height || width != base_width)
                {
                    return Err(FormulaEvalError::Value);
                }
                let mut total = 0.0_f64;
                for index in 0..first_values.len() {
                    let mut product = 1.0;
                    for (_, _, values) in arguments {
                        product *= values[index];
                    }
                    total += product;
                }
                if total.is_finite() {
                    Ok(total)
                } else {
                    Err(FormulaEvalError::Num)
                }
            };
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    return finish(arguments.as_slice());
                }

                let checkpoint = self.index;
                if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
                    self.index = next_index;
                    self.skip_whitespace();
                    if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')')) {
                        let capacity = rect
                            .checked_cell_count_usize()
                            .map_err(|_| FormulaEvalError::Num)?;
                        let mut values = Vec::with_capacity(capacity);
                        for row in rect.row_first..=rect.row_last {
                            for col in rect.col_first..=rect.col_last {
                                match self.evaluator.cell_value_or_blank(
                                    target_sheet_id,
                                    row,
                                    col,
                                )? {
                                    CellValue::Number(number) => values.push(number),
                                    CellValue::Error(error) => {
                                        return Err(formula_eval_error_from_cell_error(error));
                                    }
                                    CellValue::Blank
                                    | CellValue::Bool(_)
                                    | CellValue::Text(_)
                                    | CellValue::IsoDateTime(_)
                                    | CellValue::RichText(_) => {
                                        values.push(0.0);
                                    }
                                }
                            }
                        }
                        arguments.push((rect.height(), rect.width(), values));
                    } else {
                        self.index = checkpoint;
                        arguments.push((1, 1, vec![self.parse_comparison()?]));
                    }
                } else {
                    arguments.push((1, 1, vec![self.parse_comparison()?]));
                }

                self.skip_whitespace();
                if self.consume_char(',') {
                    continue;
                }
                if self.consume_char(')') {
                    return finish(arguments.as_slice());
                }
                return Err(FormulaEvalError::Unsupported);
            }
        }
        if name.eq_ignore_ascii_case("SUMXMY2")
            || name.eq_ignore_ascii_case("SUMX2MY2")
            || name.eq_ignore_ascii_case("SUMX2PY2")
        {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Err(FormulaEvalError::Value);
            }
            let first_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let second_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if first_values.len() != second_values.len() {
                return Err(FormulaEvalError::NA);
            }
            let mut total = 0.0_f64;
            for (first_value, second_value) in first_values.iter().zip(second_values.iter()) {
                if name.eq_ignore_ascii_case("SUMXMY2") {
                    let difference = first_value - second_value;
                    total += difference * difference;
                } else if name.eq_ignore_ascii_case("SUMX2MY2") {
                    total += first_value * first_value - second_value * second_value;
                } else {
                    total += first_value * first_value + second_value * second_value;
                }
            }
            return Ok(total);
        }
        if name.eq_ignore_ascii_case("FORECAST") || name.eq_ignore_ascii_case("FORECAST.LINEAR") {
            let x = self.parse_comparison()?;
            if !x.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let known_y = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let known_x = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if known_y.len() != known_x.len() || known_y.is_empty() {
                return Err(FormulaEvalError::NA);
            }
            let count = known_y.len() as f64;
            let mean_y = known_y.iter().sum::<f64>() / count;
            let mean_x = known_x.iter().sum::<f64>() / count;
            let mut sum_xy_deviation = 0.0_f64;
            let mut sum_x_deviation_square = 0.0_f64;
            for (y_value, x_value) in known_y.iter().zip(known_x.iter()) {
                let y_deviation = y_value - mean_y;
                let x_deviation = x_value - mean_x;
                sum_xy_deviation += y_deviation * x_deviation;
                sum_x_deviation_square += x_deviation * x_deviation;
            }
            if sum_x_deviation_square == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            return Ok(mean_y + sum_xy_deviation / sum_x_deviation_square * (x - mean_x));
        }
        if name.eq_ignore_ascii_case("FORECAST.ETS")
            || name.eq_ignore_ascii_case("FORECAST.ETS.CONFINT")
            || name.eq_ignore_ascii_case("FORECAST.ETS.SEASONALITY")
            || name.eq_ignore_ascii_case("FORECAST.ETS.STAT")
        {
            return self.parse_forecast_ets_function(name);
        }
        if name.eq_ignore_ascii_case("LINEST") || name.eq_ignore_ascii_case("LOGEST") {
            return self.parse_regression_coefficient_function(name.eq_ignore_ascii_case("LOGEST"));
        }
        if name.eq_ignore_ascii_case("TREND") || name.eq_ignore_ascii_case("GROWTH") {
            return self.parse_regression_prediction_function(name.eq_ignore_ascii_case("GROWTH"));
        }
        if name.eq_ignore_ascii_case("CORREL")
            || name.eq_ignore_ascii_case("PEARSON")
            || name.eq_ignore_ascii_case("COVAR")
            || name.eq_ignore_ascii_case("COVARIANCE.P")
            || name.eq_ignore_ascii_case("COVARIANCE.S")
            || name.eq_ignore_ascii_case("SLOPE")
            || name.eq_ignore_ascii_case("INTERCEPT")
            || name.eq_ignore_ascii_case("RSQ")
            || name.eq_ignore_ascii_case("STEYX")
        {
            self.skip_whitespace();
            let first_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            let second_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if first_values.len() != second_values.len() {
                return Err(FormulaEvalError::NA);
            }
            let count = first_values.len();
            if count == 0 {
                return if name.eq_ignore_ascii_case("CORREL")
                    || name.eq_ignore_ascii_case("COVAR")
                    || name.eq_ignore_ascii_case("COVARIANCE.P")
                    || name.eq_ignore_ascii_case("COVARIANCE.S")
                {
                    Err(FormulaEvalError::Div0)
                } else {
                    Err(FormulaEvalError::NA)
                };
            }
            if name.eq_ignore_ascii_case("COVARIANCE.S") && count < 2 {
                return Err(FormulaEvalError::Div0);
            }

            let first_mean = first_values.iter().sum::<f64>() / count as f64;
            let second_mean = second_values.iter().sum::<f64>() / count as f64;
            let mut sum_first_second_deviation = 0.0_f64;
            let mut sum_first_deviation_square = 0.0_f64;
            let mut sum_second_deviation_square = 0.0_f64;
            for (first_value, second_value) in first_values.iter().zip(second_values.iter()) {
                let first_deviation = first_value - first_mean;
                let second_deviation = second_value - second_mean;
                sum_first_second_deviation += first_deviation * second_deviation;
                sum_first_deviation_square += first_deviation * first_deviation;
                sum_second_deviation_square += second_deviation * second_deviation;
            }

            if name.eq_ignore_ascii_case("COVAR") || name.eq_ignore_ascii_case("COVARIANCE.P") {
                return Ok(sum_first_second_deviation / count as f64);
            }
            if name.eq_ignore_ascii_case("COVARIANCE.S") {
                return Ok(sum_first_second_deviation / (count - 1) as f64);
            }
            if name.eq_ignore_ascii_case("SLOPE") || name.eq_ignore_ascii_case("INTERCEPT") {
                if sum_second_deviation_square == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let slope = sum_first_second_deviation / sum_second_deviation_square;
                if name.eq_ignore_ascii_case("SLOPE") {
                    return Ok(slope);
                }
                return Ok(first_mean - slope * second_mean);
            }
            if name.eq_ignore_ascii_case("STEYX") {
                if count < 3 || sum_second_deviation_square == 0.0 {
                    return Err(FormulaEvalError::Div0);
                }
                let residual_square_sum = sum_first_deviation_square
                    - sum_first_second_deviation * sum_first_second_deviation
                        / sum_second_deviation_square;
                return Ok((residual_square_sum.max(0.0) / (count - 2) as f64).sqrt());
            }

            let denominator = sum_first_deviation_square * sum_second_deviation_square;
            if denominator == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            let correlation = sum_first_second_deviation / denominator.sqrt();
            if name.eq_ignore_ascii_case("RSQ") {
                return Ok(correlation * correlation);
            }
            return Ok(correlation);
        }
        if name.eq_ignore_ascii_case("PROB") {
            let x_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let probabilities = self.parse_aggregate_argument()?;
            if x_values.len() != probabilities.len() {
                return Err(FormulaEvalError::NA);
            }
            let mut probability_sum = 0.0_f64;
            for probability in &probabilities {
                if !probability.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                if *probability <= 0.0 || *probability > 1.0 {
                    return Err(FormulaEvalError::Num);
                }
                probability_sum += probability;
            }
            if (probability_sum - 1.0).abs() > 1e-7 {
                return Err(FormulaEvalError::Num);
            }

            self.skip_whitespace();
            let (lower_limit, upper_limit) = if self.consume_char(')') {
                (0.0, None)
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let lower_limit = self.parse_comparison()?;
                if !lower_limit.is_finite() {
                    return Err(FormulaEvalError::Value);
                }
                self.skip_whitespace();
                if self.consume_char(')') {
                    (lower_limit, None)
                } else {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    let upper_limit = self.parse_comparison()?;
                    if !upper_limit.is_finite() {
                        return Err(FormulaEvalError::Value);
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    if upper_limit < lower_limit {
                        return Err(FormulaEvalError::Num);
                    }
                    (lower_limit, Some(upper_limit))
                }
            };

            let total = x_values
                .iter()
                .zip(probabilities.iter())
                .filter_map(|(x_value, probability)| {
                    let matches = if let Some(upper_limit) = upper_limit {
                        *x_value >= lower_limit && *x_value <= upper_limit
                    } else {
                        *x_value == lower_limit
                    };
                    matches.then_some(*probability)
                })
                .sum::<f64>();
            return Ok(total);
        }
        if name.eq_ignore_ascii_case("PERCENTOF") {
            let subset = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let all_values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let denominator = all_values.iter().sum::<f64>();
            if denominator == 0.0 {
                return Err(FormulaEvalError::Div0);
            }
            return formula_checked_numeric_result(subset.iter().sum::<f64>() / denominator);
        }
        if name.eq_ignore_ascii_case("PERCENTILE")
            || name.eq_ignore_ascii_case("PERCENTILE.INC")
            || name.eq_ignore_ascii_case("PERCENTILE.EXC")
        {
            let values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let k = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return percentile_value(values, k, name.eq_ignore_ascii_case("PERCENTILE.EXC"));
        }
        if name.eq_ignore_ascii_case("PERCENTRANK")
            || name.eq_ignore_ascii_case("PERCENTRANK.INC")
            || name.eq_ignore_ascii_case("PERCENTRANK.EXC")
        {
            let values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let x = self.parse_comparison()?;
            self.skip_whitespace();
            let significance = if self.consume_char(')') {
                3
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let significance = formula_integer_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
                significance
            };
            return percent_rank_value(
                values,
                x,
                significance,
                name.eq_ignore_ascii_case("PERCENTRANK.EXC"),
            );
        }
        if name.eq_ignore_ascii_case("QUARTILE")
            || name.eq_ignore_ascii_case("QUARTILE.INC")
            || name.eq_ignore_ascii_case("QUARTILE.EXC")
        {
            let values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let quart = self.parse_comparison()?;
            if !quart.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            if quart < i64::MIN as f64 || quart > i64::MAX as f64 {
                return Err(FormulaEvalError::Num);
            }
            let quart = quart.trunc() as i64;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let exclusive = name.eq_ignore_ascii_case("QUARTILE.EXC");
            if exclusive {
                if !(1..=3).contains(&quart) {
                    return Err(FormulaEvalError::Num);
                }
            } else if !(0..=4).contains(&quart) {
                return Err(FormulaEvalError::Num);
            }
            return percentile_value(values, quart as f64 / 4.0, exclusive);
        }
        if name.eq_ignore_ascii_case("TRIMMEAN") {
            let mut values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let percent = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if values.is_empty() {
                return Err(FormulaEvalError::Div0);
            }
            if !percent.is_finite() || !(0.0..=1.0).contains(&percent) {
                return Err(FormulaEvalError::Num);
            }
            values.sort_by(|left, right| left.total_cmp(right));
            let trim_count = (values.len() as f64 * percent).floor() as usize;
            let trim_each_side = (trim_count - trim_count % 2) / 2;
            let remaining = values.len().saturating_sub(trim_each_side * 2);
            if remaining == 0 {
                return Err(FormulaEvalError::Num);
            }
            return Ok(values[trim_each_side..trim_each_side + remaining]
                .iter()
                .sum::<f64>()
                / remaining as f64);
        }
        if name.eq_ignore_ascii_case("LARGE") || name.eq_ignore_ascii_case("SMALL") {
            let mut values = self.parse_aggregate_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let k = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            if values.is_empty() || k < 1 || k > values.len() as i64 {
                return Err(FormulaEvalError::Num);
            }
            values.sort_by(|left, right| left.total_cmp(right));
            let index = if name.eq_ignore_ascii_case("LARGE") {
                values.len() - k as usize
            } else {
                k as usize - 1
            };
            return Ok(values[index]);
        }
        if name.eq_ignore_ascii_case("RANK")
            || name.eq_ignore_ascii_case("RANK.EQ")
            || name.eq_ignore_ascii_case("RANK.AVG")
        {
            let number = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (target_sheet_id, rect) = self.parse_reference_argument()?;
            let values = self
                .evaluator
                .numeric_values_in_rect(target_sheet_id, rect)?;
            self.skip_whitespace();
            let ascending = if self.consume_char(')') {
                false
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let order = self.parse_comparison()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
                order != 0.0
            };
            let tie_count = values.iter().filter(|value| **value == number).count();
            if tie_count == 0 {
                return Err(FormulaEvalError::NA);
            }
            let ahead_count = values
                .iter()
                .filter(|value| {
                    if ascending {
                        **value < number
                    } else {
                        **value > number
                    }
                })
                .count();
            let rank = ahead_count as f64 + 1.0;
            if name.eq_ignore_ascii_case("RANK.AVG") {
                return Ok(rank + (tie_count as f64 - 1.0) / 2.0);
            }
            return Ok(rank);
        }
        let function =
            FormulaAggregateFunction::from_name(name).ok_or(FormulaEvalError::Unsupported)?;
        let mut values = Vec::new();
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return function.evaluate(values.as_slice());
            }
            values.extend(self.parse_aggregate_argument()?);
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return function.evaluate(values.as_slice());
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn consume_remaining_optional_value_arguments(&mut self) -> Result<(), FormulaEvalError> {
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(());
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                continue;
            }
            let checkpoint = self.index;
            if let Some((_, _, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
            } else {
                self.index = checkpoint;
                self.parse_value_probe_argument()?;
            }
        }
    }

    fn consume_all_value_arguments(&mut self) -> Result<(), FormulaEvalError> {
        let mut needs_separator = false;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(());
            }
            if needs_separator {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if self.consume_char(')') {
                    return Ok(());
                }
            }
            if self.peek_char().is_some_and(|ch| ch == ',') {
                needs_separator = true;
                continue;
            }
            let checkpoint = self.index;
            if let Some((_, _, next_index)) = self.try_parse_reference()? {
                self.index = next_index;
            } else {
                self.index = checkpoint;
                self.parse_value_probe_argument()?;
            }
            needs_separator = true;
        }
    }

    fn parse_scalar_function(
        &mut self,
        function: FormulaScalarFunction,
    ) -> Result<f64, FormulaEvalError> {
        let mut args = Vec::new();
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return function.evaluate(args.as_slice(), self.evaluator.context);
            }
            args.push(self.parse_comparison()?);
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return function.evaluate(args.as_slice(), self.evaluator.context);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_logical_function(
        &mut self,
        function: FormulaLogicalFunction,
    ) -> Result<f64, FormulaEvalError> {
        let mut saw_value = false;
        let mut true_count = 0_u64;
        let mut false_count = 0_u64;
        macro_rules! record_value {
            ($value:expr, $from_reference:expr) => {
                match $value {
                    FormulaValueProbe::Bool(value) => {
                        saw_value = true;
                        if value {
                            true_count += 1;
                        } else {
                            false_count += 1;
                        }
                    }
                    FormulaValueProbe::Number(value) => {
                        saw_value = true;
                        if value != 0.0 {
                            true_count += 1;
                        } else {
                            false_count += 1;
                        }
                    }
                    FormulaValueProbe::Blank | FormulaValueProbe::Text(_) if $from_reference => {}
                    FormulaValueProbe::Blank => {
                        saw_value = true;
                        false_count += 1;
                    }
                    FormulaValueProbe::Text(_) => return Err(FormulaEvalError::Value),
                    FormulaValueProbe::Error(error) => return Err(error),
                    FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
                        return Err(FormulaEvalError::Value);
                    }
                }
            };
        }
        macro_rules! finish_logical {
            () => {{
                if !saw_value {
                    return Err(FormulaEvalError::Value);
                }
                Ok(match function {
                    FormulaLogicalFunction::And => {
                        if false_count == 0 {
                            1.0
                        } else {
                            0.0
                        }
                    }
                    FormulaLogicalFunction::Or => {
                        if true_count > 0 {
                            1.0
                        } else {
                            0.0
                        }
                    }
                    FormulaLogicalFunction::Xor => {
                        if true_count % 2 == 1 {
                            1.0
                        } else {
                            0.0
                        }
                    }
                })
            }};
        }
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return finish_logical!();
            }

            if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
                for (target_sheet_id, rect) in reference.areas() {
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            let value =
                                self.evaluator
                                    .cell_value_or_blank(*target_sheet_id, row, col)?;
                            record_value!(formula_value_probe_from_cell_value(value), true);
                        }
                    }
                }
            } else {
                self.skip_whitespace();
                if self.parse_string_literal()?.is_some() {
                    return Err(FormulaEvalError::Value);
                }
                match self.parse_catchable_argument()? {
                    Ok(value) => record_value!(FormulaValueProbe::Number(value), false),
                    Err(error) => record_value!(FormulaValueProbe::Error(error), false),
                }
            }

            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return finish_logical!();
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    fn parse_catchable_argument(
        &mut self,
    ) -> Result<Result<f64, FormulaEvalError>, FormulaEvalError> {
        match self.parse_comparison() {
            Ok(value) => Ok(Ok(value)),
            Err(FormulaEvalError::Unsupported) => Err(FormulaEvalError::Unsupported),
            Err(error) => Ok(Err(error)),
        }
    }

    fn parse_text_value_argument(&mut self) -> Result<String, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(text) = self.parse_string_literal()? {
            return Ok(text);
        }
        let checkpoint = self.index;
        if let Some((reference, next_index)) = self.try_parse_reference_set()? {
            self.index = next_index;
            let (target_sheet_id, rect) = reference.single_area()?;
            if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                return Err(FormulaEvalError::Value);
            }
            let value = self.evaluator.cell_value_or_blank(
                target_sheet_id,
                rect.row_first,
                rect.col_first,
            )?;
            return formula_text_from_value_probe(formula_value_probe_from_cell_value(value));
        }
        self.index = checkpoint;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                if let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? {
                    return formula_text_from_value_probe(value);
                }
                if identifier.eq_ignore_ascii_case("LAMBDA")
                    || formula_text_function_name(identifier.as_str())
                {
                    match self.parse_text_function(identifier.as_str()) {
                        Ok(text) => return Ok(text),
                        Err(FormulaEvalError::Unsupported) => self.index = checkpoint,
                        Err(error) => return Err(error),
                    }
                } else {
                    self.index = checkpoint;
                }
            } else {
                if identifier.eq_ignore_ascii_case("TRUE") {
                    return Ok("TRUE".to_string());
                }
                if identifier.eq_ignore_ascii_case("FALSE") {
                    return Ok("FALSE".to_string());
                }
                if let Some(value) = self.binding_value(identifier.as_str())
                    && self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')'))
                {
                    return formula_text_from_value_probe(value);
                }
                if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')'))
                    && let Some(value) = self.defined_name_value_probe(identifier.as_str())?
                {
                    return formula_text_from_value_probe(value);
                }
                self.index = checkpoint;
            }
        }
        formula_text_from_number(self.parse_comparison()?)
    }

    fn parse_text_values_argument(&mut self) -> Result<Vec<String>, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(reference) = self.parse_reference_set_before_boundary(&[',', ')'])? {
            let mut values = Vec::new();
            for (target_sheet_id, rect) in reference.areas() {
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        let value =
                            self.evaluator
                                .cell_value_or_blank(*target_sheet_id, row, col)?;
                        values.push(formula_text_from_value_probe(
                            formula_value_probe_from_cell_value(value),
                        )?);
                    }
                }
            }
            return Ok(values);
        }
        Ok(vec![self.parse_text_value_argument()?])
    }

    fn parse_value_probe_argument(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        self.skip_whitespace();
        if let Some(text) = self.parse_string_literal()? {
            return Ok(FormulaValueProbe::Text(text));
        }
        let checkpoint = self.index;
        if let Some((reference, next_index)) = self.try_parse_reference_set()? {
            self.index = next_index;
            let (target_sheet_id, rect) = reference.single_area()?;
            if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                return Err(FormulaEvalError::Value);
            }
            let value = self.evaluator.cell_value_or_blank(
                target_sheet_id,
                rect.row_first,
                rect.col_first,
            )?;
            return Ok(formula_value_probe_from_cell_value(value));
        }
        self.index = checkpoint;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                if let Some(value) = self.parse_bound_lambda_call_value(identifier.as_str())? {
                    return Ok(value);
                }
                if identifier.eq_ignore_ascii_case("LAMBDA") {
                    let lambda = self.parse_lambda_value_function()?;
                    self.skip_whitespace();
                    if self.consume_char('(') {
                        return self.parse_lambda_call_arguments(lambda);
                    }
                    return Ok(lambda);
                }
                if formula_text_function_name(identifier.as_str()) {
                    match self.parse_text_function(identifier.as_str()) {
                        Ok(text) => return Ok(FormulaValueProbe::Text(text)),
                        Err(FormulaEvalError::Unsupported) => self.index = checkpoint,
                        Err(error) => return Err(error),
                    }
                } else {
                    self.index = checkpoint;
                }
            } else {
                if identifier.eq_ignore_ascii_case("TRUE") {
                    return Ok(FormulaValueProbe::Bool(true));
                }
                if identifier.eq_ignore_ascii_case("FALSE") {
                    return Ok(FormulaValueProbe::Bool(false));
                }
                if let Some(value) = self.binding_value(identifier.as_str())
                    && self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')'))
                {
                    return Ok(value);
                }
                if self.peek_char().is_none_or(|ch| matches!(ch, ',' | ')'))
                    && let Some(value) = self.defined_name_value_probe(identifier.as_str())?
                {
                    return Ok(value);
                }
                self.index = checkpoint;
            }
        }
        match self.parse_comparison() {
            Ok(value) => Ok(FormulaValueProbe::Number(value)),
            Err(FormulaEvalError::Unsupported) => Err(FormulaEvalError::Unsupported),
            Err(error) => Ok(FormulaValueProbe::Error(error)),
        }
    }

    fn parse_number(&mut self) -> Result<Option<f64>, FormulaEvalError> {
        let start = self.index;
        let mut saw_digit = false;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() {
                saw_digit = true;
                self.index += ch.len_utf8();
            } else if ch == '.' {
                self.index += ch.len_utf8();
            } else {
                break;
            }
        }
        if !saw_digit {
            self.index = start;
            return Ok(None);
        }
        let number = self.input[start..self.index]
            .parse::<f64>()
            .map_err(|_| FormulaEvalError::Value)?;
        Ok(Some(number))
    }

    fn parse_string_literal(&mut self) -> Result<Option<String>, FormulaEvalError> {
        if !self.consume_char('"') {
            return Ok(None);
        }
        let mut value = String::new();
        while let Some(ch) = self.peek_char() {
            self.index += ch.len_utf8();
            if ch == '"' {
                if self.peek_char() == Some('"') {
                    value.push('"');
                    self.index += 1;
                    continue;
                }
                return Ok(Some(value));
            }
            value.push(ch);
        }
        Err(FormulaEvalError::Unsupported)
    }

    fn parse_identifier(&mut self) -> Option<String> {
        let start = self.index;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_alphabetic()
                || ch == '_'
                || ch == '.'
                || (self.index > start && ch.is_ascii_digit())
            {
                self.index += ch.len_utf8();
            } else {
                break;
            }
        }
        (self.index > start).then(|| self.input[start..self.index].to_string())
    }

    fn parse_identifier_at(&self, start: usize) -> Option<(String, usize)> {
        let mut cursor = start;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch.is_ascii_alphabetic()
                || ch == '_'
                || ch == '.'
                || (cursor > start && ch.is_ascii_digit())
            {
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        (cursor > start).then(|| (self.input[start..cursor].to_string(), cursor))
    }

    fn skip_whitespace(&mut self) {
        while let Some(ch) = self.peek_char() {
            if ch.is_whitespace() {
                self.index += ch.len_utf8();
            } else {
                break;
            }
        }
    }

    fn consume_char(&mut self, expected: char) -> bool {
        if self.peek_char() == Some(expected) {
            self.index += expected.len_utf8();
            true
        } else {
            false
        }
    }

    fn consume_comparison_operator(&mut self) -> Option<FormulaComparisonOperator> {
        let remaining = &self.input[self.index..];
        for (token, operator) in [
            ("<>", FormulaComparisonOperator::NotEqual),
            ("<=", FormulaComparisonOperator::LessThanOrEqual),
            (">=", FormulaComparisonOperator::GreaterThanOrEqual),
            ("=", FormulaComparisonOperator::Equal),
            ("<", FormulaComparisonOperator::LessThan),
            (">", FormulaComparisonOperator::GreaterThan),
        ] {
            if remaining.starts_with(token) {
                self.index += token.len();
                return Some(operator);
            }
        }
        None
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.index..].chars().next()
    }
}
