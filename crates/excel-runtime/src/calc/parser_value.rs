//! The typed expression grammar: Excel's operators over typed values.
//!
//! Precedence, loosest first: comparisons, `&`, `+ -`, `* /`, `^`, postfix `%`, then unary
//! `-`/`+`, so `=-2^2` is 4. Operands are typed values (literals, cells, names, and function
//! results), and operators coerce them the way Excel does: `="3"+1` is 4, `=1=1` is TRUE, and
//! `="a"="A"` is TRUE. A construct this grammar does not model returns `Unsupported`, and the
//! evaluator falls back to the older numeric and text grammars.

use super::*;

/// Functions whose numeric implementations return a logical as 1 or 0.
const LOGICAL_FUNCTIONS: [&str; 21] = [
    "AND",
    "OR",
    "XOR",
    "NOT",
    "TRUE",
    "FALSE",
    "ISBLANK",
    "ISERR",
    "ISERROR",
    "ISNA",
    "ISNUMBER",
    "ISLOGICAL",
    "ISNONTEXT",
    "ISTEXT",
    "ISREF",
    "ISFORMULA",
    "ISEVEN",
    "ISODD",
    "ISOMITTED",
    "EXACT",
    "REGEXTEST",
];

/// Orders two values as Excel's comparison operators do: numbers before text before logicals,
/// text without regard to case, and a blank as the empty value of the other operand's type.
pub(super) fn formula_compare_values(
    left: &FormulaValueProbe,
    right: &FormulaValueProbe,
) -> Result<Ordering, FormulaEvalError> {
    use FormulaValueProbe::{Blank, Bool, Number, Text};
    for value in [left, right] {
        match value {
            FormulaValueProbe::Error(error) => return Err(*error),
            FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
                return Err(FormulaEvalError::Unsupported);
            }
            _ => {}
        }
    }
    let rank = |value: &FormulaValueProbe| match value {
        Number(_) => 0,
        Text(_) => 1,
        _ => 2,
    };
    Ok(match (left, right) {
        (Blank, Blank) => Ordering::Equal,
        (Blank, Number(number)) => 0.0_f64.total_cmp(number),
        (Number(number), Blank) => number.total_cmp(&0.0),
        (Blank, Text(text)) => "".cmp(text.as_str()),
        (Text(text), Blank) => text.as_str().cmp(""),
        (Blank, Bool(value)) => false.cmp(value),
        (Bool(value), Blank) => value.cmp(&false),
        (Number(left), Number(right)) => left.partial_cmp(right).unwrap_or(Ordering::Equal),
        (Text(left), Text(right)) => left.to_lowercase().cmp(&right.to_lowercase()),
        (Bool(left), Bool(right)) => left.cmp(right),
        _ => rank(left).cmp(&rank(right)),
    })
}

/// An evaluation error as a value, so the expression around it keeps parsing and functions such
/// as `IFERROR` can catch it. Only `Unsupported` stays an `Err`.
fn formula_error_value(
    result: Result<FormulaValueProbe, FormulaEvalError>,
) -> Result<FormulaValueProbe, FormulaEvalError> {
    match result {
        Err(FormulaEvalError::Unsupported) => Err(FormulaEvalError::Unsupported),
        Err(error) => Ok(FormulaValueProbe::Error(error)),
        value => value,
    }
}

/// The result of an arithmetic operator, or `#NUM!` when it leaves the finite numbers.
fn formula_finite_number(value: f64) -> Result<FormulaValueProbe, FormulaEvalError> {
    if value.is_finite() {
        Ok(FormulaValueProbe::Number(value))
    } else {
        Err(FormulaEvalError::Num)
    }
}

/// `base ^ exponent` with Excel's errors: `0^0` and roots of negative numbers are `#NUM!`, and a
/// negative power of zero is `#DIV/0!`.
pub(super) fn formula_power(base: f64, exponent: f64) -> Result<f64, FormulaEvalError> {
    if base == 0.0 && exponent == 0.0 {
        return Err(FormulaEvalError::Num);
    }
    if base == 0.0 && exponent < 0.0 {
        return Err(FormulaEvalError::Div0);
    }
    let value = base.powf(exponent);
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Num)
    }
}

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    /// A whole formula in the typed grammar.
    pub(super) fn parse_typed_formula(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let value = self.parse_typed_comparison()?;
        self.skip_whitespace();
        if self.index == self.input.len() {
            Ok(value)
        } else {
            Err(FormulaEvalError::Unsupported)
        }
    }

    pub(super) fn parse_typed_comparison(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut value = self.parse_typed_concatenation()?;
        loop {
            self.skip_whitespace();
            let Some(operator) = self.consume_comparison_operator() else {
                return Ok(value);
            };
            let right = self.parse_typed_concatenation()?;
            value = formula_error_value(formula_compare_values(&value, &right).map(|ordering| {
                FormulaValueProbe::Bool(match operator {
                    FormulaComparisonOperator::Equal => ordering == Ordering::Equal,
                    FormulaComparisonOperator::NotEqual => ordering != Ordering::Equal,
                    FormulaComparisonOperator::LessThan => ordering == Ordering::Less,
                    FormulaComparisonOperator::LessThanOrEqual => ordering != Ordering::Greater,
                    FormulaComparisonOperator::GreaterThan => ordering == Ordering::Greater,
                    FormulaComparisonOperator::GreaterThanOrEqual => ordering != Ordering::Less,
                })
            }))?;
        }
    }

    fn parse_typed_concatenation(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut value = self.parse_typed_additive()?;
        loop {
            self.skip_whitespace();
            if !self.consume_char('&') {
                return Ok(value);
            }
            let right = self.parse_typed_additive()?;
            value = formula_error_value((|| {
                let mut text = self.typed_text(value)?;
                text.push_str(&self.typed_text(right)?);
                Ok(FormulaValueProbe::Text(text))
            })())?;
        }
    }

    fn parse_typed_additive(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut value = self.parse_typed_multiplicative()?;
        loop {
            self.skip_whitespace();
            let subtract = if self.consume_char('+') {
                false
            } else if self.consume_char('-') {
                true
            } else {
                return Ok(value);
            };
            let right = self.parse_typed_multiplicative()?;
            value = self.typed_arithmetic(value, right, |left, right| {
                Ok(if subtract { left - right } else { left + right })
            })?;
        }
    }

    fn parse_typed_multiplicative(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut value = self.parse_typed_power()?;
        loop {
            self.skip_whitespace();
            let divide = if self.consume_char('*') {
                false
            } else if self.consume_char('/') {
                true
            } else {
                return Ok(value);
            };
            let right = self.parse_typed_power()?;
            value = self.typed_arithmetic(value, right, |left, right| {
                if !divide {
                    Ok(left * right)
                } else if right == 0.0 {
                    Err(FormulaEvalError::Div0)
                } else {
                    Ok(left / right)
                }
            })?;
        }
    }

    fn parse_typed_power(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut value = self.parse_typed_percent()?;
        loop {
            self.skip_whitespace();
            if !self.consume_char('^') {
                return Ok(value);
            }
            let exponent = self.parse_typed_percent()?;
            value = self.typed_arithmetic(value, exponent, formula_power)?;
        }
    }

    fn parse_typed_percent(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        let mut value = self.parse_typed_unary()?;
        loop {
            self.skip_whitespace();
            if !self.consume_char('%') {
                return Ok(value);
            }
            value = formula_error_value(
                self.typed_number(value)
                    .map(|number| FormulaValueProbe::Number(number / 100.0)),
            )?;
        }
    }

    fn parse_typed_unary(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char('-') {
            let value = self.parse_typed_unary()?;
            return formula_error_value(
                self.typed_number(value)
                    .map(|number| FormulaValueProbe::Number(-number)),
            );
        }
        // Unary plus leaves its operand as it is: `=+"a"` is text.
        if self.consume_char('+') {
            return self.parse_typed_unary();
        }
        self.parse_typed_primary()
    }

    fn parse_typed_primary(&mut self) -> Result<FormulaValueProbe, FormulaEvalError> {
        if self.consume_char('(') {
            let value = self.parse_typed_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return Ok(value);
        }
        if let Some(text) = self.parse_string_literal()? {
            return Ok(FormulaValueProbe::Text(text));
        }
        if let Some(number) = self.parse_number()? {
            return Ok(FormulaValueProbe::Number(number));
        }
        if let Some(error) = self.parse_error_literal() {
            return Ok(FormulaValueProbe::Error(error));
        }
        let checkpoint = self.index;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(') {
                let arguments_start = self.index;
                return match self.parse_typed_function(identifier.as_str()) {
                    Err(FormulaEvalError::Unsupported) => Err(FormulaEvalError::Unsupported),
                    // A function that fails is its error, wherever in its arguments it stopped.
                    Err(error) => {
                        self.index = self.call_end(arguments_start)?;
                        Ok(FormulaValueProbe::Error(error))
                    }
                    value => value,
                };
            }
        }
        self.index = checkpoint;
        if let Some((reference, next_index)) = self.try_parse_reference_set()? {
            self.index = next_index;
            let (sheet_id, rect) = reference.single_area()?;
            if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                return Ok(FormulaValueProbe::Error(FormulaEvalError::Value));
            }
            return formula_error_value(
                self.evaluator
                    .cell_value_or_blank(sheet_id, rect.row_first, rect.col_first)
                    .map(formula_value_probe_from_cell_value),
            );
        }
        if let Some(identifier) = self.parse_identifier() {
            if identifier.eq_ignore_ascii_case("TRUE") {
                return Ok(FormulaValueProbe::Bool(true));
            }
            if identifier.eq_ignore_ascii_case("FALSE") {
                return Ok(FormulaValueProbe::Bool(false));
            }
            if let Some(value) = self.binding_value(identifier.as_str()) {
                return Ok(value);
            }
            if let Some(value) = self.defined_name_value_probe(identifier.as_str())? {
                return Ok(value);
            }
            return Ok(FormulaValueProbe::Error(FormulaEvalError::Name));
        }
        Err(FormulaEvalError::Unsupported)
    }

    /// The index just past the `)` that closes a call whose arguments start at `start`, skipping
    /// string literals, quoted sheet names, and nested parentheses and brackets.
    fn call_end(&self, start: usize) -> Result<usize, FormulaEvalError> {
        let bytes = self.input.as_bytes();
        let mut depth = 1usize;
        let mut index = start;
        while index < bytes.len() {
            match bytes[index] {
                quote @ (b'"' | b'\'') => {
                    index += 1;
                    while index < bytes.len() {
                        if bytes[index] == quote {
                            if bytes.get(index + 1) == Some(&quote) {
                                index += 1;
                            } else {
                                break;
                            }
                        }
                        index += 1;
                    }
                }
                b'(' | b'[' | b'{' => depth += 1,
                b')' | b']' | b'}' => {
                    depth -= 1;
                    if depth == 0 {
                        return Ok(index + 1);
                    }
                }
                _ => {}
            }
            index += 1;
        }
        Err(FormulaEvalError::Unsupported)
    }

    /// A binary arithmetic operator: an error operand is the result, left first; otherwise both
    /// operands coerce to numbers and `operation` combines them.
    fn typed_arithmetic(
        &self,
        left: FormulaValueProbe,
        right: FormulaValueProbe,
        operation: impl Fn(f64, f64) -> Result<f64, FormulaEvalError>,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        formula_error_value((|| {
            let left = self.typed_number(left)?;
            let right = self.typed_number(right)?;
            formula_finite_number(operation(left, right)?)
        })())
    }

    /// A function call whose name and `(` are consumed, as a typed value: functions with typed
    /// implementations return their values, text functions return text, logical functions return
    /// TRUE or FALSE, and every other function returns its number.
    fn parse_typed_function(&mut self, name: &str) -> Result<FormulaValueProbe, FormulaEvalError> {
        let arguments_start = self.index;
        if let Some(value) = self.parse_bound_lambda_call_value(name)? {
            return Ok(value);
        }
        if name.eq_ignore_ascii_case("LAMBDA") {
            let lambda = self.parse_lambda_value_function()?;
            self.skip_whitespace();
            if self.consume_char('(') {
                return self.parse_lambda_call_arguments(lambda);
            }
            return Ok(lambda);
        }
        let upper = name.to_ascii_uppercase();
        let typed = match upper.as_str() {
            "IF" => Some(self.parse_if_value_function()),
            "IFERROR" => Some(self.parse_error_fallback_value_function(false)),
            "IFNA" => Some(self.parse_error_fallback_value_function(true)),
            "IFS" => Some(self.parse_ifs_value_function()),
            "SWITCH" => Some(self.parse_switch_value_function()),
            "CHOOSE" => Some(self.parse_choose_value_function()),
            "INDEX" => Some(self.parse_index_value_function()),
            "LOOKUP" => Some(self.parse_lookup_value_function()),
            "VLOOKUP" => Some(self.parse_table_lookup_value_function(false)),
            "HLOOKUP" => Some(self.parse_table_lookup_value_function(true)),
            "XLOOKUP" => Some(self.parse_xlookup_value_function()),
            "CELL" => Some(self.parse_cell_value_function()),
            "INFO" => Some(self.parse_info_value_function()),
            "LET" => Some(self.parse_let_value_function()),
            "MAKEARRAY" => Some(self.parse_makearray_value_function()),
            "REDUCE" | "SCAN" => Some(self.parse_reduce_scan_value_function(name)),
            "GETPIVOTDATA" => Some(self.parse_getpivotdata_value_function()),
            "INDIRECT" | "OFFSET" | "TRIMRANGE" => {
                Some(self.parse_reference_projection_value_function(name))
            }
            "DAVERAGE" | "DCOUNT" | "DCOUNTA" | "DGET" | "DMAX" | "DMIN" | "DPRODUCT"
            | "DSTDEV" | "DSTDEVP" | "DSUM" | "DVAR" | "DVARP" => {
                Some(self.parse_database_value_function(name))
            }
            _ => None,
        };
        if let Some(value) = typed {
            return value;
        }
        if formula_text_function_name(name) {
            match self.parse_text_function(name) {
                Ok(text) => return Ok(FormulaValueProbe::Text(text)),
                Err(FormulaEvalError::Unsupported) => self.index = arguments_start,
                Err(error) => return Err(error),
            }
        }
        let number = self.parse_function(name)?;
        Ok(if LOGICAL_FUNCTIONS.contains(&upper.as_str()) {
            FormulaValueProbe::Bool(number != 0.0)
        } else {
            FormulaValueProbe::Number(number)
        })
    }

    /// An operand of an arithmetic operator: blanks are 0, logicals 1 or 0, and text converts as
    /// `VALUE` converts it or is `#VALUE!`.
    pub(super) fn typed_number(&self, value: FormulaValueProbe) -> Result<f64, FormulaEvalError> {
        match value {
            FormulaValueProbe::Text(text) => formula_value_text(
                self.evaluator.context.date_system(),
                self.evaluator.context.locale(),
                text.as_str(),
            )
            .map_err(|_| FormulaEvalError::Value),
            value => formula_number_from_value_probe(value),
        }
    }

    /// An operand of `&`: numbers in their general text form, logicals as TRUE or FALSE.
    pub(super) fn typed_text(&self, value: FormulaValueProbe) -> Result<String, FormulaEvalError> {
        match value {
            FormulaValueProbe::Error(error) => Err(error),
            value => formula_value_to_text(value, false),
        }
    }
}
