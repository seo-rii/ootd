//! Text rendering, byte/character slicing, regular expressions, and value-to-text conversion.

use super::*;

pub(crate) fn render_range_text_value(value: &OmValue) -> String {
    match value {
        OmValue::Missing | OmValue::Empty | OmValue::Null => String::new(),
        OmValue::Bool(true) => "TRUE".to_string(),
        OmValue::Bool(false) => "FALSE".to_string(),
        OmValue::Number(number) => number.to_string(),
        OmValue::Text(text) => text.clone(),
        OmValue::Error(error) => formula_cell_error_text(error).to_string(),
        OmValue::Object(_) | OmValue::Array(_) => String::new(),
    }
}

pub(crate) fn formula_cell_error_text(error: &CellError) -> &str {
    error.as_lexical_str()
}

pub(crate) fn format_formula_string_literal(value: &str) -> String {
    format!("\"{}\"", value.replace('"', "\"\""))
}

pub(crate) fn worksheet_function_formula_name(member: &str) -> OmResult<String> {
    let mut name = String::with_capacity(member.len());
    for ch in member.chars() {
        if ch.is_ascii_alphanumeric() {
            name.push(ch.to_ascii_uppercase());
        } else if ch == '.' || ch == '_' {
            name.push('.');
        } else {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                format!("WorksheetFunction.{member} is not a valid worksheet function name"),
            ));
        }
    }
    if name.is_empty() || !name.as_bytes()[0].is_ascii_alphabetic() {
        return Err(OmError::new(
            OmErrorCode::NotFound,
            format!("WorksheetFunction.{member} is not a valid worksheet function name"),
        ));
    }
    Ok(name)
}

pub(super) fn formula_eval_error_text(error: FormulaEvalError) -> &'static str {
    match error {
        FormulaEvalError::Unsupported => "#VALUE!",
        FormulaEvalError::Circular => "#CALC!",
        FormulaEvalError::Null => "#NULL!",
        FormulaEvalError::Div0 => "#DIV/0!",
        FormulaEvalError::Value => "#VALUE!",
        FormulaEvalError::Ref => "#REF!",
        FormulaEvalError::Name => "#NAME?",
        FormulaEvalError::NA => "#N/A",
        FormulaEvalError::Num => "#NUM!",
        FormulaEvalError::GettingData => "#GETTING_DATA",
        FormulaEvalError::Spill => "#SPILL!",
        FormulaEvalError::Calc => "#CALC!",
        FormulaEvalError::Field => "#FIELD!",
        FormulaEvalError::Blocked => "#BLOCKED!",
        FormulaEvalError::Busy => "#BUSY!",
        FormulaEvalError::Connect => "#CONNECT!",
        FormulaEvalError::Python => "#PYTHON!",
        FormulaEvalError::Timeout => "#TIMEOUT!",
        FormulaEvalError::Unknown => "#UNKNOWN!",
    }
}

pub(super) fn formula_strict_text_literal(text: &str) -> String {
    format!("\"{}\"", text.replace('"', "\"\""))
}

pub(super) fn formula_value_to_text(
    value: FormulaValueProbe,
    strict: bool,
) -> Result<String, FormulaEvalError> {
    match value {
        FormulaValueProbe::Blank => Ok(String::new()),
        FormulaValueProbe::Bool(value) => Ok(if value { "TRUE" } else { "FALSE" }.to_string()),
        FormulaValueProbe::Number(value) => formula_text_from_number(value),
        FormulaValueProbe::Text(value) if strict => Ok(formula_strict_text_literal(value.as_str())),
        FormulaValueProbe::Text(value) => Ok(value),
        FormulaValueProbe::Error(error) => Ok(formula_eval_error_text(error).to_string()),
        FormulaValueProbe::Omitted | FormulaValueProbe::Lambda { .. } => {
            Err(FormulaEvalError::Value)
        }
    }
}

pub(super) fn formula_text_byte_width(ch: char) -> usize {
    if ch.is_ascii() { 1 } else { 2 }
}

pub(super) fn formula_text_byte_len(text: &str) -> usize {
    text.chars().map(formula_text_byte_width).sum()
}

pub(super) fn formula_text_byte_slice(text: &str, start: usize, count: usize) -> String {
    if count == 0 {
        return String::new();
    }
    let end = start.saturating_add(count);
    let mut position = 1_usize;
    let mut output = String::new();
    for ch in text.chars() {
        let width = formula_text_byte_width(ch);
        let next_position = position + width;
        if position >= end {
            break;
        }
        if position >= start && next_position <= end {
            output.push(ch);
        }
        position = next_position;
    }
    output
}

pub(super) fn formula_text_char_position_to_byte_position(
    text: &str,
    char_position: usize,
) -> usize {
    let units = text
        .chars()
        .take(char_position.saturating_sub(1))
        .map(formula_text_byte_width)
        .sum::<usize>();
    units + 1
}

pub(super) fn formula_encode_url(text: &str) -> String {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    let mut output = String::new();
    for byte in text.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'-' | b'_' | b'.' | b'~') {
            output.push(byte as char);
        } else {
            output.push('%');
            output.push(HEX[(byte >> 4) as usize] as char);
            output.push(HEX[(byte & 0x0f) as usize] as char);
        }
    }
    output
}

pub(super) fn formula_regex_case_insensitive(
    case_sensitivity: i64,
) -> Result<bool, FormulaEvalError> {
    match case_sensitivity {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(FormulaEvalError::Value),
    }
}

pub(super) fn formula_regex_from_pattern(
    pattern: &str,
    case_insensitive: bool,
) -> Result<Regex, FormulaEvalError> {
    RegexBuilder::new(pattern)
        .case_insensitive(case_insensitive)
        .build()
        .map_err(|_| FormulaEvalError::Value)
}

pub(super) fn formula_selected_text_from_value_probe(
    value: FormulaValueProbe,
) -> Result<String, FormulaEvalError> {
    match value {
        FormulaValueProbe::Text(value) => Ok(value),
        FormulaValueProbe::Error(error) => Err(error),
        FormulaValueProbe::Blank
        | FormulaValueProbe::Bool(_)
        | FormulaValueProbe::Number(_)
        | FormulaValueProbe::Omitted
        | FormulaValueProbe::Lambda { .. } => Err(FormulaEvalError::Unsupported),
    }
}

pub(super) fn formula_array_projection_function_name(name: &str) -> bool {
    name.eq_ignore_ascii_case("BYCOL")
        || name.eq_ignore_ascii_case("BYROW")
        || name.eq_ignore_ascii_case("CHOOSECOLS")
        || name.eq_ignore_ascii_case("CHOOSEROWS")
        || name.eq_ignore_ascii_case("DROP")
        || name.eq_ignore_ascii_case("EXPAND")
        || name.eq_ignore_ascii_case("FILTER")
        || name.eq_ignore_ascii_case("GROUPBY")
        || name.eq_ignore_ascii_case("HSTACK")
        || name.eq_ignore_ascii_case("MAP")
        || name.eq_ignore_ascii_case("PIVOTBY")
        || name.eq_ignore_ascii_case("SORT")
        || name.eq_ignore_ascii_case("SORTBY")
        || name.eq_ignore_ascii_case("TAKE")
        || name.eq_ignore_ascii_case("TOCOL")
        || name.eq_ignore_ascii_case("TOROW")
        || name.eq_ignore_ascii_case("TRANSPOSE")
        || name.eq_ignore_ascii_case("UNIQUE")
        || name.eq_ignore_ascii_case("VSTACK")
        || name.eq_ignore_ascii_case("WRAPCOLS")
        || name.eq_ignore_ascii_case("WRAPROWS")
}

pub(super) fn formula_text_function_name(name: &str) -> bool {
    formula_array_projection_function_name(name)
        || name.eq_ignore_ascii_case("ADDRESS")
        || name.eq_ignore_ascii_case("ARRAYTOTEXT")
        || name.eq_ignore_ascii_case("ASC")
        || name.eq_ignore_ascii_case("BAHTTEXT")
        || name.eq_ignore_ascii_case("CONCAT")
        || name.eq_ignore_ascii_case("CONCATENATE")
        || name.eq_ignore_ascii_case("DGET")
        || name.eq_ignore_ascii_case("DBCS")
        || name.eq_ignore_ascii_case("ENCODEURL")
        || name.eq_ignore_ascii_case("LEFT")
        || name.eq_ignore_ascii_case("LEFTB")
        || name.eq_ignore_ascii_case("RIGHT")
        || name.eq_ignore_ascii_case("RIGHTB")
        || name.eq_ignore_ascii_case("MID")
        || name.eq_ignore_ascii_case("MIDB")
        || name.eq_ignore_ascii_case("BASE")
        || name.eq_ignore_ascii_case("BIN2HEX")
        || name.eq_ignore_ascii_case("BIN2OCT")
        || name.eq_ignore_ascii_case("CELL")
        || name.eq_ignore_ascii_case("CHAR")
        || name.eq_ignore_ascii_case("CLEAN")
        || name.eq_ignore_ascii_case("COMPLEX")
        || name.eq_ignore_ascii_case("CUBEKPIMEMBER")
        || name.eq_ignore_ascii_case("CUBEMEMBER")
        || name.eq_ignore_ascii_case("CUBEMEMBERPROPERTY")
        || name.eq_ignore_ascii_case("CUBERANKEDMEMBER")
        || name.eq_ignore_ascii_case("CUBESET")
        || name.eq_ignore_ascii_case("DEC2BIN")
        || name.eq_ignore_ascii_case("DEC2HEX")
        || name.eq_ignore_ascii_case("DEC2OCT")
        || name.eq_ignore_ascii_case("DOLLAR")
        || name.eq_ignore_ascii_case("DETECTLANGUAGE")
        || name.eq_ignore_ascii_case("FIXED")
        || name.eq_ignore_ascii_case("FILTERXML")
        || name.eq_ignore_ascii_case("GETPIVOTDATA")
        || name.eq_ignore_ascii_case("HYPERLINK")
        || name.eq_ignore_ascii_case("HEX2BIN")
        || name.eq_ignore_ascii_case("HEX2OCT")
        || name.eq_ignore_ascii_case("IMAGE")
        || name.eq_ignore_ascii_case("IMCONJUGATE")
        || name.eq_ignore_ascii_case("IMCOS")
        || name.eq_ignore_ascii_case("IMCOSH")
        || name.eq_ignore_ascii_case("IMCOT")
        || name.eq_ignore_ascii_case("IMCSC")
        || name.eq_ignore_ascii_case("IMCSCH")
        || name.eq_ignore_ascii_case("IMDIV")
        || name.eq_ignore_ascii_case("IMEXP")
        || name.eq_ignore_ascii_case("INFO")
        || name.eq_ignore_ascii_case("INDIRECT")
        || name.eq_ignore_ascii_case("IMLN")
        || name.eq_ignore_ascii_case("LAMBDA")
        || name.eq_ignore_ascii_case("LET")
        || name.eq_ignore_ascii_case("MAKEARRAY")
        || name.eq_ignore_ascii_case("IMLOG10")
        || name.eq_ignore_ascii_case("IMLOG2")
        || name.eq_ignore_ascii_case("IMPOWER")
        || name.eq_ignore_ascii_case("IMPRODUCT")
        || name.eq_ignore_ascii_case("IMSEC")
        || name.eq_ignore_ascii_case("IMSECH")
        || name.eq_ignore_ascii_case("IMSIN")
        || name.eq_ignore_ascii_case("IMSINH")
        || name.eq_ignore_ascii_case("IMSQRT")
        || name.eq_ignore_ascii_case("IMSUB")
        || name.eq_ignore_ascii_case("IMSUM")
        || name.eq_ignore_ascii_case("IMTAN")
        || name.eq_ignore_ascii_case("JIS")
        || name.eq_ignore_ascii_case("OCT2BIN")
        || name.eq_ignore_ascii_case("OCT2HEX")
        || name.eq_ignore_ascii_case("OFFSET")
        || name.eq_ignore_ascii_case("PHONETIC")
        || name.eq_ignore_ascii_case("REDUCE")
        || name.eq_ignore_ascii_case("ROMAN")
        || name.eq_ignore_ascii_case("SCAN")
        || name.eq_ignore_ascii_case("T")
        || name.eq_ignore_ascii_case("TEXT")
        || name.eq_ignore_ascii_case("TEXTSPLIT")
        || name.eq_ignore_ascii_case("UNICHAR")
        || name.eq_ignore_ascii_case("UPPER")
        || name.eq_ignore_ascii_case("LOWER")
        || name.eq_ignore_ascii_case("PROPER")
        || name.eq_ignore_ascii_case("REGEXEXTRACT")
        || name.eq_ignore_ascii_case("REGEXREPLACE")
        || name.eq_ignore_ascii_case("TRIM")
        || name.eq_ignore_ascii_case("TEXTJOIN")
        || name.eq_ignore_ascii_case("TEXTBEFORE")
        || name.eq_ignore_ascii_case("TEXTAFTER")
        || name.eq_ignore_ascii_case("TRIMRANGE")
        || name.eq_ignore_ascii_case("TRANSLATE")
        || name.eq_ignore_ascii_case("REPT")
        || name.eq_ignore_ascii_case("REPLACE")
        || name.eq_ignore_ascii_case("REPLACEB")
        || name.eq_ignore_ascii_case("SUBSTITUTE")
        || name.eq_ignore_ascii_case("VALUETOTEXT")
        || name.eq_ignore_ascii_case("FORMULATEXT")
        || name.eq_ignore_ascii_case("IF")
        || name.eq_ignore_ascii_case("IFS")
        || name.eq_ignore_ascii_case("SWITCH")
        || name.eq_ignore_ascii_case("CHOOSE")
        || name.eq_ignore_ascii_case("INDEX")
        || name.eq_ignore_ascii_case("LOOKUP")
        || name.eq_ignore_ascii_case("VLOOKUP")
        || name.eq_ignore_ascii_case("HLOOKUP")
        || name.eq_ignore_ascii_case("WEBSERVICE")
        || name.eq_ignore_ascii_case("XLOOKUP")
}

pub(super) fn formula_text_from_number(value: f64) -> Result<String, FormulaEvalError> {
    if !value.is_finite() {
        return Err(FormulaEvalError::Value);
    }
    if value == 0.0 {
        return Ok("0".to_string());
    }
    if value.fract() == 0.0 && value >= i64::MIN as f64 && value <= i64::MAX as f64 {
        return Ok((value as i64).to_string());
    }
    Ok(value.to_string())
}

pub(super) fn formula_numbervalue(
    text: &str,
    decimal_separator: &str,
    group_separator: &str,
) -> Result<f64, FormulaEvalError> {
    if decimal_separator.chars().count() != 1
        || group_separator.chars().count() != 1
        || decimal_separator == group_separator
    {
        return Err(FormulaEvalError::Value);
    }
    let decimal_separator = decimal_separator
        .chars()
        .next()
        .ok_or(FormulaEvalError::Value)?;
    let group_separator = group_separator
        .chars()
        .next()
        .ok_or(FormulaEvalError::Value)?;

    let mut body = text.trim();
    if body.is_empty() {
        return Err(FormulaEvalError::Value);
    }
    let mut multiplier = 1.0;
    while let Some(stripped) = body.strip_suffix('%') {
        multiplier /= 100.0;
        body = stripped.trim_end();
    }
    if body.is_empty() {
        return Err(FormulaEvalError::Value);
    }
    let mut normalized = String::with_capacity(body.len());
    for ch in body.chars() {
        if ch == group_separator {
            continue;
        }
        if ch == decimal_separator {
            normalized.push('.');
        } else if !ch.is_whitespace() {
            normalized.push(ch);
        }
    }
    let value = normalized
        .parse::<f64>()
        .map_err(|_| FormulaEvalError::Value)?
        * multiplier;
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Value)
    }
}

pub(super) fn formula_value_text(
    date_system: DateSystem,
    locale: RuntimeLocale,
    text: &str,
) -> Result<f64, FormulaEvalError> {
    let mut body = text.trim();
    if body.is_empty() {
        return Err(FormulaEvalError::Value);
    }

    let mut accounting_negative = false;
    if let Some(inner) = body
        .strip_prefix('(')
        .and_then(|value| value.strip_suffix(')'))
    {
        accounting_negative = true;
        body = inner.trim();
        if body.is_empty() {
            return Err(FormulaEvalError::Value);
        }
    } else if body.contains('(') || body.contains(')') {
        return Err(FormulaEvalError::Value);
    }

    let mut explicit_negative = false;
    if let Some(rest) = body.strip_prefix('+') {
        body = rest.trim_start();
    } else if let Some(rest) = body.strip_prefix('-') {
        explicit_negative = true;
        body = rest.trim_start();
    }
    if accounting_negative && explicit_negative {
        return Err(FormulaEvalError::Value);
    }

    if let Some(rest) = body.strip_prefix('$') {
        body = rest.trim_start();
    }
    if let Some(rest) = body.strip_prefix('+') {
        if explicit_negative {
            return Err(FormulaEvalError::Value);
        }
        body = rest.trim_start();
    } else if let Some(rest) = body.strip_prefix('-') {
        if explicit_negative || accounting_negative {
            return Err(FormulaEvalError::Value);
        }
        explicit_negative = true;
        body = rest.trim_start();
    }
    if body.contains('$') || body.contains('(') || body.contains(')') {
        return Err(FormulaEvalError::Value);
    }

    let decimal = locale.decimal_separator.to_string();
    let group = locale.group_separator.to_string();
    let mut value = match formula_numbervalue(body, &decimal, &group) {
        Ok(value) => value,
        Err(FormulaEvalError::Value) if !accounting_negative && !explicit_negative => {
            if let Ok(value) = formula_datevalue_text(date_system, locale, body) {
                value
            } else if let Ok(value) = formula_timevalue_text(body) {
                value
            } else {
                let mut parsed = None;
                for (index, ch) in body.char_indices() {
                    if !ch.is_whitespace() {
                        continue;
                    }
                    let date_text = body[..index].trim_end();
                    let time_text = body[index..].trim_start();
                    if date_text.is_empty() || time_text.is_empty() {
                        continue;
                    }
                    if let (Ok(date), Ok(time)) = (
                        formula_datevalue_text(date_system, locale, date_text),
                        formula_timevalue_text(time_text),
                    ) {
                        parsed = Some(date + time);
                        break;
                    }
                }
                parsed.ok_or(FormulaEvalError::Value)?
            }
        }
        Err(error) => return Err(error),
    };
    if accounting_negative ^ explicit_negative {
        value = -value;
    }
    if value.is_finite() {
        Ok(value)
    } else {
        Err(FormulaEvalError::Value)
    }
}

pub(super) fn formula_proper_text(text: &str) -> String {
    let mut output = String::new();
    let mut capitalize_next = true;
    for ch in text.chars() {
        if ch.is_alphabetic() {
            if capitalize_next {
                output.extend(ch.to_uppercase());
            } else {
                output.extend(ch.to_lowercase());
            }
            capitalize_next = false;
        } else {
            output.push(ch);
            capitalize_next = true;
        }
    }
    output
}

pub(super) fn formula_text_delimiter_matches(
    text: &str,
    delimiter: &str,
    case_insensitive: bool,
) -> Vec<(usize, usize)> {
    if delimiter.is_empty() {
        return Vec::new();
    }
    let mut matches = Vec::new();
    for (start, _) in text.char_indices() {
        let end = start + delimiter.len();
        if end <= text.len() && text.is_char_boundary(end) {
            let candidate = &text[start..end];
            if candidate == delimiter
                || (case_insensitive && candidate.eq_ignore_ascii_case(delimiter))
            {
                matches.push((start, end));
            }
        }
    }
    matches
}

pub(super) fn formula_detect_language_tag(text: &str) -> &'static str {
    if text
        .chars()
        .any(|ch| ('\u{AC00}'..='\u{D7AF}').contains(&ch))
    {
        return "ko";
    }
    if text
        .chars()
        .any(|ch| ('\u{3040}'..='\u{30FF}').contains(&ch))
    {
        return "ja";
    }
    if text
        .chars()
        .any(|ch| ('\u{4E00}'..='\u{9FFF}').contains(&ch))
    {
        return "zh";
    }
    if text
        .chars()
        .any(|ch| ('\u{0400}'..='\u{04FF}').contains(&ch))
    {
        return "ru";
    }
    if text
        .chars()
        .any(|ch| ('\u{0600}'..='\u{06FF}').contains(&ch))
    {
        return "ar";
    }
    if text.chars().any(|ch| ch.is_ascii_alphabetic()) {
        return "en";
    }
    "und"
}

pub(crate) fn formula_sheet_address_qualifier(sheet_name: &str) -> String {
    if excel_reference_qualifier_needs_quotes(sheet_name) {
        format!("'{}'!", sheet_name.replace('\'', "''"))
    } else {
        format!("{sheet_name}!")
    }
}

pub(super) fn formula_fixed_number_text(
    number: f64,
    decimals: i64,
    use_commas: bool,
) -> Result<String, FormulaEvalError> {
    if !number.is_finite() || !(-127..=127).contains(&decimals) {
        return Err(FormulaEvalError::Value);
    }
    let rounded = if decimals >= 0 {
        let factor = formula_round_factor(decimals as f64)?;
        let scaled = number * factor;
        if !scaled.is_finite() {
            return Err(FormulaEvalError::Num);
        }
        round_half_away_from_zero(scaled) / factor
    } else {
        let factor = formula_round_factor((-decimals) as f64)?;
        let scaled = number / factor;
        if !scaled.is_finite() {
            return Err(FormulaEvalError::Num);
        }
        round_half_away_from_zero(scaled) * factor
    };
    if !rounded.is_finite() {
        return Err(FormulaEvalError::Num);
    }
    let negative = rounded < 0.0;
    let precision = decimals.max(0) as usize;
    let mut body = format!("{:.*}", precision, rounded.abs());
    if use_commas {
        let grouped = {
            let (integer, fraction) = body
                .split_once('.')
                .map(|(integer, fraction)| (integer, Some(fraction)))
                .unwrap_or((body.as_str(), None));
            let mut integer_grouped = String::new();
            for (index, ch) in integer.chars().rev().enumerate() {
                if index > 0 && index % 3 == 0 {
                    integer_grouped.push(',');
                }
                integer_grouped.push(ch);
            }
            let mut grouped = integer_grouped.chars().rev().collect::<String>();
            if let Some(fraction) = fraction {
                grouped.push('.');
                grouped.push_str(fraction);
            }
            grouped
        };
        body = grouped;
    }
    if negative {
        Ok(format!("-{body}"))
    } else {
        Ok(body)
    }
}
