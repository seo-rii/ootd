//! Conversion between the SpreadsheetML file formula grammar and the grammar users type.
//!
//! Functions added after Excel 2007 are stored with an `_xlfn.` prefix (`_xlfn._xlws.` for the
//! worksheet-only `FILTER` and `SORT`), `LET`/`LAMBDA` parameters with `_xlpm.`, the implicit
//! intersection operator `@x` as `_xlfn.SINGLE(x)`, and the spill operator `A1#` as
//! `_xlfn.ANCHORARRAY(A1)`. The model keeps the typed grammar; only names in
//! [`FUTURE_FUNCTIONS`] are converted, so unknown `_xlfn.` names survive unchanged.

/// Functions that the file grammar stores with the `_xlfn.` prefix.
const FUTURE_FUNCTIONS: &[&str] = &[
    "ACOT",
    "ACOTH",
    "AGGREGATE",
    "ANCHORARRAY",
    "ARABIC",
    "ARRAYTOTEXT",
    "BASE",
    "BETA.DIST",
    "BETA.INV",
    "BINOM.DIST",
    "BINOM.DIST.RANGE",
    "BINOM.INV",
    "BITAND",
    "BITLSHIFT",
    "BITOR",
    "BITRSHIFT",
    "BITXOR",
    "BYCOL",
    "BYROW",
    "CEILING.MATH",
    "CEILING.PRECISE",
    "CHISQ.DIST",
    "CHISQ.DIST.RT",
    "CHISQ.INV",
    "CHISQ.INV.RT",
    "CHISQ.TEST",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "COMBINA",
    "CONCAT",
    "CONFIDENCE.NORM",
    "CONFIDENCE.T",
    "COT",
    "COTH",
    "COVARIANCE.P",
    "COVARIANCE.S",
    "CSC",
    "CSCH",
    "DAYS",
    "DECIMAL",
    "DETECTLANGUAGE",
    "DROP",
    "ECMA.CEILING",
    "ENCODEURL",
    "ERF.PRECISE",
    "ERFC.PRECISE",
    "EXPAND",
    "EXPON.DIST",
    "F.DIST",
    "F.DIST.RT",
    "F.INV",
    "F.INV.RT",
    "F.TEST",
    "FIELDVALUE",
    "FILTER",
    "FILTERXML",
    "FLOOR.MATH",
    "FLOOR.PRECISE",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.SEASONALITY",
    "FORECAST.ETS.STAT",
    "FORECAST.LINEAR",
    "FORMULATEXT",
    "GAMMA",
    "GAMMA.DIST",
    "GAMMA.INV",
    "GAMMALN.PRECISE",
    "GAUSS",
    "GROUPBY",
    "HSTACK",
    "HYPGEOM.DIST",
    "IFNA",
    "IFS",
    "IMAGE",
    "IMCOSH",
    "IMCOT",
    "IMCSC",
    "IMCSCH",
    "IMSEC",
    "IMSECH",
    "IMSINH",
    "IMTAN",
    "ISFORMULA",
    "ISO.CEILING",
    "ISOMITTED",
    "ISOWEEKNUM",
    "LAMBDA",
    "LET",
    "LOGNORM.DIST",
    "LOGNORM.INV",
    "MAKEARRAY",
    "MAP",
    "MAXIFS",
    "MINIFS",
    "MODE.MULT",
    "MODE.SNGL",
    "MUNIT",
    "NEGBINOM.DIST",
    "NETWORKDAYS.INTL",
    "NORM.DIST",
    "NORM.INV",
    "NORM.S.DIST",
    "NORM.S.INV",
    "NUMBERVALUE",
    "PDURATION",
    "PERCENTILE.EXC",
    "PERCENTILE.INC",
    "PERCENTOF",
    "PERCENTRANK.EXC",
    "PERCENTRANK.INC",
    "PERMUTATIONA",
    "PHI",
    "PIVOTBY",
    "POISSON.DIST",
    "QUARTILE.EXC",
    "QUARTILE.INC",
    "RANDARRAY",
    "RANK.AVG",
    "RANK.EQ",
    "REDUCE",
    "REGEXEXTRACT",
    "REGEXREPLACE",
    "REGEXTEST",
    "RRI",
    "SCAN",
    "SEC",
    "SECH",
    "SEQUENCE",
    "SHEET",
    "SHEETS",
    "SINGLE",
    "SKEW.P",
    "SORT",
    "SORTBY",
    "STDEV.P",
    "STDEV.S",
    "STOCKHISTORY",
    "SWITCH",
    "T.DIST",
    "T.DIST.2T",
    "T.DIST.RT",
    "T.INV",
    "T.INV.2T",
    "T.TEST",
    "TAKE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTJOIN",
    "TEXTSPLIT",
    "TOCOL",
    "TOROW",
    "TRANSLATE",
    "TRIMRANGE",
    "UNICHAR",
    "UNICODE",
    "UNIQUE",
    "VALUETOTEXT",
    "VAR.P",
    "VAR.S",
    "VSTACK",
    "WEBSERVICE",
    "WEIBULL.DIST",
    "WORKDAY.INTL",
    "WRAPCOLS",
    "WRAPROWS",
    "XLOOKUP",
    "XMATCH",
    "XOR",
    "Z.TEST",
];
/// Future functions that are additionally stored with the `_xlws.` worksheet-only prefix.
const WORKSHEET_ONLY_FUNCTIONS: &[&str] = &["FILTER", "SORT"];
const FUTURE_PREFIX: &str = "_xlfn.";
const WORKSHEET_PREFIX: &str = "_xlws.";
const PARAMETER_PREFIX: &str = "_xlpm.";

fn is_future_function(name: &str) -> bool {
    FUTURE_FUNCTIONS
        .iter()
        .any(|candidate| candidate.eq_ignore_ascii_case(name))
}

fn is_identifier_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_' || ch == '\\'
}

fn is_identifier_char(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.' | '\\')
}

/// Copies a quoted run (`"…"` string or `'…'` sheet name, with doubled-quote escapes) or a
/// bracketed structured/external reference starting at `start`, returning the end index.
fn quoted_or_bracketed_end(text: &str, start: usize) -> Option<usize> {
    let bytes = text.as_bytes();
    match bytes[start] {
        quote @ (b'"' | b'\'') => {
            let mut index = start + 1;
            while index < bytes.len() {
                if bytes[index] == quote {
                    if bytes.get(index + 1) == Some(&quote) {
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
            let mut index = start;
            while index < bytes.len() {
                match bytes[index] {
                    b'[' => depth += 1,
                    b']' => {
                        depth -= 1;
                        if depth == 0 {
                            return Some(index + 1);
                        }
                    }
                    b'\'' if bytes.get(index + 1).is_some() => index += 1,
                    _ => {}
                }
                index += 1;
            }
            Some(bytes.len())
        }
        _ => None,
    }
}

/// The index just past the `)` matching the `(` at `open`, or `None` if unbalanced.
fn matching_paren_end(text: &str, open: usize) -> Option<usize> {
    let bytes = text.as_bytes();
    let mut depth = 0usize;
    let mut index = open;
    while index < bytes.len() {
        if let Some(end) = quoted_or_bracketed_end(text, index) {
            index = end;
            continue;
        }
        match bytes[index] {
            b'(' => depth += 1,
            b')' => {
                depth -= 1;
                if depth == 0 {
                    return Some(index + 1);
                }
            }
            _ => {}
        }
        index += 1;
    }
    None
}

/// Splits call arguments on top-level commas.
fn split_arguments(inner: &str) -> Vec<&str> {
    let bytes = inner.as_bytes();
    let mut arguments = Vec::new();
    let mut depth = 0usize;
    let mut start = 0usize;
    let mut index = 0usize;
    while index < bytes.len() {
        if let Some(end) = quoted_or_bracketed_end(inner, index) {
            index = end;
            continue;
        }
        match bytes[index] {
            b'(' | b'{' => depth += 1,
            b')' | b'}' => depth = depth.saturating_sub(1),
            b',' if depth == 0 => {
                arguments.push(&inner[start..index]);
                start = index + 1;
            }
            _ => {}
        }
        index += 1;
    }
    arguments.push(&inner[start..]);
    arguments
}

/// Whether an operand can follow `@` or precede `#` without parentheses: a single reference,
/// name, or function call.
fn is_simple_operand(operand: &str) -> bool {
    let operand = operand.trim();
    if operand.is_empty() {
        return false;
    }
    let mut index = 0usize;
    let bytes = operand.as_bytes();
    if bytes[0] == b'\'' {
        match quoted_or_bracketed_end(operand, 0) {
            Some(end) if operand[end..].starts_with('!') => index = end + 1,
            _ => return false,
        }
    }
    while index < bytes.len() {
        let ch = operand[index..]
            .chars()
            .next()
            .expect("index is a char boundary");
        if ch == '(' {
            return matching_paren_end(operand, index) == Some(operand.len());
        }
        if ch == '[' {
            match quoted_or_bracketed_end(operand, index) {
                Some(end) => {
                    index = end;
                    continue;
                }
                None => return false,
            }
        }
        if !(is_identifier_char(ch) || matches!(ch, '$' | ':' | '!')) {
            return false;
        }
        index += ch.len_utf8();
    }
    true
}

/// Converts file-grammar formula text to the typed grammar.
pub(crate) fn file_formula_to_model(text: &str) -> String {
    let mut output = String::with_capacity(text.len());
    let mut index = 0usize;
    while index < text.len() {
        if let Some(end) = quoted_or_bracketed_end(text, index) {
            output.push_str(&text[index..end]);
            index = end;
            continue;
        }
        let ch = text[index..]
            .chars()
            .next()
            .expect("index is a char boundary");
        let previous_is_identifier = output.chars().next_back().is_some_and(is_identifier_char);
        if !is_identifier_start(ch) || previous_is_identifier {
            output.push(ch);
            index += ch.len_utf8();
            continue;
        }
        let end = text[index..]
            .char_indices()
            .find(|(_, candidate)| !is_identifier_char(*candidate))
            .map_or(text.len(), |(offset, _)| index + offset);
        let identifier = &text[index..end];
        let is_call = text[end..].starts_with('(');
        if let Some(parameter) = identifier.strip_prefix(PARAMETER_PREFIX)
            && !parameter.is_empty()
        {
            output.push_str(parameter);
            index = end;
            continue;
        }
        let Some(future) = identifier.strip_prefix(FUTURE_PREFIX) else {
            output.push_str(identifier);
            index = end;
            continue;
        };
        let name = future.strip_prefix(WORKSHEET_PREFIX).unwrap_or(future);
        if !is_call || !is_future_function(name) {
            output.push_str(identifier);
            index = end;
            continue;
        }
        let special = ["SINGLE", "ANCHORARRAY"]
            .into_iter()
            .find(|special| name.eq_ignore_ascii_case(special));
        if let Some(special) = special
            && let Some(close) = matching_paren_end(text, end)
        {
            let inner = file_formula_to_model(&text[end + 1..close - 1]);
            if split_arguments(&inner).len() == 1 {
                let simple = is_simple_operand(&inner);
                match special {
                    "SINGLE" if simple => output.push_str(&format!("@{}", inner.trim())),
                    "SINGLE" => output.push_str(&format!("@({inner})")),
                    _ if simple => output.push_str(&format!("{}#", inner.trim())),
                    _ => output.push_str(&format!("{name}({inner})")),
                }
                index = close;
                continue;
            }
        }
        output.push_str(name);
        index = end;
    }
    output
}

/// Converts typed-grammar formula text to the file grammar.
pub(crate) fn model_formula_to_file(text: &str) -> String {
    convert_model_span(text, &[])
}

fn convert_model_span(text: &str, parameters: &[String]) -> String {
    let mut output = String::with_capacity(text.len() + 8);
    let mut index = 0usize;
    while index < text.len() {
        if let Some(end) = quoted_or_bracketed_end(text, index) {
            // A quoted sheet name or bracketed reference may be the operand of `#`.
            let token_end = spill_operand_end(text, index).unwrap_or(end);
            if token_end > end && text[token_end..].starts_with('#') {
                output.push_str(&format!(
                    "{FUTURE_PREFIX}ANCHORARRAY({})",
                    &text[index..token_end]
                ));
                index = token_end + 1;
                continue;
            }
            output.push_str(&text[index..end]);
            index = end;
            continue;
        }
        let ch = text[index..]
            .chars()
            .next()
            .expect("index is a char boundary");
        if ch == '@' {
            let operand_start = index + 1;
            let (operand, operand_end) = if text[operand_start..].starts_with('(') {
                match matching_paren_end(text, operand_start) {
                    Some(close) => (&text[operand_start + 1..close - 1], close),
                    None => (&text[operand_start..], text.len()),
                }
            } else {
                let end = spill_operand_end(text, operand_start).unwrap_or(operand_start);
                (&text[operand_start..end], end)
            };
            if !operand.is_empty() {
                output.push_str(&format!(
                    "{FUTURE_PREFIX}SINGLE({})",
                    convert_model_span(operand, parameters)
                ));
                index = operand_end;
                continue;
            }
        }
        let previous_is_identifier = output.chars().next_back().is_some_and(is_identifier_char);
        if !is_identifier_start(ch) || previous_is_identifier {
            output.push(ch);
            index += ch.len_utf8();
            continue;
        }
        let end = text[index..]
            .char_indices()
            .find(|(_, candidate)| !is_identifier_char(*candidate))
            .map_or(text.len(), |(offset, _)| index + offset);
        let identifier = &text[index..end];
        let is_parameter = parameters
            .iter()
            .any(|parameter| parameter.eq_ignore_ascii_case(identifier));
        if text[end..].starts_with('(') {
            let Some(close) = matching_paren_end(text, end) else {
                output.push_str(&text[index..]);
                break;
            };
            let prefix = if is_parameter {
                PARAMETER_PREFIX.to_string()
            } else if is_future_function(identifier) {
                if WORKSHEET_ONLY_FUNCTIONS
                    .iter()
                    .any(|candidate| candidate.eq_ignore_ascii_case(identifier))
                {
                    format!("{FUTURE_PREFIX}{WORKSHEET_PREFIX}")
                } else {
                    FUTURE_PREFIX.to_string()
                }
            } else {
                String::new()
            };
            output.push_str(&prefix);
            output.push_str(identifier);
            output.push('(');
            output.push_str(&convert_call_arguments(
                identifier,
                &text[end + 1..close - 1],
                parameters,
            ));
            output.push(')');
            index = close;
            continue;
        }
        let spill_end = spill_operand_end(text, index).unwrap_or(end);
        if text[spill_end..].starts_with('#') && !is_parameter {
            output.push_str(&format!(
                "{FUTURE_PREFIX}ANCHORARRAY({})",
                &text[index..spill_end]
            ));
            index = spill_end + 1;
            continue;
        }
        if is_parameter {
            output.push_str(PARAMETER_PREFIX);
        }
        output.push_str(identifier);
        index = end;
    }
    output
}

/// Converts the arguments of one call, bringing `LET` and `LAMBDA` parameter names into scope
/// for the arguments that can see them.
fn convert_call_arguments(function: &str, inner: &str, parameters: &[String]) -> String {
    let arguments = split_arguments(inner);
    let is_let = function.eq_ignore_ascii_case("LET");
    let is_lambda = function.eq_ignore_ascii_case("LAMBDA");
    let mut scope = parameters.to_vec();
    if is_lambda {
        scope.extend(
            arguments[..arguments.len().saturating_sub(1)]
                .iter()
                .map(|argument| argument.trim().to_string()),
        );
    }
    let mut converted = Vec::with_capacity(arguments.len());
    for (position, argument) in arguments.iter().enumerate() {
        let declares_let_name = is_let && position % 2 == 0 && position + 1 < arguments.len();
        if declares_let_name {
            scope.push(argument.trim().to_string());
        }
        converted.push(convert_model_span(argument, &scope));
    }
    converted.join(",")
}

/// The end of a reference-like token starting at `start` (optionally sheet-qualified), used as
/// the operand of `@` or `#`.
fn spill_operand_end(text: &str, start: usize) -> Option<usize> {
    let bytes = text.as_bytes();
    let mut index = start;
    if bytes.get(index) == Some(&b'\'') {
        let end = quoted_or_bracketed_end(text, index)?;
        if !text[end..].starts_with('!') {
            return None;
        }
        index = end + 1;
    }
    let token_start = index;
    while index < bytes.len() {
        let ch = text[index..]
            .chars()
            .next()
            .expect("index is a char boundary");
        if is_identifier_char(ch) || matches!(ch, '$' | ':' | '!') {
            index += ch.len_utf8();
        } else if ch == '[' {
            index = quoted_or_bracketed_end(text, index)?;
        } else if ch == '(' && index > token_start {
            index = matching_paren_end(text, index)?;
            break;
        } else {
            break;
        }
    }
    (index > start).then_some(index)
}

#[cfg(test)]
mod tests {
    use super::{file_formula_to_model, model_formula_to_file};

    #[test]
    fn future_function_prefixes_round_trip() {
        for (file, model) in [
            ("_xlfn.XLOOKUP(A1,B:B,C:C)", "XLOOKUP(A1,B:B,C:C)"),
            ("_xlfn._xlws.SORT(A1:A9,1,-1)", "SORT(A1:A9,1,-1)"),
            ("_xlfn._xlws.FILTER(A1:B9,B1:B9>0)", "FILTER(A1:B9,B1:B9>0)"),
            (
                "SUM(_xlfn.IFS(A1>0,1,TRUE,2),3)",
                "SUM(IFS(A1>0,1,TRUE,2),3)",
            ),
            ("_xlfn.CONCAT(\"_xlfn.X(\",A1)", "CONCAT(\"_xlfn.X(\",A1)"),
            (
                "'_xlfn.Sheet'!A1+_xlfn.DAYS(B1,C1)",
                "'_xlfn.Sheet'!A1+DAYS(B1,C1)",
            ),
            ("_xlfn.FUTUREVENDOR(A1)", "_xlfn.FUTUREVENDOR(A1)"),
            ("SUM(A1:A3)", "SUM(A1:A3)"),
            ("_xlfn.SINGLE(A1:A10)", "@A1:A10"),
            ("_xlfn.SINGLE(Sheet2!A:A)+1", "@Sheet2!A:A+1"),
            ("_xlfn.SINGLE(A1:A2+1)", "@(A1:A2+1)"),
            ("_xlfn.SINGLE(INDEX(A1:B2,1,0))", "@INDEX(A1:B2,1,0)"),
            ("SUM(_xlfn.ANCHORARRAY(D1))", "SUM(D1#)"),
            ("_xlfn.ANCHORARRAY('My Sheet'!D1)", "'My Sheet'!D1#"),
            (
                "_xlfn.LET(_xlpm.x,1,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)",
                "LET(x,1,y,x+1,x*y)",
            ),
            (
                "_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))",
                "MAP(A1:A3,LAMBDA(v,v*2))",
            ),
            (
                "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n+1),_xlpm.f(2))",
                "LET(f,LAMBDA(n,n+1),f(2))",
            ),
            ("Table1[@Col]+1", "Table1[@Col]+1"),
            ("\"a@b#c\"&A1", "\"a@b#c\"&A1"),
        ] {
            assert_eq!(file_formula_to_model(file), model, "load {file}");
            assert_eq!(model_formula_to_file(model), file, "save {model}");
        }
    }

    #[test]
    fn typed_parameters_do_not_capture_outer_names() {
        assert_eq!(
            model_formula_to_file("x+LET(x,1,x)"),
            "x+_xlfn.LET(_xlpm.x,1,_xlpm.x)"
        );
        assert_eq!(
            model_formula_to_file("LET(a,b,b,2,a+b)"),
            "_xlfn.LET(_xlpm.a,b,_xlpm.b,2,_xlpm.a+_xlpm.b)"
        );
    }
}
