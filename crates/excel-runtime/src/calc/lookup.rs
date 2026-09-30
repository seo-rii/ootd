//! Exact, approximate, wildcard, and binary lookup matching.

use super::*;

pub(super) fn lookup_match_index_in_values(
    lookup_value: &FormulaValueProbe,
    values: &[FormulaValueProbe],
    mode: FormulaLookupMode,
) -> Result<usize, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = lookup_value {
        return Err(*error);
    }
    if matches!(mode, FormulaLookupMode::Exact) {
        for (index, candidate) in values.iter().enumerate() {
            if formula_value_probe_exact_match(lookup_value, candidate)? {
                return Ok(index);
            }
        }
        return Err(FormulaEvalError::NA);
    }

    let mut best = None;
    for (index, candidate) in values.iter().enumerate() {
        let ordering = formula_value_probe_ordering(candidate, lookup_value)?;
        let Some(ordering) = ordering else {
            continue;
        };
        match mode {
            FormulaLookupMode::ApproxAscending if ordering != Ordering::Greater => {
                best = Some(index);
            }
            FormulaLookupMode::ApproxDescending if ordering != Ordering::Less => {
                best = Some(index);
            }
            FormulaLookupMode::Exact
            | FormulaLookupMode::ApproxAscending
            | FormulaLookupMode::ApproxDescending => {}
        }
    }
    best.ok_or(FormulaEvalError::NA)
}

pub(super) fn xlookup_match_index_in_values(
    lookup_value: &FormulaValueProbe,
    values: &[FormulaValueProbe],
    mode: FormulaXLookupMatchMode,
    search_mode: FormulaXLookupSearchMode,
) -> Result<usize, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = lookup_value {
        return Err(*error);
    }
    match search_mode {
        FormulaXLookupSearchMode::Forward => {
            xlookup_linear_match_index_in_values(lookup_value, values, mode, false)
        }
        FormulaXLookupSearchMode::Reverse => {
            xlookup_linear_match_index_in_values(lookup_value, values, mode, true)
        }
        FormulaXLookupSearchMode::BinaryAscending => {
            xlookup_binary_match_index_in_values(lookup_value, values, mode, true)
        }
        FormulaXLookupSearchMode::BinaryDescending => {
            xlookup_binary_match_index_in_values(lookup_value, values, mode, false)
        }
    }
}

pub(super) fn xlookup_linear_match_index_in_values(
    lookup_value: &FormulaValueProbe,
    values: &[FormulaValueProbe],
    mode: FormulaXLookupMatchMode,
    reverse_search: bool,
) -> Result<usize, FormulaEvalError> {
    let indexes = if reverse_search {
        (0..values.len()).rev().collect::<Vec<_>>()
    } else {
        (0..values.len()).collect::<Vec<_>>()
    };
    for index in indexes.iter().copied() {
        if match mode {
            FormulaXLookupMatchMode::Wildcard => {
                formula_value_probe_wildcard_match(lookup_value, &values[index])?
            }
            FormulaXLookupMatchMode::Exact
            | FormulaXLookupMatchMode::ExactOrNextSmaller
            | FormulaXLookupMatchMode::ExactOrNextLarger => {
                formula_value_probe_exact_match(lookup_value, &values[index])?
            }
        } {
            return Ok(index);
        }
    }
    if matches!(
        mode,
        FormulaXLookupMatchMode::Exact | FormulaXLookupMatchMode::Wildcard
    ) {
        return Err(FormulaEvalError::NA);
    }

    let mut best = None::<(usize, FormulaValueProbe)>;
    for index in indexes {
        let candidate = &values[index];
        let Some(ordering) = formula_value_probe_ordering(candidate, lookup_value)? else {
            continue;
        };
        let is_viable = match mode {
            FormulaXLookupMatchMode::Exact => false,
            FormulaXLookupMatchMode::ExactOrNextSmaller => ordering == Ordering::Less,
            FormulaXLookupMatchMode::ExactOrNextLarger => ordering == Ordering::Greater,
            FormulaXLookupMatchMode::Wildcard => false,
        };
        if !is_viable {
            continue;
        }
        let replace = match &best {
            None => true,
            Some((_, current)) => match formula_value_probe_ordering(candidate, current)? {
                Some(candidate_to_current) => match mode {
                    FormulaXLookupMatchMode::Exact => false,
                    FormulaXLookupMatchMode::ExactOrNextSmaller => {
                        candidate_to_current == Ordering::Greater
                    }
                    FormulaXLookupMatchMode::ExactOrNextLarger => {
                        candidate_to_current == Ordering::Less
                    }
                    FormulaXLookupMatchMode::Wildcard => false,
                },
                None => false,
            },
        };
        if replace {
            best = Some((index, candidate.clone()));
        }
    }
    best.map(|(index, _)| index).ok_or(FormulaEvalError::NA)
}

pub(super) fn xlookup_binary_match_index_in_values(
    lookup_value: &FormulaValueProbe,
    values: &[FormulaValueProbe],
    mode: FormulaXLookupMatchMode,
    ascending: bool,
) -> Result<usize, FormulaEvalError> {
    if matches!(mode, FormulaXLookupMatchMode::Wildcard) {
        return Err(FormulaEvalError::Value);
    }

    let mut comparable = Vec::with_capacity(values.len());
    for (index, candidate) in values.iter().enumerate() {
        if formula_value_probe_ordering(candidate, lookup_value)?.is_some() {
            comparable.push((index, candidate));
        }
    }
    if comparable.is_empty() {
        return Err(FormulaEvalError::NA);
    }

    let mut low = 0_usize;
    let mut high = comparable.len();
    while low < high {
        let mid = low + (high - low) / 2;
        let ordering = formula_value_probe_ordering(comparable[mid].1, lookup_value)?
            .ok_or(FormulaEvalError::NA)?;
        let before_lookup = if ascending {
            ordering == Ordering::Less
        } else {
            ordering == Ordering::Greater
        };
        if before_lookup {
            low = mid + 1;
        } else {
            high = mid;
        }
    }

    if let Some((index, candidate)) = comparable.get(low)
        && formula_value_probe_exact_match(lookup_value, candidate)?
    {
        return Ok(*index);
    }

    let candidate = match mode {
        FormulaXLookupMatchMode::Exact => None,
        FormulaXLookupMatchMode::ExactOrNextSmaller if ascending => {
            low.checked_sub(1).and_then(|index| comparable.get(index))
        }
        FormulaXLookupMatchMode::ExactOrNextSmaller => comparable.get(low),
        FormulaXLookupMatchMode::ExactOrNextLarger if ascending => comparable.get(low),
        FormulaXLookupMatchMode::ExactOrNextLarger => {
            low.checked_sub(1).and_then(|index| comparable.get(index))
        }
        FormulaXLookupMatchMode::Wildcard => None,
    };
    candidate
        .map(|(index, _)| *index)
        .ok_or(FormulaEvalError::NA)
}

pub(super) fn formula_value_probe_exact_match(
    lookup_value: &FormulaValueProbe,
    candidate: &FormulaValueProbe,
) -> Result<bool, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = lookup_value {
        return Err(*error);
    }
    if let FormulaValueProbe::Error(error) = candidate {
        return Err(*error);
    }
    Ok(match (lookup_value, candidate) {
        (FormulaValueProbe::Blank, FormulaValueProbe::Blank) => true,
        (FormulaValueProbe::Bool(left), FormulaValueProbe::Bool(right)) => left == right,
        (FormulaValueProbe::Number(left), FormulaValueProbe::Number(right)) => left == right,
        (FormulaValueProbe::Text(left), FormulaValueProbe::Text(right)) => {
            left.eq_ignore_ascii_case(right)
        }
        _ => false,
    })
}

pub(super) fn formula_value_probe_wildcard_match(
    lookup_value: &FormulaValueProbe,
    candidate: &FormulaValueProbe,
) -> Result<bool, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = lookup_value {
        return Err(*error);
    }
    if let FormulaValueProbe::Error(error) = candidate {
        return Err(*error);
    }
    Ok(match (lookup_value, candidate) {
        (FormulaValueProbe::Text(pattern), FormulaValueProbe::Text(value)) => {
            formula_wildcard_matches(pattern, value, true)
        }
        _ => false,
    })
}

pub(super) fn formula_value_probe_ordering(
    left: &FormulaValueProbe,
    right: &FormulaValueProbe,
) -> Result<Option<Ordering>, FormulaEvalError> {
    if let FormulaValueProbe::Error(error) = left {
        return Err(*error);
    }
    if let FormulaValueProbe::Error(error) = right {
        return Err(*error);
    }
    Ok(match (left, right) {
        (FormulaValueProbe::Bool(left), FormulaValueProbe::Bool(right)) => Some(left.cmp(right)),
        (FormulaValueProbe::Number(left), FormulaValueProbe::Number(right)) => {
            if !left.is_finite() || !right.is_finite() {
                return Err(FormulaEvalError::Value);
            }
            left.partial_cmp(right)
        }
        (FormulaValueProbe::Text(left), FormulaValueProbe::Text(right)) => {
            Some(left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase()))
        }
        _ => None,
    })
}

pub(super) fn formula_wildcard_matches(pattern: &str, value: &str, case_insensitive: bool) -> bool {
    let tokens = formula_wildcard_tokens(pattern, case_insensitive);
    let chars = formula_wildcard_chars(value, case_insensitive);
    formula_wildcard_tokens_match(tokens.as_slice(), chars.as_slice(), false)
}

pub(super) fn formula_wildcard_find(
    pattern: &str,
    value: &str,
    start: usize,
    case_insensitive: bool,
) -> Option<usize> {
    let tokens = formula_wildcard_tokens(pattern, case_insensitive);
    let chars = formula_wildcard_chars(value, case_insensitive);
    for index in start.saturating_sub(1)..=chars.len() {
        if formula_wildcard_tokens_match(tokens.as_slice(), &chars[index..], true) {
            return Some(index + 1);
        }
    }
    None
}

pub(super) fn formula_wildcard_tokens(
    pattern: &str,
    case_insensitive: bool,
) -> Vec<FormulaWildcardToken> {
    let mut tokens = Vec::new();
    let mut chars = pattern.chars();
    while let Some(ch) = chars.next() {
        match ch {
            '*' => tokens.push(FormulaWildcardToken::AnySequence),
            '?' => tokens.push(FormulaWildcardToken::AnyChar),
            '~' => {
                let literal = chars.next().unwrap_or('~');
                tokens.push(FormulaWildcardToken::Literal(formula_wildcard_char(
                    literal,
                    case_insensitive,
                )));
            }
            literal => tokens.push(FormulaWildcardToken::Literal(formula_wildcard_char(
                literal,
                case_insensitive,
            ))),
        }
    }
    tokens
}

pub(super) fn formula_wildcard_chars(value: &str, case_insensitive: bool) -> Vec<char> {
    value
        .chars()
        .map(|ch| formula_wildcard_char(ch, case_insensitive))
        .collect()
}

pub(super) fn formula_wildcard_char(ch: char, case_insensitive: bool) -> char {
    if case_insensitive {
        ch.to_ascii_lowercase()
    } else {
        ch
    }
}

pub(super) fn formula_wildcard_tokens_match(
    tokens: &[FormulaWildcardToken],
    chars: &[char],
    accept_prefix: bool,
) -> bool {
    fn matches_from(
        tokens: &[FormulaWildcardToken],
        chars: &[char],
        accept_prefix: bool,
        token_index: usize,
        char_index: usize,
        memo: &mut [Vec<Option<bool>>],
    ) -> bool {
        if let Some(value) = memo[token_index][char_index] {
            return value;
        }
        let matched = if token_index == tokens.len() {
            accept_prefix || char_index == chars.len()
        } else {
            match tokens[token_index] {
                FormulaWildcardToken::Literal(expected) => {
                    char_index < chars.len()
                        && chars[char_index] == expected
                        && matches_from(
                            tokens,
                            chars,
                            accept_prefix,
                            token_index + 1,
                            char_index + 1,
                            memo,
                        )
                }
                FormulaWildcardToken::AnyChar => {
                    char_index < chars.len()
                        && matches_from(
                            tokens,
                            chars,
                            accept_prefix,
                            token_index + 1,
                            char_index + 1,
                            memo,
                        )
                }
                FormulaWildcardToken::AnySequence => {
                    matches_from(
                        tokens,
                        chars,
                        accept_prefix,
                        token_index + 1,
                        char_index,
                        memo,
                    ) || (char_index < chars.len()
                        && matches_from(
                            tokens,
                            chars,
                            accept_prefix,
                            token_index,
                            char_index + 1,
                            memo,
                        ))
                }
            }
        };
        memo[token_index][char_index] = Some(matched);
        matched
    }

    let mut memo = vec![vec![None; chars.len() + 1]; tokens.len() + 1];
    matches_from(tokens, chars, accept_prefix, 0, 0, &mut memo)
}

pub(super) fn parse_formula_criteria_numeric_literal(
    input: &str,
) -> Option<(FormulaComparisonOperator, f64)> {
    for (prefix, operator) in [
        ("<>", FormulaComparisonOperator::NotEqual),
        ("<=", FormulaComparisonOperator::LessThanOrEqual),
        (">=", FormulaComparisonOperator::GreaterThanOrEqual),
        ("=", FormulaComparisonOperator::Equal),
        ("<", FormulaComparisonOperator::LessThan),
        (">", FormulaComparisonOperator::GreaterThan),
    ] {
        if let Some(rest) = input.strip_prefix(prefix) {
            let operand = rest.trim();
            if operand.is_empty() {
                return None;
            }
            return operand.parse::<f64>().ok().map(|value| (operator, value));
        }
    }
    input
        .parse::<f64>()
        .ok()
        .map(|value| (FormulaComparisonOperator::Equal, value))
}
