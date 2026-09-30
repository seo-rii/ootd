//! The formula evaluator: cell evaluation, memoization, dependency ordering, and range reads.

use super::*;

/// Whether `&` appears outside string literals, quoted sheet names, and bracketed references.
fn formula_has_top_level_concatenation(formula: &str) -> bool {
    let bytes = formula.as_bytes();
    let mut index = 0usize;
    let mut bracket_depth = 0usize;
    while index < bytes.len() {
        match bytes[index] {
            quote @ (b'"' | b'\'') if bracket_depth == 0 => {
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
            b'[' => bracket_depth += 1,
            b']' => bracket_depth = bracket_depth.saturating_sub(1),
            b'&' if bracket_depth == 0 => return true,
            _ => {}
        }
        index += 1;
    }
    false
}

/// The single cell an implicit intersection selects from `rect` at `position`.
fn implicit_intersection_cell(rect: Rect, position: Option<(u32, u32)>) -> Option<(u32, u32)> {
    if rect.row_first == rect.row_last && rect.col_first == rect.col_last {
        return Some((rect.row_first, rect.col_first));
    }
    let (row, col) = position?;
    if rect.col_first == rect.col_last && (rect.row_first..=rect.row_last).contains(&row) {
        return Some((row, rect.col_first));
    }
    if rect.row_first == rect.row_last && (rect.col_first..=rect.col_last).contains(&col) {
        return Some((rect.row_first, col));
    }
    None
}

pub(crate) struct FormulaEvaluator<'a> {
    pub(super) state: &'a WorkbookState,
    pub(super) context: &'a CalcContext,
    pub(super) visiting: BTreeSet<(SheetId, u32, u32)>,
    pub(super) resolving_names: BTreeSet<DefinedNameId>,
}

impl<'a> FormulaEvaluator<'a> {
    pub(crate) fn new(state: &'a WorkbookState, context: &'a CalcContext) -> Self {
        Self {
            state,
            context,
            visiting: BTreeSet::new(),
            resolving_names: BTreeSet::new(),
        }
    }

    pub(crate) fn evaluate_formula_cell(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Option<CellValue> {
        match self.evaluate_formula_cell_result(sheet_id, row, col) {
            Ok(value) => Some(value),
            Err(error) => error.into_cell_value(),
        }
    }

    pub(crate) fn evaluate_formula_cell_result(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<CellValue, FormulaEvalError> {
        self.evaluate_precedents_in_dependency_order(sheet_id, row, col);
        self.evaluate_cell(sheet_id, row, col)
    }

    /// The stored value of a cell, which iterative calculation reads across a circular reference.
    fn previous_cell_value(&self, key: (SheetId, u32, u32)) -> CellValue {
        self.state
            .worksheet_data()
            .get(&key.0)
            .and_then(|worksheet| worksheet.cells.get(&(key.1, key.2)))
            .map_or(CellValue::Blank, |cell| cell.value.clone())
    }

    /// The formula text of a cell in A1 form, or `None` for a value cell.
    pub(super) fn cell_formula_a1_text(&self, key: (SheetId, u32, u32)) -> Option<String> {
        let (sheet_id, row, col) = key;
        let formula = self
            .state
            .worksheet_data()
            .get(&sheet_id)?
            .cells
            .get(&(row, col))?
            .formula
            .as_ref()?;
        Some(if formula.is_r1c1 {
            convert_formula_r1c1_to_a1(&formula.text, row, col)
        } else {
            formula.text.clone()
        })
    }

    /// Formula cells a formula references lexically, through A1, 3D, and defined-name references.
    /// Function-call names and string literals are skipped; references computed at run time
    /// (`INDIRECT`, `OFFSET`) are resolved by ordinary evaluation instead.
    pub(super) fn lexical_formula_precedents(
        &mut self,
        key: (SheetId, u32, u32),
    ) -> Vec<(SheetId, u32, u32)> {
        let Some(formula_text) = self.cell_formula_a1_text(key) else {
            return Vec::new();
        };
        let (sheet_id, row, col) = key;
        let mut areas = Vec::new();
        {
            let mut parser = FormulaParser::new(&formula_text, self, sheet_id, Some((row, col)));
            let mut previous = None::<char>;
            while let Some(ch) = parser.input[parser.index..].chars().next() {
                if ch == '"' {
                    parser.index += 1;
                    while let Some(inner) = parser.input[parser.index..].chars().next() {
                        parser.index += inner.len_utf8();
                        if inner == '"' {
                            if parser.input[parser.index..].starts_with('"') {
                                parser.index += 1;
                            } else {
                                break;
                            }
                        }
                    }
                    previous = Some('"');
                    continue;
                }
                let at_boundary = previous.is_none_or(|previous| {
                    !previous.is_alphanumeric() && !matches!(previous, '_' | '.' | '$')
                });
                if at_boundary
                    && let Ok(Some((reference, next_index))) = parser.try_parse_reference_set()
                    && next_index > parser.index
                    && !parser.input[next_index..].trim_start().starts_with('(')
                {
                    areas.extend(reference.areas);
                    previous = parser.input[..next_index].chars().next_back();
                    parser.index = next_index;
                    continue;
                }
                previous = Some(ch);
                parser.index += ch.len_utf8();
            }
        }
        let mut precedents = Vec::new();
        for (area_sheet_id, rect) in areas {
            let Some(worksheet) = self.state.worksheet_data().get(&area_sheet_id) else {
                continue;
            };
            for ((cell_row, cell_col), cell) in worksheet
                .cells
                .range((rect.row_first, rect.col_first)..=(rect.row_last, rect.col_last))
            {
                if (rect.col_first..=rect.col_last).contains(cell_col) && cell.formula.is_some() {
                    precedents.push((area_sheet_id, *cell_row, *cell_col));
                }
            }
        }
        precedents
    }

    /// Evaluates the formula precedents of a cell before the cell itself, walking the lexical
    /// dependency graph with an explicit stack. Every precedent result is memoized in the
    /// calculation context, so evaluating the cell afterwards never recurses through a long
    /// chain.
    ///
    /// A lexical back edge does not prove a cycle, because reference-only arguments such as
    /// `AREAS(A1:B2)` or `ROWS(A1:A2)` never read the referenced values. The back-edge target
    /// therefore reads as provisionally circular only while the rest of its cycle is evaluated,
    /// exactly as recursive evaluation would observe it through its visiting guard, and the
    /// target itself is evaluated normally once its precedents are known.
    pub(super) fn evaluate_precedents_in_dependency_order(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) {
        let root = (sheet_id, row, col);
        let mut path_members = BTreeSet::<(SheetId, u32, u32)>::new();
        let mut provisional = BTreeSet::<(SheetId, u32, u32)>::new();
        let mut finished = BTreeSet::<(SheetId, u32, u32)>::new();
        let mut stack = vec![(root, false)];
        while let Some((key, expanded)) = stack.pop() {
            if expanded {
                path_members.remove(&key);
                finished.insert(key);
                if provisional.remove(&key) {
                    self.context.forget_cell_result(key);
                }
                if key != root && self.context.cell_result(key).is_none() {
                    let _ = self.evaluate_cell(key.0, key.1, key.2);
                }
                continue;
            }
            if finished.contains(&key) || self.context.cell_result(key).is_some() {
                continue;
            }
            if path_members.contains(&key) {
                if provisional.insert(key) {
                    let provisional_result = if self.context.iterative() {
                        Ok(self.previous_cell_value(key))
                    } else {
                        Err(FormulaEvalError::Circular)
                    };
                    self.context.remember_cell_result(key, &provisional_result);
                }
                continue;
            }
            path_members.insert(key);
            stack.push((key, true));
            for precedent in self.lexical_formula_precedents(key).into_iter().rev() {
                if !finished.contains(&precedent) && self.context.cell_result(precedent).is_none() {
                    stack.push((precedent, false));
                }
            }
        }
    }

    pub(crate) fn evaluate_dynamic_array_formula_cell_result(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<FormulaArrayResult, FormulaEvalError> {
        self.evaluate_precedents_in_dependency_order(sheet_id, row, col);
        if let Some(Err(FormulaEvalError::Circular)) =
            self.context.cell_result((sheet_id, row, col))
        {
            return Err(FormulaEvalError::Circular);
        }
        let formula = self
            .state
            .worksheet_data()
            .get(&sheet_id)
            .and_then(|worksheet| worksheet.cells.get(&(row, col)))
            .and_then(|cell| cell.formula.as_ref())
            .ok_or(FormulaEvalError::Unsupported)?;
        if !self.visiting.insert((sheet_id, row, col)) {
            return Err(FormulaEvalError::Circular);
        }
        let formula_text = if formula.is_r1c1 {
            convert_formula_r1c1_to_a1(&formula.text, row, col)
        } else {
            formula.text.clone()
        };
        let formula_text =
            self.resolve_implicit_intersections(sheet_id, &formula_text, Some((row, col)));
        let result = {
            let mut parser = FormulaParser::new(&formula_text, self, sheet_id, Some((row, col)));
            parser.parse_dynamic_array_formula()
        };
        let result = match result {
            Ok(result) => Ok(result),
            Err(FormulaEvalError::Unsupported) => {
                match self.evaluate_formula_text(sheet_id, &formula_text, Some((row, col))) {
                    Ok(value) => Ok(FormulaArrayResult::single(value)),
                    Err(FormulaEvalError::Unsupported) => {
                        let mut parser =
                            FormulaParser::new(&formula_text, self, sheet_id, Some((row, col)));
                        parser
                            .parse_value_probe_formula()
                            .and_then(|probe| {
                                formula_cell_value_from_probe(probe)
                                    .ok_or(FormulaEvalError::Unsupported)
                            })
                            .map(FormulaArrayResult::single)
                    }
                    Err(error) => Err(error),
                }
            }
            Err(error) => Err(error),
        };
        self.visiting.remove(&(sheet_id, row, col));
        result
    }

    pub(super) fn evaluate_cell(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<CellValue, FormulaEvalError> {
        let Some(cell) = self
            .state
            .worksheet_data()
            .get(&sheet_id)
            .and_then(|worksheet| worksheet.cells.get(&(row, col)))
        else {
            return Ok(CellValue::Blank);
        };
        let Some(formula) = cell.formula.as_ref() else {
            return Ok(cell.value.clone());
        };
        // Each formula cell is evaluated once per calculation cycle, so every dependent observes
        // the same result (including volatile draws) and shared precedents are not re-evaluated.
        if let Some(result) = self.context.cell_result((sheet_id, row, col)) {
            return result;
        }
        if !self.visiting.insert((sheet_id, row, col)) {
            if self.context.iterative() {
                return Ok(cell.value.clone());
            }
            return Err(FormulaEvalError::Circular);
        }
        let formula_text = if formula.is_r1c1 {
            convert_formula_r1c1_to_a1(&formula.text, row, col)
        } else {
            formula.text.clone()
        };
        let result = self.evaluate_formula_text(sheet_id, &formula_text, Some((row, col)));
        self.visiting.remove(&(sheet_id, row, col));
        self.context
            .remember_cell_result((sheet_id, row, col), &result);
        result
    }

    /// Replaces every implicit-intersection operand `@reference` with the single cell the
    /// reference intersects at `current_position`: the cell itself, the cell in the current row of
    /// a one-column reference, or the cell in the current column of a one-row reference. A
    /// reference with no such cell becomes `#VALUE!`, and `@` before any other operand is
    /// dropped because scalar evaluation already yields a single value.
    pub(super) fn resolve_implicit_intersections(
        &mut self,
        sheet_id: SheetId,
        formula_text: &str,
        current_position: Option<(u32, u32)>,
    ) -> String {
        if !formula_text.contains('@') {
            return formula_text.to_string();
        }
        let bytes = formula_text.as_bytes();
        let mut output = String::with_capacity(formula_text.len());
        let mut index = 0usize;
        while index < bytes.len() {
            match bytes[index] {
                quote @ (b'"' | b'\'') => {
                    let start = index;
                    index += 1;
                    while index < bytes.len() {
                        if bytes[index] == quote {
                            if bytes.get(index + 1) == Some(&quote) {
                                index += 2;
                                continue;
                            }
                            index += 1;
                            break;
                        }
                        index += 1;
                    }
                    output.push_str(&formula_text[start..index]);
                }
                b'[' => {
                    let start = index;
                    let mut depth = 0usize;
                    while index < bytes.len() {
                        match bytes[index] {
                            b'[' => depth += 1,
                            b']' => {
                                depth -= 1;
                                if depth == 0 {
                                    index += 1;
                                    break;
                                }
                            }
                            _ => {}
                        }
                        index += 1;
                    }
                    output.push_str(&formula_text[start..index]);
                }
                b'@' => {
                    let operand_start = index + 1;
                    let parsed = {
                        let mut parser =
                            FormulaParser::new(formula_text, self, sheet_id, current_position);
                        parser.index = operand_start;
                        parser.try_parse_reference_set()
                    };
                    match parsed {
                        Ok(Some((reference, next_index))) if next_index > operand_start => {
                            let token = &formula_text[operand_start..next_index];
                            let qualifier = token.rfind('!').map_or("", |bang| &token[..=bang]);
                            let intersection =
                                reference.single_area().ok().and_then(|(_, rect)| {
                                    implicit_intersection_cell(rect, current_position)
                                });
                            match intersection {
                                Some((row, col)) => {
                                    output.push_str(qualifier);
                                    output.push_str(&format_cell_address(row, col, false, false));
                                }
                                None => output.push_str("#VALUE!"),
                            }
                            index = next_index;
                        }
                        _ => index = operand_start,
                    }
                }
                _ => {
                    let ch = formula_text[index..]
                        .chars()
                        .next()
                        .expect("index is a char boundary");
                    output.push(ch);
                    index += ch.len_utf8();
                }
            }
        }
        output
    }

    pub(crate) fn evaluate_formula_text(
        &mut self,
        sheet_id: SheetId,
        formula_text: &str,
        current_position: Option<(u32, u32)>,
    ) -> Result<CellValue, FormulaEvalError> {
        let resolved =
            self.resolve_implicit_intersections(sheet_id, formula_text, current_position);
        let formula_text = resolved.as_str();
        let text_result = {
            let mut parser = FormulaParser::new(formula_text, self, sheet_id, current_position);
            parser.parse_text_formula()
        };
        match text_result {
            Ok(text) => return Ok(CellValue::Text(text)),
            Err(FormulaEvalError::Unsupported) => {}
            Err(error) => return Err(error),
        }
        let evaluate_typed = |evaluator: &mut Self| {
            let probe = FormulaParser::new(formula_text, evaluator, sheet_id, current_position)
                .parse_value_probe_formula()?;
            match probe {
                FormulaValueProbe::Error(error) => Err(error),
                probe => formula_cell_value_from_probe(probe).ok_or(FormulaEvalError::Unsupported),
            }
        };
        // A top-level `&` makes the result text, which the numeric parser cannot produce.
        if formula_has_top_level_concatenation(formula_text) {
            return evaluate_typed(self);
        }
        let numeric_result = FormulaParser::new(formula_text, self, sheet_id, current_position)
            .parse_formula()
            .map(CellValue::Number);
        if !matches!(numeric_result, Err(FormulaEvalError::Unsupported)) {
            return numeric_result;
        }
        evaluate_typed(self)
    }

    pub(super) fn evaluate_formula_value_probe_text(
        &mut self,
        sheet_id: SheetId,
        formula_text: &str,
        current_position: Option<(u32, u32)>,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        FormulaParser::new(formula_text, self, sheet_id, current_position)
            .parse_value_probe_formula()
    }

    pub(super) fn numeric_cell_value(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<f64, FormulaEvalError> {
        match self.evaluate_cell(sheet_id, row, col) {
            Ok(CellValue::Blank) => Ok(0.0),
            Ok(CellValue::Number(number)) => Ok(number),
            Ok(CellValue::Bool(value)) => Ok(if value { 1.0 } else { 0.0 }),
            Ok(CellValue::Text(_) | CellValue::IsoDateTime(_) | CellValue::RichText(_)) => {
                Err(FormulaEvalError::Value)
            }
            Ok(CellValue::Error(error)) => Err(formula_eval_error_from_cell_error(error)),
            Err(FormulaEvalError::Unsupported) => self
                .state
                .worksheet_data()
                .get(&sheet_id)
                .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                .map(|cell| match &cell.value {
                    CellValue::Blank => Ok(0.0),
                    CellValue::Number(number) => Ok(*number),
                    CellValue::Bool(value) => Ok(if *value { 1.0 } else { 0.0 }),
                    CellValue::Text(_) | CellValue::IsoDateTime(_) | CellValue::RichText(_) => {
                        Err(FormulaEvalError::Value)
                    }
                    CellValue::Error(error) => {
                        Err(formula_eval_error_from_cell_error(error.clone()))
                    }
                })
                .unwrap_or(Ok(0.0)),
            Err(error) => Err(error),
        }
    }

    pub(super) fn numeric_values_in_rect(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<Vec<f64>, FormulaEvalError> {
        let Some(worksheet) = self.state.worksheet_data().get(&sheet_id) else {
            return Err(FormulaEvalError::Ref);
        };
        let keys = worksheet
            .cells
            .keys()
            .copied()
            .filter(|(row, col)| {
                (rect.row_first..=rect.row_last).contains(row)
                    && (rect.col_first..=rect.col_last).contains(col)
            })
            .collect::<Vec<_>>();
        let mut values = Vec::new();
        for (row, col) in keys {
            match self.evaluate_cell(sheet_id, row, col) {
                Ok(CellValue::Number(number)) => values.push(number),
                Ok(CellValue::Error(error)) => {
                    return Err(formula_eval_error_from_cell_error(error));
                }
                Ok(
                    CellValue::Blank
                    | CellValue::Bool(_)
                    | CellValue::Text(_)
                    | CellValue::IsoDateTime(_)
                    | CellValue::RichText(_),
                ) => {}
                Err(FormulaEvalError::Unsupported) => {
                    if let Some(cell) = self
                        .state
                        .worksheet_data()
                        .get(&sheet_id)
                        .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                    {
                        match &cell.value {
                            CellValue::Number(number) => values.push(*number),
                            CellValue::Error(error) => {
                                return Err(formula_eval_error_from_cell_error(error.clone()));
                            }
                            CellValue::Blank
                            | CellValue::Bool(_)
                            | CellValue::Text(_)
                            | CellValue::IsoDateTime(_)
                            | CellValue::RichText(_) => {}
                        }
                    }
                }
                Err(error) => return Err(error),
            }
        }
        Ok(values)
    }

    pub(super) fn numeric_values_in_reference(
        &mut self,
        reference: &FormulaReference,
    ) -> Result<Vec<f64>, FormulaEvalError> {
        let mut values = Vec::new();
        for (sheet_id, rect) in reference.areas() {
            values.extend(self.numeric_values_in_rect(*sheet_id, *rect)?);
        }
        Ok(values)
    }

    pub(super) fn counta_values_in_rect(
        &self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<u64, FormulaEvalError> {
        let Some(worksheet) = self.state.worksheet_data().get(&sheet_id) else {
            return Err(FormulaEvalError::Ref);
        };
        Ok(worksheet
            .cells
            .iter()
            .filter(|((row, col), cell)| {
                (rect.row_first..=rect.row_last).contains(row)
                    && (rect.col_first..=rect.col_last).contains(col)
                    && (cell.formula.is_some() || !matches!(cell.value, CellValue::Blank))
            })
            .count() as u64)
    }

    pub(super) fn counta_values_in_reference(
        &self,
        reference: &FormulaReference,
    ) -> Result<u64, FormulaEvalError> {
        let mut count = 0_u64;
        for (sheet_id, rect) in reference.areas() {
            count += self.counta_values_in_rect(*sheet_id, *rect)?;
        }
        Ok(count)
    }

    pub(super) fn subtotal_numeric_values_in_rect(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<Vec<f64>, FormulaEvalError> {
        let Some(worksheet) = self.state.worksheet_data().get(&sheet_id) else {
            return Err(FormulaEvalError::Ref);
        };
        let keys = worksheet
            .cells
            .keys()
            .copied()
            .filter(|(row, col)| {
                (rect.row_first..=rect.row_last).contains(row)
                    && (rect.col_first..=rect.col_last).contains(col)
            })
            .collect::<Vec<_>>();
        let mut values = Vec::new();
        for (row, col) in keys {
            if self
                .state
                .worksheet_data()
                .get(&sheet_id)
                .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                .and_then(|cell| cell.formula.as_ref())
                .is_some_and(|formula| formula_source_has_top_level_function(formula, "SUBTOTAL"))
            {
                continue;
            }
            match self.evaluate_cell(sheet_id, row, col) {
                Ok(CellValue::Number(number)) => values.push(number),
                Ok(CellValue::Error(error)) => {
                    return Err(formula_eval_error_from_cell_error(error));
                }
                Ok(
                    CellValue::Blank
                    | CellValue::Bool(_)
                    | CellValue::Text(_)
                    | CellValue::IsoDateTime(_)
                    | CellValue::RichText(_),
                ) => {}
                Err(FormulaEvalError::Unsupported) => {
                    if let Some(cell) = self
                        .state
                        .worksheet_data()
                        .get(&sheet_id)
                        .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                    {
                        match &cell.value {
                            CellValue::Number(number) => values.push(*number),
                            CellValue::Error(error) => {
                                return Err(formula_eval_error_from_cell_error(error.clone()));
                            }
                            CellValue::Blank
                            | CellValue::Bool(_)
                            | CellValue::Text(_)
                            | CellValue::IsoDateTime(_)
                            | CellValue::RichText(_) => {}
                        }
                    }
                }
                Err(error) => return Err(error),
            }
        }
        Ok(values)
    }

    pub(super) fn subtotal_counta_values_in_rect(
        &self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<u64, FormulaEvalError> {
        let Some(worksheet) = self.state.worksheet_data().get(&sheet_id) else {
            return Err(FormulaEvalError::Ref);
        };
        Ok(worksheet
            .cells
            .iter()
            .filter(|((row, col), cell)| {
                (rect.row_first..=rect.row_last).contains(row)
                    && (rect.col_first..=rect.col_last).contains(col)
                    && !cell.formula.as_ref().is_some_and(|formula| {
                        formula_source_has_top_level_function(formula, "SUBTOTAL")
                    })
                    && (cell.formula.is_some() || !matches!(cell.value, CellValue::Blank))
            })
            .count() as u64)
    }

    pub(super) fn countblank_values_in_rect(
        &self,
        sheet_id: SheetId,
        rect: Rect,
    ) -> Result<u64, FormulaEvalError> {
        let total = rect
            .checked_cell_count()
            .map_err(|_| FormulaEvalError::Ref)?;
        Ok(total - self.counta_values_in_rect(sheet_id, rect)?)
    }

    pub(super) fn countblank_values_in_reference(
        &self,
        reference: &FormulaReference,
    ) -> Result<u64, FormulaEvalError> {
        let mut count = 0_u64;
        for (sheet_id, rect) in reference.areas() {
            count += self.countblank_values_in_rect(*sheet_id, *rect)?;
        }
        Ok(count)
    }

    pub(super) fn cell_value_or_blank(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<CellValue, FormulaEvalError> {
        if !self.state.worksheet_data().contains_key(&sheet_id) {
            return Err(FormulaEvalError::Ref);
        }
        match self.evaluate_cell(sheet_id, row, col) {
            Ok(value) => Ok(value),
            Err(FormulaEvalError::Unsupported) => Ok(self
                .state
                .worksheet_data()
                .get(&sheet_id)
                .and_then(|worksheet| worksheet.cells.get(&(row, col)))
                .map(|cell| cell.value.clone())
                .unwrap_or(CellValue::Blank)),
            Err(error) => Err(error),
        }
    }

    pub(super) fn lookup_values_in_rect(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        orientation: FormulaLookupOrientation,
    ) -> Result<Vec<FormulaValueProbe>, FormulaEvalError> {
        if !self.state.worksheet_data().contains_key(&sheet_id) {
            return Err(FormulaEvalError::Ref);
        }
        let mut values = Vec::new();
        match orientation {
            FormulaLookupOrientation::FirstColumn => {
                for row in rect.row_first..=rect.row_last {
                    let value = self.cell_value_or_blank(sheet_id, row, rect.col_first)?;
                    values.push(formula_value_probe_from_cell_value(value));
                }
            }
            FormulaLookupOrientation::FirstRow => {
                for col in rect.col_first..=rect.col_last {
                    let value = self.cell_value_or_blank(sheet_id, rect.row_first, col)?;
                    values.push(formula_value_probe_from_cell_value(value));
                }
            }
        }
        Ok(values)
    }

    pub(super) fn lookup_result_at(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        Ok(formula_value_probe_from_cell_value(
            self.cell_value_or_blank(sheet_id, row, col)?,
        ))
    }

    pub(super) fn formula_source_at(
        &self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
    ) -> Result<Option<FormulaSource>, FormulaEvalError> {
        if !self.state.worksheet_data().contains_key(&sheet_id) {
            return Err(FormulaEvalError::Ref);
        }
        Ok(self
            .state
            .worksheet_data()
            .get(&sheet_id)
            .and_then(|worksheet| worksheet.cells.get(&(row, col)))
            .and_then(|cell| cell.formula.clone()))
    }

    pub(super) fn countif_values_in_rect(
        &mut self,
        sheet_id: SheetId,
        rect: Rect,
        criteria: &FormulaCriteria,
    ) -> Result<u64, FormulaEvalError> {
        let mut count = 0_u64;
        for row in rect.row_first..=rect.row_last {
            for col in rect.col_first..=rect.col_last {
                let value = self.cell_value_or_blank(sheet_id, row, col)?;
                if let CellValue::Error(error) = value {
                    return Err(formula_eval_error_from_cell_error(error));
                }
                if criteria.matches(&value) {
                    count += 1;
                }
            }
        }
        Ok(count)
    }

    pub(super) fn sumif_values_in_rect(
        &mut self,
        criteria_sheet_id: SheetId,
        criteria_rect: Rect,
        criteria: &FormulaCriteria,
        sum_sheet_id: SheetId,
        sum_rect: Rect,
    ) -> Result<f64, FormulaEvalError> {
        Ok(self
            .conditional_sum_and_count_in_rect(
                criteria_sheet_id,
                criteria_rect,
                criteria,
                sum_sheet_id,
                sum_rect,
            )?
            .0)
    }

    pub(super) fn averageif_values_in_rect(
        &mut self,
        criteria_sheet_id: SheetId,
        criteria_rect: Rect,
        criteria: &FormulaCriteria,
        average_sheet_id: SheetId,
        average_rect: Rect,
    ) -> Result<f64, FormulaEvalError> {
        let (total, count) = self.conditional_sum_and_count_in_rect(
            criteria_sheet_id,
            criteria_rect,
            criteria,
            average_sheet_id,
            average_rect,
        )?;
        if count == 0 {
            return Err(FormulaEvalError::Div0);
        }
        Ok(total / count as f64)
    }

    pub(super) fn conditional_sum_and_count_in_rect(
        &mut self,
        criteria_sheet_id: SheetId,
        criteria_rect: Rect,
        criteria: &FormulaCriteria,
        value_sheet_id: SheetId,
        value_rect: Rect,
    ) -> Result<(f64, u64), FormulaEvalError> {
        let mut total = 0.0;
        let mut count = 0_u64;
        for row in criteria_rect.row_first..=criteria_rect.row_last {
            for col in criteria_rect.col_first..=criteria_rect.col_last {
                let criteria_value = self.cell_value_or_blank(criteria_sheet_id, row, col)?;
                if let CellValue::Error(error) = criteria_value {
                    return Err(formula_eval_error_from_cell_error(error));
                }
                if !criteria.matches(&criteria_value) {
                    continue;
                }
                let value_row = value_rect.row_first + (row - criteria_rect.row_first);
                let value_col = value_rect.col_first + (col - criteria_rect.col_first);
                let value = self.cell_value_or_blank(value_sheet_id, value_row, value_col)?;
                match value {
                    CellValue::Number(number) => {
                        total += number;
                        count += 1;
                    }
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
        Ok((total, count))
    }

    pub(super) fn countifs_values_in_rects(
        &mut self,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<u64, FormulaEvalError> {
        let base_rect = self.validate_multi_criteria_shapes(None, criteria_ranges)?;
        let mut count = 0_u64;
        for row in base_rect.row_first..=base_rect.row_last {
            for col in base_rect.col_first..=base_rect.col_last {
                let mut matches_all = true;
                for criteria_range in criteria_ranges {
                    if !self.criteria_range_matches_at_offset(
                        base_rect,
                        row,
                        col,
                        criteria_range,
                    )? {
                        matches_all = false;
                        break;
                    }
                }
                if matches_all {
                    count += 1;
                }
            }
        }
        Ok(count)
    }

    pub(super) fn sumifs_values_in_rect(
        &mut self,
        sum_sheet_id: SheetId,
        sum_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<f64, FormulaEvalError> {
        Ok(self
            .multi_criteria_sum_and_count_in_rect(sum_sheet_id, sum_rect, criteria_ranges)?
            .0)
    }

    pub(super) fn averageifs_values_in_rect(
        &mut self,
        average_sheet_id: SheetId,
        average_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<f64, FormulaEvalError> {
        let (total, count) = self.multi_criteria_sum_and_count_in_rect(
            average_sheet_id,
            average_rect,
            criteria_ranges,
        )?;
        if count == 0 {
            return Err(FormulaEvalError::Div0);
        }
        Ok(total / count as f64)
    }

    pub(super) fn multi_criteria_sum_and_count_in_rect(
        &mut self,
        value_sheet_id: SheetId,
        value_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<(f64, u64), FormulaEvalError> {
        let base_rect = self.validate_multi_criteria_shapes(Some(value_rect), criteria_ranges)?;
        let mut total = 0.0;
        let mut count = 0_u64;
        for row in base_rect.row_first..=base_rect.row_last {
            for col in base_rect.col_first..=base_rect.col_last {
                let mut matches_all = true;
                for criteria_range in criteria_ranges {
                    if !self.criteria_range_matches_at_offset(
                        base_rect,
                        row,
                        col,
                        criteria_range,
                    )? {
                        matches_all = false;
                        break;
                    }
                }
                if !matches_all {
                    continue;
                }
                let value_row = value_rect.row_first + (row - base_rect.row_first);
                let value_col = value_rect.col_first + (col - base_rect.col_first);
                let value = self.cell_value_or_blank(value_sheet_id, value_row, value_col)?;
                match value {
                    CellValue::Number(number) => {
                        total += number;
                        count += 1;
                    }
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
        Ok((total, count))
    }

    pub(super) fn minifs_values_in_rect(
        &mut self,
        min_sheet_id: SheetId,
        min_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<f64, FormulaEvalError> {
        self.multi_criteria_extreme_value_in_rect(min_sheet_id, min_rect, criteria_ranges, true)
    }

    pub(super) fn maxifs_values_in_rect(
        &mut self,
        max_sheet_id: SheetId,
        max_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<f64, FormulaEvalError> {
        self.multi_criteria_extreme_value_in_rect(max_sheet_id, max_rect, criteria_ranges, false)
    }

    pub(super) fn multi_criteria_extreme_value_in_rect(
        &mut self,
        value_sheet_id: SheetId,
        value_rect: Rect,
        criteria_ranges: &[FormulaCriteriaRange],
        want_min: bool,
    ) -> Result<f64, FormulaEvalError> {
        let base_rect = self.validate_multi_criteria_shapes(Some(value_rect), criteria_ranges)?;
        let mut best = None::<f64>;
        for row in base_rect.row_first..=base_rect.row_last {
            for col in base_rect.col_first..=base_rect.col_last {
                let mut matches_all = true;
                for criteria_range in criteria_ranges {
                    if !self.criteria_range_matches_at_offset(
                        base_rect,
                        row,
                        col,
                        criteria_range,
                    )? {
                        matches_all = false;
                        break;
                    }
                }
                if !matches_all {
                    continue;
                }
                let value_row = value_rect.row_first + (row - base_rect.row_first);
                let value_col = value_rect.col_first + (col - base_rect.col_first);
                let value = self.cell_value_or_blank(value_sheet_id, value_row, value_col)?;
                match value {
                    CellValue::Number(number) => {
                        best = Some(match best {
                            Some(current) if want_min => current.min(number),
                            Some(current) => current.max(number),
                            None => number,
                        });
                    }
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
        Ok(best.unwrap_or(0.0))
    }

    pub(super) fn criteria_range_matches_at_offset(
        &mut self,
        base_rect: Rect,
        row: u32,
        col: u32,
        criteria_range: &FormulaCriteriaRange,
    ) -> Result<bool, FormulaEvalError> {
        let criteria_row = criteria_range.rect.row_first + (row - base_rect.row_first);
        let criteria_col = criteria_range.rect.col_first + (col - base_rect.col_first);
        let value =
            self.cell_value_or_blank(criteria_range.sheet_id, criteria_row, criteria_col)?;
        if let CellValue::Error(error) = value {
            return Err(formula_eval_error_from_cell_error(error));
        }
        Ok(criteria_range.criteria.matches(&value))
    }

    pub(super) fn validate_multi_criteria_shapes(
        &self,
        value_rect: Option<Rect>,
        criteria_ranges: &[FormulaCriteriaRange],
    ) -> Result<Rect, FormulaEvalError> {
        let Some(base_rect) = criteria_ranges
            .first()
            .map(|criteria_range| criteria_range.rect)
        else {
            return Err(FormulaEvalError::Value);
        };
        if criteria_ranges.iter().any(|criteria_range| {
            criteria_range.rect.width() != base_rect.width()
                || criteria_range.rect.height() != base_rect.height()
        }) {
            return Err(FormulaEvalError::Value);
        }
        if let Some(value_rect) = value_rect {
            if value_rect.width() != base_rect.width() || value_rect.height() != base_rect.height()
            {
                return Err(FormulaEvalError::Value);
            }
        }
        Ok(base_rect)
    }
}
