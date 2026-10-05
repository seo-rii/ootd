//! Formula parser methods for references and lookups: A1/3D/defined-name references, INDEX/MATCH/XLOOKUP/VLOOKUP, ROW/COLUMN/SHEET, and array projection and generation.

use super::*;

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    pub(super) fn parse_cell_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let info_type = self.parse_text_value_argument()?.to_ascii_lowercase();
        self.skip_whitespace();
        let (target_sheet_id, row, col) = if self.consume_char(')') {
            let Some((row, col)) = self.current_position else {
                return Err(FormulaEvalError::Value);
            };
            (self.sheet_id, row, col)
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (target_sheet_id, rect) = self.parse_reference_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            (target_sheet_id, rect.row_first, rect.col_first)
        };

        match info_type.as_str() {
            "address" => Ok(FormulaValueProbe::Text(format_cell_address(
                row, col, true, true,
            ))),
            "col" => Ok(FormulaValueProbe::Number(col as f64)),
            "row" => Ok(FormulaValueProbe::Number(row as f64)),
            "contents" => {
                let value = self
                    .evaluator
                    .cell_value_or_blank(target_sheet_id, row, col)?;
                Ok(formula_value_probe_from_cell_value(value))
            }
            "filename" => Ok(FormulaValueProbe::Text(String::new())),
            "format" => Ok(FormulaValueProbe::Text("G".to_string())),
            "color" | "parentheses" | "protect" => Ok(FormulaValueProbe::Number(0.0)),
            "prefix" => Ok(FormulaValueProbe::Text(String::new())),
            "type" => {
                let value = self
                    .evaluator
                    .cell_value_or_blank(target_sheet_id, row, col)?;
                Ok(FormulaValueProbe::Text(
                    match value {
                        CellValue::Blank => "b",
                        CellValue::Text(_) | CellValue::RichText(_) => "l",
                        CellValue::Bool(_)
                        | CellValue::Number(_)
                        | CellValue::Error(_)
                        | CellValue::IsoDateTime(_) => "v",
                    }
                    .to_string(),
                ))
            }
            "width" => Ok(FormulaValueProbe::Number(8.0)),
            _ => Err(FormulaEvalError::Value),
        }
    }

    pub(super) fn parse_info_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let info_type = self.parse_text_value_argument()?.to_ascii_lowercase();
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        match info_type.as_str() {
            "directory" => Ok(FormulaValueProbe::Text(String::new())),
            "numfile" => Ok(FormulaValueProbe::Number(
                self.evaluator.state.worksheets().len() as f64,
            )),
            "origin" => Ok(FormulaValueProbe::Text("$A:$A".to_string())),
            "osversion" => Ok(FormulaValueProbe::Text(std::env::consts::OS.to_string())),
            "recalc" => Ok(FormulaValueProbe::Text("Automatic".to_string())),
            "release" => Ok(FormulaValueProbe::Text(APPLICATION_VERSION.to_string())),
            "system" => Ok(FormulaValueProbe::Text(
                if cfg!(target_os = "macos") {
                    "mac"
                } else {
                    "pcdos"
                }
                .to_string(),
            )),
            _ => Err(FormulaEvalError::Value),
        }
    }

    pub(super) fn parse_column_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            let Some((_, col)) = self.current_position else {
                return Err(FormulaEvalError::Value);
            };
            return Ok(col as f64);
        }
        let (_, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(rect.col_first as f64)
    }

    pub(super) fn parse_columns_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (_, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(rect.width() as f64)
    }

    pub(super) fn parse_row_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            let Some((row, _)) = self.current_position else {
                return Err(FormulaEvalError::Value);
            };
            return Ok(row as f64);
        }
        let (_, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(rect.row_first as f64)
    }

    pub(super) fn parse_rows_function(&mut self) -> Result<f64, FormulaEvalError> {
        let (_, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(rect.height() as f64)
    }

    pub(super) fn parse_areas_function(&mut self) -> Result<f64, FormulaEvalError> {
        let reference = self.parse_reference_set_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(reference.len() as f64)
    }

    pub(super) fn parse_sheet_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return self.sheet_index(self.sheet_id);
        }
        let (target_sheet_id, _) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.sheet_index(target_sheet_id)
    }

    pub(super) fn parse_sheets_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(self.evaluator.state.worksheets().len() as f64);
        }
        let reference = self.parse_reference_set_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let sheet_ids = reference
            .areas()
            .iter()
            .map(|(sheet_id, _)| *sheet_id)
            .collect::<BTreeSet<_>>();
        Ok(sheet_ids.len() as f64)
    }

    pub(super) fn sheet_index(&self, sheet_id: SheetId) -> Result<f64, FormulaEvalError> {
        self.evaluator
            .state
            .worksheets()
            .iter()
            .position(|worksheet| worksheet.id == sheet_id)
            .map(|index| index as f64 + 1.0)
            .ok_or(FormulaEvalError::Ref)
    }

    pub(super) fn parse_index_function(&mut self) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_index_value_function()?)
    }

    pub(super) fn parse_index_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let reference = self.parse_index_reference_function()?;
        let (sheet_id, rect) = reference.single_area()?;
        self.evaluator
            .lookup_result_at(sheet_id, rect.row_first, rect.col_first)
    }

    pub(super) fn parse_index_reference_function(
        &mut self,
    ) -> Result<FormulaReference, FormulaEvalError> {
        let reference = self.parse_reference_set_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let row_index = formula_integer_argument(self.parse_comparison()?)?;
        if row_index < 0 {
            return Err(FormulaEvalError::Value);
        }
        let mut col_index = 1_i64;
        self.skip_whitespace();
        if self.consume_char(',') {
            col_index = formula_integer_argument(self.parse_comparison()?)?;
            if col_index < 0 {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
        }
        let mut area_index = 1_i64;
        if self.consume_char(',') {
            area_index = formula_integer_argument(self.parse_comparison()?)?;
            if area_index < 1 {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
        }
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let area_offset = usize::try_from(area_index - 1).map_err(|_| FormulaEvalError::Ref)?;
        let (sheet_id, rect) = reference
            .areas()
            .get(area_offset)
            .copied()
            .ok_or(FormulaEvalError::Ref)?;
        if row_index > i64::from(rect.height()) || col_index > i64::from(rect.width()) {
            return Err(FormulaEvalError::Ref);
        }
        let (row_first, row_last) = if row_index == 0 {
            (rect.row_first, rect.row_last)
        } else {
            let row = rect.row_first + row_index as u32 - 1;
            (row, row)
        };
        let (col_first, col_last) = if col_index == 0 {
            (rect.col_first, rect.col_last)
        } else {
            let col = rect.col_first + col_index as u32 - 1;
            (col, col)
        };
        Ok(FormulaReference::single(
            sheet_id,
            Rect {
                row_first,
                row_last,
                col_first,
                col_last,
            },
        ))
    }

    pub(super) fn parse_match_function(&mut self) -> Result<f64, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (sheet_id, rect) = self.parse_reference_argument()?;
        let mode = self
            .parse_optional_lookup_mode_argument(true)?
            .unwrap_or(FormulaLookupMode::ApproxAscending);
        if rect.width() > 1 && rect.height() > 1 {
            return Err(FormulaEvalError::Value);
        }
        let orientation = if rect.height() == 1 {
            FormulaLookupOrientation::FirstRow
        } else {
            FormulaLookupOrientation::FirstColumn
        };
        let values = self
            .evaluator
            .lookup_values_in_rect(sheet_id, rect, orientation)?;
        let index = lookup_match_index_in_values(&lookup_value, values.as_slice(), mode)?;
        Ok(index as f64 + 1.0)
    }

    pub(super) fn parse_xmatch_function(&mut self) -> Result<f64, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (sheet_id, rect, orientation) = self.parse_lookup_vector_reference_argument()?;
        let mut mode = FormulaXLookupMatchMode::Exact;
        let mut search_mode = FormulaXLookupSearchMode::Forward;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            mode = formula_xlookup_match_mode_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                search_mode = formula_xlookup_search_mode_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        let values = self
            .evaluator
            .lookup_values_in_rect(sheet_id, rect, orientation)?;
        let index =
            xlookup_match_index_in_values(&lookup_value, values.as_slice(), mode, search_mode)?;
        Ok(index as f64 + 1.0)
    }

    pub(super) fn parse_vlookup_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_table_lookup_function(false)
    }

    pub(super) fn parse_hlookup_function(&mut self) -> Result<f64, FormulaEvalError> {
        self.parse_table_lookup_function(true)
    }

    pub(super) fn parse_lookup_function(&mut self) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_lookup_value_function()?)
    }

    pub(super) fn parse_lookup_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (lookup_sheet_id, lookup_rect, lookup_orientation) =
            self.parse_lookup_vector_reference_argument()?;
        let lookup_len = match lookup_orientation {
            FormulaLookupOrientation::FirstColumn => lookup_rect.height(),
            FormulaLookupOrientation::FirstRow => lookup_rect.width(),
        };

        self.skip_whitespace();
        let (return_sheet_id, return_rect, return_orientation) = if self.consume_char(')') {
            (lookup_sheet_id, lookup_rect, lookup_orientation)
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (return_sheet_id, return_rect, return_orientation) =
                self.parse_lookup_vector_reference_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            (return_sheet_id, return_rect, return_orientation)
        };
        let return_len = match return_orientation {
            FormulaLookupOrientation::FirstColumn => return_rect.height(),
            FormulaLookupOrientation::FirstRow => return_rect.width(),
        };
        if lookup_len != return_len {
            return Err(FormulaEvalError::Value);
        }

        let lookup_values = self.evaluator.lookup_values_in_rect(
            lookup_sheet_id,
            lookup_rect,
            lookup_orientation,
        )?;
        let match_index = lookup_match_index_in_values(
            &lookup_value,
            lookup_values.as_slice(),
            FormulaLookupMode::ApproxAscending,
        )?;
        let (row, col) = match return_orientation {
            FormulaLookupOrientation::FirstColumn => (
                return_rect.row_first + match_index as u32,
                return_rect.col_first,
            ),
            FormulaLookupOrientation::FirstRow => (
                return_rect.row_first,
                return_rect.col_first + match_index as u32,
            ),
        };
        self.evaluator.lookup_result_at(return_sheet_id, row, col)
    }

    pub(super) fn parse_table_lookup_function(
        &mut self,
        horizontal: bool,
    ) -> Result<f64, FormulaEvalError> {
        formula_number_from_value_probe(self.parse_table_lookup_value_function(horizontal)?)
    }

    pub(super) fn parse_table_lookup_value_function(
        &mut self,
        horizontal: bool,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (sheet_id, rect) = self.parse_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let result_index = formula_integer_argument(self.parse_comparison()?)?;
        if result_index < 1 {
            return Err(FormulaEvalError::Value);
        }
        let max_result_index = if horizontal {
            rect.height()
        } else {
            rect.width()
        };
        if result_index > i64::from(max_result_index) {
            return Err(FormulaEvalError::Ref);
        }
        let mode = self
            .parse_optional_lookup_mode_argument(false)?
            .unwrap_or(FormulaLookupMode::ApproxAscending);
        let orientation = if horizontal {
            FormulaLookupOrientation::FirstRow
        } else {
            FormulaLookupOrientation::FirstColumn
        };
        let values = self
            .evaluator
            .lookup_values_in_rect(sheet_id, rect, orientation)?;
        let match_index = lookup_match_index_in_values(&lookup_value, values.as_slice(), mode)?;
        let row = if horizontal {
            rect.row_first + result_index as u32 - 1
        } else {
            rect.row_first + match_index as u32
        };
        let col = if horizontal {
            rect.col_first + match_index as u32
        } else {
            rect.col_first + result_index as u32 - 1
        };
        self.evaluator.lookup_result_at(sheet_id, row, col)
    }

    pub(super) fn parse_xlookup_value_function(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let lookup_value = self.parse_lookup_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (lookup_sheet_id, lookup_rect, lookup_orientation) =
            self.parse_lookup_vector_reference_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let (return_sheet_id, return_rect, return_orientation) =
            self.parse_lookup_vector_reference_argument()?;
        let lookup_len = match lookup_orientation {
            FormulaLookupOrientation::FirstColumn => lookup_rect.height(),
            FormulaLookupOrientation::FirstRow => lookup_rect.width(),
        };
        let return_len = match return_orientation {
            FormulaLookupOrientation::FirstColumn => return_rect.height(),
            FormulaLookupOrientation::FirstRow => return_rect.width(),
        };
        if lookup_len != return_len {
            return Err(FormulaEvalError::Value);
        }

        let mut if_not_found = None;
        let mut mode = FormulaXLookupMatchMode::Exact;
        let mut search_mode = FormulaXLookupSearchMode::Forward;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            if_not_found = Some(self.parse_value_probe_argument()?);
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                mode = formula_xlookup_match_mode_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    search_mode = formula_xlookup_search_mode_argument(self.parse_comparison()?)?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        }

        let lookup_values = self.evaluator.lookup_values_in_rect(
            lookup_sheet_id,
            lookup_rect,
            lookup_orientation,
        )?;
        let match_index = match xlookup_match_index_in_values(
            &lookup_value,
            lookup_values.as_slice(),
            mode,
            search_mode,
        ) {
            Ok(index) => index,
            Err(FormulaEvalError::NA) => {
                return if_not_found.ok_or(FormulaEvalError::NA);
            }
            Err(error) => return Err(error),
        };
        let (row, col) = match return_orientation {
            FormulaLookupOrientation::FirstColumn => (
                return_rect.row_first + match_index as u32,
                return_rect.col_first,
            ),
            FormulaLookupOrientation::FirstRow => (
                return_rect.row_first,
                return_rect.col_first + match_index as u32,
            ),
        };
        self.evaluator.lookup_result_at(return_sheet_id, row, col)
    }

    pub(super) fn parse_reference_projection_value_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let reference = self.parse_reference_projection_reference_function(name)?;
        let Some((sheet_id, rect)) = reference.areas().first().copied() else {
            return Err(FormulaEvalError::Ref);
        };
        let value = self
            .evaluator
            .cell_value_or_blank(sheet_id, rect.row_first, rect.col_first)?;
        Ok(formula_value_probe_from_cell_value(value))
    }

    pub(super) fn parse_reference_projection_reference_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaReference, FormulaEvalError> {
        if name.eq_ignore_ascii_case("INDEX") {
            return self.parse_index_reference_function();
        }

        if name.eq_ignore_ascii_case("INDIRECT") {
            let mut reference_text = self.parse_text_value_argument()?;
            let mut a1_style = true;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                a1_style = self.parse_comparison()? != 0.0;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            if !a1_style {
                let (base_row, base_col) = self.current_position.unwrap_or((1, 1));
                reference_text =
                    convert_formula_r1c1_to_a1(reference_text.as_str(), base_row, base_col);
            }
            return parse_formula_reference_text(
                reference_text.as_str(),
                self.sheet_id,
                self.evaluator.state,
            );
        }

        if name.eq_ignore_ascii_case("OFFSET") {
            let (sheet_id, rect) = self.parse_reference_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let row_offset = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let col_offset = formula_integer_argument(self.parse_comparison()?)?;
            let mut height = i64::from(rect.height());
            let mut width = i64::from(rect.width());
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                height = formula_integer_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    width = formula_integer_argument(self.parse_comparison()?)?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            if height < 1 || width < 1 {
                return Err(FormulaEvalError::Ref);
            }
            let row = i64::from(rect.row_first) + row_offset;
            let col = i64::from(rect.col_first) + col_offset;
            if row < 1
                || col < 1
                || row + height - 1 > i64::from(EXCEL_MAX_ROW_INDEX)
                || col + width - 1 > i64::from(EXCEL_MAX_COLUMN_INDEX)
            {
                return Err(FormulaEvalError::Ref);
            }
            return Ok(FormulaReference::single(
                sheet_id,
                Rect {
                    row_first: row as u32,
                    row_last: (row + height - 1) as u32,
                    col_first: col as u32,
                    col_last: (col + width - 1) as u32,
                },
            ));
        }

        if name.eq_ignore_ascii_case("TRIMRANGE") {
            let (sheet_id, rect) = self.parse_reference_argument()?;
            let mut trim_rows = 3_i64;
            let mut trim_cols = 3_i64;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                trim_rows = formula_integer_argument(self.parse_comparison()?)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    trim_cols = formula_integer_argument(self.parse_comparison()?)?;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
            if !(0..=3).contains(&trim_rows) || !(0..=3).contains(&trim_cols) {
                return Err(FormulaEvalError::Value);
            }

            let mut row_first = rect.row_first;
            let mut row_last = rect.row_last;
            if matches!(trim_rows, 1 | 3) {
                while row_first <= row_last {
                    let mut has_value = false;
                    for col in rect.col_first..=rect.col_last {
                        if !matches!(
                            self.evaluator
                                .cell_value_or_blank(sheet_id, row_first, col)?,
                            CellValue::Blank
                        ) {
                            has_value = true;
                            break;
                        }
                    }
                    if has_value {
                        break;
                    }
                    row_first += 1;
                }
            }
            if matches!(trim_rows, 2 | 3) {
                while row_first <= row_last {
                    let mut has_value = false;
                    for col in rect.col_first..=rect.col_last {
                        if !matches!(
                            self.evaluator
                                .cell_value_or_blank(sheet_id, row_last, col)?,
                            CellValue::Blank
                        ) {
                            has_value = true;
                            break;
                        }
                    }
                    if has_value {
                        break;
                    }
                    row_last = row_last.saturating_sub(1);
                }
            }
            if row_first > row_last {
                return Err(FormulaEvalError::Calc);
            }

            let mut col_first = rect.col_first;
            let mut col_last = rect.col_last;
            if matches!(trim_cols, 1 | 3) {
                while col_first <= col_last {
                    let mut has_value = false;
                    for row in row_first..=row_last {
                        if !matches!(
                            self.evaluator
                                .cell_value_or_blank(sheet_id, row, col_first)?,
                            CellValue::Blank
                        ) {
                            has_value = true;
                            break;
                        }
                    }
                    if has_value {
                        break;
                    }
                    col_first += 1;
                }
            }
            if matches!(trim_cols, 2 | 3) {
                while col_first <= col_last {
                    let mut has_value = false;
                    for row in row_first..=row_last {
                        if !matches!(
                            self.evaluator
                                .cell_value_or_blank(sheet_id, row, col_last)?,
                            CellValue::Blank
                        ) {
                            has_value = true;
                            break;
                        }
                    }
                    if has_value {
                        break;
                    }
                    col_last = col_last.saturating_sub(1);
                }
            }
            if col_first > col_last {
                return Err(FormulaEvalError::Calc);
            }
            return Ok(FormulaReference::single(
                sheet_id,
                Rect {
                    row_first,
                    row_last,
                    col_first,
                    col_last,
                },
            ));
        }

        Err(FormulaEvalError::Unsupported)
    }

    pub(super) fn parse_array_projection_value_function(
        &mut self,
        name: &str,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let (sheet_id, rect) = self.parse_reference_argument()?;
        let mut row = rect.row_first;
        let mut col = rect.col_first;

        macro_rules! finish_with_cell {
            () => {{
                let value = self.evaluator.cell_value_or_blank(sheet_id, row, col)?;
                return Ok(formula_value_probe_from_cell_value(value));
            }};
        }
        macro_rules! parse_dimension_count {
            () => {{
                let value = formula_integer_argument(self.parse_comparison()?)?;
                if value < 1 {
                    return Err(FormulaEvalError::Value);
                }
                value
            }};
        }
        macro_rules! parse_selected_offset {
            ($index:expr, $size:expr) => {{
                let index = $index;
                let size = i64::from($size);
                if index == 0 || index.abs() > size {
                    return Err(FormulaEvalError::Value);
                }
                if index > 0 {
                    u32::try_from(index - 1).map_err(|_| FormulaEvalError::Value)?
                } else {
                    u32::try_from(size + index).map_err(|_| FormulaEvalError::Value)?
                }
            }};
        }
        let consume_remaining_arguments = |parser: &mut Self| -> Result<(), FormulaEvalError> {
            loop {
                parser.skip_whitespace();
                if parser.consume_char(')') {
                    return Ok(());
                }
                if !parser.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                parser.skip_whitespace();
                let checkpoint = parser.index;
                if let Some((_, _, next_index)) = parser.try_parse_reference()? {
                    parser.index = next_index;
                } else {
                    parser.index = checkpoint;
                    parser.parse_value_probe_argument()?;
                }
            }
        };

        if name.eq_ignore_ascii_case("GROUPBY") {
            return self.parse_groupby_value_function(sheet_id, rect);
        }

        if name.eq_ignore_ascii_case("PIVOTBY") {
            return self.parse_pivotby_value_function(sheet_id, rect);
        }

        if name.eq_ignore_ascii_case("BYROW") || name.eq_ignore_ascii_case("BYCOL") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let lambda = self.parse_lambda_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let value = self.evaluator.cell_value_or_blank(sheet_id, row, col)?;
            return self
                .evaluate_lambda_value(lambda, vec![formula_value_probe_from_cell_value(value)]);
        }

        if name.eq_ignore_ascii_case("MAP") {
            let mut references = vec![(sheet_id, rect)];
            loop {
                self.skip_whitespace();
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if let Some(lambda) = self.try_parse_lambda_argument()? {
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    let mut arguments = Vec::with_capacity(references.len());
                    for (target_sheet_id, target_rect) in references {
                        if target_rect.width() != rect.width()
                            || target_rect.height() != rect.height()
                        {
                            return Err(FormulaEvalError::Value);
                        }
                        let value = self.evaluator.cell_value_or_blank(
                            target_sheet_id,
                            target_rect.row_first,
                            target_rect.col_first,
                        )?;
                        arguments.push(formula_value_probe_from_cell_value(value));
                    }
                    return self.evaluate_lambda_value(lambda, arguments);
                }
                references.push(self.parse_reference_argument()?);
            }
        }

        if name.eq_ignore_ascii_case("FILTER") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let (include_sheet_id, include_rect) = self.parse_reference_argument()?;
            let mut if_empty = None;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                if_empty = Some(self.parse_value_probe_argument()?);
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
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

            if include_rect.height() == rect.height() && include_rect.width() == 1 {
                for row_offset in 0..rect.height() {
                    let include = self.evaluator.cell_value_or_blank(
                        include_sheet_id,
                        include_rect.row_first + row_offset,
                        include_rect.col_first,
                    )?;
                    if include_value(include)? {
                        row = rect.row_first + row_offset;
                        finish_with_cell!();
                    }
                }
            } else if include_rect.width() == rect.width() && include_rect.height() == 1 {
                for col_offset in 0..rect.width() {
                    let include = self.evaluator.cell_value_or_blank(
                        include_sheet_id,
                        include_rect.row_first,
                        include_rect.col_first + col_offset,
                    )?;
                    if include_value(include)? {
                        col = rect.col_first + col_offset;
                        finish_with_cell!();
                    }
                }
            } else if include_rect.height() == rect.height() && include_rect.width() == rect.width()
            {
                for row_offset in 0..rect.height() {
                    for col_offset in 0..rect.width() {
                        let include = self.evaluator.cell_value_or_blank(
                            include_sheet_id,
                            include_rect.row_first + row_offset,
                            include_rect.col_first + col_offset,
                        )?;
                        if include_value(include)? {
                            row = rect.row_first + row_offset;
                            col = rect.col_first + col_offset;
                            finish_with_cell!();
                        }
                    }
                }
            } else {
                return Err(FormulaEvalError::Value);
            }
            return if_empty.ok_or(FormulaEvalError::Calc);
        }

        if name.eq_ignore_ascii_case("SORT") || name.eq_ignore_ascii_case("SORTBY") {
            let mut sort_index = 1_i64;
            let mut sort_order = 1_i64;
            let mut by_col = false;
            let mut by_sheet_id = sheet_id;
            let mut by_rect = rect;

            if name.eq_ignore_ascii_case("SORT") {
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
                                by_col = self.parse_comparison()? != 0.0;
                            }
                            self.skip_whitespace();
                            if !self.consume_char(')') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                        }
                    }
                }
            } else {
                self.skip_whitespace();
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let by_reference = self.parse_reference_argument()?;
                by_sheet_id = by_reference.0;
                by_rect = by_reference.1;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        sort_order = formula_integer_argument(self.parse_comparison()?)?;
                    }
                    loop {
                        self.skip_whitespace();
                        if self.consume_char(')') {
                            break;
                        }
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        let checkpoint = self.index;
                        if let Some((_, _, next_index)) = self.try_parse_reference()? {
                            self.index = next_index;
                        } else {
                            self.index = checkpoint;
                            self.parse_value_probe_argument()?;
                        }
                    }
                }
            }

            if sort_index < 1 || !matches!(sort_order, -1 | 1) {
                return Err(FormulaEvalError::Value);
            }
            let descending = sort_order == -1;
            let compare_values = |left: &FormulaValueProbe,
                                  right: &FormulaValueProbe|
             -> Result<Ordering, FormulaEvalError> {
                if let FormulaValueProbe::Error(error) = left {
                    return Err(*error);
                }
                if let FormulaValueProbe::Error(error) = right {
                    return Err(*error);
                }
                if let Some(ordering) = formula_value_probe_ordering(left, right)? {
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
                Ok(rank(left).cmp(&rank(right)))
            };
            let sort_indexes = |indexes: &mut Vec<usize>,
                                keys: &[FormulaValueProbe]|
             -> Result<(), FormulaEvalError> {
                for index in 1..indexes.len() {
                    let current = indexes[index];
                    let mut position = index;
                    while position > 0 {
                        let ordering =
                            compare_values(&keys[indexes[position - 1]], &keys[current])?;
                        let should_shift = if descending {
                            ordering == Ordering::Less
                        } else {
                            ordering == Ordering::Greater
                        };
                        if !should_shift {
                            break;
                        }
                        indexes[position] = indexes[position - 1];
                        position -= 1;
                    }
                    indexes[position] = current;
                }
                Ok(())
            };

            if name.eq_ignore_ascii_case("SORT") {
                if by_col {
                    if sort_index > i64::from(rect.height()) {
                        return Err(FormulaEvalError::Value);
                    }
                    let key_row = rect.row_first + sort_index as u32 - 1;
                    let mut keys = Vec::new();
                    for candidate_col in rect.col_first..=rect.col_last {
                        let value =
                            self.evaluator
                                .cell_value_or_blank(sheet_id, key_row, candidate_col)?;
                        keys.push(formula_value_probe_from_cell_value(value));
                    }
                    let mut indexes = (0..keys.len()).collect::<Vec<_>>();
                    sort_indexes(&mut indexes, keys.as_slice())?;
                    col = rect.col_first
                        + u32::try_from(indexes[0]).map_err(|_| FormulaEvalError::Value)?;
                } else {
                    if sort_index > i64::from(rect.width()) {
                        return Err(FormulaEvalError::Value);
                    }
                    let key_col = rect.col_first + sort_index as u32 - 1;
                    let mut keys = Vec::new();
                    for candidate_row in rect.row_first..=rect.row_last {
                        let value =
                            self.evaluator
                                .cell_value_or_blank(sheet_id, candidate_row, key_col)?;
                        keys.push(formula_value_probe_from_cell_value(value));
                    }
                    let mut indexes = (0..keys.len()).collect::<Vec<_>>();
                    sort_indexes(&mut indexes, keys.as_slice())?;
                    row = rect.row_first
                        + u32::try_from(indexes[0]).map_err(|_| FormulaEvalError::Value)?;
                }
            } else if by_rect.height() == rect.height() && by_rect.width() == 1 {
                let mut keys = Vec::new();
                for candidate_row in by_rect.row_first..=by_rect.row_last {
                    let value = self.evaluator.cell_value_or_blank(
                        by_sheet_id,
                        candidate_row,
                        by_rect.col_first,
                    )?;
                    keys.push(formula_value_probe_from_cell_value(value));
                }
                let mut indexes = (0..keys.len()).collect::<Vec<_>>();
                sort_indexes(&mut indexes, keys.as_slice())?;
                row = rect.row_first
                    + u32::try_from(indexes[0]).map_err(|_| FormulaEvalError::Value)?;
            } else if by_rect.width() == rect.width() && by_rect.height() == 1 {
                let mut keys = Vec::new();
                for candidate_col in by_rect.col_first..=by_rect.col_last {
                    let value = self.evaluator.cell_value_or_blank(
                        by_sheet_id,
                        by_rect.row_first,
                        candidate_col,
                    )?;
                    keys.push(formula_value_probe_from_cell_value(value));
                }
                let mut indexes = (0..keys.len()).collect::<Vec<_>>();
                sort_indexes(&mut indexes, keys.as_slice())?;
                col = rect.col_first
                    + u32::try_from(indexes[0]).map_err(|_| FormulaEvalError::Value)?;
            } else {
                return Err(FormulaEvalError::Value);
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("UNIQUE") {
            let mut by_col = false;
            let mut exactly_once = false;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    by_col = self.parse_comparison()? != 0.0;
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
            if !exactly_once {
                finish_with_cell!();
            }

            if by_col {
                for candidate_col in rect.col_first..=rect.col_last {
                    let mut count = 0_u32;
                    for compare_col in rect.col_first..=rect.col_last {
                        let mut same = true;
                        for candidate_row in rect.row_first..=rect.row_last {
                            let left = self.evaluator.cell_value_or_blank(
                                sheet_id,
                                candidate_row,
                                candidate_col,
                            )?;
                            let right = self.evaluator.cell_value_or_blank(
                                sheet_id,
                                candidate_row,
                                compare_col,
                            )?;
                            if !formula_value_probe_exact_match(
                                &formula_value_probe_from_cell_value(left),
                                &formula_value_probe_from_cell_value(right),
                            )? {
                                same = false;
                                break;
                            }
                        }
                        if same {
                            count += 1;
                        }
                    }
                    if count == 1 {
                        col = candidate_col;
                        finish_with_cell!();
                    }
                }
            } else {
                for candidate_row in rect.row_first..=rect.row_last {
                    let mut count = 0_u32;
                    for compare_row in rect.row_first..=rect.row_last {
                        let mut same = true;
                        for candidate_col in rect.col_first..=rect.col_last {
                            let left = self.evaluator.cell_value_or_blank(
                                sheet_id,
                                candidate_row,
                                candidate_col,
                            )?;
                            let right = self.evaluator.cell_value_or_blank(
                                sheet_id,
                                compare_row,
                                candidate_col,
                            )?;
                            if !formula_value_probe_exact_match(
                                &formula_value_probe_from_cell_value(left),
                                &formula_value_probe_from_cell_value(right),
                            )? {
                                same = false;
                                break;
                            }
                        }
                        if same {
                            count += 1;
                        }
                    }
                    if count == 1 {
                        row = candidate_row;
                        finish_with_cell!();
                    }
                }
            }
            return Err(FormulaEvalError::Calc);
        }

        if name.eq_ignore_ascii_case("TAKE") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let rows = formula_integer_argument(self.parse_comparison()?)?;
            if rows == 0 || rows.unsigned_abs() > u64::from(rect.height()) {
                return Err(FormulaEvalError::Calc);
            }
            if rows < 0 {
                row = rect.row_last
                    - u32::try_from(rows.unsigned_abs()).map_err(|_| FormulaEvalError::Calc)?
                    + 1;
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let columns = formula_integer_argument(self.parse_comparison()?)?;
                if columns == 0 || columns.unsigned_abs() > u64::from(rect.width()) {
                    return Err(FormulaEvalError::Calc);
                }
                if columns < 0 {
                    col = rect.col_last
                        - u32::try_from(columns.unsigned_abs())
                            .map_err(|_| FormulaEvalError::Calc)?
                        + 1;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("DROP") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let rows = formula_integer_argument(self.parse_comparison()?)?;
            if rows.unsigned_abs() >= u64::from(rect.height()) {
                return Err(FormulaEvalError::Calc);
            }
            if rows > 0 {
                row = rect.row_first + u32::try_from(rows).map_err(|_| FormulaEvalError::Calc)?;
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let columns = formula_integer_argument(self.parse_comparison()?)?;
                if columns.unsigned_abs() >= u64::from(rect.width()) {
                    return Err(FormulaEvalError::Calc);
                }
                if columns > 0 {
                    col = rect.col_first
                        + u32::try_from(columns).map_err(|_| FormulaEvalError::Calc)?;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("CHOOSECOLS") || name.eq_ignore_ascii_case("CHOOSEROWS") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let selected = formula_integer_argument(self.parse_comparison()?)?;
            if name.eq_ignore_ascii_case("CHOOSECOLS") {
                col = rect.col_first + parse_selected_offset!(selected, rect.width());
            } else {
                row = rect.row_first + parse_selected_offset!(selected, rect.height());
            }
            consume_remaining_arguments(self)?;
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("EXPAND") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let rows = parse_dimension_count!();
            if rows < i64::from(rect.height()) {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    let columns = parse_dimension_count!();
                    if columns < i64::from(rect.width()) {
                        return Err(FormulaEvalError::Value);
                    }
                }
                consume_remaining_arguments(self)?;
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("HSTACK") || name.eq_ignore_ascii_case("VSTACK") {
            consume_remaining_arguments(self)?;
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("TRANSPOSE") {
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("WRAPCOLS") || name.eq_ignore_ascii_case("WRAPROWS") {
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            parse_dimension_count!();
            consume_remaining_arguments(self)?;
            finish_with_cell!();
        }

        if name.eq_ignore_ascii_case("TOCOL") || name.eq_ignore_ascii_case("TOROW") {
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
                    if !(0..=3).contains(&ignore) {
                        return Err(FormulaEvalError::Value);
                    }
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
            let include_value = |value: &CellValue| -> bool {
                let ignore_blank = matches!(ignore, 1 | 3);
                let ignore_error = matches!(ignore, 2 | 3);
                !(ignore_blank && matches!(value, CellValue::Blank))
                    && !(ignore_error && matches!(value, CellValue::Error(_)))
            };
            if scan_by_column {
                for candidate_col in rect.col_first..=rect.col_last {
                    for candidate_row in rect.row_first..=rect.row_last {
                        let value = self.evaluator.cell_value_or_blank(
                            sheet_id,
                            candidate_row,
                            candidate_col,
                        )?;
                        if include_value(&value) {
                            return Ok(formula_value_probe_from_cell_value(value));
                        }
                    }
                }
            } else {
                for candidate_row in rect.row_first..=rect.row_last {
                    for candidate_col in rect.col_first..=rect.col_last {
                        let value = self.evaluator.cell_value_or_blank(
                            sheet_id,
                            candidate_row,
                            candidate_col,
                        )?;
                        if include_value(&value) {
                            return Ok(formula_value_probe_from_cell_value(value));
                        }
                    }
                }
            }
            return Err(FormulaEvalError::Calc);
        }

        Err(FormulaEvalError::Unsupported)
    }

    pub(super) fn parse_sequence_function(&mut self) -> Result<f64, FormulaEvalError> {
        let rows = formula_integer_argument(self.parse_comparison()?)?;
        if rows < 1 {
            return Err(FormulaEvalError::Value);
        }
        let mut columns = 1_i64;
        let mut start = 1.0_f64;
        let mut step = 1.0_f64;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                columns = formula_integer_argument(self.parse_comparison()?)?;
                if columns < 1 {
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
        if !start.is_finite() || !step.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let _ = columns;
        Ok(start)
    }

    pub(super) fn parse_randarray_function(&mut self) -> Result<f64, FormulaEvalError> {
        let mut rows = 1_i64;
        let mut columns = 1_i64;
        let mut min = 0.0_f64;
        let mut max = 1.0_f64;
        let mut whole_number = false;
        self.skip_whitespace();
        if !self.consume_char(')') {
            rows = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    columns = formula_integer_argument(self.parse_comparison()?)?;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        min = self.parse_comparison()?;
                    }
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        self.skip_whitespace();
                        if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                            max = self.parse_comparison()?;
                        }
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            if !self.consume_char(',') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                            whole_number = self.parse_comparison()? != 0.0;
                            self.skip_whitespace();
                            if !self.consume_char(')') {
                                return Err(FormulaEvalError::Unsupported);
                            }
                        }
                    }
                }
            }
        }
        if rows < 1 || columns < 1 {
            return Err(FormulaEvalError::Value);
        }
        if !min.is_finite() || !max.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        if min > max {
            return Err(FormulaEvalError::Value);
        }
        let context = self.evaluator.context;
        if whole_number {
            context.rand_between(min, max)
        } else {
            formula_checked_numeric_result(min + (max - min) * context.rand())
        }
    }

    pub(super) fn parse_reference_argument(&mut self) -> Result<(SheetId, Rect), FormulaEvalError> {
        Ok(self.parse_reference_set_argument()?.single_area()?)
    }

    pub(super) fn parse_reference_set_before_boundary(
        &mut self,
        boundaries: &[char],
    ) -> Result<Option<FormulaReference>, FormulaEvalError> {
        let checkpoint = self.index;
        match self.parse_reference_set_argument() {
            Ok(reference) => {
                self.skip_whitespace();
                if self.peek_char().is_none_or(|ch| boundaries.contains(&ch)) {
                    Ok(Some(reference))
                } else {
                    self.index = checkpoint;
                    Ok(None)
                }
            }
            Err(FormulaEvalError::Value | FormulaEvalError::Unsupported) => {
                self.index = checkpoint;
                Ok(None)
            }
            Err(error) => {
                self.index = checkpoint;
                Err(error)
            }
        }
    }

    pub(super) fn parse_reference_set_argument(
        &mut self,
    ) -> Result<FormulaReference, FormulaEvalError> {
        self.skip_whitespace();
        let checkpoint = self.index;
        if let Some(identifier) = self.parse_identifier() {
            self.skip_whitespace();
            if self.consume_char('(')
                && (identifier.eq_ignore_ascii_case("INDIRECT")
                    || identifier.eq_ignore_ascii_case("INDEX")
                    || identifier.eq_ignore_ascii_case("OFFSET")
                    || identifier.eq_ignore_ascii_case("TRIMRANGE"))
            {
                return self.parse_reference_projection_reference_function(identifier.as_str());
            }
        }
        self.index = checkpoint;
        let Some((reference, next_index)) = self.try_parse_reference_set()? else {
            return Err(FormulaEvalError::Value);
        };
        self.index = next_index;
        Ok(reference)
    }

    pub(super) fn parse_lookup_vector_reference_argument(
        &mut self,
    ) -> Result<(SheetId, Rect, FormulaLookupOrientation), FormulaEvalError> {
        let (sheet_id, rect) = self.parse_reference_argument()?;
        let orientation = if rect.width() == 1 {
            FormulaLookupOrientation::FirstColumn
        } else if rect.height() == 1 {
            FormulaLookupOrientation::FirstRow
        } else {
            return Err(FormulaEvalError::Value);
        };
        Ok((sheet_id, rect, orientation))
    }

    pub(super) fn parse_lookup_value_argument(
        &mut self,
    ) -> Result<FormulaValueProbe, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        if let FormulaValueProbe::Error(error) = value {
            return Err(error);
        }
        Ok(value)
    }

    pub(super) fn parse_optional_lookup_mode_argument(
        &mut self,
        allow_descending: bool,
    ) -> Result<Option<FormulaLookupMode>, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(None);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let mode_value = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(Some(if mode_value == 0 {
            FormulaLookupMode::Exact
        } else if allow_descending && mode_value < 0 {
            FormulaLookupMode::ApproxDescending
        } else {
            FormulaLookupMode::ApproxAscending
        }))
    }

    pub(super) fn try_parse_reference_set(
        &self,
    ) -> Result<Option<(FormulaReference, usize)>, FormulaEvalError> {
        if let Some((reference, next_index)) = self.try_parse_3d_reference()? {
            return Ok(Some((reference, next_index)));
        }
        if let Some((sheet_id, rect, next_index)) = self.try_parse_a1_reference()? {
            return Ok(Some((FormulaReference::single(sheet_id, rect), next_index)));
        }
        self.try_parse_named_reference()
    }

    pub(super) fn try_parse_reference(
        &self,
    ) -> Result<Option<(SheetId, Rect, usize)>, FormulaEvalError> {
        let Some((reference, next_index)) = self.try_parse_reference_set()? else {
            return Ok(None);
        };
        let (sheet_id, rect) = reference.single_area()?;
        Ok(Some((sheet_id, rect, next_index)))
    }

    pub(super) fn try_parse_a1_reference(
        &self,
    ) -> Result<Option<(SheetId, Rect, usize)>, FormulaEvalError> {
        let (sheet_id, start) = self.try_parse_sheet_qualifier()?;
        let mut cursor = start;
        let mut saw_reference_char = false;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch.is_ascii_alphanumeric() || ch == '$' || ch == ':' {
                saw_reference_char = true;
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        if !saw_reference_char {
            return Ok(None);
        }
        let has_spill_operator = self.input[cursor..].starts_with('#');
        let next_index = cursor + usize::from(has_spill_operator);
        let next_is_boundary = self.input[next_index..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if !next_is_boundary {
            return Ok(None);
        }
        let token = &self.input[start..cursor];
        let Some(mut rect) = parse_rect_a1(token).ok() else {
            return Ok(None);
        };
        if has_spill_operator {
            if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                return Err(FormulaEvalError::Ref);
            }
            rect = self
                .evaluator
                .state
                .worksheet_data()
                .get(&sheet_id)
                .and_then(|worksheet| {
                    worksheet
                        .spill_ranges
                        .get(&(rect.row_first, rect.col_first))
                })
                .copied()
                .ok_or(FormulaEvalError::Ref)?;
        }
        Ok(Some((sheet_id, rect, next_index)))
    }

    pub(super) fn try_parse_3d_reference(
        &self,
    ) -> Result<Option<(FormulaReference, usize)>, FormulaEvalError> {
        let Some((start_sheet_id, end_sheet_id, reference_start)) =
            self.try_parse_3d_sheet_span_prefix()?
        else {
            return Ok(None);
        };
        let mut cursor = reference_start;
        let mut saw_reference_char = false;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch.is_ascii_alphanumeric() || ch == '$' || ch == ':' {
                saw_reference_char = true;
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        if !saw_reference_char {
            return Ok(None);
        }
        let next_is_boundary = self.input[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if !next_is_boundary {
            return Ok(None);
        }
        let Some(rect) = parse_rect_a1(&self.input[reference_start..cursor]).ok() else {
            return Ok(None);
        };
        let sheets = formula_sheets_in_3d_span(self.evaluator.state, start_sheet_id, end_sheet_id)?;
        let areas = sheets
            .into_iter()
            .map(|sheet_id| (sheet_id, rect))
            .collect::<Vec<_>>();
        Ok(Some((
            FormulaReference::with_explicit_area_count(1, areas)?,
            cursor,
        )))
    }

    pub(super) fn try_parse_3d_sheet_span_prefix(
        &self,
    ) -> Result<Option<(SheetId, SheetId, usize)>, FormulaEvalError> {
        if self.peek_char() == Some('\'') {
            let mut cursor = self.index + 1;
            let mut sheet_span = String::new();
            while cursor < self.input.len() {
                let ch = self.input[cursor..]
                    .chars()
                    .next()
                    .ok_or(FormulaEvalError::Unsupported)?;
                if ch == '\'' {
                    let next_cursor = cursor + ch.len_utf8();
                    if self.input[next_cursor..].starts_with('\'') {
                        sheet_span.push('\'');
                        cursor = next_cursor + 1;
                        continue;
                    }
                    if !self.input[next_cursor..].starts_with('!') {
                        return Ok(None);
                    }
                    let Some((start_sheet, end_sheet)) = sheet_span.split_once(':') else {
                        return Ok(None);
                    };
                    let start_sheet_id = self
                        .resolve_sheet_name(start_sheet)
                        .ok_or(FormulaEvalError::Ref)?;
                    let end_sheet_id = self
                        .resolve_sheet_name(end_sheet)
                        .ok_or(FormulaEvalError::Ref)?;
                    return Ok(Some((start_sheet_id, end_sheet_id, next_cursor + 1)));
                }
                sheet_span.push(ch);
                cursor += ch.len_utf8();
            }
            return Ok(None);
        }

        let mut cursor = self.index;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch == '!' {
                let sheet_span = &self.input[self.index..cursor];
                let Some((start_sheet, end_sheet)) = sheet_span.split_once(':') else {
                    return Ok(None);
                };
                if start_sheet.is_empty() || end_sheet.is_empty() {
                    return Err(FormulaEvalError::Ref);
                }
                let start_sheet_id = self
                    .resolve_sheet_name(start_sheet)
                    .ok_or(FormulaEvalError::Ref)?;
                let end_sheet_id = self
                    .resolve_sheet_name(end_sheet)
                    .ok_or(FormulaEvalError::Ref)?;
                return Ok(Some((start_sheet_id, end_sheet_id, cursor + 1)));
            }
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || ch == ':' {
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        Ok(None)
    }

    pub(super) fn try_parse_named_reference(
        &self,
    ) -> Result<Option<(FormulaReference, usize)>, FormulaEvalError> {
        let (qualified_sheet_id, start) = self.try_parse_sheet_qualifier()?;
        let qualified = start != self.index;
        let Some((name, cursor)) = self.parse_identifier_at(start) else {
            return Ok(None);
        };
        let next_is_boundary = self.input[cursor..]
            .chars()
            .next()
            .is_none_or(|ch| !ch.is_ascii_alphanumeric() && ch != '_' && ch != '.');
        if !next_is_boundary {
            return Ok(None);
        }

        let defined_name = if qualified {
            self.evaluator
                .state
                .defined_names
                .lookup_in_scope(NameScope::Worksheet(qualified_sheet_id), name.as_str())
        } else {
            self.evaluator
                .state
                .defined_names
                .lookup(Some(self.sheet_id), name.as_str())
        };
        let Some(defined_name) = defined_name else {
            return Ok(None);
        };
        let Some(reference) = self.defined_name_reference(defined_name) else {
            return Ok(None);
        };
        Ok(Some((reference, cursor)))
    }

    pub(super) fn defined_name_reference(
        &self,
        defined_name: &office_common::DefinedName,
    ) -> Option<FormulaReference> {
        if defined_name.refers_to.is_r1c1 {
            return None;
        }
        let default_sheet_id = match defined_name.scope {
            NameScope::Workbook => self.sheet_id,
            NameScope::Worksheet(sheet_id) => sheet_id,
        };
        parse_formula_reference_text(
            self.name_formula_at_caller(defined_name.refers_to.text.as_str())
                .as_str(),
            default_sheet_id,
            self.evaluator.state,
        )
        .ok()
    }

    /// A defined name's formula with its relative references resolved at the calling cell.
    /// Name formulas are stored relative to A1, so they move by the caller's offset from A1 and
    /// wrap around the grid. Without a calling cell (`Application.Evaluate`) A1 is the caller.
    pub(super) fn name_formula_at_caller(&self, refers_to: &str) -> String {
        let (row, col) = self.current_position.unwrap_or((1, 1));
        shift_formula_a1_references_wrapping(refers_to, i64::from(row) - 1, i64::from(col) - 1)
    }

    pub(super) fn defined_name_value_probe(
        &mut self,
        name: &str,
    ) -> Result<Option<FormulaValueProbe>, FormulaEvalError> {
        let Some((name_id, scope, refers_to_text, is_r1c1)) = self
            .evaluator
            .state
            .defined_names
            .lookup(Some(self.sheet_id), name)
            .map(|defined_name| {
                (
                    defined_name.id,
                    defined_name.scope,
                    defined_name.refers_to.text.clone(),
                    defined_name.refers_to.is_r1c1,
                )
            })
        else {
            return Ok(None);
        };

        if !self.evaluator.resolving_names.insert(name_id) {
            return Err(FormulaEvalError::Calc);
        }
        let refers_to_text = if is_r1c1 {
            refers_to_text
        } else {
            self.name_formula_at_caller(refers_to_text.as_str())
        };
        let result = (|| {
            let default_sheet_id = match scope {
                NameScope::Workbook => self.sheet_id,
                NameScope::Worksheet(sheet_id) => sheet_id,
            };
            if !is_r1c1
                && let Ok(reference) = parse_formula_reference_text(
                    refers_to_text.as_str(),
                    default_sheet_id,
                    self.evaluator.state,
                )
            {
                let (sheet_id, rect) = reference.single_area()?;
                if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
                    return Err(FormulaEvalError::Value);
                }
                let value =
                    self.evaluator
                        .cell_value_or_blank(sheet_id, rect.row_first, rect.col_first)?;
                return Ok(formula_value_probe_from_cell_value(value));
            }

            let formula_text = if is_r1c1 {
                let (base_row, base_col) = self.current_position.unwrap_or((1, 1));
                convert_formula_r1c1_to_a1(refers_to_text.as_str(), base_row, base_col)
            } else {
                refers_to_text
            };
            let value = self.evaluator.evaluate_formula_value_probe_text(
                self.sheet_id,
                formula_text.as_str(),
                self.current_position,
            )?;
            Ok(value)
        })();
        self.evaluator.resolving_names.remove(&name_id);
        result.map(Some)
    }

    pub(super) fn try_parse_sheet_qualifier(&self) -> Result<(SheetId, usize), FormulaEvalError> {
        if self.peek_char() == Some('\'') {
            let mut cursor = self.index + 1;
            let mut sheet_name = String::new();
            while cursor < self.input.len() {
                let ch = self.input[cursor..]
                    .chars()
                    .next()
                    .ok_or(FormulaEvalError::Unsupported)?;
                if ch == '\'' {
                    let next_cursor = cursor + ch.len_utf8();
                    if self.input[next_cursor..].starts_with('\'') {
                        sheet_name.push('\'');
                        cursor = next_cursor + 1;
                        continue;
                    }
                    if !self.input[next_cursor..].starts_with('!') {
                        return Ok((self.sheet_id, self.index));
                    }
                    let sheet_id = self
                        .resolve_sheet_name(sheet_name.as_str())
                        .ok_or(FormulaEvalError::Ref)?;
                    return Ok((sheet_id, next_cursor + 1));
                }
                sheet_name.push(ch);
                cursor += ch.len_utf8();
            }
            return Ok((self.sheet_id, self.index));
        }

        let mut cursor = self.index;
        while let Some(ch) = self.input[cursor..].chars().next() {
            if ch == '!' {
                let sheet_name = &self.input[self.index..cursor];
                if sheet_name.is_empty() {
                    return Err(FormulaEvalError::Ref);
                }
                let sheet_id = self
                    .resolve_sheet_name(sheet_name)
                    .ok_or(FormulaEvalError::Ref)?;
                return Ok((sheet_id, cursor + 1));
            }
            if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' {
                cursor += ch.len_utf8();
            } else {
                break;
            }
        }
        Ok((self.sheet_id, self.index))
    }

    pub(super) fn resolve_sheet_name(&self, name: &str) -> Option<SheetId> {
        self.evaluator
            .state
            .worksheets()
            .iter()
            .find(|worksheet| worksheet.name.eq_ignore_ascii_case(name))
            .map(|worksheet| worksheet.id)
    }
}
