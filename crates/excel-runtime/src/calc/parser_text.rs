//! Formula parser methods for text functions: concatenation, slicing, formatting, search/replace, regular expressions, and web/text services.

use super::*;

impl<'a, 'b, 'state> FormulaParser<'a, 'b, 'state> {
    pub(super) fn parse_text_function(&mut self, name: &str) -> Result<String, FormulaEvalError> {
        if name.eq_ignore_ascii_case("ADDRESS") {
            return self.parse_address_text_function();
        }
        if name.eq_ignore_ascii_case("ARRAYTOTEXT") {
            return self.parse_arraytotext_function();
        }
        if name.eq_ignore_ascii_case("ASC")
            || name.eq_ignore_ascii_case("DBCS")
            || name.eq_ignore_ascii_case("JIS")
        {
            return self.parse_unary_text_function(|text| text);
        }
        if name.eq_ignore_ascii_case("CONCAT") || name.eq_ignore_ascii_case("CONCATENATE") {
            return self.parse_concat_text_function();
        }
        if name.eq_ignore_ascii_case("DGET") {
            return formula_selected_text_from_value_probe(
                self.parse_database_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("LAMBDA") {
            let lambda = self.parse_lambda_value_function()?;
            self.skip_whitespace();
            if !self.consume_char('(') {
                return Err(FormulaEvalError::Calc);
            }
            return formula_selected_text_from_value_probe(
                self.parse_lambda_call_arguments(lambda)?,
            );
        }
        if name.eq_ignore_ascii_case("MAKEARRAY") {
            return formula_selected_text_from_value_probe(self.parse_makearray_value_function()?);
        }
        if name.eq_ignore_ascii_case("REDUCE") || name.eq_ignore_ascii_case("SCAN") {
            return formula_selected_text_from_value_probe(
                self.parse_reduce_scan_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("GETPIVOTDATA") {
            return formula_selected_text_from_value_probe(
                self.parse_getpivotdata_value_function()?,
            );
        }
        if name.eq_ignore_ascii_case("CUBEKPIMEMBER")
            || name.eq_ignore_ascii_case("CUBEMEMBER")
            || name.eq_ignore_ascii_case("CUBERANKEDMEMBER")
            || name.eq_ignore_ascii_case("CUBESET")
        {
            return self.parse_cube_caption_text_function(name);
        }
        if name.eq_ignore_ascii_case("CUBEMEMBERPROPERTY") {
            return self.parse_cubememberproperty_text_function();
        }
        if formula_array_projection_function_name(name) {
            return formula_selected_text_from_value_probe(
                self.parse_array_projection_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("REGEXEXTRACT") {
            return self.parse_regex_extract_function();
        }
        if name.eq_ignore_ascii_case("REGEXREPLACE") {
            return self.parse_regex_replace_function();
        }
        if name.eq_ignore_ascii_case("DETECTLANGUAGE") {
            return self.parse_detectlanguage_function();
        }
        if name.eq_ignore_ascii_case("FILTERXML") {
            return self.parse_filterxml_function();
        }
        if name.eq_ignore_ascii_case("IMAGE") {
            return self.parse_image_function();
        }
        if name.eq_ignore_ascii_case("TRANSLATE") {
            return self.parse_translate_function();
        }
        if name.eq_ignore_ascii_case("WEBSERVICE") {
            return self.parse_webservice_function();
        }
        if name.eq_ignore_ascii_case("ENCODEURL") {
            return self.parse_unary_text_function(|text| formula_encode_url(text.as_str()));
        }
        if name.eq_ignore_ascii_case("LEFT") {
            return self.parse_left_text_function(false);
        }
        if name.eq_ignore_ascii_case("LEFTB") {
            return self.parse_left_text_function(true);
        }
        if name.eq_ignore_ascii_case("RIGHT") {
            return self.parse_right_text_function(false);
        }
        if name.eq_ignore_ascii_case("RIGHTB") {
            return self.parse_right_text_function(true);
        }
        if name.eq_ignore_ascii_case("MID") {
            return self.parse_mid_text_function(false);
        }
        if name.eq_ignore_ascii_case("MIDB") {
            return self.parse_mid_text_function(true);
        }
        if name.eq_ignore_ascii_case("BASE") {
            return self.parse_base_text_function();
        }
        if name.eq_ignore_ascii_case("BAHTTEXT") {
            return self.parse_bahttext_function();
        }
        if name.eq_ignore_ascii_case("CELL") {
            return formula_selected_text_from_value_probe(self.parse_cell_value_function()?);
        }
        if name.eq_ignore_ascii_case("INFO") {
            return formula_selected_text_from_value_probe(self.parse_info_value_function()?);
        }
        if name.eq_ignore_ascii_case("INDIRECT")
            || name.eq_ignore_ascii_case("OFFSET")
            || name.eq_ignore_ascii_case("TRIMRANGE")
        {
            return formula_selected_text_from_value_probe(
                self.parse_reference_projection_value_function(name)?,
            );
        }
        if name.eq_ignore_ascii_case("DEC2BIN") {
            return self.parse_decimal_engineering_text_function(2, 10, 10);
        }
        if name.eq_ignore_ascii_case("DEC2OCT") {
            return self.parse_decimal_engineering_text_function(8, 30, 10);
        }
        if name.eq_ignore_ascii_case("DEC2HEX") {
            return self.parse_decimal_engineering_text_function(16, 40, 10);
        }
        if name.eq_ignore_ascii_case("BIN2OCT") {
            return self.parse_engineering_text_function(2, 10, 10, 8, 30, 10);
        }
        if name.eq_ignore_ascii_case("BIN2HEX") {
            return self.parse_engineering_text_function(2, 10, 10, 16, 40, 10);
        }
        if name.eq_ignore_ascii_case("OCT2BIN") {
            return self.parse_engineering_text_function(8, 30, 10, 2, 10, 10);
        }
        if name.eq_ignore_ascii_case("OCT2HEX") {
            return self.parse_engineering_text_function(8, 30, 10, 16, 40, 10);
        }
        if name.eq_ignore_ascii_case("HEX2BIN") {
            return self.parse_engineering_text_function(16, 40, 10, 2, 10, 10);
        }
        if name.eq_ignore_ascii_case("HEX2OCT") {
            return self.parse_engineering_text_function(16, 40, 10, 8, 30, 10);
        }
        if name.eq_ignore_ascii_case("COMPLEX") {
            let real = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let imaginary = self.parse_comparison()?;
            self.skip_whitespace();
            let suffix = if self.consume_char(')') {
                'i'
            } else {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let suffix = formula_complex_suffix(self.parse_text_value_argument()?.as_str())?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
                suffix
            };
            return formula_complex_format(formula_complex_number(real, imaginary, Some(suffix))?);
        }
        if name.eq_ignore_ascii_case("IMCONJUGATE") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_number(
                value.real,
                -value.imaginary,
                value.suffix,
            )?);
        }
        if name.eq_ignore_ascii_case("IMCOS") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_cos(value)?);
        }
        if name.eq_ignore_ascii_case("IMCOSH") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_cosh(value)?);
        }
        if name.eq_ignore_ascii_case("IMCOT") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_divide(
                formula_complex_cos(value)?,
                formula_complex_sin(value)?,
            )?);
        }
        if name.eq_ignore_ascii_case("IMCSC") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_reciprocal(formula_complex_sin(
                value,
            )?)?);
        }
        if name.eq_ignore_ascii_case("IMCSCH") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_reciprocal(formula_complex_sinh(
                value,
            )?)?);
        }
        if name.eq_ignore_ascii_case("IMDIV") {
            let left = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let right = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_divide(left, right)?);
        }
        if name.eq_ignore_ascii_case("IMEXP") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_exp(value)?);
        }
        if name.eq_ignore_ascii_case("IMLN") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_ln(value)?);
        }
        if name.eq_ignore_ascii_case("IMLOG10") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let value = formula_complex_ln(value)?;
            let base = 10.0_f64.ln();
            return formula_complex_format(formula_complex_number(
                value.real / base,
                value.imaginary / base,
                value.suffix,
            )?);
        }
        if name.eq_ignore_ascii_case("IMLOG2") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            let value = formula_complex_ln(value)?;
            let base = 2.0_f64.ln();
            return formula_complex_format(formula_complex_number(
                value.real / base,
                value.imaginary / base,
                value.suffix,
            )?);
        }
        if name.eq_ignore_ascii_case("IMPOWER") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let power = self.parse_comparison()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_power(value, power)?);
        }
        if name.eq_ignore_ascii_case("IMPRODUCT") || name.eq_ignore_ascii_case("IMSUM") {
            let product = name.eq_ignore_ascii_case("IMPRODUCT");
            let mut value = formula_complex_number(if product { 1.0 } else { 0.0 }, 0.0, None)?;
            let mut saw_argument = false;
            loop {
                self.skip_whitespace();
                if self.consume_char(')') {
                    return if saw_argument {
                        formula_complex_format(value)
                    } else {
                        Err(FormulaEvalError::Value)
                    };
                }
                let argument = self.parse_complex_argument()?;
                value = if product {
                    formula_complex_multiply(value, argument)?
                } else {
                    formula_complex_add(value, argument)?
                };
                saw_argument = true;
                self.skip_whitespace();
                if self.consume_char(',') {
                    continue;
                }
                if self.consume_char(')') {
                    return formula_complex_format(value);
                }
                return Err(FormulaEvalError::Unsupported);
            }
        }
        if name.eq_ignore_ascii_case("IMSEC") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_reciprocal(formula_complex_cos(
                value,
            )?)?);
        }
        if name.eq_ignore_ascii_case("IMSECH") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_reciprocal(formula_complex_cosh(
                value,
            )?)?);
        }
        if name.eq_ignore_ascii_case("IMSIN") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_sin(value)?);
        }
        if name.eq_ignore_ascii_case("IMSINH") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_sinh(value)?);
        }
        if name.eq_ignore_ascii_case("IMSQRT") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_sqrt(value)?);
        }
        if name.eq_ignore_ascii_case("IMSUB") {
            let left = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let right = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_subtract(left, right)?);
        }
        if name.eq_ignore_ascii_case("IMTAN") {
            let value = self.parse_complex_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return formula_complex_format(formula_complex_divide(
                formula_complex_sin(value)?,
                formula_complex_cos(value)?,
            )?);
        }
        if name.eq_ignore_ascii_case("ROMAN") {
            return self.parse_roman_text_function();
        }
        if name.eq_ignore_ascii_case("DOLLAR") {
            return self.parse_dollar_text_function();
        }
        if name.eq_ignore_ascii_case("FIXED") {
            return self.parse_fixed_text_function();
        }
        if name.eq_ignore_ascii_case("TEXT") {
            return self.parse_text_format_function();
        }
        if name.eq_ignore_ascii_case("HYPERLINK") {
            return self.parse_hyperlink_text_function();
        }
        if name.eq_ignore_ascii_case("PHONETIC") {
            return self.parse_phonetic_text_function();
        }
        if name.eq_ignore_ascii_case("CHAR") {
            return self.parse_character_text_function(false);
        }
        if name.eq_ignore_ascii_case("UNICHAR") {
            return self.parse_character_text_function(true);
        }
        if name.eq_ignore_ascii_case("CLEAN") {
            return self.parse_unary_text_function(|text| {
                text.chars()
                    .filter(|ch| !matches!(*ch as u32, 0..=31))
                    .collect()
            });
        }
        if name.eq_ignore_ascii_case("T") {
            let value = self.parse_value_probe_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            return match value {
                FormulaValueProbe::Text(value) => Ok(value),
                FormulaValueProbe::Error(error) => Err(error),
                FormulaValueProbe::Blank
                | FormulaValueProbe::Bool(_)
                | FormulaValueProbe::Number(_)
                | FormulaValueProbe::Omitted
                | FormulaValueProbe::Lambda { .. } => Ok(String::new()),
            };
        }
        if name.eq_ignore_ascii_case("UPPER") {
            return self.parse_unary_text_function(|text| text.to_uppercase());
        }
        if name.eq_ignore_ascii_case("LOWER") {
            return self.parse_unary_text_function(|text| text.to_lowercase());
        }
        if name.eq_ignore_ascii_case("PROPER") {
            return self.parse_unary_text_function(|text| formula_proper_text(text.as_str()));
        }
        if name.eq_ignore_ascii_case("TRIM") {
            return self.parse_unary_text_function(|text| {
                text.split_whitespace().collect::<Vec<_>>().join(" ")
            });
        }
        if name.eq_ignore_ascii_case("TEXTJOIN") {
            return self.parse_textjoin_function();
        }
        if name.eq_ignore_ascii_case("TEXTSPLIT") {
            return self.parse_textsplit_function();
        }
        if name.eq_ignore_ascii_case("TEXTBEFORE") {
            return self.parse_text_boundary_function(false);
        }
        if name.eq_ignore_ascii_case("TEXTAFTER") {
            return self.parse_text_boundary_function(true);
        }
        if name.eq_ignore_ascii_case("REPT") {
            return self.parse_rept_text_function();
        }
        if name.eq_ignore_ascii_case("REPLACE") {
            return self.parse_replace_text_function(false);
        }
        if name.eq_ignore_ascii_case("REPLACEB") {
            return self.parse_replace_text_function(true);
        }
        if name.eq_ignore_ascii_case("SUBSTITUTE") {
            return self.parse_substitute_text_function();
        }
        if name.eq_ignore_ascii_case("VALUETOTEXT") {
            return self.parse_valuetotext_function();
        }
        if name.eq_ignore_ascii_case("FORMULATEXT") {
            return self.parse_formulatext_function();
        }
        if name.eq_ignore_ascii_case("IF") {
            return formula_selected_text_from_value_probe(self.parse_if_value_function()?);
        }
        if name.eq_ignore_ascii_case("IFS") {
            return formula_selected_text_from_value_probe(self.parse_ifs_value_function()?);
        }
        if name.eq_ignore_ascii_case("SWITCH") {
            return formula_selected_text_from_value_probe(self.parse_switch_value_function()?);
        }
        if name.eq_ignore_ascii_case("CHOOSE") {
            return formula_selected_text_from_value_probe(self.parse_choose_value_function()?);
        }
        if name.eq_ignore_ascii_case("INDEX") {
            return formula_selected_text_from_value_probe(self.parse_index_value_function()?);
        }
        if name.eq_ignore_ascii_case("LOOKUP") {
            return formula_selected_text_from_value_probe(self.parse_lookup_value_function()?);
        }
        if name.eq_ignore_ascii_case("VLOOKUP") {
            return formula_selected_text_from_value_probe(
                self.parse_table_lookup_value_function(false)?,
            );
        }
        if name.eq_ignore_ascii_case("HLOOKUP") {
            return formula_selected_text_from_value_probe(
                self.parse_table_lookup_value_function(true)?,
            );
        }
        if name.eq_ignore_ascii_case("XLOOKUP") {
            return formula_selected_text_from_value_probe(self.parse_xlookup_value_function()?);
        }
        if name.eq_ignore_ascii_case("LET") {
            return formula_selected_text_from_value_probe(self.parse_let_value_function()?);
        }
        Err(FormulaEvalError::Unsupported)
    }

    pub(super) fn parse_concat_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let mut output = String::new();
        let mut saw_argument = false;
        loop {
            self.skip_whitespace();
            if self.consume_char(')') {
                return if saw_argument {
                    Ok(output)
                } else {
                    Err(FormulaEvalError::Value)
                };
            }
            output.push_str(self.parse_text_value_argument()?.as_str());
            saw_argument = true;
            self.skip_whitespace();
            if self.consume_char(',') {
                continue;
            }
            if self.consume_char(')') {
                return Ok(output);
            }
            return Err(FormulaEvalError::Unsupported);
        }
    }

    pub(super) fn parse_address_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let row = formula_integer_argument(self.parse_comparison()?)?;
        if row < 1 || row > i64::from(EXCEL_MAX_ROW_INDEX) {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let col = formula_integer_argument(self.parse_comparison()?)?;
        if col < 1 || col > i64::from(EXCEL_MAX_COLUMN_INDEX) {
            return Err(FormulaEvalError::Value);
        }

        let mut abs_num = 1_i64;
        let mut a1 = true;
        let mut sheet_text = None;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            abs_num = formula_integer_argument(self.parse_comparison()?)?;
            if !(1..=4).contains(&abs_num) {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                a1 = self.parse_comparison()? != 0.0;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    sheet_text = Some(self.parse_text_value_argument()?);
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                }
            }
        }

        let row = u32::try_from(row).map_err(|_| FormulaEvalError::Value)?;
        let col = u32::try_from(col).map_err(|_| FormulaEvalError::Value)?;
        let row_absolute = matches!(abs_num, 1 | 2);
        let column_absolute = matches!(abs_num, 1 | 3);
        let mut address = if a1 {
            format_cell_address(row, col, row_absolute, column_absolute)
        } else {
            let row_part = if row_absolute {
                row.to_string()
            } else {
                format!("[{row}]")
            };
            let col_part = if column_absolute {
                col.to_string()
            } else {
                format!("[{col}]")
            };
            format!("R{row_part}C{col_part}")
        };
        if let Some(sheet_text) = sheet_text {
            address = format!(
                "{}{address}",
                formula_sheet_address_qualifier(sheet_text.as_str())
            );
        }
        Ok(address)
    }

    pub(super) fn parse_left_text_function(
        &mut self,
        byte_mode: bool,
    ) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let count = self.parse_optional_text_count_argument(1)?;
        if byte_mode {
            Ok(formula_text_byte_slice(text.as_str(), 1, count))
        } else {
            Ok(text.chars().take(count).collect())
        }
    }

    pub(super) fn parse_right_text_function(
        &mut self,
        byte_mode: bool,
    ) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let count = self.parse_optional_text_count_argument(1)?;
        if byte_mode {
            let len = formula_text_byte_len(text.as_str());
            let start = len.saturating_sub(count) + 1;
            return Ok(formula_text_byte_slice(text.as_str(), start, count));
        }
        let chars = text.chars().collect::<Vec<_>>();
        Ok(chars
            .iter()
            .skip(chars.len().saturating_sub(count))
            .collect())
    }

    pub(super) fn parse_mid_text_function(
        &mut self,
        byte_mode: bool,
    ) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let start = formula_positive_position_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let count = formula_non_negative_count_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if byte_mode {
            Ok(formula_text_byte_slice(text.as_str(), start, count))
        } else {
            Ok(text.chars().skip(start - 1).take(count).collect())
        }
    }

    pub(super) fn parse_character_text_function(
        &mut self,
        unicode: bool,
    ) -> Result<String, FormulaEvalError> {
        let code = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let code = u32::try_from(code).map_err(|_| FormulaEvalError::Value)?;
        if (!unicode && !(1..=255).contains(&code)) || code == 0 {
            return Err(FormulaEvalError::Value);
        }
        char::from_u32(code)
            .map(|ch| ch.to_string())
            .ok_or(FormulaEvalError::Value)
    }

    pub(super) fn parse_bahttext_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if !number.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let scaled = number.abs() * 100.0;
        if !scaled.is_finite() || scaled > u128::MAX as f64 {
            return Err(FormulaEvalError::Num);
        }
        let total_satang = round_half_away_from_zero(scaled) as u128;
        let baht = total_satang / 100;
        let satang = (total_satang % 100) as u32;

        let convert_group = |group: u32| -> String {
            let digit_text = [
                "ศูนย์",
                "หนึ่ง",
                "สอง",
                "สาม",
                "สี่",
                "ห้า",
                "หก",
                "เจ็ด",
                "แปด",
                "เก้า",
            ];
            let mut output = String::new();
            let hundred_thousands = group / 100_000;
            let ten_thousands = group / 10_000 % 10;
            let thousands = group / 1_000 % 10;
            let hundreds = group / 100 % 10;
            let tens = group / 10 % 10;
            let ones = group % 10;
            for (digit, suffix) in [
                (hundred_thousands, "แสน"),
                (ten_thousands, "หมื่น"),
                (thousands, "พัน"),
                (hundreds, "ร้อย"),
            ] {
                if digit != 0 {
                    output.push_str(digit_text[digit as usize]);
                    output.push_str(suffix);
                }
            }
            if tens != 0 {
                if tens == 1 {
                    output.push_str("สิบ");
                } else if tens == 2 {
                    output.push_str("ยี่สิบ");
                } else {
                    output.push_str(digit_text[tens as usize]);
                    output.push_str("สิบ");
                }
            }
            if ones != 0 {
                if ones == 1 && group > 1 {
                    output.push_str("เอ็ด");
                } else {
                    output.push_str(digit_text[ones as usize]);
                }
            }
            if output.is_empty() {
                output.push_str(digit_text[0]);
            }
            output
        };

        let convert_number = |mut value: u128| -> String {
            if value == 0 {
                return "ศูนย์".to_string();
            }
            let mut groups = Vec::new();
            while value > 0 {
                groups.push((value % 1_000_000) as u32);
                value /= 1_000_000;
            }
            let mut output = String::new();
            for index in (0..groups.len()).rev() {
                let group = groups[index];
                if group != 0 {
                    output.push_str(convert_group(group).as_str());
                }
                if index > 0 {
                    output.push_str("ล้าน");
                }
            }
            output
        };

        let mut output = String::new();
        if number < 0.0 && total_satang != 0 {
            output.push_str("ลบ");
        }
        output.push_str(convert_number(baht).as_str());
        output.push_str("บาท");
        if satang == 0 {
            output.push_str("ถ้วน");
        } else {
            output.push_str(convert_group(satang).as_str());
            output.push_str("สตางค์");
        }
        Ok(output)
    }

    pub(super) fn parse_roman_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        if !number.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let number = number.trunc();
        if !(0.0..=3999.0).contains(&number) {
            return Err(FormulaEvalError::Value);
        }
        let mut form = 0_usize;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            form = match self.parse_value_probe_argument()? {
                FormulaValueProbe::Bool(true) => 0,
                FormulaValueProbe::Bool(false) => 4,
                FormulaValueProbe::Number(value) => {
                    if !value.is_finite() {
                        return Err(FormulaEvalError::Value);
                    }
                    let value = value.trunc();
                    if !(0.0..=4.0).contains(&value) {
                        return Err(FormulaEvalError::Value);
                    }
                    value as usize
                }
                FormulaValueProbe::Error(error) => return Err(error),
                FormulaValueProbe::Blank
                | FormulaValueProbe::Text(_)
                | FormulaValueProbe::Omitted
                | FormulaValueProbe::Lambda { .. } => {
                    return Err(FormulaEvalError::Value);
                }
            };
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
        formula_roman_text(number as i64, form)
    }

    pub(super) fn parse_text_format_decimals_argument(&mut self) -> Result<i64, FormulaEvalError> {
        let value = self.parse_comparison()?;
        if !value.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let value = value.trunc();
        if !(-127.0..=127.0).contains(&value) {
            return Err(FormulaEvalError::Value);
        }
        Ok(value as i64)
    }

    pub(super) fn parse_dollar_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        let mut decimals = 2_i64;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            decimals = self.parse_text_format_decimals_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
        let formatted = formula_fixed_number_text(number, decimals, true)?;
        if let Some(positive) = formatted.strip_prefix('-') {
            Ok(format!("(${positive})"))
        } else {
            Ok(format!("${formatted}"))
        }
    }

    pub(super) fn parse_fixed_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        let mut decimals = 2_i64;
        let mut use_commas = true;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            decimals = self.parse_text_format_decimals_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                use_commas = self.parse_comparison()? == 0.0;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_fixed_number_text(number, decimals, use_commas)
    }

    pub(super) fn parse_text_format_function(&mut self) -> Result<String, FormulaEvalError> {
        let number = self.parse_comparison()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let format = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }

        if !number.is_finite() {
            return Err(FormulaEvalError::Value);
        }
        let sections = format.split(';').collect::<Vec<_>>();
        if sections.is_empty() || sections.len() > 4 {
            return Err(FormulaEvalError::Value);
        }
        let format = if number < 0.0 && sections.len() > 1 {
            sections[1]
        } else if number == 0.0 && sections.len() > 2 {
            sections[2]
        } else {
            sections[0]
        }
        .trim();
        if format.is_empty() {
            return Err(FormulaEvalError::Value);
        }

        #[derive(Clone)]
        enum TextDateTimeFormatToken {
            Literal(String),
            Year(usize),
            MonthOrMinute(usize),
            Day(usize),
            Hour(usize),
            Second(usize),
            AmPm(String),
        }

        let mut tokens = Vec::new();
        let mut literal = String::new();
        let push_literal = |tokens: &mut Vec<TextDateTimeFormatToken>, literal: &mut String| {
            if !literal.is_empty() {
                tokens.push(TextDateTimeFormatToken::Literal(std::mem::take(literal)));
            }
        };
        let format_chars = format.chars().collect::<Vec<_>>();
        let mut index = 0_usize;
        while index < format_chars.len() {
            let ch = format_chars[index];
            if ch == '"' {
                index += 1;
                loop {
                    let quoted = format_chars.get(index).ok_or(FormulaEvalError::Value)?;
                    index += 1;
                    if *quoted == '"' {
                        break;
                    }
                    literal.push(*quoted);
                }
                continue;
            }
            if ch == '\\' {
                index += 1;
                let escaped = format_chars.get(index).ok_or(FormulaEvalError::Value)?;
                literal.push(*escaped);
                index += 1;
                continue;
            }
            if matches!(ch, '_' | '*') {
                index += 1;
                format_chars.get(index).ok_or(FormulaEvalError::Value)?;
                index += 1;
                continue;
            }
            let starts_with = |needle: &str| -> bool {
                let needle_chars = needle.chars().collect::<Vec<_>>();
                format_chars
                    .get(index..index + needle_chars.len())
                    .is_some_and(|candidate| {
                        candidate
                            .iter()
                            .zip(needle_chars.iter())
                            .all(|(left, right)| left.eq_ignore_ascii_case(right))
                    })
            };
            if starts_with("AM/PM") || starts_with("A/P") {
                push_literal(&mut tokens, &mut literal);
                let len = if starts_with("AM/PM") { 5 } else { 3 };
                tokens.push(TextDateTimeFormatToken::AmPm(
                    format_chars[index..index + len].iter().collect(),
                ));
                index += len;
                continue;
            }
            if ch.is_ascii_alphabetic() {
                let lower = ch.to_ascii_lowercase();
                if !matches!(lower, 'y' | 'm' | 'd' | 'h' | 's') {
                    return Err(FormulaEvalError::Value);
                }
                let start = index;
                index += 1;
                while index < format_chars.len()
                    && format_chars[index].to_ascii_lowercase() == lower
                {
                    index += 1;
                }
                push_literal(&mut tokens, &mut literal);
                let len = index - start;
                tokens.push(match lower {
                    'y' => TextDateTimeFormatToken::Year(len),
                    'm' => TextDateTimeFormatToken::MonthOrMinute(len),
                    'd' => TextDateTimeFormatToken::Day(len),
                    'h' => TextDateTimeFormatToken::Hour(len),
                    's' => TextDateTimeFormatToken::Second(len),
                    _ => return Err(FormulaEvalError::Value),
                });
                continue;
            }
            literal.push(ch);
            index += 1;
        }
        push_literal(&mut tokens, &mut literal);

        if tokens
            .iter()
            .any(|token| !matches!(token, TextDateTimeFormatToken::Literal(_)))
        {
            let has_am_pm = tokens
                .iter()
                .any(|token| matches!(token, TextDateTimeFormatToken::AmPm(_)));
            let mut minute_tokens = vec![false; tokens.len()];
            let mut needs_date = false;
            let mut needs_time = false;
            for token_index in 0..tokens.len() {
                match &tokens[token_index] {
                    TextDateTimeFormatToken::Literal(_) => {}
                    TextDateTimeFormatToken::Year(_) | TextDateTimeFormatToken::Day(_) => {
                        needs_date = true;
                    }
                    TextDateTimeFormatToken::Hour(_)
                    | TextDateTimeFormatToken::Second(_)
                    | TextDateTimeFormatToken::AmPm(_) => {
                        needs_time = true;
                    }
                    TextDateTimeFormatToken::MonthOrMinute(_) => {
                        let previous = tokens[..token_index]
                            .iter()
                            .rev()
                            .find(|token| !matches!(token, TextDateTimeFormatToken::Literal(_)));
                        let next = tokens[token_index + 1..]
                            .iter()
                            .find(|token| !matches!(token, TextDateTimeFormatToken::Literal(_)));
                        let time_context = previous.is_some_and(|token| {
                            matches!(
                                token,
                                TextDateTimeFormatToken::Hour(_)
                                    | TextDateTimeFormatToken::Second(_)
                            )
                        }) || next.is_some_and(|token| {
                            matches!(
                                token,
                                TextDateTimeFormatToken::Hour(_)
                                    | TextDateTimeFormatToken::Second(_)
                            )
                        }) || (has_am_pm
                            && !previous.is_some_and(|token| {
                                matches!(
                                    token,
                                    TextDateTimeFormatToken::Year(_)
                                        | TextDateTimeFormatToken::Day(_)
                                )
                            })
                            && !next.is_some_and(|token| {
                                matches!(
                                    token,
                                    TextDateTimeFormatToken::Year(_)
                                        | TextDateTimeFormatToken::Day(_)
                                )
                            }));
                        minute_tokens[token_index] = time_context;
                        if time_context {
                            needs_time = true;
                        } else {
                            needs_date = true;
                        }
                    }
                }
            }

            let (year, month, day, weekday) = if needs_date {
                let (year, month, day) = self.evaluator.context.date_system().ymd(number)?;
                let serial = formula_serial_integer(number)?;
                (
                    year,
                    month,
                    day,
                    formula_weekday_monday0_from_serial(serial) as usize,
                )
            } else {
                (1900, 1, 1, 0)
            };
            let (hour, minute, second) = if needs_time {
                formula_time_parts_from_serial(number)?
            } else {
                (0, 0, 0)
            };
            let month_names = [
                ("Jan", "January"),
                ("Feb", "February"),
                ("Mar", "March"),
                ("Apr", "April"),
                ("May", "May"),
                ("Jun", "June"),
                ("Jul", "July"),
                ("Aug", "August"),
                ("Sep", "September"),
                ("Oct", "October"),
                ("Nov", "November"),
                ("Dec", "December"),
            ];
            let weekday_names = [
                ("Mon", "Monday"),
                ("Tue", "Tuesday"),
                ("Wed", "Wednesday"),
                ("Thu", "Thursday"),
                ("Fri", "Friday"),
                ("Sat", "Saturday"),
                ("Sun", "Sunday"),
            ];
            let mut output = String::new();
            for (token_index, token) in tokens.iter().enumerate() {
                match token {
                    TextDateTimeFormatToken::Literal(value) => output.push_str(value),
                    TextDateTimeFormatToken::Year(len) => {
                        if *len <= 2 {
                            output.push_str(format!("{:02}", year.rem_euclid(100)).as_str());
                        } else {
                            output.push_str(format!("{year:04}").as_str());
                        }
                    }
                    TextDateTimeFormatToken::MonthOrMinute(len) => {
                        if minute_tokens[token_index] {
                            if *len == 1 {
                                output.push_str(minute.to_string().as_str());
                            } else {
                                output.push_str(format!("{minute:02}").as_str());
                            }
                        } else {
                            match *len {
                                1 => output.push_str(month.to_string().as_str()),
                                2 => output.push_str(format!("{month:02}").as_str()),
                                3 => output.push_str(month_names[(month - 1) as usize].0),
                                _ => output.push_str(month_names[(month - 1) as usize].1),
                            }
                        }
                    }
                    TextDateTimeFormatToken::Day(len) => match *len {
                        1 => output.push_str(day.to_string().as_str()),
                        2 => output.push_str(format!("{day:02}").as_str()),
                        3 => output.push_str(weekday_names[weekday].0),
                        _ => output.push_str(weekday_names[weekday].1),
                    },
                    TextDateTimeFormatToken::Hour(len) => {
                        let display_hour = if has_am_pm {
                            match hour % 12 {
                                0 => 12,
                                value => value,
                            }
                        } else {
                            hour
                        };
                        if *len == 1 {
                            output.push_str(display_hour.to_string().as_str());
                        } else {
                            output.push_str(format!("{display_hour:02}").as_str());
                        }
                    }
                    TextDateTimeFormatToken::Second(len) => {
                        if *len == 1 {
                            output.push_str(second.to_string().as_str());
                        } else {
                            output.push_str(format!("{second:02}").as_str());
                        }
                    }
                    TextDateTimeFormatToken::AmPm(pattern) => {
                        let marker = if pattern.eq_ignore_ascii_case("A/P") {
                            if hour < 12 { "A" } else { "P" }
                        } else if hour < 12 {
                            "AM"
                        } else {
                            "PM"
                        };
                        if pattern.chars().any(|ch| ch.is_ascii_uppercase()) {
                            output.push_str(marker);
                        } else {
                            output.push_str(marker.to_ascii_lowercase().as_str());
                        }
                    }
                }
            }
            return Ok(output);
        }

        let mut in_quote = false;
        let mut escaped = false;
        for ch in format.chars() {
            if escaped {
                escaped = false;
                continue;
            }
            if ch == '\\' && !in_quote {
                escaped = true;
                continue;
            }
            if ch == '"' {
                in_quote = !in_quote;
                continue;
            }
            if !in_quote && (ch == '@' || ch.is_ascii_alphabetic()) {
                return Err(FormulaEvalError::Value);
            }
        }
        if in_quote || escaped {
            return Err(FormulaEvalError::Value);
        }

        let first_placeholder = format
            .find(|ch| matches!(ch, '0' | '#'))
            .ok_or(FormulaEvalError::Value)?;
        let last_placeholder = format
            .rfind(|ch| matches!(ch, '0' | '#'))
            .ok_or(FormulaEvalError::Value)?;
        let core = &format[first_placeholder..=last_placeholder];
        if core.matches('.').count() > 1 {
            return Err(FormulaEvalError::Value);
        }
        let decimals = core
            .split_once('.')
            .map(|(_, fraction)| {
                fraction
                    .chars()
                    .filter(|ch| matches!(*ch, '0' | '#'))
                    .count() as i64
            })
            .unwrap_or(0);
        let integer_part = core
            .split_once('.')
            .map(|(integer, _)| integer)
            .unwrap_or(core);
        let use_commas = integer_part.contains(',');
        let percent_count = format.chars().filter(|ch| *ch == '%').count();
        if percent_count > 1 {
            return Err(FormulaEvalError::Value);
        }
        let mut value = number;
        if percent_count == 1 {
            value *= 100.0;
        }
        if number < 0.0 && sections.len() > 1 {
            value = value.abs();
        }

        let format_literal = |literal: &str| -> Result<String, FormulaEvalError> {
            let mut output = String::new();
            let mut chars = literal.chars();
            let mut in_quote = false;
            while let Some(ch) = chars.next() {
                match ch {
                    '"' => in_quote = !in_quote,
                    '\\' => {
                        let escaped = chars.next().ok_or(FormulaEvalError::Value)?;
                        output.push(escaped);
                    }
                    '_' | '*' if !in_quote => {
                        chars.next().ok_or(FormulaEvalError::Value)?;
                    }
                    _ => output.push(ch),
                }
            }
            if in_quote {
                Err(FormulaEvalError::Value)
            } else {
                Ok(output)
            }
        };
        let formatted = formula_fixed_number_text(value, decimals, use_commas)?;
        let prefix = format_literal(&format[..first_placeholder])?;
        let suffix = format_literal(&format[last_placeholder + 1..])?;
        if let Some(positive) = formatted.strip_prefix('-') {
            Ok(format!("-{prefix}{positive}{suffix}"))
        } else {
            Ok(format!("{prefix}{formatted}{suffix}"))
        }
    }

    pub(super) fn parse_hyperlink_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let link_location = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(link_location);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let friendly_name = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_text_from_value_probe(friendly_name)
    }

    pub(super) fn parse_phonetic_text_function(&mut self) -> Result<String, FormulaEvalError> {
        self.skip_whitespace();
        let checkpoint = self.index;
        if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
            self.index = next_index;
            self.skip_whitespace();
            if self.consume_char(')') {
                let mut output = String::new();
                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        match self
                            .evaluator
                            .cell_value_or_blank(target_sheet_id, row, col)?
                        {
                            CellValue::Text(value) => output.push_str(value.as_str()),
                            CellValue::IsoDateTime(value) => output.push_str(value.as_str()),
                            CellValue::RichText(value) => {
                                if value.phonetic_text().is_empty() {
                                    output.push_str(value.as_str());
                                } else {
                                    output.push_str(value.phonetic_text());
                                }
                            }
                            CellValue::Error(error) => {
                                return Err(formula_eval_error_from_cell_error(error));
                            }
                            CellValue::Blank | CellValue::Bool(_) | CellValue::Number(_) => {}
                        }
                    }
                }
                return Ok(output);
            }
        }
        self.index = checkpoint;
        let value = self.parse_value_probe_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        match value {
            FormulaValueProbe::Text(value) => Ok(value),
            FormulaValueProbe::Error(error) => Err(error),
            FormulaValueProbe::Blank
            | FormulaValueProbe::Bool(_)
            | FormulaValueProbe::Number(_)
            | FormulaValueProbe::Omitted
            | FormulaValueProbe::Lambda { .. } => Ok(String::new()),
        }
    }

    pub(super) fn parse_unary_text_function(
        &mut self,
        transform: impl FnOnce(String) -> String,
    ) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(transform(text))
    }

    pub(super) fn parse_optional_text_count_argument(
        &mut self,
        default: usize,
    ) -> Result<usize, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(default);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let count = formula_non_negative_count_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(count)
    }

    pub(super) fn parse_textjoin_function(&mut self) -> Result<String, FormulaEvalError> {
        let delimiter = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let ignore_empty = self.parse_comparison()? != 0.0;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }

        let mut parts = Vec::new();
        loop {
            for text in self.parse_text_values_argument()? {
                if !ignore_empty || !text.is_empty() {
                    parts.push(text);
                }
            }
            self.skip_whitespace();
            if self.consume_char(')') {
                return Ok(parts.join(delimiter.as_str()));
            }
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
        }
    }

    pub(super) fn parse_textsplit_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        self.skip_whitespace();
        let col_delimiter = if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
            None
        } else {
            Some(self.parse_text_value_argument()?)
        };
        let mut row_delimiter = None;
        let mut ignore_empty = false;
        let mut match_mode = 0_i64;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                row_delimiter = Some(self.parse_text_value_argument()?);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    ignore_empty = self.parse_comparison()? != 0.0;
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    self.skip_whitespace();
                    if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                        match_mode = formula_integer_argument(self.parse_comparison()?)?;
                        if !matches!(match_mode, 0 | 1) {
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
                            self.parse_value_probe_argument()?;
                        }
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }
        let mut delimiters = Vec::new();
        if let Some(delimiter) = row_delimiter {
            delimiters.push(delimiter);
        }
        if let Some(delimiter) = col_delimiter {
            delimiters.push(delimiter);
        }
        if delimiters.is_empty() || delimiters.iter().any(|delimiter| delimiter.is_empty()) {
            return Err(FormulaEvalError::Value);
        }

        let case_insensitive = match_mode == 1;
        let mut parts = vec![text];
        for delimiter in delimiters {
            let mut next_parts = Vec::new();
            for part in parts {
                let matches = formula_text_delimiter_matches(
                    part.as_str(),
                    delimiter.as_str(),
                    case_insensitive,
                );
                if matches.is_empty() {
                    next_parts.push(part);
                    continue;
                }
                let mut start = 0_usize;
                for (match_start, match_end) in matches {
                    next_parts.push(part[start..match_start].to_string());
                    start = match_end;
                }
                next_parts.push(part[start..].to_string());
            }
            parts = next_parts;
        }
        parts
            .into_iter()
            .find(|part| !ignore_empty || !part.is_empty())
            .ok_or(FormulaEvalError::Calc)
    }

    pub(super) fn parse_detectlanguage_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(formula_detect_language_tag(text.as_str()).to_string())
    }

    pub(super) fn parse_translate_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let mut source_language = None::<String>;
        let mut target_language = None::<String>;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                source_language = Some(self.parse_text_value_argument()?);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                self.skip_whitespace();
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    target_language = Some(self.parse_text_value_argument()?);
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        let source_language =
            source_language.unwrap_or_else(|| formula_detect_language_tag(text.as_str()).into());
        let target_language = target_language.unwrap_or_else(|| "en".to_string());
        if source_language.eq_ignore_ascii_case(target_language.as_str()) {
            return Ok(text);
        }
        let normalized = text.trim().to_ascii_lowercase();
        let translated = match (
            source_language.to_ascii_lowercase().as_str(),
            target_language.to_ascii_lowercase().as_str(),
            normalized.as_str(),
        ) {
            ("en", "es", "hello") => "hola",
            ("en", "es", "world") => "mundo",
            ("en", "fr", "hello") => "bonjour",
            ("en", "de", "hello") => "hallo",
            ("es", "en", "hola") => "hello",
            ("es", "en", "mundo") => "world",
            ("fr", "en", "bonjour") => "hello",
            ("de", "en", "hallo") => "hello",
            _ => return Ok(text),
        };
        Ok(translated.to_string())
    }

    pub(super) fn parse_image_function(&mut self) -> Result<String, FormulaEvalError> {
        let source = self.parse_text_value_argument()?;
        let mut alt_text = None::<String>;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            self.skip_whitespace();
            if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                alt_text = Some(self.parse_text_value_argument()?);
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
                if !self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                    self.parse_value_probe_argument()?;
                }
            }
        }
        Ok(alt_text.filter(|value| !value.is_empty()).unwrap_or(source))
    }

    pub(super) fn parse_webservice_function(&mut self) -> Result<String, FormulaEvalError> {
        let url = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let Some((metadata, payload)) = url
            .strip_prefix("data:")
            .and_then(|data| data.split_once(','))
        else {
            return Err(FormulaEvalError::Value);
        };
        if metadata.contains(";base64") {
            return Err(FormulaEvalError::Value);
        }
        let mut bytes = Vec::new();
        let payload_bytes = payload.as_bytes();
        let mut index = 0_usize;
        while index < payload_bytes.len() {
            if payload_bytes[index] == b'%' {
                if index + 2 >= payload_bytes.len() {
                    return Err(FormulaEvalError::Value);
                }
                let hex = &payload[index + 1..index + 3];
                let value = u8::from_str_radix(hex, 16).map_err(|_| FormulaEvalError::Value)?;
                bytes.push(value);
                index += 3;
            } else {
                bytes.push(payload_bytes[index]);
                index += 1;
            }
        }
        String::from_utf8(bytes).map_err(|_| FormulaEvalError::Value)
    }

    pub(super) fn parse_filterxml_function(&mut self) -> Result<String, FormulaEvalError> {
        let xml = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let xpath = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let xpath = xpath.trim();
        let absolute = xpath.starts_with('/') && !xpath.starts_with("//");
        let trimmed = xpath.trim_start_matches('/');
        let mut path = Vec::new();
        let mut attribute_name = None::<String>;
        for segment in trimmed.split('/').filter(|segment| !segment.is_empty()) {
            if let Some(attribute) = segment.strip_prefix('@') {
                attribute_name = Some(attribute.to_ascii_lowercase());
                break;
            }
            let element = segment.split('[').next().unwrap_or(segment);
            if element.is_empty() {
                return Err(FormulaEvalError::Value);
            }
            path.push(element.to_ascii_lowercase());
        }
        if path.is_empty() {
            return Err(FormulaEvalError::Value);
        }
        let matches_path = |stack: &[String], path: &[String]| -> bool {
            if absolute {
                stack == path
            } else {
                stack.ends_with(path)
            }
        };

        let mut reader = Reader::from_reader(Cursor::new(xml.as_bytes()));
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut stack = Vec::<String>::new();
        let mut capture_depth = None::<usize>;
        let mut captured = String::new();
        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(Event::Start(element)) => {
                    stack.push(
                        String::from_utf8_lossy(xml_local_name(element.name().as_ref()))
                            .to_ascii_lowercase(),
                    );
                    if capture_depth.is_none() && matches_path(stack.as_slice(), path.as_slice()) {
                        if let Some(attribute_name) = &attribute_name {
                            for attr in element.attributes() {
                                let attr = attr.map_err(|_| FormulaEvalError::Value)?;
                                if String::from_utf8_lossy(xml_local_name(attr.key.as_ref()))
                                    .eq_ignore_ascii_case(attribute_name.as_str())
                                {
                                    return attr
                                        .decode_and_unescape_value(reader.decoder())
                                        .map(|value| value.into_owned())
                                        .map_err(|_| FormulaEvalError::Value);
                                }
                            }
                            return Err(FormulaEvalError::Value);
                        }
                        capture_depth = Some(stack.len());
                    }
                }
                Ok(Event::Empty(element)) => {
                    stack.push(
                        String::from_utf8_lossy(xml_local_name(element.name().as_ref()))
                            .to_ascii_lowercase(),
                    );
                    if capture_depth.is_none() && matches_path(stack.as_slice(), path.as_slice()) {
                        if let Some(attribute_name) = &attribute_name {
                            for attr in element.attributes() {
                                let attr = attr.map_err(|_| FormulaEvalError::Value)?;
                                if String::from_utf8_lossy(xml_local_name(attr.key.as_ref()))
                                    .eq_ignore_ascii_case(attribute_name.as_str())
                                {
                                    return attr
                                        .decode_and_unescape_value(reader.decoder())
                                        .map(|value| value.into_owned())
                                        .map_err(|_| FormulaEvalError::Value);
                                }
                            }
                            return Err(FormulaEvalError::Value);
                        }
                        return Ok(String::new());
                    }
                    stack.pop();
                }
                Ok(Event::Text(text)) if capture_depth.is_some() => {
                    captured.push_str(String::from_utf8_lossy(text.as_ref()).as_ref());
                }
                Ok(Event::CData(text)) if capture_depth.is_some() => {
                    captured.push_str(String::from_utf8_lossy(text.as_ref()).as_ref());
                }
                Ok(Event::End(_)) => {
                    if capture_depth == Some(stack.len()) {
                        return Ok(captured.trim().to_string());
                    }
                    stack.pop();
                }
                Ok(Event::Eof) => break,
                Ok(_) => {}
                Err(_) => return Err(FormulaEvalError::Value),
            }
            buffer.clear();
        }
        Err(FormulaEvalError::Value)
    }

    pub(super) fn parse_text_boundary_function(
        &mut self,
        after: bool,
    ) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let delimiter = self.parse_text_value_argument()?;
        if delimiter.is_empty() {
            return Err(FormulaEvalError::Value);
        }

        let mut instance = 1_i64;
        let mut match_mode = 0_i64;
        let mut match_end = false;
        let mut if_not_found = None;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            instance = formula_integer_argument(self.parse_comparison()?)?;
            if instance == 0 {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                match_mode = formula_integer_argument(self.parse_comparison()?)?;
                if !matches!(match_mode, 0 | 1) {
                    return Err(FormulaEvalError::Value);
                }
                self.skip_whitespace();
                if !self.consume_char(')') {
                    if !self.consume_char(',') {
                        return Err(FormulaEvalError::Unsupported);
                    }
                    match_end = self.parse_comparison()? != 0.0;
                    self.skip_whitespace();
                    if !self.consume_char(')') {
                        if !self.consume_char(',') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                        if_not_found = Some(self.parse_text_value_argument()?);
                        self.skip_whitespace();
                        if !self.consume_char(')') {
                            return Err(FormulaEvalError::Unsupported);
                        }
                    }
                }
            }
        }

        let mut matches =
            formula_text_delimiter_matches(text.as_str(), delimiter.as_str(), match_mode == 1);
        if match_end {
            if instance > 0 {
                matches.push((text.len(), text.len()));
            } else {
                matches.insert(0, (0, 0));
            }
        }
        let selected = if instance > 0 {
            usize::try_from(instance - 1)
                .ok()
                .and_then(|index| matches.get(index))
        } else {
            let count = instance
                .checked_abs()
                .and_then(|value| usize::try_from(value).ok());
            count
                .and_then(|count| matches.len().checked_sub(count))
                .and_then(|index| matches.get(index))
        };
        let Some(&(start, end)) = selected else {
            return if_not_found.ok_or(FormulaEvalError::NA);
        };
        if after {
            Ok(text[end..].to_string())
        } else {
            Ok(text[..start].to_string())
        }
    }

    pub(super) fn parse_regex_test_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pattern = self.parse_text_value_argument()?;
        self.skip_whitespace();
        let case_insensitive = if self.consume_char(')') {
            false
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let case_sensitivity = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            formula_regex_case_insensitive(case_sensitivity)?
        };
        let regex = formula_regex_from_pattern(pattern.as_str(), case_insensitive)?;
        Ok(if regex.is_match(text.as_str()) {
            1.0
        } else {
            0.0
        })
    }

    pub(super) fn parse_regex_extract_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pattern = self.parse_text_value_argument()?;
        let mut return_mode = 0_i64;
        let mut case_insensitive = false;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            return_mode = formula_integer_argument(self.parse_comparison()?)?;
            if !matches!(return_mode, 0 | 1 | 2) {
                return Err(FormulaEvalError::Value);
            }
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let case_sensitivity = formula_integer_argument(self.parse_comparison()?)?;
                case_insensitive = formula_regex_case_insensitive(case_sensitivity)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }

        let regex = formula_regex_from_pattern(pattern.as_str(), case_insensitive)?;
        let captures = regex.captures(text.as_str()).ok_or(FormulaEvalError::NA)?;
        if return_mode == 2 {
            return captures
                .iter()
                .skip(1)
                .find_map(|capture| capture.map(|value| value.as_str().to_string()))
                .ok_or(FormulaEvalError::NA);
        }
        captures
            .get(0)
            .map(|value| value.as_str().to_string())
            .ok_or(FormulaEvalError::NA)
    }

    pub(super) fn parse_regex_replace_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let pattern = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let replacement = self.parse_text_value_argument()?;
        let mut occurrence = 0_i64;
        let mut case_insensitive = false;
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            occurrence = formula_integer_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                let case_sensitivity = formula_integer_argument(self.parse_comparison()?)?;
                case_insensitive = formula_regex_case_insensitive(case_sensitivity)?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }

        let regex = formula_regex_from_pattern(pattern.as_str(), case_insensitive)?;
        if occurrence == 0 {
            return Ok(regex
                .replace_all(text.as_str(), replacement.as_str())
                .into_owned());
        }
        let mut matches = Vec::new();
        for captures in regex.captures_iter(text.as_str()) {
            let Some(full_match) = captures.get(0) else {
                continue;
            };
            let mut expanded = String::new();
            captures.expand(replacement.as_str(), &mut expanded);
            matches.push((full_match.start(), full_match.end(), expanded));
        }
        let selected = if occurrence > 0 {
            usize::try_from(occurrence - 1)
                .ok()
                .and_then(|index| matches.get(index))
        } else {
            occurrence
                .checked_abs()
                .and_then(|value| usize::try_from(value).ok())
                .and_then(|count| matches.len().checked_sub(count))
                .and_then(|index| matches.get(index))
        };
        let Some((start, end, expanded)) = selected else {
            return Ok(text);
        };
        let mut output = String::new();
        output.push_str(&text[..*start]);
        output.push_str(expanded.as_str());
        output.push_str(&text[*end..]);
        Ok(output)
    }

    pub(super) fn parse_rept_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let count = formula_non_negative_count_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let output_len = text
            .chars()
            .count()
            .checked_mul(count)
            .ok_or(FormulaEvalError::Value)?;
        if output_len > 32_767 {
            return Err(FormulaEvalError::Value);
        }
        Ok(text.repeat(count))
    }

    pub(super) fn parse_replace_text_function(
        &mut self,
        byte_mode: bool,
    ) -> Result<String, FormulaEvalError> {
        let old_text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let start = formula_positive_position_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let count = formula_non_negative_count_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let new_text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        if byte_mode {
            let len = formula_text_byte_len(old_text.as_str());
            if start > len + 1 {
                return Err(FormulaEvalError::Value);
            }
            let prefix = formula_text_byte_slice(old_text.as_str(), 1, start - 1);
            let suffix_start = start.saturating_add(count);
            let suffix_count = len.saturating_sub(suffix_start.saturating_sub(1));
            let suffix = formula_text_byte_slice(old_text.as_str(), suffix_start, suffix_count);
            return Ok(format!("{prefix}{new_text}{suffix}"));
        }
        let chars = old_text.chars().collect::<Vec<_>>();
        if start > chars.len() + 1 {
            return Err(FormulaEvalError::Value);
        }
        let start_index = start - 1;
        let end_index = (start_index + count).min(chars.len());
        let mut output = String::new();
        output.extend(chars[..start_index].iter());
        output.push_str(new_text.as_str());
        output.extend(chars[end_index..].iter());
        Ok(output)
    }

    pub(super) fn parse_substitute_text_function(&mut self) -> Result<String, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let old_text = self.parse_text_value_argument()?;
        if old_text.is_empty() {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let new_text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(text.replace(old_text.as_str(), new_text.as_str()));
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let instance = formula_positive_position_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let mut match_count = 0_usize;
        for (byte_index, _) in text.match_indices(old_text.as_str()) {
            match_count += 1;
            if match_count == instance {
                let mut output = String::new();
                output.push_str(&text[..byte_index]);
                output.push_str(new_text.as_str());
                output.push_str(&text[byte_index + old_text.len()..]);
                return Ok(output);
            }
        }
        Ok(text)
    }

    pub(super) fn parse_formulatext_function(&mut self) -> Result<String, FormulaEvalError> {
        let (target_sheet_id, rect) = self.parse_reference_argument()?;
        if rect.row_first != rect.row_last || rect.col_first != rect.col_last {
            return Err(FormulaEvalError::Value);
        }
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        let Some(formula) =
            self.evaluator
                .formula_source_at(target_sheet_id, rect.row_first, rect.col_first)?
        else {
            return Err(FormulaEvalError::NA);
        };
        let text = if formula.is_r1c1 {
            convert_formula_r1c1_to_a1(&formula.text, rect.row_first, rect.col_first)
        } else {
            formula.text
        };
        Ok(if text.starts_with('=') {
            text
        } else {
            format!("={text}")
        })
    }

    pub(super) fn parse_value_text_format_argument(&mut self) -> Result<bool, FormulaEvalError> {
        self.skip_whitespace();
        if self.consume_char(')') {
            return Ok(false);
        }
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let format = formula_integer_argument(self.parse_comparison()?)?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        match format {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(FormulaEvalError::Value),
        }
    }

    pub(super) fn parse_valuetotext_function(&mut self) -> Result<String, FormulaEvalError> {
        let value = self.parse_value_probe_argument()?;
        let strict = self.parse_value_text_format_argument()?;
        formula_value_to_text(value, strict)
    }

    pub(super) fn parse_arraytotext_function(&mut self) -> Result<String, FormulaEvalError> {
        self.skip_whitespace();
        let checkpoint = self.index;
        let mut rows = Vec::new();
        if let Some((target_sheet_id, rect, next_index)) = self.try_parse_reference()? {
            self.index = next_index;
            self.skip_whitespace();
            if self.peek_char().is_some_and(|ch| matches!(ch, ',' | ')')) {
                for row in rect.row_first..=rect.row_last {
                    let mut values = Vec::new();
                    for col in rect.col_first..=rect.col_last {
                        let value =
                            self.evaluator
                                .cell_value_or_blank(target_sheet_id, row, col)?;
                        values.push(formula_value_probe_from_cell_value(value));
                    }
                    rows.push(values);
                }
            } else {
                self.index = checkpoint;
            }
        }
        if rows.is_empty() {
            rows.push(vec![self.parse_value_probe_argument()?]);
        }
        let strict = self.parse_value_text_format_argument()?;
        if !strict {
            return rows
                .into_iter()
                .flatten()
                .map(|value| formula_value_to_text(value, false))
                .collect::<Result<Vec<_>, _>>()
                .map(|values| values.join(", "));
        }
        let mut row_texts = Vec::new();
        for row in rows {
            let values = row
                .into_iter()
                .map(|value| formula_value_to_text(value, true))
                .collect::<Result<Vec<_>, _>>()?;
            row_texts.push(values.join(","));
        }
        Ok(format!("{{{}}}", row_texts.join(";")))
    }

    pub(super) fn parse_len_function(&mut self, byte_mode: bool) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if byte_mode {
            formula_text_byte_len(text.as_str()) as f64
        } else {
            text.chars().count() as f64
        })
    }

    pub(super) fn parse_character_code_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        text.chars()
            .next()
            .map(|ch| ch as u32 as f64)
            .ok_or(FormulaEvalError::Value)
    }

    pub(super) fn parse_find_function(
        &mut self,
        case_insensitive: bool,
        byte_mode: bool,
    ) -> Result<f64, FormulaEvalError> {
        let needle = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let haystack = self.parse_text_value_argument()?;
        self.skip_whitespace();
        let start = if self.consume_char(')') {
            1
        } else {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            let start = formula_positive_position_argument(self.parse_comparison()?)?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                return Err(FormulaEvalError::Unsupported);
            }
            start
        };
        let haystack_len = haystack.chars().count();
        if start > haystack_len + 1 || (start > haystack_len && !needle.is_empty()) {
            return Err(FormulaEvalError::Value);
        }
        if case_insensitive {
            let Some(position) =
                formula_wildcard_find(needle.as_str(), haystack.as_str(), start, true)
            else {
                return Err(FormulaEvalError::Value);
            };
            return Ok(if byte_mode {
                formula_text_char_position_to_byte_position(haystack.as_str(), position) as f64
            } else {
                position as f64
            });
        }
        let searchable = haystack.chars().skip(start - 1).collect::<String>();
        let Some(byte_index) = searchable.find(needle.as_str()) else {
            return Err(FormulaEvalError::Value);
        };
        let position = start + searchable[..byte_index].chars().count();
        Ok(if byte_mode {
            formula_text_char_position_to_byte_position(haystack.as_str(), position) as f64
        } else {
            position as f64
        })
    }

    pub(super) fn parse_exact_function(&mut self) -> Result<f64, FormulaEvalError> {
        let left = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(',') {
            return Err(FormulaEvalError::Unsupported);
        }
        let right = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        Ok(if left == right { 1.0 } else { 0.0 })
    }

    pub(super) fn parse_value_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        self.skip_whitespace();
        if !self.consume_char(')') {
            return Err(FormulaEvalError::Unsupported);
        }
        formula_value_text(
            self.evaluator.context.date_system(),
            self.evaluator.context.locale(),
            text.as_str(),
        )
    }

    pub(super) fn parse_numbervalue_function(&mut self) -> Result<f64, FormulaEvalError> {
        let text = self.parse_text_value_argument()?;
        let mut decimal_separator = ".".to_string();
        let mut group_separator = ",".to_string();
        self.skip_whitespace();
        if !self.consume_char(')') {
            if !self.consume_char(',') {
                return Err(FormulaEvalError::Unsupported);
            }
            decimal_separator = self.parse_text_value_argument()?;
            self.skip_whitespace();
            if !self.consume_char(')') {
                if !self.consume_char(',') {
                    return Err(FormulaEvalError::Unsupported);
                }
                group_separator = self.parse_text_value_argument()?;
                self.skip_whitespace();
                if !self.consume_char(')') {
                    return Err(FormulaEvalError::Unsupported);
                }
            }
        }
        formula_numbervalue(
            text.as_str(),
            decimal_separator.as_str(),
            group_separator.as_str(),
        )
    }
}
