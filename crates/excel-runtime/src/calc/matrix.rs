//! Matrix determinant/inverse and linear regression helpers.

use super::*;

pub(super) fn formula_matrix_determinant(
    mut matrix: Vec<Vec<f64>>,
) -> Result<f64, FormulaEvalError> {
    let size = matrix.len();
    if size == 0 || matrix.iter().any(|row| row.len() != size) {
        return Err(FormulaEvalError::Value);
    }
    if matrix.iter().flatten().any(|value| !value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }

    let mut determinant = 1.0_f64;
    for pivot_index in 0..size {
        let mut pivot_row = pivot_index;
        let mut pivot_abs = matrix[pivot_index][pivot_index].abs();
        for (row_index, row) in matrix.iter().enumerate().skip(pivot_index + 1) {
            let candidate_abs = row[pivot_index].abs();
            if candidate_abs > pivot_abs {
                pivot_abs = candidate_abs;
                pivot_row = row_index;
            }
        }
        if pivot_abs <= 1e-12 {
            return Ok(0.0);
        }
        if pivot_row != pivot_index {
            matrix.swap(pivot_index, pivot_row);
            determinant = -determinant;
        }
        let pivot = matrix[pivot_index][pivot_index];
        determinant *= pivot;
        if !determinant.is_finite() {
            return Err(FormulaEvalError::Num);
        }
        for row_index in pivot_index + 1..size {
            let factor = matrix[row_index][pivot_index] / pivot;
            for col_index in pivot_index + 1..size {
                matrix[row_index][col_index] -= factor * matrix[pivot_index][col_index];
            }
        }
    }
    formula_checked_numeric_result(determinant)
}

pub(super) fn formula_matrix_inverse_top_left(
    mut matrix: Vec<Vec<f64>>,
) -> Result<f64, FormulaEvalError> {
    let size = matrix.len();
    if size == 0 || matrix.iter().any(|row| row.len() != size) {
        return Err(FormulaEvalError::Value);
    }
    if matrix.iter().flatten().any(|value| !value.is_finite()) {
        return Err(FormulaEvalError::Value);
    }
    let mut inverse = vec![vec![0.0_f64; size]; size];
    for (index, row) in inverse.iter_mut().enumerate() {
        row[index] = 1.0;
    }

    for pivot_index in 0..size {
        let mut pivot_row = pivot_index;
        let mut pivot_abs = matrix[pivot_index][pivot_index].abs();
        for (row_index, row) in matrix.iter().enumerate().skip(pivot_index + 1) {
            let candidate_abs = row[pivot_index].abs();
            if candidate_abs > pivot_abs {
                pivot_abs = candidate_abs;
                pivot_row = row_index;
            }
        }
        if pivot_abs <= 1e-12 {
            return Err(FormulaEvalError::Num);
        }
        if pivot_row != pivot_index {
            matrix.swap(pivot_index, pivot_row);
            inverse.swap(pivot_index, pivot_row);
        }
        let pivot = matrix[pivot_index][pivot_index];
        for col_index in 0..size {
            matrix[pivot_index][col_index] /= pivot;
            inverse[pivot_index][col_index] /= pivot;
        }
        for row_index in 0..size {
            if row_index == pivot_index {
                continue;
            }
            let factor = matrix[row_index][pivot_index];
            for col_index in 0..size {
                matrix[row_index][col_index] -= factor * matrix[pivot_index][col_index];
                inverse[row_index][col_index] -= factor * inverse[pivot_index][col_index];
            }
        }
    }
    formula_checked_numeric_result(inverse[0][0])
}

pub(super) fn formula_regression_default_x(count: usize) -> Vec<f64> {
    (1..=count).map(|value| value as f64).collect()
}

pub(super) fn formula_regression_slope_intercept(
    known_y: &[f64],
    known_x: &[f64],
    constant: bool,
    exponential: bool,
) -> Result<(f64, f64), FormulaEvalError> {
    if known_y.len() != known_x.len() || known_y.is_empty() {
        return Err(FormulaEvalError::NA);
    }
    if known_y
        .iter()
        .chain(known_x.iter())
        .any(|value| !value.is_finite())
    {
        return Err(FormulaEvalError::Value);
    }
    let y_values = if exponential {
        let mut transformed = Vec::with_capacity(known_y.len());
        for value in known_y {
            if *value <= 0.0 {
                return Err(FormulaEvalError::Num);
            }
            transformed.push(value.ln());
        }
        transformed
    } else {
        known_y.to_vec()
    };

    if constant {
        let count = y_values.len() as f64;
        let mean_y = y_values.iter().sum::<f64>() / count;
        let mean_x = known_x.iter().sum::<f64>() / count;
        let mut sum_xy_deviation = 0.0_f64;
        let mut sum_x_deviation_square = 0.0_f64;
        for (y_value, x_value) in y_values.iter().zip(known_x.iter()) {
            let y_deviation = y_value - mean_y;
            let x_deviation = x_value - mean_x;
            sum_xy_deviation += y_deviation * x_deviation;
            sum_x_deviation_square += x_deviation * x_deviation;
        }
        if sum_x_deviation_square == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        let slope = sum_xy_deviation / sum_x_deviation_square;
        formula_checked_numeric_result(mean_y - slope * mean_x).map(|intercept| (slope, intercept))
    } else {
        let mut sum_xy = 0.0_f64;
        let mut sum_x_square = 0.0_f64;
        for (y_value, x_value) in y_values.iter().zip(known_x.iter()) {
            sum_xy += y_value * x_value;
            sum_x_square += x_value * x_value;
        }
        if sum_x_square == 0.0 {
            return Err(FormulaEvalError::Div0);
        }
        formula_checked_numeric_result(sum_xy / sum_x_square).map(|slope| (slope, 0.0))
    }
}
