mod cells;
mod table;

pub(super) use cells::{
    cell_reference, collect_support_part_dimension_coords, format_cell_error,
    parse_worksheet_cells, rewrite_worksheet_xml,
};
pub(super) use table::resolve_table_structural_owners;

#[cfg(test)]
pub(super) use table::parse_table_structural_owner;

#[cfg(test)]
pub(super) use cells::{
    compute_dimension_ref, compute_dimension_ref_with_preserved,
    extend_dimension_coords_from_reference, parse_cell_error, parse_cell_reference,
};
