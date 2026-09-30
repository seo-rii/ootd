mod cell_metadata;
mod cells;
mod table;

pub use cell_metadata::WorkbookCellMetadata;
pub(super) use cell_metadata::{
    SHEET_METADATA_CONTENT_TYPE, dynamic_array_cell_metadata_xml, parse_dynamic_array_cell_metadata,
};
pub(super) use cells::{
    cell_reference, collect_support_part_dimension_coords, format_cell_error,
    parse_worksheet_cells_with_cell_metadata, rewrite_worksheet_xml_with_cell_metadata,
};
#[cfg(test)]
pub(super) use cells::{parse_worksheet_cells, rewrite_worksheet_xml};
pub(super) use table::resolve_table_structural_owners;

#[cfg(test)]
pub(super) use table::parse_table_structural_owner;

#[cfg(test)]
pub(super) use cells::{
    compute_dimension_ref, compute_dimension_ref_with_preserved,
    extend_dimension_coords_from_reference, parse_cell_error, parse_cell_reference,
};
