/// Builds JSON request strings for all CLI engine operations.
/// Replaces the C# `CliEngineRequestBuilder` + `DispatcherLibrary.IJsonWrap`/`JsonFactory`.
use serde_json::{json, Value};

// ─── Engine action ID constants ─────────────────────────────────────────────
pub const ACTION_IMPORT_WORKBOOK: i32 = 179;
pub const ACTION_OPEN_WORKBOOK: i32 = 272;
pub const ACTION_CREATE_WORKBOOK: i32 = 128;
pub const ACTION_CREATE_WORKSHEET: i32 = 104;
pub const ACTION_EXPORT_WORKBOOK: i32 = 270;
pub const ACTION_CONTENT_API: i32 = 0;
pub const ACTION_ACTIVE_CELL_INFO: i32 = 17988;
pub const ACTION_EDIT_CELL_INFO: i32 = 8552;
pub const ACTION_CLOSE_WORKBOOK: i32 = 185;

// Sheet management
pub const ACTION_DELETE_WORKSHEET: i32 = 100;
pub const ACTION_RENAME_WORKSHEET: i32 = 102;
pub const ACTION_REORDER_WORKSHEET: i32 = 101;
pub const ACTION_HIDE_WORKSHEET: i32 = 103;
pub const ACTION_UNHIDE_WORKSHEET: i32 = 105;
pub const ACTION_DUPLICATE_WORKSHEET: i32 = 768;

// Row & column structure
pub const ACTION_INSERT_ROW: i32 = 42;
pub const ACTION_DELETE_ROW: i32 = 44;
pub const ACTION_INSERT_COLUMN: i32 = 43;
pub const ACTION_DELETE_COLUMN: i32 = 45;
pub const ACTION_HIDE_ROW: i32 = 109;
pub const ACTION_UNHIDE_ROW: i32 = 110;
pub const ACTION_HIDE_COLUMN: i32 = 324;
pub const ACTION_UNHIDE_COLUMN: i32 = 325;
pub const ACTION_RESIZE_ROW: i32 = 10;
pub const ACTION_RESIZE_COLUMN: i32 = 41;

// Clear
pub const ACTION_CLEAR_ALL: i32 = 61;
pub const ACTION_CLEAR_CONTENT: i32 = 32;
pub const ACTION_CLEAR_FORMAT: i32 = 31;

// Merge
pub const ACTION_MERGECELLS: i32 = 149;
pub const ACTION_UNMERGE: i32 = 37;

// Clipboard
pub const ACTION_SERVERCLIP_COPY: i32 = 760;
pub const ACTION_CUT: i32 = 761;
pub const ACTION_SERVERCLIP_PASTE: i32 = 762;

// Undo / Redo
pub const ACTION_UNDO: i32 = 163;
pub const ACTION_REDO: i32 = 164;

// Find / Replace
pub const ACTION_FIND: i32 = 254;
pub const ACTION_REPLACE: i32 = 255;
pub const ACTION_REPLACEALL: i32 = 256;

// Sort & Filter
pub const ACTION_SORT: i32 = 60;
pub const ACTION_CREATE_FILTER: i32 = 260;
pub const ACTION_REMOVE_FILTER: i32 = 262;

// Named ranges
pub const ACTION_ADDDEFINEDNAME: i32 = 137;
pub const ACTION_DELETEDEFINEDNAME: i32 = 139;
pub const ACTION_MANAGEDEFINEDNAME: i32 = 140;

// Freeze
pub const ACTION_FREEZE: i32 = 162;
pub const ACTION_UNFREEZE: i32 = 196;

// ─── Shared JSON builders ────────────────────────────────────────────────────

fn build_active_cell(row: i32, col: i32) -> Value {
    json!({
        "active_row": row,
        "active_column": col
    })
}

fn build_range_object(start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> Value {
    json!({
        "start_row": start_row,
        "start_column": start_col,
        "end_row": end_row,
        "end_column": end_col
    })
}

fn build_active_info(sheet_id: &str, row: i32, col: i32) -> Value {
    json!({
        "active_sheet_id": sheet_id,
        "active_cell": build_active_cell(row, col),
        "active_range_list": [build_range_object(row, col, row, col)]
    })
}

fn build_sheet_range_list(
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> Value {
    json!([{
        "sheet_id": sheet_id,
        "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
    }])
}

// ─── File operations ─────────────────────────────────────────────────────────

pub fn build_import_workbook(file_path: &str) -> String {
    json!({
        "action_id": ACTION_IMPORT_WORKBOOK,
        "file_path": file_path
    })
    .to_string()
}

pub fn build_open_workbook(file_path: &str) -> String {
    json!({
        "action_id": ACTION_OPEN_WORKBOOK,
        "file_path": file_path
    })
    .to_string()
}

pub fn build_create_workbook(doc_name: &str) -> String {
    json!({
        "action_id": ACTION_CREATE_WORKBOOK,
        "file_path": doc_name
    })
    .to_string()
}

pub fn build_close_workbook(rid: &str) -> String {
    json!({
        "action_id": ACTION_CLOSE_WORKBOOK,
        "rid": rid,
        "is_force_close": true
    })
    .to_string()
}

pub fn build_export_workbook(
    rid: &str,
    sheet_id: &str,
    directory_path: &str,
    file_name: &str,
    file_type: i32,
    is_save_as: bool,
) -> String {
    let mut v = json!({
        "action_id": ACTION_EXPORT_WORKBOOK,
        "rid": rid,
        "sheet_id": sheet_id,
        "file_path": directory_path,
        "file_name": file_name,
        "file_type": file_type
    });
    if is_save_as {
        v["is_save_as"] = json!(true);
    }
    v.to_string()
}

// ─── Sheet management ────────────────────────────────────────────────────────

pub fn build_add_sheet(rid: &str) -> String {
    json!({
        "action_id": ACTION_CREATE_WORKSHEET,
        "rid": rid
    })
    .to_string()
}

pub fn build_delete_sheet(rid: &str, sheet_id: &str) -> String {
    json!({
        "action_id": ACTION_DELETE_WORKSHEET,
        "rid": rid,
        "sheet_id": sheet_id
    })
    .to_string()
}

pub fn build_rename_sheet(rid: &str, sheet_id: &str, new_name: &str) -> String {
    json!({
        "action_id": ACTION_RENAME_WORKSHEET,
        "rid": rid,
        "sheet_id": sheet_id,
        "new_sheet_name": new_name
    })
    .to_string()
}

pub fn build_reorder_sheet(rid: &str, sheet_id: &str, new_position: i32) -> String {
    json!({
        "action_id": ACTION_REORDER_WORKSHEET,
        "rid": rid,
        "sheet_id": sheet_id,
        "new_position": new_position
    })
    .to_string()
}

pub fn build_duplicate_sheet(rid: &str, active_sheet_id: &str) -> String {
    json!({
        "action_id": ACTION_DUPLICATE_WORKSHEET,
        "rid": rid,
        "active_sheet_id": active_sheet_id
    })
    .to_string()
}

pub fn build_hide_sheet(rid: &str, sheet_ids: &[&str]) -> String {
    json!({
        "action_id": ACTION_HIDE_WORKSHEET,
        "rid": rid,
        "sheet_list": sheet_ids
    })
    .to_string()
}

pub fn build_unhide_sheet(rid: &str, sheet_ids: &[&str]) -> String {
    json!({
        "action_id": ACTION_UNHIDE_WORKSHEET,
        "rid": rid,
        "sheet_list": sheet_ids,
        "is_expose_all": false
    })
    .to_string()
}

// ─── Cell operations ─────────────────────────────────────────────────────────

pub fn build_set_cell_value(
    rid: &str,
    sheet_id: &str,
    row: i32,
    col: i32,
    value: &str,
    is_formula: bool,
) -> String {
    json!({
        "action_id": ACTION_CONTENT_API,
        "rid": rid,
        "sheet_range_list": [{
            "sheet_id": sheet_id,
            "start_row": row,
            "start_column": col,
            "end_row": row,
            "end_column": col
        }],
        "content": value,
        "is_cse": is_formula,
        "active_info": build_active_info(sheet_id, row, col)
    })
    .to_string()
}

pub fn build_get_cell_info(rid: &str, sheet_id: &str, row: i32, col: i32) -> String {
    json!({
        "action_id": ACTION_ACTIVE_CELL_INFO,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_cell": build_active_cell(row, col)
    })
    .to_string()
}

pub fn build_evaluate_formula(rid: &str, sheet_id: &str, formula: &str) -> String {
    json!({
        "action_id": ACTION_EDIT_CELL_INFO,
        "rid": rid,
        "sheet_id": sheet_id,
        "content": formula
    })
    .to_string()
}

// ─── Fetch requests (sent via FetchJson / DocFetchJson) ──────────────────────

pub fn build_doc_fetch(rid: &str) -> String {
    // doc_meta bitmask: sheetsList(1)|sheetMeta(2)|activeSheetInfo(4)|meta(8)|
    // definedName(16)|defaultStyleInfo(32)|workbookViewInfo(64)|documentSettings(128) = 255
    json!({
        "rid": rid,
        "doc_meta": 255
    })
    .to_string()
}

pub fn build_initial_sheet_fetch(
    rid: &str,
    sheet_id: &str,
    max_row: i32,
    max_col: i32,
) -> String {
    // doc_meta = sheetMeta(2)
    // sheet_meta = rowheader(2)|columnHeader(4)|defaultRowHeight(8192)|
    //   defaultColumnWidth(262144)|sheetViewInfo(16384)|freezeInfo(1024)|
    //   activeInfo(16)|filter(2048)|hiddenRows(256)|hiddenColumns(512)|
    //   sparklineInfo(1048576) = 1338962
    let doc_meta = 2;
    let sheet_meta = 2 | 4 | 8192 | 262144 | 16384 | 1024 | 16 | 2048 | 256 | 512 | 1048576;

    json!({
        "rid": rid,
        "doc_meta": doc_meta,
        "meta": [{"sheet_meta": sheet_meta, "cell_meta": 0}],
        "ranges": [{"boundary": [0, 0, max_row, max_col], "sheet_id": sheet_id}]
    })
    .to_string()
}

pub fn build_cell_fetch(rid: &str, sheet_id: &str, row: i32, col: i32) -> String {
    // cell_meta = actualValue(1)|displayValue(2)|formulaValue(4) = 7
    json!({
        "rid": rid,
        "doc_meta": 2,
        "meta": [{"sheet_meta": 1, "cell_meta": 7}],
        "ranges": [{"boundary": [row, col, row, col], "sheet_id": sheet_id}]
    })
    .to_string()
}

pub fn build_range_cell_fetch(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    // cell_meta = actualValue(1)|displayValue(2) = 3
    json!({
        "rid": rid,
        "doc_meta": 2,
        "meta": [{"sheet_meta": 1, "cell_meta": 3}],
        "ranges": [{"boundary": [start_row, start_col, end_row, end_col], "sheet_id": sheet_id}]
    })
    .to_string()
}

pub fn build_fetch_workbook_info(rid: &str) -> String {
    json!({ "rid": rid }).to_string()
}

pub fn build_fetch_workbook_info_with_sheet(rid: &str, sheet_id: &str) -> String {
    json!({ "rid": rid, "sheet_id": sheet_id }).to_string()
}

// ─── Row & column operations ─────────────────────────────────────────────────

pub fn build_insert_row(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
    is_above: bool,
) -> String {
    json!({
        "action_id": ACTION_INSERT_ROW,
        "rid": rid,
        "sheet_id": sheet_id,
        "range_list": [build_range_object(start_row, start_col, end_row, end_col)],
        "is_insert_above": is_above,
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_delete_row(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_DELETE_ROW,
        "rid": rid,
        "sheet_id": sheet_id,
        "range_list": [build_range_object(start_row, start_col, end_row, end_col)],
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_insert_column(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
    is_before: bool,
) -> String {
    json!({
        "action_id": ACTION_INSERT_COLUMN,
        "rid": rid,
        "sheet_id": sheet_id,
        "range_list": [build_range_object(start_row, start_col, end_row, end_col)],
        "is_insert_before": is_before,
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_delete_column(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_DELETE_COLUMN,
        "rid": rid,
        "sheet_id": sheet_id,
        "range_list": [build_range_object(start_row, start_col, end_row, end_col)],
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_hide_row(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    end_row: i32,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_HIDE_ROW,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, 0, end_row, 0),
        "active_info": build_active_info(sheet_id, active_row, active_col),
        "column_width": 0
    })
    .to_string()
}

pub fn build_unhide_row(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    end_row: i32,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_UNHIDE_ROW,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, 0, end_row, 0),
        "active_info": build_active_info(sheet_id, active_row, active_col),
        "column_width": 0
    })
    .to_string()
}

pub fn build_hide_column(
    rid: &str,
    sheet_id: &str,
    start_col: i32,
    end_col: i32,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_HIDE_COLUMN,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, 0, start_col, 0, end_col),
        "active_info": build_active_info(sheet_id, active_row, active_col),
        "row_height": 0
    })
    .to_string()
}

pub fn build_unhide_column(
    rid: &str,
    sheet_id: &str,
    start_col: i32,
    end_col: i32,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_UNHIDE_COLUMN,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, 0, start_col, 0, end_col),
        "active_info": build_active_info(sheet_id, active_row, active_col),
        "row_height": 0
    })
    .to_string()
}

pub fn build_resize_row(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    end_row: i32,
    height: i32,
    auto_fit: bool,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_RESIZE_ROW,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, 0, end_row, 0),
        "is_auto_fit": auto_fit,
        "active_info": build_active_info(sheet_id, active_row, active_col),
        "row_height": height
    })
    .to_string()
}

pub fn build_resize_column(
    rid: &str,
    sheet_id: &str,
    start_col: i32,
    end_col: i32,
    width: i32,
    auto_fit: bool,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_RESIZE_COLUMN,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, 0, start_col, 0, end_col),
        "is_auto_fit": auto_fit,
        "active_info": build_active_info(sheet_id, active_row, active_col),
        "column_width": width
    })
    .to_string()
}

// ─── Clear ───────────────────────────────────────────────────────────────────

pub fn build_clear(
    rid: &str,
    sheet_id: &str,
    action_id: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": action_id,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

// ─── Merge ───────────────────────────────────────────────────────────────────

pub fn build_merge_cells(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_MERGECELLS,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col),
        "is_forced": true
    })
    .to_string()
}

pub fn build_unmerge_cells(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_UNMERGE,
        "rid": rid,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

// ─── Clipboard ───────────────────────────────────────────────────────────────

pub fn build_copy(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SERVERCLIP_COPY,
        "rid": rid,
        "sheet_id": sheet_id,
        "range_list": [build_range_object(start_row, start_col, end_row, end_col)],
        "copy_selected_type": 0,
        "object_id_list": []
    })
    .to_string()
}

pub fn build_cut(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    // Note: cut uses a single "range" object, not an array — different from copy.
    json!({
        "action_id": ACTION_CUT,
        "rid": rid,
        "sheet_id": sheet_id,
        "range": build_range_object(start_row, start_col, end_row, end_col),
        "copy_selected_type": 0,
        "object_id_list": []
    })
    .to_string()
}

pub fn build_paste(rid: &str, sheet_id: &str, row: i32, col: i32, paste_type: i32) -> String {
    json!({
        "action_id": ACTION_SERVERCLIP_PASTE,
        "rid": rid,
        "sheet_id": sheet_id,
        "paste_type": paste_type,
        "range_list": [build_range_object(row, col, row, col)],
        "active_info": build_active_info(sheet_id, row, col)
    })
    .to_string()
}

// ─── Undo / Redo ─────────────────────────────────────────────────────────────

pub fn build_undo(rid: &str) -> String {
    json!({
        "action_id": ACTION_UNDO,
        "rid": rid
    })
    .to_string()
}

pub fn build_redo(rid: &str) -> String {
    json!({
        "action_id": ACTION_REDO,
        "rid": rid
    })
    .to_string()
}

// ─── Find / Replace ──────────────────────────────────────────────────────────

pub fn build_find(
    rid: &str,
    sheet_id: &str,
    search_text: &str,
    active_row: i32,
    active_col: i32,
    is_exact: bool,
    is_case_sensitive: bool,
    is_next: bool,
) -> String {
    json!({
        "action_id": ACTION_FIND,
        "rid": rid,
        "sheet_id": sheet_id,
        "cell_content": search_text,
        "active_row": active_row,
        "active_column": active_col,
        "range_list": [build_range_object(active_row, active_col, active_row, active_col)],
        "get_number_of_matches": true,
        "is_exact_match": is_exact,
        "is_next": is_next,
        "is_row_wise": true,
        "is_value_string": true,
        "traverse_mode": 1,
        "is_case_sensitive": is_case_sensitive
    })
    .to_string()
}

pub fn build_replace(
    rid: &str,
    sheet_id: &str,
    search_text: &str,
    replace_text: &str,
    active_row: i32,
    active_col: i32,
    is_exact: bool,
    is_case_sensitive: bool,
    replace_all: bool,
) -> String {
    let action_id = if replace_all {
        ACTION_REPLACEALL
    } else {
        ACTION_REPLACE
    };
    json!({
        "action_id": action_id,
        "rid": rid,
        "sheet_id": sheet_id,
        "cell_content": search_text,
        "replace_string": replace_text,
        "active_row": active_row,
        "active_column": active_col,
        "range_list": [build_range_object(active_row, active_col, active_row, active_col)],
        "get_number_of_matches": true,
        "is_exact_match": is_exact,
        "is_next": true,
        "is_row_wise": true,
        "is_value_string": true,
        "traverse_mode": 1,
        "is_case_sensitive": is_case_sensitive,
        "is_include_formulas": false,
        "active_info": build_active_info(sheet_id, active_row, active_col)
    })
    .to_string()
}

// ─── Sort & Filter ───────────────────────────────────────────────────────────

pub fn build_sort(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
    sort_column: i32,
    is_ascending: bool,
    has_header: bool,
) -> String {
    json!({
        "action_id": ACTION_SORT,
        "rid": rid,
        "sheet_id": sheet_id,
        "start_row": start_row,
        "start_column": start_col,
        "end_row": end_row,
        "end_column": end_col,
        "is_sort_columns": false,
        "is_case_sensitive": false,
        "contains_header": has_header,
        "sort_conditions": [{
            "sort_base": sort_column,
            "sort_type": 0,
            "sort_order": is_ascending
        }],
        "is_auto_expand": true,
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_create_filter(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_CREATE_FILTER,
        "rid": rid,
        "sheet_id": sheet_id,
        "start_row": start_row,
        "start_column": start_col,
        "end_row": end_row,
        "end_column": end_col,
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_remove_filter(rid: &str, sheet_id: &str) -> String {
    json!({
        "action_id": ACTION_REMOVE_FILTER,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_info": build_active_info(sheet_id, 0, 0),
        "table_id": ""
    })
    .to_string()
}

// ─── Named ranges ────────────────────────────────────────────────────────────

pub fn build_add_defined_name(
    rid: &str,
    sheet_id: &str,
    name: &str,
    expression: &str,
    comment: &str,
) -> String {
    json!({
        "action_id": ACTION_ADDDEFINEDNAME,
        "rid": rid,
        "sheet_id": sheet_id,
        "scope_sheet_id": sheet_id,
        "defined_name": name,
        "active_row_index": 0,
        "active_column_index": 0,
        "expression": expression,
        "comment": comment
    })
    .to_string()
}

pub fn build_delete_defined_name(rid: &str, sheet_id: &str, name: &str) -> String {
    json!({
        "action_id": ACTION_DELETEDEFINEDNAME,
        "rid": rid,
        "sheet_id": sheet_id,
        "scope_sheet_id": sheet_id,
        "defined_name": name
    })
    .to_string()
}

pub fn build_manage_defined_names(rid: &str) -> String {
    json!({
        "action_id": ACTION_MANAGEDEFINEDNAME,
        "rid": rid,
        "active_row_index": 0,
        "active_column_index": 0
    })
    .to_string()
}

// ─── Freeze / Unfreeze ───────────────────────────────────────────────────────

pub fn build_freeze(rid: &str, sheet_id: &str, row: i32, col: i32) -> String {
    json!({
        "action_id": ACTION_FREEZE,
        "rid": rid,
        "sheet_id": sheet_id,
        "range": build_range_object(row, col, row, col)
    })
    .to_string()
}

pub fn build_unfreeze(rid: &str, sheet_id: &str) -> String {
    json!({
        "action_id": ACTION_UNFREEZE,
        "rid": rid,
        "sheet_id": sheet_id
    })
    .to_string()
}
