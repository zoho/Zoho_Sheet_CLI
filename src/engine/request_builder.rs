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

// Table
pub const ACTION_INSERT_TABLE: i32 = 10073;
pub const ACTION_SELECT_TABLE_RANGE: i32 = 10132;
pub const ACTION_CHANGE_TABLE_NAME: i32 = 10133;
pub const ACTION_MANAGE_TABLE: i32 = 10134;
pub const ACTION_DELETE_TABLE: i32 = 10135;
pub const ACTION_CHANGE_TABLE_STYLE: i32 = 10136;
pub const ACTION_SET_DEFAULT_TABLE_STYLE: i32 = 10137;
pub const ACTION_INSERT_TABLE_ROW: i32 = 10144;
pub const ACTION_INSERT_TABLE_COLUMN: i32 = 10145;
pub const ACTION_DELETE_TABLE_ROW: i32 = 10146;
pub const ACTION_DELETE_TABLE_COLUMN: i32 = 10147;
pub const ACTION_CHANGE_TABLE_SOURCE: i32 = 10148;
pub const ACTION_CHANGE_TABLE_OPTIONS: i32 = 10150;

// Pivot Table
pub const ACTION_CREATE_PIVOT_TABLE: i32 = 913;
pub const ACTION_DELETE_PIVOT_TABLE: i32 = 914;
pub const ACTION_APPLY_PIVOT_GROUPING: i32 = 915;
pub const ACTION_MODIFY_PIVOT_PROPERTIES: i32 = 916;
pub const ACTION_CHANGE_PIVOT_FIELD_TYPE: i32 = 917;
pub const ACTION_REMOVE_PIVOT_FIELD: i32 = 918;
pub const ACTION_REMOVE_PIVOT_FILTER: i32 = 919;
pub const ACTION_REMOVE_PIVOT_SORT: i32 = 920;
pub const ACTION_SELECT_PIVOT_FIELD: i32 = 921;
pub const ACTION_REMOVE_GROUP: i32 = 922;
pub const ACTION_MODIFY_VALUE_AGGREGATION_TYPE: i32 = 923;
pub const ACTION_MODIFY_VALUE_SHOW_DATA_AS: i32 = 924;
pub const ACTION_PIVOT_TABLE_INFO: i32 = 925;
pub const ACTION_PIVOT_CELL_INFO: i32 = 926;
pub const ACTION_PIVOT_FILTER_INFO: i32 = 927;
pub const ACTION_APPLY_PIVOT_DATE_GROUPING: i32 = 928;
pub const ACTION_MOVE_PIVOT_TABLE: i32 = 929;
pub const ACTION_REFRESH_PIVOT_TABLE: i32 = 930;
pub const ACTION_CHANGE_PIVOT_TABLE_SOURCE: i32 = 931;
pub const ACTION_APPLY_PIVOT_FILTER: i32 = 911;
pub const ACTION_APPLY_PIVOT_SORT: i32 = 912;
pub const ACTION_COPY_PIVOT_TABLE: i32 = 933;
pub const ACTION_EDIT_PIVOT_NAME: i32 = 934;
pub const ACTION_REFRESH_PIVOT_TABLE_ON_LOAD: i32 = 935;

// Chart
pub const ACTION_CLONE_CHART: i32 = 71;
pub const ACTION_DELETE_CHART: i32 = 73;
pub const ACTION_UPDATE_CHART_POSITION: i32 = 74;
pub const ACTION_INSERT_CHART: i32 = 75;
pub const ACTION_MANAGE_CHART_WITH_RANGE: i32 = 76;
pub const ACTION_MANAGE_CHART_WITH_ID: i32 = 80;
pub const ACTION_CUSTOMIZE_CHART_PROPERTY_TWO: i32 = 652;
pub const ACTION_RECOMMEND_CHART: i32 = 695;
pub const ACTION_MANAGE_TOP_BOTTOM: i32 = 697;
pub const ACTION_UPDATE_CHART_TYPE: i32 = 1122;
pub const ACTION_CUSTOMIZE_CHART_PROPERTY_ONE: i32 = 6001;
pub const ACTION_MOVE_CHART: i32 = 6002;

// Font formatting
pub const ACTION_SET_ITALIC: i32 = 1;
pub const ACTION_SET_UNDERLINE: i32 = 2;
pub const ACTION_SET_FONT_SIZE: i32 = 4;
pub const ACTION_SET_FONT_COLOR: i32 = 8;
pub const ACTION_SET_DOUBLE_UNDERLINE: i32 = 13;
pub const ACTION_SET_BOLD: i32 = 38;
pub const ACTION_STRIKE_THROUGH: i32 = 77;
pub const ACTION_SET_SUPERSCRIPT: i32 = 78;
pub const ACTION_SET_SUBSCRIPT: i32 = 79;

// Cell formatting – alignment
pub const ACTION_VERTICAL_ALIGNMENT: i32 = 5;
pub const ACTION_HORIZONTAL_ALIGNMENT: i32 = 6;
pub const ACTION_FILL_COLOR: i32 = 7;
pub const ACTION_WRAP_TEXT: i32 = 9;
pub const ACTION_TEXT_ROTATION: i32 = 14;
pub const ACTION_SET_BORDER: i32 = 36;
pub const ACTION_DEFAULT_FORMAT: i32 = 275;
pub const ACTION_INCREASE_INDENT: i32 = 10052;
pub const ACTION_DECREASE_INDENT: i32 = 10053;

// Cell formatting – number formatting
pub const ACTION_INCREASE_DECIMAL: i32 = 172;
pub const ACTION_DECREASE_DECIMAL: i32 = 173;
pub const ACTION_MANAGE_NUMBER_FORMAT: i32 = 3090;
pub const ACTION_MANAGE_CUSTOM_FORMAT: i32 = 4000;
pub const ACTION_GET_NUMBER_FORMAT_INFO: i32 = 5526;
pub const ACTION_PREVIEW_NUMBER_FORMAT: i32 = 10058;
pub const ACTION_APPLY_NUMBER_FORMAT: i32 = 10059;

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

pub fn build_open_workbook(file_path: &str, file_type: Option<i32>) -> String {
    let mut v = json!({
        "action_id": ACTION_OPEN_WORKBOOK,
        "file_path": file_path
    });
    if let Some(ft) = file_type {
        v["file_type"] = json!(ft);
    }
    v.to_string()
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
        "sheet_id": sheet_id,
        "active_sheet_id": sheet_id,
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

// ─── Table operations ────────────────────────────────────────────────────────

pub fn build_insert_table(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
    has_headers: bool,
) -> String {
    json!({
        "action_id": ACTION_INSERT_TABLE,
        "rid": rid,
        "sheet_range_list": [{
            "sheet_id": sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }],
        "is_contain_headers": has_headers,
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_select_table_range(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SELECT_TABLE_RANGE,
        "rid": rid,
        "sheet_range_list": [{
            "sheet_id": sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }]
    })
    .to_string()
}

pub fn build_delete_table(rid: &str, table_id: &str, keep_format: bool) -> String {
    json!({
        "action_id": ACTION_DELETE_TABLE,
        "rid": rid,
        "table_id": table_id,
        "is_keep_table_format": keep_format
    })
    .to_string()
}

pub fn build_delete_table_column(
    rid: &str,
    table_id: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_DELETE_TABLE_COLUMN,
        "rid": rid,
        "table_id": table_id,
        "sheet_range_list": [{
            "sheet_id": sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }]
    })
    .to_string()
}

pub fn build_delete_table_row(
    rid: &str,
    table_id: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_DELETE_TABLE_ROW,
        "rid": rid,
        "table_id": table_id,
        "sheet_range_list": [{
            "sheet_id": sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }]
    })
    .to_string()
}

pub fn build_change_table_name(
    rid: &str,
    sheet_id: &str,
    table_id: &str,
    new_name: &str,
) -> String {
    json!({
        "action_id": ACTION_CHANGE_TABLE_NAME,
        "rid": rid,
        "sheet_id": sheet_id,
        "table_id": table_id,
        "table_name": new_name
    })
    .to_string()
}

pub fn build_change_table_options(
    rid: &str,
    table_id: &str,
    sheet_id: &str,
    setting_type: i32,
    is_enabled: bool,
) -> String {
    json!({
        "action_id": ACTION_CHANGE_TABLE_OPTIONS,
        "rid": rid,
        "table_id": table_id,
        "sheet_id": sheet_id,
        "table_settings": {
            "setting_type": setting_type,
            "is_enabled": is_enabled
        }
    })
    .to_string()
}

pub fn build_change_table_source(
    rid: &str,
    table_id: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_CHANGE_TABLE_SOURCE,
        "rid": rid,
        "table_id": table_id,
        "sheet_range_list": [{
            "sheet_id": sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }]
    })
    .to_string()
}

pub fn build_change_table_style_pattern(
    rid: &str,
    table_id: &str,
    style_pattern: i32,
    keep_cell_format: bool,
) -> String {
    json!({
        "action_id": ACTION_CHANGE_TABLE_STYLE,
        "rid": rid,
        "table_id": table_id,
        "table_style": {
            "table_style_pattern": style_pattern,
            "color": {
                "theme_color": 1,
                "tint": 0.0
            }
        },
        "is_keep_cell_format": keep_cell_format
    })
    .to_string()
}

pub fn build_insert_table_column(
    rid: &str,
    table_id: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
    is_after: bool,
) -> String {
    json!({
        "action_id": ACTION_INSERT_TABLE_COLUMN,
        "rid": rid,
        "table_id": table_id,
        "sheet_range_list": [{
            "sheet_id": sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }],
        "is_insert_after": is_after
    })
    .to_string()
}

pub fn build_insert_table_row(
    rid: &str,
    table_id: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
    is_below: bool,
) -> String {
    json!({
        "action_id": ACTION_INSERT_TABLE_ROW,
        "rid": rid,
        "table_id": table_id,
        "sheet_range_list": [{
            "sheet_id": sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }],
        "is_insert_below": is_below
    })
    .to_string()
}

pub fn build_manage_table(rid: &str, table_id: &str) -> String {
    json!({
        "action_id": ACTION_MANAGE_TABLE,
        "rid": rid,
        "table_id": table_id
    })
    .to_string()
}

/// Builds a fetch request for table list using sheet_meta filter (2048).
/// Filter objects in the response contain `table_id` when the filter belongs to a table.
pub fn build_table_list_fetch(rid: &str, sheet_id: &str) -> String {
    // sheet_meta = filter(2048) — filter objects contain table_id when from a table
    // Use full sheet boundary to ensure all tables are captured
    json!({
        "rid": rid,
        "doc_meta": 2,
        "meta": [{"sheet_meta": 2048, "cell_meta": 0}],
        "ranges": [{"boundary": [0, 0, 262143, 4095], "sheet_id": sheet_id}]
    })
    .to_string()
}

pub fn build_set_default_table_style(rid: &str, style_pattern: i32) -> String {
    json!({
        "action_id": ACTION_SET_DEFAULT_TABLE_STYLE,
        "rid": rid,
        "table_style": {
            "table_style_pattern": style_pattern,
            "color": {
                "theme_color": 1,
                "tint": 0.0
            }
        }
    })
    .to_string()
}

// ─── Font formatting ─────────────────────────────────────────────────────────

pub fn build_set_bold(
    rid: &str,
    sheet_id: &str,
    is_bold: bool,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_BOLD,
        "rid": rid,
        "is_bold": is_bold,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_set_italic(
    rid: &str,
    sheet_id: &str,
    is_italic: bool,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_ITALIC,
        "rid": rid,
        "is_italic": is_italic,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_set_underline(
    rid: &str,
    sheet_id: &str,
    is_underline: bool,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_UNDERLINE,
        "rid": rid,
        "is_underline": is_underline,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_set_double_underline(
    rid: &str,
    sheet_id: &str,
    is_double_underline: bool,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_DOUBLE_UNDERLINE,
        "rid": rid,
        "is_double_underline": is_double_underline,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_strike_through(
    rid: &str,
    sheet_id: &str,
    is_strike_through: bool,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_STRIKE_THROUGH,
        "rid": rid,
        "is_strike_through": is_strike_through,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_set_superscript(
    rid: &str,
    sheet_id: &str,
    is_superscript: bool,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_SUPERSCRIPT,
        "rid": rid,
        "is_superscript": is_superscript,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_set_subscript(
    rid: &str,
    sheet_id: &str,
    is_subscript: bool,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_SUBSCRIPT,
        "rid": rid,
        "is_subscript": is_subscript,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_set_font_size(
    rid: &str,
    sheet_id: &str,
    font_size: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_FONT_SIZE,
        "rid": rid,
        "font_size": font_size,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_set_font_color_rgb(
    rid: &str,
    sheet_id: &str,
    red: i32,
    green: i32,
    blue: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_FONT_COLOR,
        "rid": rid,
        "is_automatic": false,
        "font_color": {
            "red": red,
            "green": green,
            "blue": blue
        },
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_set_font_color_auto(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_FONT_COLOR,
        "rid": rid,
        "is_automatic": true,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

// ─── Cell formatting – alignment ─────────────────────────────────────────────

pub fn build_horizontal_alignment(
    rid: &str,
    sheet_id: &str,
    alignment_type: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_HORIZONTAL_ALIGNMENT,
        "rid": rid,
        "horizontal_alignment_type": alignment_type,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_vertical_alignment(
    rid: &str,
    sheet_id: &str,
    alignment_type: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_VERTICAL_ALIGNMENT,
        "rid": rid,
        "vertical_alignment_type": alignment_type,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_wrap_text(
    rid: &str,
    sheet_id: &str,
    wrap_type: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_WRAP_TEXT,
        "rid": rid,
        "wrap_type": wrap_type,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_text_rotation(
    rid: &str,
    sheet_id: &str,
    angle: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_TEXT_ROTATION,
        "rid": rid,
        "text_rotation_angle": angle,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_increase_indent(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_INCREASE_INDENT,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_row_index": start_row,
        "active_column_index": start_col,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_decrease_indent(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_DECREASE_INDENT,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_row_index": start_row,
        "active_column_index": start_col,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

// ─── Cell formatting – fill color ────────────────────────────────────────────

pub fn build_fill_color_rgb(
    rid: &str,
    sheet_id: &str,
    red: i32,
    green: i32,
    blue: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_FILL_COLOR,
        "rid": rid,
        "no_fill": false,
        "fill_color": {
            "red": red,
            "green": green,
            "blue": blue
        },
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_fill_color_none(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_FILL_COLOR,
        "rid": rid,
        "no_fill": true,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

// ─── Cell formatting – border ────────────────────────────────────────────────

pub fn build_set_border(
    rid: &str,
    sheet_id: &str,
    border_type: i32,
    border_line_style: i32,
    red: i32,
    green: i32,
    blue: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_SET_BORDER,
        "rid": rid,
        "border_type": border_type,
        "border_line_style": border_line_style,
        "border_color": {
            "red": red,
            "green": green,
            "blue": blue
        },
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

// ─── Cell formatting – default format ────────────────────────────────────────

pub fn build_default_format(rid: &str, format_json: Value) -> String {
    let mut v = format_json;
    v["action_id"] = json!(ACTION_DEFAULT_FORMAT);
    v["rid"] = json!(rid);
    v.to_string()
}

// ─── Number formatting ───────────────────────────────────────────────────────

pub fn build_apply_number_format(
    rid: &str,
    sheet_id: &str,
    number_format_text: &str,
    number_format_type: i32,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_APPLY_NUMBER_FORMAT,
        "rid": rid,
        "number_format_text": number_format_text,
        "number_format_type": number_format_type,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_increase_decimal(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_INCREASE_DECIMAL,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_row_index": start_row,
        "active_column_index": start_col,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_decrease_decimal(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_DECREASE_DECIMAL,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_row_index": start_row,
        "active_column_index": start_col,
        "sheet_range_list": build_sheet_range_list(sheet_id, start_row, start_col, end_row, end_col),
        "active_info": build_active_info(sheet_id, start_row, start_col)
    })
    .to_string()
}

pub fn build_preview_number_format(
    rid: &str,
    sheet_id: &str,
    number_format_text: &str,
    number_format_type: i32,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_PREVIEW_NUMBER_FORMAT,
        "rid": rid,
        "sheet_id": sheet_id,
        "number_format_text": number_format_text,
        "number_format_type": number_format_type,
        "active_row_index": active_row,
        "active_column_index": active_col
    })
    .to_string()
}

pub fn build_get_number_format_info(
    rid: &str,
    sheet_id: &str,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_GET_NUMBER_FORMAT_INFO,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_row_index": active_row,
        "active_column_index": active_col
    })
    .to_string()
}

pub fn build_manage_custom_format(rid: &str) -> String {
    json!({
        "action_id": ACTION_MANAGE_CUSTOM_FORMAT,
        "rid": rid
    })
    .to_string()
}

pub fn build_manage_number_format(rid: &str) -> String {
    json!({
        "action_id": ACTION_MANAGE_NUMBER_FORMAT,
        "rid": rid
    })
    .to_string()
}

// ─── Pivot Table operations ──────────────────────────────────────────────────

pub fn build_create_pivot_table_new_sheet(
    rid: &str,
    sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
) -> String {
    json!({
        "action_id": ACTION_CREATE_PIVOT_TABLE,
        "rid": rid,
        "source_range": [{
            "sheet_id": sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }],
        "is_new_sheet": true
    })
    .to_string()
}

pub fn build_create_pivot_table_at_dest(
    rid: &str,
    source_sheet_id: &str,
    start_row: i32,
    start_col: i32,
    end_row: i32,
    end_col: i32,
    dest_sheet_id: &str,
    dest_row: i32,
    dest_col: i32,
) -> String {
    json!({
        "action_id": ACTION_CREATE_PIVOT_TABLE,
        "rid": rid,
        "source_range": [{
            "sheet_id": source_sheet_id,
            "range_list": [build_range_object(start_row, start_col, end_row, end_col)]
        }],
        "is_new_sheet": false,
        "destination_range": [{
            "sheet_id": dest_sheet_id,
            "range_list": [build_range_object(dest_row, dest_col, dest_row, dest_col)]
        }]
    })
    .to_string()
}

pub fn build_delete_pivot_table(rid: &str, sheet_id: &str, pivot_id: &str) -> String {
    json!({
        "action_id": ACTION_DELETE_PIVOT_TABLE,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id
    })
    .to_string()
}

pub fn build_change_pivot_field_type(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
    destination_field_index: i32,
    destination_field_type: i32,
) -> String {
    json!({
        "action_id": ACTION_CHANGE_PIVOT_FIELD_TYPE,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type,
        "destination_field_index": destination_field_index,
        "destination_field_type": destination_field_type
    })
    .to_string()
}

pub fn build_select_pivot_field(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    header_index: i32,
    pivot_field_type: i32,
    field_index: i32,
) -> String {
    json!({
        "action_id": ACTION_SELECT_PIVOT_FIELD,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "header_index": header_index,
        "pivot_field_type": pivot_field_type,
        "field_index": field_index
    })
    .to_string()
}

pub fn build_pivot_table_info(rid: &str, sheet_id: &str, pivot_id: &str) -> String {
    json!({
        "action_id": ACTION_PIVOT_TABLE_INFO,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id
    })
    .to_string()
}

pub fn build_move_pivot_table(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    destination_sheet_id: &str,
    dest_row: i32,
    dest_col: i32,
) -> String {
    json!({
        "action_id": ACTION_MOVE_PIVOT_TABLE,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "destination_sheet_id": destination_sheet_id,
        "start_cell_index": build_active_cell(dest_row, dest_col)
    })
    .to_string()
}

pub fn build_refresh_pivot_table(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
) -> String {
    json!({
        "action_id": ACTION_REFRESH_PIVOT_TABLE,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "active_info": {
            "active_sheet_id": sheet_id,
            "active_cell": build_active_cell(0, 0)
        }
    })
    .to_string()
}

pub fn build_copy_pivot_table(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    destination_sheet_id: &str,
    dest_row: i32,
    dest_col: i32,
) -> String {
    json!({
        "action_id": ACTION_COPY_PIVOT_TABLE,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "destination_sheet_id": destination_sheet_id,
        "start_cell_index": build_active_cell(dest_row, dest_col)
    })
    .to_string()
}

pub fn build_edit_pivot_name(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    new_pivot_name: &str,
) -> String {
    json!({
        "action_id": ACTION_EDIT_PIVOT_NAME,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "new_pivot_name": new_pivot_name
    })
    .to_string()
}

/// Builds a fetch request for pivot table list using sheet_meta pivot_info (524288).
/// Response contains pivot_id, is_pivot_table_empty, pivot_range for each pivot on the sheet.
pub fn build_pivot_list_fetch(rid: &str, sheet_id: &str) -> String {
    // sheet_meta = pivot_info (bit 19 = 524288)
    json!({
        "rid": rid,
        "doc_meta": 2,
        "meta": [{"sheet_meta": 524288, "cell_meta": 0}],
        "ranges": [{"boundary": [0, 0, 262143, 4095], "sheet_id": sheet_id}]
    })
    .to_string()
}

// ─── Pivot Filter / Sort / Grouping ─────────────────────────────────────────

pub fn build_apply_pivot_filter_condition(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
    criteria_id: i32,
    sub_criteria_id: i32,
    val1: &str,
    val2: &str,
    value_field_index: i32,
) -> String {
    json!({
        "action_id": ACTION_APPLY_PIVOT_FILTER,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type,
        "is_selection_filter": false,
        "condition": {
            "criteria_id": criteria_id,
            "sub_criteria_id": sub_criteria_id,
            "val1": val1,
            "val2": val2,
            "value_field_index": value_field_index
        }
    })
    .to_string()
}

pub fn build_apply_pivot_filter_selection(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
    check_mark_vector: Vec<i32>,
) -> String {
    json!({
        "action_id": ACTION_APPLY_PIVOT_FILTER,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type,
        "is_selection_filter": true,
        "check_mark_vector": check_mark_vector
    })
    .to_string()
}

pub fn build_remove_pivot_filter(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
) -> String {
    json!({
        "action_id": ACTION_REMOVE_PIVOT_FILTER,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type
    })
    .to_string()
}

pub fn build_pivot_filter_info(
    rid: &str,
    sheet_id: &str,
    active_row: i32,
    active_column: i32,
) -> String {
    json!({
        "action_id": ACTION_PIVOT_FILTER_INFO,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_cell": {
            "active_row": active_row,
            "active_column": active_column
        }
    })
    .to_string()
}

pub fn build_apply_pivot_sort(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
    is_asc_order: bool,
    sort_aggregation_index: i32,
) -> String {
    json!({
        "action_id": ACTION_APPLY_PIVOT_SORT,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type,
        "is_asc_order": is_asc_order,
        "sort_aggregation_index": sort_aggregation_index
    })
    .to_string()
}

pub fn build_remove_pivot_sort(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
) -> String {
    json!({
        "action_id": ACTION_REMOVE_PIVOT_SORT,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type
    })
    .to_string()
}

pub fn build_apply_pivot_grouping(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
    minimum: f64,
    maximum: f64,
    range: f64,
    is_min_default: bool,
    is_max_default: bool,
) -> String {
    json!({
        "action_id": ACTION_APPLY_PIVOT_GROUPING,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type,
        "minimum": minimum,
        "maximum": maximum,
        "range": range,
        "is_min_default": is_min_default,
        "is_max_default": is_max_default
    })
    .to_string()
}

pub fn build_apply_pivot_date_grouping(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
    date_grouping_types: Vec<i32>,
    minimum: f64,
    maximum: f64,
    is_min_default: bool,
    is_max_default: bool,
    no_of_days: i32,
) -> String {
    json!({
        "action_id": ACTION_APPLY_PIVOT_DATE_GROUPING,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type,
        "date_grouping_types": date_grouping_types,
        "minimum": minimum,
        "maximum": maximum,
        "is_min_default": is_min_default,
        "is_max_default": is_max_default,
        "no_of_days": no_of_days
    })
    .to_string()
}

pub fn build_remove_group(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
) -> String {
    json!({
        "action_id": ACTION_REMOVE_GROUP,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type
    })
    .to_string()
}

pub fn build_remove_pivot_field(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    pivot_field_type: i32,
) -> String {
    json!({
        "action_id": ACTION_REMOVE_PIVOT_FIELD,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "pivot_field_type": pivot_field_type
    })
    .to_string()
}

pub fn build_modify_pivot_properties(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    pivot_property: i32,
    is_enabled: bool,
) -> String {
    json!({
        "action_id": ACTION_MODIFY_PIVOT_PROPERTIES,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "pivot_property": pivot_property,
        "is_enabled": is_enabled
    })
    .to_string()
}

pub fn build_modify_value_aggregation_type(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    summarise_by: i32,
) -> String {
    json!({
        "action_id": ACTION_MODIFY_VALUE_AGGREGATION_TYPE,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "summarise_by": summarise_by
    })
    .to_string()
}

pub fn build_modify_value_show_data_as(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    field_index: i32,
    show_data_as: i32,
) -> String {
    json!({
        "action_id": ACTION_MODIFY_VALUE_SHOW_DATA_AS,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "field_index": field_index,
        "show_data_as": show_data_as
    })
    .to_string()
}

pub fn build_change_pivot_table_source(
    rid: &str,
    sheet_id: &str,
    pivot_id: &str,
    destination_sheet_id: &str,
    start_row: i32,
    end_row: i32,
    start_column: i32,
    end_column: i32,
) -> String {
    json!({
        "action_id": ACTION_CHANGE_PIVOT_TABLE_SOURCE,
        "rid": rid,
        "sheet_id": sheet_id,
        "pivot_id": pivot_id,
        "destination_sheet_id": destination_sheet_id,
        "range": {
            "start_row": start_row,
            "end_row": end_row,
            "start_column": start_column,
            "end_column": end_column
        },
        "active_info": {
            "active_sheet_id": sheet_id,
            "active_cell": build_active_cell(0, 0)
        }
    })
    .to_string()
}

pub fn build_pivot_cell_info(
    rid: &str,
    sheet_id: &str,
    active_row: i32,
    active_column: i32,
) -> String {
    json!({
        "action_id": ACTION_PIVOT_CELL_INFO,
        "rid": rid,
        "sheet_id": sheet_id,
        "active_cell": {
            "active_row": active_row,
            "active_column": active_column
        }
    })
    .to_string()
}

pub fn build_refresh_pivot_table_on_load(rid: &str, sheet_id: &str) -> String {
    json!({
        "action_id": ACTION_REFRESH_PIVOT_TABLE_ON_LOAD,
        "rid": rid,
        "sheet_id": sheet_id,
        "is_forced": false,
        "active_info": {
            "active_sheet_id": sheet_id,
            "active_cell": build_active_cell(0, 0)
        }
    })
    .to_string()
}

// ─── Chart ───────────────────────────────────────────────────────────────────

pub fn build_recommend_chart(
    rid: &str,
    sheet_id: &str,
    range_list: Vec<serde_json::Value>,
) -> String {
    json!({
        "action_id": ACTION_RECOMMEND_CHART,
        "rid": rid,
        "sheet_id": sheet_id,
        "range_list": range_list
    })
    .to_string()
}

pub fn build_insert_chart(
    rid: &str,
    sheet_id: &str,
    chart_type: i32,
    chart_sub_type: Option<i32>,
    range_list: Vec<serde_json::Value>,
    start_x: i32,
    start_y: i32,
    end_x: i32,
    end_y: i32,
    position_type: i32,
    active_row: i32,
    active_col: i32,
) -> String {
    let mut payload = serde_json::Map::new();
    payload.insert("action_id".into(), json!(ACTION_INSERT_CHART));
    payload.insert("rid".into(), json!(rid));
    payload.insert("sheet_id".into(), json!(sheet_id));
    payload.insert("chart_type".into(), json!(chart_type));
    if let Some(sub_type) = chart_sub_type {
        payload.insert("chart_sub_type".into(), json!(sub_type));
    }
    payload.insert("range_list".into(), json!(range_list));
    payload.insert(
        "offset_position".into(),
        json!({
            "start_x": start_x as f64,
            "start_y": start_y as f64,
            "end_x": end_x as f64,
            "end_y": end_y as f64
        }),
    );
    payload.insert(
        "range_position".into(),
        json!({
            "start_row": active_row,
            "start_column": active_col,
            "end_row": active_row + 15,
            "end_column": active_col + 5
        }),
    );
    payload.insert("start_row".into(), json!(active_row));
    payload.insert("start_column".into(), json!(active_col));
    payload.insert("start_row_offset".into(), json!(start_x as f64));
    payload.insert("start_column_offset".into(), json!(start_y as f64));
    payload.insert("height".into(), json!((end_y - start_y).max(10) as f64));
    payload.insert("width".into(), json!((end_x - start_x).max(10) as f64));
    payload.insert("position_type".into(), json!(position_type));
    payload.insert(
        "active_info".into(),
        json!(build_active_info(sheet_id, active_row, active_col)),
    );

    serde_json::Value::Object(payload)
        .to_string()
}

pub fn build_delete_chart(
    rid: &str,
    sheet_id: &str,
    chart_id: &str,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_DELETE_CHART,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id": chart_id,
        "active_info": build_active_info(sheet_id, active_row, active_col)
    })
    .to_string()
}

pub fn build_update_chart_position(
    rid: &str,
    sheet_id: &str,
    chart_id: &str,
    start_x: i32,
    start_y: i32,
    end_x: i32,
    end_y: i32,
    position_type: i32,
) -> String {
    let (offset_pos, range_pos, top_start_row, top_start_col) = if position_type == 1 {
        // Range-based: coordinates are row/col values
        let range = serde_json::json!({
            "start_row": start_y,
            "start_column": start_x,
            "end_row": end_y,
            "end_column": end_x
        });
        let offset = serde_json::json!({
            "start_x": 0.0_f64,
            "start_y": 0.0_f64,
            "end_x": 0.0_f64,
            "end_y": 0.0_f64
        });
        (offset, range, start_y, start_x)
    } else {
        // Pixel-based: coordinates are pixel offsets
        let offset = serde_json::json!({
            "start_x": start_x as f64,
            "start_y": start_y as f64,
            "end_x": end_x as f64,
            "end_y": end_y as f64
        });
        let range = serde_json::json!({
            "start_row": 0,
            "start_column": 0,
            "end_row": 15,
            "end_column": 5
        });
        (offset, range, 0, 0)
    };
    json!({
        "action_id": ACTION_UPDATE_CHART_POSITION,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id": chart_id,
        "offset_position": offset_pos,
        "range_position": range_pos,
        "start_row": top_start_row,
        "start_column": top_start_col,
        "position_type": position_type,
        "active_info": build_active_info(sheet_id, top_start_row, top_start_col)
    })
    .to_string()
}

pub fn build_update_chart_type(
    rid: &str,
    sheet_id: &str,
    chart_id: &str,
    chart_type: i32,
    chart_sub_type: Option<i32>,
    active_row: i32,
    active_col: i32,
) -> String {
    let mut chart_properties = serde_json::Map::new();
    chart_properties.insert("chart_type".into(), json!(chart_type));
    if let Some(sub_type) = chart_sub_type {
        chart_properties.insert("chart_sub_type".into(), json!(sub_type));
    }

    json!({
        "action_id": ACTION_UPDATE_CHART_TYPE,
        "sub_action_id": 211,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id": chart_id,
        "chart_properties": chart_properties,
        "active_info": build_active_info(sheet_id, active_row, active_col)
    })
    .to_string()
}

pub fn build_manage_chart_with_range(
    rid: &str,
    sheet_id: &str,
    range_list: Vec<serde_json::Value>,
) -> String {
    json!({
        "action_id": ACTION_MANAGE_CHART_WITH_RANGE,
        "rid": rid,
        "sheet_id": sheet_id,
        "range_list": range_list
    })
    .to_string()
}

pub fn build_manage_chart_with_id(
    rid: &str,
    sheet_id: &str,
    chart_id_list: Vec<String>,
) -> String {
    json!({
        "action_id": ACTION_MANAGE_CHART_WITH_ID,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id_list": chart_id_list
    })
    .to_string()
}

pub fn build_customize_chart_property_two(
    rid: &str,
    sheet_id: &str,
    chart_id: &str,
    chart_properties: serde_json::Value,
) -> String {
    json!({
        "action_id": ACTION_CUSTOMIZE_CHART_PROPERTY_TWO,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id": chart_id,
        "chart_properties": chart_properties,
        "active_info": {
            "active_sheet_id": sheet_id,
            "active_chart_id": chart_id
        }
    })
    .to_string()
}

pub fn build_move_chart(
    rid: &str,
    sheet_id: &str,
    destination_sheet_id: &str,
    chart_id: &str,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_MOVE_CHART,
        "rid": rid,
        "sheet_id": sheet_id,
        "destination_sheet_id": destination_sheet_id,
        "chart_id": chart_id,
        "active_info": build_active_info(sheet_id, active_row, active_col)
    })
    .to_string()
}

pub fn build_clone_chart(
    rid: &str,
    sheet_id: &str,
    chart_id: &str,
    active_row: i32,
    active_col: i32,
) -> String {
    json!({
        "action_id": ACTION_CLONE_CHART,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id": chart_id,
        "active_info": build_active_info(sheet_id, active_row, active_col)
    })
    .to_string()
}

pub fn build_customize_chart_property_one(
    rid: &str,
    sheet_id: &str,
    chart_id: &str,
    sub_action_id: i32,
    chart_properties: serde_json::Value,
) -> String {
    json!({
        "action_id": ACTION_CUSTOMIZE_CHART_PROPERTY_ONE,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id": chart_id,
        "sub_action_id": sub_action_id,
        "chart_properties": chart_properties
    })
    .to_string()
}

pub fn build_customize_chart_with_subaction(
    rid: &str,
    sheet_id: &str,
    chart_id: &str,
    action_id: i32,
    sub_action_id: i32,
    chart_properties: serde_json::Value,
) -> String {
    json!({
        "action_id": action_id,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id": chart_id,
        "sub_action_id": sub_action_id,
        "chart_properties": chart_properties
    })
    .to_string()
}

pub fn build_manage_top_bottom(
    rid: &str,
    sheet_id: &str,
    chart_id: &str,
) -> String {
    json!({
        "action_id": ACTION_MANAGE_TOP_BOTTOM,
        "rid": rid,
        "sheet_id": sheet_id,
        "chart_id": chart_id,
        "active_info": build_active_info(sheet_id, 0, 0)
    })
    .to_string()
}
