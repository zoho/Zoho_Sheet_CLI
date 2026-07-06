/// Parses JSON response strings from the engine into typed result objects.
/// Replaces the C# `CliEngineResponseParser` + `DispatcherLibrary.IJsonWrap`/`JsonFactory`.
use serde_json::Value;

// ─── Result types ────────────────────────────────────────────────────────────

#[derive(Debug, Default)]
pub struct WorkbookOpenResult {
    pub rid: Option<String>,
    pub workbook_name: Option<String>,
    pub sheet_count: i32,
    pub active_sheet_id: Option<String>,
    pub status_code: i32,
    pub status_message: Option<String>,
}

#[derive(Debug, Default, Clone)]
pub struct SheetInfo {
    pub sheet_name: String,
    pub index: i32,
    pub sheet_id: String,
}

#[derive(Debug, Default)]
pub struct CellValueResult {
    pub display_value: String,
    pub raw_value: String,
    pub formula: String,
    pub status_code: i32,
    pub status_message: Option<String>,
}

#[derive(Debug, Default)]
pub struct FormulaEvalResult {
    pub result_value: String,
    pub result_type: String,
    pub status_code: i32,
    pub status_message: Option<String>,
}

#[derive(Debug, Default)]
pub struct SetCellResult {
    pub computed_value: String,
    pub status_code: i32,
    pub status_message: Option<String>,
}

#[derive(Debug, Default)]
pub struct ExportResult {
    pub success: bool,
    pub exported_path: String,
    pub file_size_bytes: i64,
    pub status_code: i32,
    pub status_message: Option<String>,
}

#[derive(Debug, Default)]
pub struct EngineStatusResult {
    pub status_code: i32,
    pub status_message: Option<String>,
}

#[derive(Debug, Default)]
pub struct FindReplaceResult {
    pub status_code: i32,
    pub status_message: Option<String>,
    pub match_count: i32,
    pub found_row: i32,
    pub found_col: i32,
    pub has_meta: bool,
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

fn parse_root(json_str: &str) -> Option<Value> {
    serde_json::from_str(json_str).ok()
}

fn get_str(v: &Value, key: &str) -> Option<String> {
    v.get(key).and_then(|x| x.as_str()).map(|s| s.to_string())
}

fn get_int(v: &Value, key: &str) -> Option<i32> {
    v.get(key).and_then(|x| x.as_i64()).map(|n| n as i32)
}

fn get_i64(v: &Value, key: &str) -> Option<i64> {
    v.get(key).and_then(|x| x.as_i64())
}

fn get_bool(v: &Value, key: &str) -> Option<bool> {
    v.get(key).and_then(|x| x.as_bool())
}

fn get_obj<'a>(v: &'a Value, key: &str) -> Option<&'a Value> {
    v.get(key).filter(|x| x.is_object())
}

fn get_arr<'a>(v: &'a Value, key: &str) -> Option<&'a Vec<Value>> {
    v.get(key).and_then(|x| x.as_array())
}

fn parse_status(json: &Value) -> (i32, Option<String>) {
    if let Some(status) = get_obj(json, "response_status") {
        let code = get_int(status, "status_code").unwrap_or(-1);
        let msg = get_str(status, "status_message");
        (code, msg)
    } else {
        (-1, None)
    }
}

fn unwrap_response<'a>(json: &'a Value) -> &'a Value {
    get_obj(json, "response").unwrap_or(json)
}

// ─── Parsers ─────────────────────────────────────────────────────────────────

/// Parses the response from an import/create workbook operation.
pub fn parse_workbook_open(json_str: &str) -> Option<WorkbookOpenResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);

    let mut rid = None;
    let mut workbook_name = None;
    let mut sheet_count = 0i32;
    let mut active_sheet_id: Option<String> = None;

    if let Some(resp) = get_obj(&json, "response") {
        rid = get_str(resp, "rid");
        workbook_name = get_str(resp, "workbook_name");
        sheet_count = get_int(resp, "sheet_count").unwrap_or(0);

        // Try multiple locations for active sheet id
        active_sheet_id = get_str(resp, "active_sheet_id")
            .or_else(|| get_str(resp, "sheet_id"))
            .or_else(|| {
                get_obj(resp, "active_info")
                    .and_then(|ai| get_str(ai, "active_sheet_id"))
            });

        // Fallback: first entry in sheet_list
        if active_sheet_id.is_none() {
            if let Some(arr) = get_arr(resp, "sheet_list") {
                if let Some(first) = arr.first() {
                    if first.is_object() {
                        active_sheet_id = get_str(first, "sheet_id")
                            .or_else(|| get_str(first, "associated_name"));
                    } else if let Some(inner) = first.as_array() {
                        if let Some(s) = inner.first().and_then(|v| v.as_str()) {
                            active_sheet_id = Some(s.to_string());
                        }
                    }
                }
            }
        }

        // Fallback: worksheets.meta[0][0]
        if active_sheet_id.is_none() {
            if let Some(ws) = get_obj(resp, "worksheets") {
                if let Some(meta) = get_arr(ws, "meta") {
                    if let Some(first) = meta.first().and_then(|v| v.as_array()) {
                        if let Some(s) = first.first().and_then(|v| v.as_str()) {
                            active_sheet_id = Some(s.to_string());
                        }
                    }
                }
            }
        }
    }

    Some(WorkbookOpenResult {
        rid,
        workbook_name,
        sheet_count,
        active_sheet_id,
        status_code,
        status_message,
    })
}

/// Parses the response from a sheet list query.
pub fn parse_sheet_list(json_str: &str) -> Vec<SheetInfo> {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return Vec::new(),
    };

    let data = if let Some(resp) = get_obj(&json, "response") {
        resp
    } else {
        &json
    };

    let mut sheets = extract_sheets_from_data(data);

    if sheets.is_empty() {
        if let Some(wb_info) = get_obj(data, "workbook_info") {
            sheets = extract_sheets_from_data(wb_info);
        }
    }

    sheets
}

fn extract_sheets_from_data(data: &Value) -> Vec<SheetInfo> {
    let mut sheets = Vec::new();

    // Try "sheet_list"
    if let Some(arr) = get_arr(data, "sheet_list") {
        for (i, item) in arr.iter().enumerate() {
            if item.is_object() {
                let name = get_str(item, "sheet_name")
                    .or_else(|| get_str(item, "sheetName"))
                    .unwrap_or_else(|| format!("Sheet{}", i));
                let id = get_str(item, "sheet_id")
                    .or_else(|| get_str(item, "associated_name"))
                    .or_else(|| get_str(item, "sheetId"))
                    .or_else(|| get_str(item, "associatedName"))
                    .or_else(|| get_str(item, "id"))
                    .or_else(|| get_str(item, "sheet_name"))
                    .unwrap_or_else(|| i.to_string());
                let index = get_int(item, "index").unwrap_or(i as i32);
                sheets.push(SheetInfo {
                    sheet_name: name,
                    sheet_id: id,
                    index,
                });
            } else if let Some(row) = item.as_array() {
                if row.len() >= 2 {
                    sheets.push(SheetInfo {
                        sheet_id: row[0].as_str().unwrap_or("").to_string(),
                        sheet_name: row[1].as_str().unwrap_or("").to_string(),
                        index: row
                            .get(2)
                            .and_then(|v| v.as_i64())
                            .map(|n| n as i32)
                            .unwrap_or(i as i32),
                    });
                }
            }
        }
    }

    // Fallback: worksheets.meta
    if sheets.is_empty() {
        if let Some(ws) = get_obj(data, "worksheets") {
            if let Some(meta) = get_arr(ws, "meta") {
                for (i, item) in meta.iter().enumerate() {
                    if let Some(row) = item.as_array() {
                        if row.len() >= 3 {
                            sheets.push(SheetInfo {
                                sheet_id: row[0].as_str().unwrap_or("").to_string(),
                                sheet_name: row[1].as_str().unwrap_or("").to_string(),
                                index: row[2].as_i64().map(|n| n as i32).unwrap_or(i as i32),
                            });
                        }
                    }
                }
            }
        }
    }

    // Fallback: "sheets_list" (with underscore-s)
    if sheets.is_empty() {
        if let Some(arr) = get_arr(data, "sheets_list") {
            for (i, item) in arr.iter().enumerate() {
                if item.is_object() {
                    let id = get_str(item, "sheet_id")
                        .or_else(|| get_str(item, "associated_name"))
                        .unwrap_or_else(|| i.to_string());
                    let name = get_str(item, "sheet_name")
                        .unwrap_or_else(|| format!("Sheet{}", i));
                    let index = get_int(item, "tab_position")
                        .or_else(|| get_int(item, "index"))
                        .unwrap_or(i as i32);
                    sheets.push(SheetInfo {
                        sheet_id: id,
                        sheet_name: name,
                        index,
                    });
                } else if let Some(row) = item.as_array() {
                    if row.len() >= 2 {
                        sheets.push(SheetInfo {
                            sheet_id: row[0].as_str().unwrap_or("").to_string(),
                            sheet_name: row[1].as_str().unwrap_or("").to_string(),
                            index: row
                                .get(2)
                                .and_then(|v| v.as_i64())
                                .map(|n| n as i32)
                                .unwrap_or(i as i32),
                        });
                    }
                }
            }
        }
    }

    sheets
}

/// Parses a FetchJson cell data response.
/// Path: workbook_info → sheets_data[0] → cell_data → cell_info[0] → cell_value
pub fn parse_cell_fetch(json_str: &str) -> Option<CellValueResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);

    let mut display_value = String::new();
    let mut raw_value = String::new();
    let mut formula_value = String::new();

    if let Some(cv) = navigate_to_cell_value(&json) {
        display_value = get_str(cv, "display_value").unwrap_or_default();
        raw_value = get_str(cv, "actual_value").unwrap_or_default();
        formula_value = get_str(cv, "formula_value").unwrap_or_default();
    }

    Some(CellValueResult {
        display_value,
        raw_value,
        formula: formula_value,
        status_code,
        status_message,
    })
}

fn navigate_to_cell_value(json: &Value) -> Option<&Value> {
    let wb_info = get_obj(json, "workbook_info")?;

    let sheet_data = get_arr(wb_info, "sheets_data")
        .and_then(|a| a.first())
        .filter(|v| v.is_object())
        .or_else(|| {
            get_arr(wb_info, "sheet_data")
                .and_then(|a| a.first())
                .filter(|v| v.is_object())
        })?;

    let cell_data = get_obj(sheet_data, "cell_data").or_else(|| {
        get_arr(sheet_data, "cell_data")
            .and_then(|a| a.first())
            .filter(|v| v.is_object())
    })?;

    let cell_info = get_arr(cell_data, "cell_info")
        .and_then(|a| a.first())
        .filter(|v| v.is_object())?;

    get_obj(cell_info, "cell_value")
}

/// Parses a range fetch response into a grid of display values.
/// Returns a Vec of (row, col, display_value) tuples.
pub fn parse_range_cell_values(json_str: &str) -> Vec<(i32, i32, String)> {
    let mut results = Vec::new();
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return results,
    };

    let wb_info = match get_obj(&json, "workbook_info") {
        Some(v) => v,
        None => return results,
    };

    let sheet_data = get_arr(wb_info, "sheets_data")
        .and_then(|a| a.first())
        .filter(|v| v.is_object())
        .or_else(|| {
            get_arr(wb_info, "sheet_data")
                .and_then(|a| a.first())
                .filter(|v| v.is_object())
        });

    let sheet_data = match sheet_data {
        Some(v) => v,
        None => return results,
    };

    let cell_data = get_obj(sheet_data, "cell_data").or_else(|| {
        get_arr(sheet_data, "cell_data")
            .and_then(|a| a.first())
            .filter(|v| v.is_object())
    });

    let cell_data = match cell_data {
        Some(v) => v,
        None => return results,
    };

    if let Some(cell_info_arr) = get_arr(cell_data, "cell_info") {
        for ci in cell_info_arr {
            if !ci.is_object() {
                continue;
            }
            let row = get_int(ci, "row").unwrap_or(-1);
            let col = get_int(ci, "col").unwrap_or(-1);
            if row < 0 || col < 0 {
                continue;
            }
            let display = get_obj(ci, "cell_value")
                .and_then(|cv| get_str(cv, "display_value"))
                .unwrap_or_default();
            results.push((row, col, display));
        }
    }

    results
}

/// Parses the response from a get-cell-info (ActiveCellInfo) operation.
pub fn parse_cell_value(json_str: &str) -> Option<CellValueResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);
    let data = unwrap_response(&json);

    Some(CellValueResult {
        display_value: get_str(data, "display_value").unwrap_or_default(),
        raw_value: get_str(data, "raw_value").unwrap_or_default(),
        formula: get_str(data, "formula").unwrap_or_default(),
        status_code,
        status_message,
    })
}

/// Parses the response from a formula evaluation operation.
pub fn parse_formula_eval(json_str: &str) -> Option<FormulaEvalResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);
    let data = unwrap_response(&json);

    Some(FormulaEvalResult {
        result_value: get_str(data, "result_value").unwrap_or_default(),
        result_type: get_str(data, "result_type").unwrap_or_else(|| "unknown".to_string()),
        status_code,
        status_message,
    })
}

/// Parses the response from a set-cell-value operation.
pub fn parse_set_cell_value(json_str: &str) -> Option<SetCellResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);
    let data = unwrap_response(&json);

    Some(SetCellResult {
        computed_value: get_str(data, "computed_value").unwrap_or_default(),
        status_code,
        status_message,
    })
}

/// Parses the response from an export workbook operation.
pub fn parse_export(json_str: &str) -> Option<ExportResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);
    let data = unwrap_response(&json);

    Some(ExportResult {
        success: get_bool(data, "success").unwrap_or(false),
        exported_path: get_str(data, "exported_path").unwrap_or_default(),
        file_size_bytes: get_i64(data, "file_size_bytes").unwrap_or(0),
        status_code,
        status_message,
    })
}

/// Parses a generic engine action response for status information.
pub fn parse_status_response(json_str: &str) -> EngineStatusResult {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => {
            return EngineStatusResult {
                status_code: -1,
                status_message: Some(
                    if json_str.is_empty() {
                        "Empty response from engine"
                    } else {
                        "Invalid JSON response from engine"
                    }
                    .to_string(),
                ),
            }
        }
    };

    if let Some(status) = get_obj(&json, "response_status") {
        EngineStatusResult {
            status_code: get_int(status, "status_code").unwrap_or(-1),
            status_message: get_str(status, "status_message"),
        }
    } else {
        EngineStatusResult {
            status_code: get_int(&json, "status_code").unwrap_or(-1),
            status_message: get_str(&json, "status_message"),
        }
    }
}

/// Extracts sheet_id from an engine action response (e.g. add_sheet).
/// Looks in common response locations: response.sheet_id, top-level sheet_id, meta.
pub fn extract_sheet_id_from_response(json_str: &str) -> Option<String> {
    let json = parse_root(json_str)?;
    // Try response -> sheet_id
    if let Some(resp) = get_obj(&json, "response") {
        if let Some(id) = get_str(resp, "sheet_id") {
            return Some(id);
        }
    }
    // Try top-level sheet_id
    if let Some(id) = get_str(&json, "sheet_id") {
        return Some(id);
    }
    // Try meta -> sheet_id or active_sheet_id
    if let Some(meta) = get_obj(&json, "meta") {
        if let Some(id) = get_str(meta, "sheet_id").or_else(|| get_str(meta, "active_sheet_id")) {
            return Some(id);
        }
    }
    None
}

/// Parses the response from a find or replace operation.
pub fn parse_find_replace(json_str: &str) -> FindReplaceResult {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => {
            return FindReplaceResult {
                status_code: -1,
                status_message: Some(
                    if json_str.is_empty() {
                        "Empty response"
                    } else {
                        "Invalid JSON"
                    }
                    .to_string(),
                ),
                ..Default::default()
            }
        }
    };

    let (status_code, status_message) = parse_status(&json);
    let data = unwrap_response(&json);

    FindReplaceResult {
        status_code,
        status_message,
        match_count: get_int(data, "number_of_occurrences").unwrap_or(0),
        found_row: get_int(data, "active_row").unwrap_or(-1),
        found_col: get_int(data, "active_column").unwrap_or(-1),
        has_meta: json.get("meta").is_some(),
    }
}

/// Helper: returns true when the status_code == 100 (engine success).
pub fn is_success(status_code: i32) -> bool {
    status_code == 100
}

/// Parses the response from an insert_table operation.
/// Returns (status_code, status_message, table_id).
pub fn parse_insert_table(json_str: &str) -> (i32, Option<String>, Option<String>) {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return (-1, Some("Invalid JSON".to_string()), None),
    };
    let (status_code, status_message) = parse_status(&json);
    let table_id = get_obj(&json, "response")
        .and_then(|r| get_str(r, "table_id"));
    (status_code, status_message, table_id)
}

/// Parses the response from a select_table_range operation.
/// Returns (status_code, is_contain_headers, range).
pub fn parse_select_table_range(
    json_str: &str,
) -> (i32, Option<String>, bool, Option<(i32, i32, i32, i32)>) {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return (-1, Some("Invalid JSON".to_string()), false, None),
    };
    let (status_code, status_message) = parse_status(&json);
    let resp = get_obj(&json, "response");
    let has_headers = resp
        .and_then(|r| get_bool(r, "is_contain_headers"))
        .unwrap_or(false);
    let range = resp.and_then(|r| get_obj(r, "range")).map(|rng| {
        (
            get_int(rng, "start_row").unwrap_or(0),
            get_int(rng, "start_column").unwrap_or(0),
            get_int(rng, "end_row").unwrap_or(0),
            get_int(rng, "end_column").unwrap_or(0),
        )
    });
    (status_code, status_message, has_headers, range)
}

/// Result type for manage_table response.
#[derive(Debug, Default)]
pub struct ManageTableResult {
    pub status_code: i32,
    pub status_message: Option<String>,
    pub table_name: String,
    pub sheet_id: String,
    pub source_start_row: i32,
    pub source_start_col: i32,
    pub source_end_row: i32,
    pub source_end_col: i32,
    pub is_header_row: bool,
    pub is_total_row: bool,
    pub is_first_column: bool,
    pub is_last_column: bool,
    pub is_banded_row: bool,
    pub is_banded_column: bool,
    pub is_show_filter_button: bool,
    pub table_style_type: String,
    pub table_color_pattern: String,
    pub column_headers: Vec<String>,
}

/// Parses the response from a manage_table operation.
pub fn parse_manage_table(json_str: &str) -> Option<ManageTableResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);
    let resp = get_obj(&json, "response")?;

    let table_name = get_str(resp, "table_name").unwrap_or_default();
    let sheet_id = get_str(resp, "sheet_id").unwrap_or_default();

    let (sr, sc, er, ec) = if let Some(src) = get_obj(resp, "table_source") {
        (
            get_int(src, "start_row").unwrap_or(0),
            get_int(src, "start_column").unwrap_or(0),
            get_int(src, "end_row").unwrap_or(0),
            get_int(src, "end_column").unwrap_or(0),
        )
    } else {
        (0, 0, 0, 0)
    };

    let style_type = get_obj(resp, "table_style")
        .and_then(|s| get_str(s, "table_style_type"))
        .unwrap_or_default();
    let color_pattern = get_obj(resp, "table_style")
        .and_then(|s| get_str(s, "table_color_pattern"))
        .unwrap_or_default();

    let mut headers = Vec::new();
    if let Some(arr) = get_arr(resp, "table_column_headers") {
        for item in arr {
            if let Some(s) = item.as_str() {
                headers.push(s.to_string());
            }
        }
    }

    Some(ManageTableResult {
        status_code,
        status_message,
        table_name,
        sheet_id,
        source_start_row: sr,
        source_start_col: sc,
        source_end_row: er,
        source_end_col: ec,
        is_header_row: get_bool(resp, "is_header_row").unwrap_or(false),
        is_total_row: get_bool(resp, "is_total_row").unwrap_or(false),
        is_first_column: get_bool(resp, "is_first_column").unwrap_or(false),
        is_last_column: get_bool(resp, "is_last_column").unwrap_or(false),
        is_banded_row: get_bool(resp, "is_banded_row").unwrap_or(false),
        is_banded_column: get_bool(resp, "is_banded_column").unwrap_or(false),
        is_show_filter_button: get_bool(resp, "is_show_filter_button").unwrap_or(false),
        table_style_type: style_type,
        table_color_pattern: color_pattern,
        column_headers: headers,
    })
}

// ─── Table list parsing ──────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct TableListEntry {
    pub table_id: String,
    pub start_row: i32,
    pub start_col: i32,
    pub end_row: i32,
    pub end_col: i32,
}

/// Parses table entries from a sheet meta fetch response (filter 2048).
/// The filter objects contain `table_id` and `filter_range` when the filter belongs to a table.
/// Response structure: workbook_info -> sheets_data[] -> filter[] -> table_id + filter_range
pub fn parse_table_list(json_str: &str) -> Vec<TableListEntry> {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return Vec::new(),
    };

    let mut entries = Vec::new();

    // Primary path: workbook_info -> sheets_data -> filter
    if let Some(wb_info) = json.get("workbook_info") {
        if let Some(sheets) = wb_info.get("sheets_data").and_then(|v| v.as_array()) {
            for sheet in sheets {
                collect_table_entries_from(sheet, &mut entries);
            }
        }
    }

    // Fallback: try response -> worksheets or top-level
    if entries.is_empty() {
        let data = unwrap_response(&json);
        if let Some(ws) = data.get("worksheets").or_else(|| data.get("worksheet")) {
            collect_table_entries_from(ws, &mut entries);
        } else {
            collect_table_entries_from(data, &mut entries);
        }
    }

    entries.sort_by(|a, b| a.table_id.cmp(&b.table_id));
    entries.dedup_by(|a, b| a.table_id == b.table_id);
    entries
}

fn extract_table_entry(item: &Value) -> Option<TableListEntry> {
    let tid = item.get("table_id").and_then(|v| v.as_str())?;
    if tid.is_empty() {
        return None;
    }
    let range = item.get("filter_range");
    let (sr, sc, er, ec) = if let Some(r) = range {
        (
            get_int(r, "start_row").unwrap_or(0),
            get_int(r, "start_column").unwrap_or(0),
            get_int(r, "end_row").unwrap_or(0),
            get_int(r, "end_column").unwrap_or(0),
        )
    } else {
        (0, 0, 0, 0)
    };
    Some(TableListEntry {
        table_id: tid.to_string(),
        start_row: sr,
        start_col: sc,
        end_row: er,
        end_col: ec,
    })
}

fn collect_table_entries_from(val: &Value, out: &mut Vec<TableListEntry>) {
    // Check "filter_info" (array of filter objects)
    if let Some(arr) = val.get("filter_info").and_then(|v| v.as_array()) {
        for item in arr {
            if let Some(entry) = extract_table_entry(item) {
                out.push(entry);
            }
        }
    }
    // Check "filter" (single or array)
    if let Some(filter_val) = val.get("filter") {
        if let Some(arr) = filter_val.as_array() {
            for item in arr {
                if let Some(entry) = extract_table_entry(item) {
                    out.push(entry);
                }
            }
        } else if let Some(entry) = extract_table_entry(filter_val) {
            out.push(entry);
        }
    }
    // Recurse into "meta" array entries (sheet-level data)
    if let Some(meta_arr) = val.get("meta").and_then(|v| v.as_array()) {
        for entry in meta_arr {
            collect_table_entries_from(entry, out);
        }
    }
    // Recurse into array values at top level
    if val.is_array() {
        if let Some(arr) = val.as_array() {
            for entry in arr {
                collect_table_entries_from(entry, out);
            }
        }
    }
    // Check inside objects with sheet data
    if let Some(obj) = val.as_object() {
        for (_key, child) in obj {
            if child.is_object() {
                if child.get("filter_info").is_some() || child.get("filter").is_some() {
                    collect_table_entries_from(child, out);
                }
            }
        }
    }
}

// ─── Pivot Table response parsers ────────────────────────────────────────────

/// Parses the response from a create_pivot_table operation.
/// Returns (status_code, status_message, pivot_id).
pub fn parse_create_pivot_table(json_str: &str) -> (i32, Option<String>, Option<String>) {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return (-1, Some("Invalid JSON".to_string()), None),
    };
    let (status_code, status_message) = parse_status(&json);
    let pivot_id = get_obj(&json, "response")
        .and_then(|r| get_str(r, "pivot_id"));
    (status_code, status_message, pivot_id)
}

/// Parses the response from a pivot_table_info operation.
#[derive(Debug, Default)]
pub struct PivotTableInfoResult {
    pub status_code: i32,
    pub status_message: Option<String>,
    pub pivot_id: String,
    pub pivot_name: String,
    pub sheet_id: String,
    pub source_range: String,
    pub headers: Vec<(String, String)>, // (header_name, data_type)
}

pub fn parse_pivot_table_info(json_str: &str) -> Option<PivotTableInfoResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);
    let resp = get_obj(&json, "response")?;

    // Engine uses "pivot_table_name" (not "pivot_name") and may not echo back pivot_id
    let pivot_name = get_str(resp, "pivot_table_name")
        .or_else(|| get_str(resp, "pivot_name"))
        .unwrap_or_default();
    let pivot_id = get_str(resp, "pivot_id").unwrap_or_default();
    let source_range = get_str(resp, "source_range").unwrap_or_default();

    // Parse header info list
    let mut headers = Vec::new();
    if let Some(builder) = resp.get("pivot_table_builder_info") {
        if let Some(list) = builder.get("header_info_list").and_then(|v| v.as_array()) {
            for item in list {
                let name = item.get("header_name").and_then(|v| v.as_str()).unwrap_or("").to_string();
                let dtype = item.get("data_type").and_then(|v| v.as_str()).unwrap_or("").to_string();
                headers.push((name, dtype));
            }
        }
    }

    Some(PivotTableInfoResult {
        status_code,
        status_message,
        pivot_id,
        pivot_name,
        sheet_id: get_str(resp, "sheet_id").unwrap_or_default(),
        source_range,
        headers,
    })
}

#[derive(Debug, Default)]
pub struct PivotFilterCondition {
    pub criteria_id: i32,
    pub sub_criteria_id: i32,
    pub val1: String,
    pub val2: String,
}

#[derive(Debug, Default)]
pub struct PivotFilterInfoResult {
    pub status_code: i32,
    pub status_message: Option<String>,
    pub pivot_id: String,
    pub active_filter_type: String,
    pub label_field_name: String,
    pub label_field_index: i32,
    pub label_field_type: String,
    pub value_field_info_list: Vec<String>,
    pub column_data: Vec<String>,
    pub check_mark_vector: Vec<i32>,
    pub custom_filter_value_field_index: i32,
    pub condition: Option<PivotFilterCondition>,
}

pub fn parse_pivot_filter_info(json_str: &str) -> Option<PivotFilterInfoResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);
    let resp = unwrap_response(&json);

    let value_field_info_list = resp
        .get("value_field_info_list")
        .and_then(|v| v.as_array())
        .map(|arr| {
            arr.iter()
                .filter_map(|x| x.as_str().map(|s| s.to_string()))
                .collect::<Vec<String>>()
        })
        .unwrap_or_default();

    let column_data = resp
        .get("column_data")
        .and_then(|v| v.as_array())
        .map(|arr| {
            arr.iter()
                .map(|x| {
                    x.as_str()
                        .map(|s| s.to_string())
                        .unwrap_or_else(|| x.to_string())
                })
                .collect::<Vec<String>>()
        })
        .unwrap_or_default();

    let check_mark_vector = resp
        .get("check_mark_vector")
        .and_then(|v| v.as_array())
        .map(|arr| {
            arr.iter()
                .filter_map(|x| x.as_i64().map(|n| n as i32))
                .collect::<Vec<i32>>()
        })
        .unwrap_or_default();

    let condition = resp.get("condition").and_then(|c| {
        if !c.is_object() {
            return None;
        }
        Some(PivotFilterCondition {
            criteria_id: get_int(c, "criteria_id").unwrap_or(-1),
            sub_criteria_id: get_int(c, "sub_criteria_id").unwrap_or(-1),
            val1: get_str(c, "val1").unwrap_or_default(),
            val2: get_str(c, "val2").unwrap_or_default(),
        })
    });

    Some(PivotFilterInfoResult {
        status_code,
        status_message,
        pivot_id: get_str(resp, "pivot_id").unwrap_or_default(),
        active_filter_type: get_str(resp, "active_filter_type").unwrap_or_default(),
        label_field_name: get_str(resp, "label_field_name").unwrap_or_default(),
        label_field_index: get_int(resp, "label_field_index").unwrap_or(-1),
        label_field_type: get_str(resp, "label_field_type").unwrap_or_default(),
        value_field_info_list,
        column_data,
        check_mark_vector,
        custom_filter_value_field_index: get_int(resp, "custom_filter_value_field_index").unwrap_or(-1),
        condition,
    })
}

#[derive(Debug, Default)]
pub struct PivotCellInfoResult {
    pub status_code: i32,
    pub status_message: Option<String>,
    pub pivot_id: String,
}

pub fn parse_pivot_cell_info(json_str: &str) -> Option<PivotCellInfoResult> {
    let json = parse_root(json_str)?;
    let (status_code, status_message) = parse_status(&json);
    let resp = unwrap_response(&json);

    Some(PivotCellInfoResult {
        status_code,
        status_message,
        pivot_id: get_str(resp, "pivot_id").unwrap_or_default(),
    })
}

// ─── Pivot Table List parsing (from sheet_meta fetch with bitmask 524288) ────

#[derive(Debug, Clone)]
pub struct PivotListEntry {
    pub pivot_id: String,
    pub is_empty: bool,
    pub start_row: i32,
    pub start_col: i32,
    pub end_row: i32,
    pub end_col: i32,
}

/// Parses pivot table entries from a sheet meta fetch response (pivot_info 524288).
/// Response structure: workbook_info -> sheets_data[] -> pivot_info[] -> pivot_id + pivot_range
pub fn parse_pivot_list(json_str: &str) -> Vec<PivotListEntry> {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return Vec::new(),
    };

    let mut entries = Vec::new();

    // Primary path: workbook_info -> sheets_data -> pivot_info
    if let Some(wb_info) = json.get("workbook_info") {
        if let Some(sheets) = wb_info.get("sheets_data").and_then(|v| v.as_array()) {
            for sheet in sheets {
                collect_pivot_entries_from(sheet, &mut entries);
            }
        }
    }

    // Fallback: try response or top-level
    if entries.is_empty() {
        let data = unwrap_response(&json);
        if let Some(ws) = data.get("worksheets").or_else(|| data.get("worksheet")) {
            collect_pivot_entries_from(ws, &mut entries);
        } else {
            collect_pivot_entries_from(data, &mut entries);
        }
    }

    entries.sort_by(|a, b| a.pivot_id.cmp(&b.pivot_id));
    entries.dedup_by(|a, b| a.pivot_id == b.pivot_id);
    entries
}

fn extract_pivot_entry(item: &Value) -> Option<PivotListEntry> {
    let pid = item.get("pivot_id").and_then(|v| v.as_str())?;
    if pid.is_empty() {
        return None;
    }
    let is_empty = item
        .get("is_pivot_table_empty")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let range = item.get("pivot_range");
    let (sr, sc, er, ec) = if let Some(r) = range {
        (
            get_int(r, "start_row").unwrap_or(0),
            get_int(r, "start_column").unwrap_or(0),
            get_int(r, "end_row").unwrap_or(0),
            get_int(r, "end_column").unwrap_or(0),
        )
    } else {
        (0, 0, 0, 0)
    };
    Some(PivotListEntry {
        pivot_id: pid.to_string(),
        is_empty,
        start_row: sr,
        start_col: sc,
        end_row: er,
        end_col: ec,
    })
}

fn collect_pivot_entries_from(val: &Value, out: &mut Vec<PivotListEntry>) {
    // Check "pivot_table" (engine returns this key) and "pivot_info" (docs name)
    for key in &["pivot_table", "pivot_info"] {
        if let Some(arr) = val.get(*key).and_then(|v| v.as_array()) {
            for item in arr {
                if let Some(entry) = extract_pivot_entry(item) {
                    out.push(entry);
                }
            }
        }
    }
    // Recurse into "meta" array entries (sheet-level data)
    if let Some(meta_arr) = val.get("meta").and_then(|v| v.as_array()) {
        for meta_item in meta_arr {
            collect_pivot_entries_from(meta_item, out);
        }
    }
}

// ─── Chart Parsers ───────────────────────────────────────────────────────────

/// Parses the response from an insert_chart operation.
/// Returns (status_code, status_message, chart_id).
pub fn parse_insert_chart(json_str: &str) -> (i32, Option<String>, Option<String>) {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return (-1, Some("Invalid JSON".to_string()), None),
    };
    let (status_code, status_message) = parse_status(&json);
    let chart_id = get_obj(&json, "response")
        .and_then(|r| get_str(r, "inserted_chart_id").or_else(|| get_str(r, "chart_id")));
    (status_code, status_message, chart_id)
}

/// Parses the response from a clone_chart operation.
/// Returns (status_code, status_message, chart_id).
pub fn parse_clone_chart(json_str: &str) -> (i32, Option<String>, Option<String>) {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return (-1, Some("Invalid JSON".to_string()), None),
    };
    let (status_code, status_message) = parse_status(&json);
    let chart_id = get_obj(&json, "response")
        .and_then(|r| get_str(r, "chart_id"));
    (status_code, status_message, chart_id)
}

/// Chart info returned by manage operations.
#[derive(Debug, Default, Clone)]
pub struct ChartInfo {
    pub chart_id: String,
    pub chart_title: String,
    pub chart_type: i32,
    pub chart_sub_type: i32,
    pub start_x: i32,
    pub start_y: i32,
    pub end_x: i32,
    pub end_y: i32,
    pub position_type: i32,
}

/// Parses the response from manage_chart operations (by range or ID).
/// Returns (status_code, status_message, chart_info_list).
pub fn parse_manage_chart(json_str: &str) -> (i32, Option<String>, Vec<ChartInfo>) {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return (-1, Some("Invalid JSON".to_string()), Vec::new()),
    };
    let (status_code, status_message) = parse_status(&json);
    let mut charts = Vec::new();

    if let Some(resp) = get_obj(&json, "response") {
        // The API returns "manage_chart" array; also support "chart_info_list" for flexibility
        let arr_opt = get_arr(resp, "manage_chart").or_else(|| get_arr(resp, "chart_info_list"));
        if let Some(arr) = arr_opt {
            for item in arr {
                let offset = item.get("offset_position");
                let title = item.get("chart_title")
                    .and_then(|ct| get_str(ct, "title_string"))
                    .unwrap_or_default();
                charts.push(ChartInfo {
                    chart_id: get_str(item, "chart_id").unwrap_or_default(),
                    chart_title: title,
                    chart_type: get_int(item, "chart_type").unwrap_or(-1),
                    chart_sub_type: get_int(item, "chart_sub_type").unwrap_or(-1),
                    start_x: offset.and_then(|o| get_int(o, "start_x")).unwrap_or(0),
                    start_y: offset.and_then(|o| get_int(o, "start_y")).unwrap_or(0),
                    end_x: offset.and_then(|o| get_int(o, "end_x")).unwrap_or(0),
                    end_y: offset.and_then(|o| get_int(o, "end_y")).unwrap_or(0),
                    position_type: get_int(item, "position_type").unwrap_or(0),
                });
            }
        }
    }

    (status_code, status_message, charts)
}

/// Recommendation entry from recommend_chart.
#[derive(Debug, Default)]
pub struct ChartRecommendation {
    pub chart_type: i32,
    pub chart_sub_type: i32,
}

/// Parses the response from a recommend_chart operation.
/// Returns (status_code, status_message, recommendations).
pub fn parse_recommend_chart(json_str: &str) -> (i32, Option<String>, Vec<ChartRecommendation>) {
    let json = match parse_root(json_str) {
        Some(v) => v,
        None => return (-1, Some("Invalid JSON".to_string()), Vec::new()),
    };
    let (status_code, status_message) = parse_status(&json);
    let mut recs = Vec::new();

    if let Some(resp) = get_obj(&json, "response") {
        let list_keys = [
            "chart_list",
            "recommend_chart_list",
            "recommended_chart_list",
            "recommend_list",
            "chart_recommend_list",
            "chart_type_list",
            "chart_recommendation_list",
        ];

        for key in list_keys {
            if let Some(arr) = get_arr(resp, key) {
                let mut parsed_from_key = Vec::new();
                for item in arr {
                    if let Some(obj_type) = get_int(item, "chart_type").or_else(|| get_int(item, "type")) {
                        let sub_type = get_int(item, "chart_sub_type")
                            .or_else(|| get_int(item, "sub_type"))
                            .or_else(|| get_int(item, "chart_subtype"))
                            .unwrap_or(-1);
                        parsed_from_key.push(ChartRecommendation {
                            chart_type: obj_type,
                            chart_sub_type: sub_type,
                        });
                    } else if let Some(type_int) = item.as_i64() {
                        // Some engines may return a plain list of chart-type enums.
                        parsed_from_key.push(ChartRecommendation {
                            chart_type: type_int as i32,
                            chart_sub_type: -1,
                        });
                    }
                }

                if !parsed_from_key.is_empty() {
                    recs = parsed_from_key;
                    break;
                }
            }
        }
    }

    (status_code, status_message, recs)
}
