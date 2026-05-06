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
