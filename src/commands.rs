/// All CLI command implementations — one per verb.
///
/// Instead of the C# DI-based ICliCommand pattern, Rust uses a dispatch function
/// that takes shared references to the engine, session, and parser.
use std::fs;
use std::path::{Path, PathBuf};

use serde_json;

use crate::engine::ffi::EngineHandle;
use crate::engine::request_builder as rb;
use crate::engine::response_parser as rp;
use crate::output;
use crate::session::CliSession;
use crate::util::cell_ref;
use crate::util::numformat;
use crate::util::date_serial;

/// Maximum allowed single argument length to guard the native engine.
const MAX_INPUT_ARG_LEN: usize = 32_768;
/// Max rows / columns for batch insert/delete.
const MAX_ROW_COL_COUNT: i32 = 10_000;
/// Supported import extensions.
const SUPPORTED_EXTENSIONS: &[&str] = &[".xlsx", ".csv", ".tsv"];
// ─── Dispatch ────────────────────────────────────────────────────────────────

/// Top-level command dispatch. Returns `true` for normal continuation,
/// `false` to exit the REPL.
pub fn dispatch(
    tokens: &[&str],
    engine: &EngineHandle,
    session: &mut CliSession,
) -> bool {
    if tokens.is_empty() {
        return true;
    }

    for t in tokens {
        if t.len() > MAX_INPUT_ARG_LEN {
            output::error(&format!(
                "Input too long ({} chars). Maximum allowed is {}.",
                t.len(),
                MAX_INPUT_ARG_LEN
            ));
            return true;
        }
    }

    let verb = tokens[0].to_lowercase();
    let args: Vec<&str> = tokens[1..].to_vec();

    match verb.as_str() {
        "open" => cmd_open(&args, engine, session),
        "close" => cmd_close(&args, engine, session),
        "save" => cmd_save(&args, engine, session),
        "cell" => cmd_cell(&args, engine, session),
        "worksheet" => cmd_sheet(&args, engine, session),
        "row" => cmd_row(&args, engine, session),
        "col" => cmd_col(&args, engine, session),
        "copy" => cmd_copy(&args, engine, session),
        "move" => cmd_move(&args, engine, session),
        "clipboardcopy" => cmd_clipboard_copy(&args, engine, session),
        "find" => cmd_find(&args, engine, session),
        "replace" => cmd_replace(&args, engine, session),
        "sort" => cmd_sort(&args, engine, session),
        "filter" => cmd_filter(&args, engine, session),
        "merge" => cmd_merge(&args, engine, session),
        "clear" => cmd_clear(&args, engine, session),
        "undo" => cmd_undo(engine, session),
        "redo" => cmd_redo(engine, session),
        "freeze" => cmd_freeze(&args, engine, session),
        "unfreeze" => cmd_unfreeze(engine, session),
        "name" => cmd_name(&args, engine, session),
        "table" => cmd_table(&args, engine, session),
        "pivot" => cmd_pivot(&args, engine, session),
        "chart" => cmd_chart(&args, engine, session),
        "format" => cmd_format(&args, engine, session),
        "help" => { print_help(&args); }
        _ => {
            output::error(&format!(
                "Unknown command: '{}'. Type 'help' for available commands.",
                verb
            ));
        }
    }
    true
}

// ─── Open ────────────────────────────────────────────────────────────────────

fn cmd_open(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: open <filepath> OR open --new <docname>");
        return;
    }

    if args[0].eq_ignore_ascii_case("--new") {
        if args.len() < 2 {
            output::error("Usage: open --new <docname>");
            return;
        }
        create_new_workbook(args[1], engine, session);
    } else {
        open_existing_file(args[0], engine, session);
    }
}

fn open_existing_file(file_path: &str, engine: &EngineHandle, session: &mut CliSession) {
    let full_path = match std::fs::canonicalize(file_path) {
        Ok(p) => p,
        Err(_) => {
            output::error(&format!("File not found: '{}'", file_path));
            return;
        }
    };

    if !full_path.exists() {
        output::error(&format!(
            "File not found: '{}'",
            full_path.file_name().unwrap_or_default().to_string_lossy()
        ));
        return;
    }

    let ext = full_path
        .extension()
        .map(|e| format!(".{}", e.to_string_lossy().to_lowercase()))
        .unwrap_or_default();
    if !SUPPORTED_EXTENSIONS.contains(&ext.as_str()) {
        output::error(&format!(
            "Unsupported file type: '{}'. Supported: .xlsx, .csv, .tsv",
            ext
        ));
        return;
    }

    // Check if already open
    if session.is_active() {
        if let Some(ref fp) = session.file_path {
            if fp.eq_ignore_ascii_case(&full_path.to_string_lossy()) {
                output::info("Already open. Use 'switch' to change sheets.");
                return;
            }
        }
    }

    // Copy to engine working dir to avoid locking the original
    let engine_dir = crate::engine::initializer::get_engine_resources_dir();
    let _ = fs::create_dir_all(&engine_dir);

    let file_name = full_path
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string();
    let base_name = full_path
        .file_stem()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string();
    let unique_suffix = &uuid::Uuid::new_v4().to_string()[..8];
    let working_copy = PathBuf::from(&engine_dir)
        .join(format!("{}_{}{}", base_name, unique_suffix, ext));

    if let Err(e) = fs::copy(&full_path, &working_copy) {
        output::error(&format!("Failed to copy file: {}", e));
        return;
    }

    let open_file_type = match ext.as_str() {
        ".csv" => Some(0),
        ".tsv" => Some(1),
        _ => None, // xlsx — engine default (2)
    };
    let request = rb::build_open_workbook(&working_copy.to_string_lossy(), open_file_type);
    match engine.process_request_json(&request) {
        Ok(response) => {
            if let Some(result) = rp::parse_workbook_open(&response) {
                if result.rid.is_none() || result.rid.as_deref() == Some("") {
                    output::error(&format!("Engine failed to open '{}'.", file_name));
                    return;
                }
                let rid = result.rid.unwrap();
                session.rid = Some(rid.clone());
                session.workbook_name = result
                    .workbook_name
                    .or_else(|| {
                        full_path
                            .file_stem()
                            .map(|s| s.to_string_lossy().to_string())
                    });
                session.sheet_count = if result.sheet_count > 0 {
                    result.sheet_count as usize
                } else {
                    1
                };
                session.active_sheet_index = 0;
                session.file_path = Some(full_path.to_string_lossy().to_string());

                populate_sheet_names(engine, session, &response, result.active_sheet_id.as_deref());

                output::success(&format!("Opened: {}", file_name));
                output::key_value("Workbook ID", session.rid.as_deref().unwrap_or(""), 2);
                output::key_value("Sheets", &format_sheet_summary(session), 2);
                output::key_value("Mode", "Offline", 2);
            } else {
                output::error(&format!("Engine failed to open '{}'.", file_name));
            }
        }
        Err(e) => output::error(&format!("Engine init failed: {}", e)),
    }
}

fn create_new_workbook(doc_name: &str, engine: &EngineHandle, session: &mut CliSession) {
    let mut resolved = doc_name.to_string();
    if !Path::new(&resolved).extension().is_some() {
        resolved.push_str(".xlsx");
    }
    let resolved_path = std::path::absolute(Path::new(&resolved))
        .unwrap_or_else(|_| PathBuf::from(&resolved));

    if let Some(dir) = resolved_path.parent() {
        let _ = fs::create_dir_all(dir);
    }

    let request = rb::build_create_workbook(&resolved_path.to_string_lossy());
    match engine.process_request_json(&request) {
        Ok(response) => {
            if let Some(result) = rp::parse_workbook_open(&response) {
                if result.rid.is_none() || result.rid.as_deref() == Some("") {
                    output::error(&format!("Engine failed to create workbook '{}'.", doc_name));
                    return;
                }
                let rid = result.rid.unwrap();
                session.rid = Some(rid.clone());
                session.workbook_name = resolved_path
                    .file_stem()
                    .map(|s| s.to_string_lossy().to_string());
                session.sheet_count = if result.sheet_count > 0 {
                    result.sheet_count as usize
                } else {
                    1
                };
                session.active_sheet_index = 0;
                session.file_path = Some(resolved_path.to_string_lossy().to_string());

                populate_sheet_names(engine, session, &response, result.active_sheet_id.as_deref());

                let file_name = resolved_path
                    .file_name()
                    .unwrap_or_default()
                    .to_string_lossy();
                output::success(&format!("Created: {}", file_name));
                output::key_value("Workbook ID", session.rid.as_deref().unwrap_or(""), 2);
                output::key_value("Sheets", &format_sheet_summary(session), 2);
                output::key_value("Mode", "Offline", 2);
            } else {
                output::error(&format!("Engine failed to create workbook '{}'.", doc_name));
            }
        }
        Err(e) => output::error(&format!("Engine init failed: {}", e)),
    }
}

fn populate_sheet_names(
    engine: &EngineHandle,
    session: &mut CliSession,
    workbook_response: &str,
    initial_active_sheet_id: Option<&str>,
) {
    session.sheet_names.clear();
    session.sheet_ids.clear();

    let rid = match &session.rid {
        Some(r) => r.clone(),
        None => return,
    };

    // Doc-level fetch to initialise engine
    let doc_fetch_req = rb::build_doc_fetch(&rid);
    let sheets = if let Ok(doc_resp) = engine.doc_fetch_json(&doc_fetch_req) {
        let mut s = rp::parse_sheet_list(&doc_resp);
        if s.is_empty() {
            s = rp::parse_sheet_list(workbook_response);
        }
        s
    } else {
        rp::parse_sheet_list(workbook_response)
    };

    if !sheets.is_empty() {
        for s in &sheets {
            session.sheet_names.push(s.sheet_name.clone());
            let id = if s.sheet_id.is_empty() {
                s.index.to_string()
            } else {
                s.sheet_id.clone()
            };
            session.sheet_ids.push(id);
        }
        session.sheet_count = sheets.len();
        if session.active_sheet_index < session.sheet_names.len() {
            session.active_sheet_name =
                Some(session.sheet_names[session.active_sheet_index].clone());
        }
        if let Some(aid) = initial_active_sheet_id {
            if !aid.is_empty() && !session.sheet_ids.is_empty() {
                session.sheet_ids[session.active_sheet_index] = aid.to_string();
            }
        }
    } else {
        let first_id = initial_active_sheet_id
            .filter(|s| !s.is_empty())
            .unwrap_or("0")
            .to_string();
        for i in 0..session.sheet_count {
            session.sheet_names.push(format!("Sheet{}", i + 1));
            session
                .sheet_ids
                .push(if i == 0 { first_id.clone() } else { i.to_string() });
        }
        session.active_sheet_name = session.sheet_names.first().cloned().or(Some("Sheet1".into()));
    }

    // Sheet-level fetch
    perform_initial_sheet_fetch(engine, session);
}

fn perform_initial_sheet_fetch(engine: &EngineHandle, session: &CliSession) {
    let rid = match &session.rid {
        Some(r) => r.clone(),
        None => return,
    };
    let sheet_id = session.get_active_sheet_id_or_default();
    let fetch = rb::build_initial_sheet_fetch(&rid, &sheet_id, 1_048_575, 16_383);
    let _ = engine.fetch_json(&fetch);
}

/// Notify the engine that the active sheet has changed by performing a sheet fetch.
fn notify_engine_active_sheet(engine: &EngineHandle, session: &CliSession) {
    let rid = match &session.rid {
        Some(r) => r.clone(),
        None => return,
    };
    let sheet_id = session.get_active_sheet_id_or_default();
    let fetch = rb::build_initial_sheet_fetch(&rid, &sheet_id, 1_048_575, 16_383);
    let _ = engine.fetch_json(&fetch);
}

fn format_sheet_summary(session: &CliSession) -> String {
    if session.sheet_names.is_empty() {
        return session.sheet_count.to_string();
    }
    let names = session.sheet_names.join(", ");
    format!("{} ({})", session.sheet_count, names)
}

fn refresh_sheet_list(engine: &EngineHandle, session: &mut CliSession) {
    let rid = match &session.rid {
        Some(r) => r.clone(),
        None => return,
    };
    let req = rb::build_doc_fetch(&rid);
    if let Ok(resp) = engine.doc_fetch_json(&req) {
        let sheets = rp::parse_sheet_list(&resp);
        if !sheets.is_empty() {
            session.sheet_names.clear();
            session.sheet_ids.clear();
            for s in &sheets {
                session.sheet_names.push(s.sheet_name.clone());
                let id = if s.sheet_id.is_empty() {
                    s.index.to_string()
                } else {
                    s.sheet_id.clone()
                };
                session.sheet_ids.push(id);
            }
            session.sheet_count = sheets.len();
            if session.active_sheet_index < session.sheet_names.len() {
                session.active_sheet_name =
                    Some(session.sheet_names[session.active_sheet_index].clone());
            }
        }
    }
}

// ─── Close ───────────────────────────────────────────────────────────────────

fn cmd_close(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if !session.is_active() {
        output::error("No workbook is currently open.");
        return;
    }

    let force = args.first().map(|a| a.eq_ignore_ascii_case("--force")).unwrap_or(false);

    if !force && session.is_dirty {
        let resp = output::confirm("You have unsaved changes. Save before closing?");
        match resp.as_str() {
            "c" => {
                output::info("Close cancelled.");
                return;
            }
            "y" => cmd_save(&[], engine, session),
            _ => {}
        }
    }

    let wb_name = session
        .workbook_name
        .clone()
        .unwrap_or_else(|| "workbook".to_string());
    let rid = session.rid.clone().unwrap();

    let request = rb::build_close_workbook(&rid);
    let status = match engine.process_request_json(&request) {
        Ok(resp) => rp::parse_status_response(&resp),
        Err(_) => rp::EngineStatusResult {
            status_code: -1,
            status_message: Some("Engine error".to_string()),
        },
    };

    session.clear();

    if rp::is_success(status.status_code) {
        output::success(&format!(
            "Closed '{}'. You can now open another workbook.",
            wb_name
        ));
    } else {
        output::warning(&format!(
            "Closed '{}' (engine status: {} — {}).",
            wb_name,
            status.status_code,
            status.status_message.unwrap_or_else(|| "unknown".to_string())
        ));
    }
}

// ─── Save ────────────────────────────────────────────────────────────────────

fn cmd_save(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if !session.is_active() {
        output::error("No workbook open. Use 'open' first.");
        return;
    }

    let mut target_path: Option<String> = None;

    let mut i = 0;
    while i < args.len() {
        if args[i].eq_ignore_ascii_case("--as") {
            if i + 1 < args.len() {
                i += 1;
                target_path = Some(args[i].to_string());
            } else {
                output::error("Usage: save --as <filepath>");
                return;
            }
        }
        i += 1;
    }

    let target = match target_path {
        Some(p) => std::path::absolute(Path::new(&p))
            .unwrap_or_else(|_| PathBuf::from(&p))
            .to_string_lossy()
            .to_string(),
        None => {
            if let Some(ref fp) = session.file_path {
                fp.clone()
            } else {
                output::error(
                    "No original file path. Use 'save --as <filepath>' to specify a destination.",
                );
                return;
            }
        }
    };

    let fmt = Path::new(&target)
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_else(|| "xlsx".to_string());

    let supported: Vec<&str> = vec!["xlsx", "csv", "tsv"];
    if !supported.contains(&fmt.as_str()) {
        output::error(&format!(
            "Unsupported format: '{}'. Supported: xlsx, csv, tsv",
            fmt
        ));
        return;
    }

    // Determine file name and directory
    let target_path_obj = Path::new(&target);
    let is_dir = target_path_obj.is_dir()
        || target.ends_with(std::path::MAIN_SEPARATOR)
        || target.ends_with('/');

    let file_name = if is_dir {
        format!(
            "{}.{}",
            session
                .workbook_name
                .as_deref()
                .unwrap_or("Workbook"),
            fmt
        )
    } else {
        let name = target_path_obj
            .file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_else(|| format!("Workbook.{}", fmt));
        if Path::new(&name).extension().is_none() {
            format!("{}.{}", name, fmt)
        } else {
            name
        }
    };

    let dir_path = if is_dir {
        target.clone()
    } else {
        target_path_obj
            .parent()
            .map(|p| p.to_string_lossy().to_string())
            .unwrap_or_else(|| ".".to_string())
    };

    let _ = fs::create_dir_all(&dir_path);
    let resolved = PathBuf::from(&dir_path).join(&file_name);
    let file_type = map_file_type(&fmt);
    let is_save_as = session
        .file_path
        .as_ref()
        .map(|fp| !fp.eq_ignore_ascii_case(&resolved.to_string_lossy()))
        .unwrap_or(true);
    let sheet_id = session.get_active_sheet_id_or_default();

    let request = rb::build_export_workbook(
        session.rid.as_deref().unwrap(),
        &sheet_id,
        &dir_path,
        &file_name,
        file_type,
        is_save_as,
    );
    match engine.process_request_json(&request) {
        Ok(response) => {
            if let Some(result) = rp::parse_export(&response) {
                let engine_success = result.success || rp::is_success(result.status_code);
                if !engine_success {
                    output::error(&format!(
                        "Save failed: engine export unsuccessful ({}).",
                        result.status_message.unwrap_or_else(|| "unknown error".to_string())
                    ));
                    return;
                }
                let file_size = if result.file_size_bytes > 0 {
                    result.file_size_bytes
                } else {
                    fs::metadata(&resolved)
                        .map(|m| m.len() as i64)
                        .unwrap_or(0)
                };
                let size_display = format_file_size(file_size);

                if is_save_as {
                    output::success(&format!("Exported: {}", file_name));
                    output::key_value("Format", &fmt.to_uppercase(), 2);
                    output::key_value("Size", &size_display, 2);
                } else {
                    output::success(&format!("Saved: {}", file_name));
                    output::key_value("Size", &size_display, 2);
                }
                session.file_path = Some(resolved.to_string_lossy().to_string());
                session.is_dirty = false;
            } else {
                output::error("Save failed: empty engine response.");
            }
        }
        Err(e) => output::error(&format!("Save failed: {}", e)),
    }
}

fn map_file_type(fmt: &str) -> i32 {
    match fmt {
        "csv" => 0,
        "tsv" => 1,
        _ => 2, // xlsx
    }
}

fn format_file_size(bytes: i64) -> String {
    if bytes < 1024 {
        format!("{} B", bytes)
    } else if bytes < 1024 * 1024 {
        format!("{:.1} KB", bytes as f64 / 1024.0)
    } else {
        format!("{:.1} MB", bytes as f64 / (1024.0 * 1024.0))
    }
}

// ─── Cell ────────────────────────────────────────────────────────────────────

fn cmd_cell(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if !session.is_active() {
        output::error("No workbook open. Use 'open' first.");
        return;
    }
    if args.len() < 2 {
        output::error("Usage: cell get <ref> | cell set <ref> <value> [--formula <expr>]");
        return;
    }
    match args[0].to_lowercase().as_str() {
        "get" => cell_get(args[1], engine, session),
        "set" => cell_set(args, engine, session),
        other => output::error(&format!("Unknown cell sub-command: '{}'. Use: get, set", other)),
    }
}

fn cell_get(cell_ref: &str, engine: &EngineHandle, session: &CliSession) {
    let (col, row) = match cell_ref::try_parse(cell_ref) {
        Some(p) => p,
        None => {
            output::error(&format!("Invalid cell reference: '{}'", cell_ref));
            return;
        }
    };

    let rid = session.rid.as_deref().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();

    let fetch_req = rb::build_cell_fetch(rid, &sheet_id, row, col);
    match engine.fetch_json(&fetch_req) {
        Ok(resp) => {
            if let Some(result) = rp::parse_cell_fetch(&resp) {
                if !rp::is_success(result.status_code) {
                    output::error(&format!(
                        "Failed to read {}: {}",
                        cell_ref.to_uppercase(),
                        result.status_message.unwrap_or_else(|| "engine error".into())
                    ));
                    return;
                }

                let fallback = format!("Sheet{}", session.active_sheet_index);
                let sheet_name = session
                    .active_sheet_name
                    .as_deref()
                    .unwrap_or(&fallback);
                let ref_display = cell_ref::to_ref(col, row);
                output::line(&format!("Cell {}  ({})", ref_display, sheet_name), 0);
                output::key_value(
                    "Display",
                    if result.display_value.is_empty() {
                        "(empty)"
                    } else {
                        &result.display_value
                    },
                    2,
                );
                output::key_value(
                    "Raw",
                    if result.raw_value.is_empty() {
                        "(empty)"
                    } else {
                        &result.raw_value
                    },
                    2,
                );
                let formula_disp = if result.formula.is_empty() || result.formula == "null" {
                    "(none)"
                } else {
                    &result.formula
                };
                output::key_value("Formula", formula_disp, 2);
            } else {
                output::error(&format!(
                    "Failed to read cell {}. Empty engine response.",
                    cell_ref.to_uppercase()
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error reading {}: {}", cell_ref, e)),
    }
}

fn cell_set(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    // args[0] = "set", args[1] = cellRef, args[2..] = value or --formula <expr>
    if args.len() < 3 {
        output::error("Usage: cell set <ref> <value> | cell set <ref> --formula <expr>");
        return;
    }

    let cell_ref = args[1];
    let (col, row) = match cell_ref::try_parse(cell_ref) {
        Some(p) => p,
        None => {
            output::error(&format!("Invalid cell reference: '{}'", cell_ref));
            return;
        }
    };

    let (is_formula, value) = if args[2].eq_ignore_ascii_case("--formula") {
        if args.len() < 4 {
            output::error("Usage: cell set <ref> --formula <expr>");
            return;
        }
        (true, args[3])
    } else {
        (false, args[2])
    };

    let rid = session.rid.as_deref().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();

    let request = rb::build_set_cell_value(rid, &sheet_id, row, col, value, is_formula);
    match engine.process_request_json(&request) {
        Ok(response) => {
            let result = rp::parse_set_cell_value(&response);
            let ref_display = cell_ref::to_ref(col, row);
            match result {
                Some(r) if rp::is_success(r.status_code) => {
                    if is_formula {
                        output::success(&format!("{} set to formula: {}", ref_display, value));
                        output::key_value(
                            "Computed value",
                            if r.computed_value.is_empty() {
                                "(pending)"
                            } else {
                                &r.computed_value
                            },
                            2,
                        );
                    } else {
                        output::success(&format!("{} set to: {}", ref_display, value));
                    }
                    session.is_dirty = true;
                }
                Some(r) => {
                    output::error(&format!(
                        "Failed to set {}: {}",
                        ref_display,
                        r.status_message.unwrap_or_else(|| "engine error".into())
                    ));
                }
                None => {
                    output::error(&format!("Failed to set {}: empty engine response", ref_display));
                }
            }
        }
        Err(e) => output::error(&format!("Engine error setting {}: {}", cell_ref, e)),
    }
}

// ─── Sheet ───────────────────────────────────────────────────────────────────

fn cmd_sheet(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if !session.is_active() {
        output::error("No workbook open. Use 'open' first.");
        return;
    }
    if args.is_empty() {
        output::error("Usage: worksheet list|switch|add|delete|rename|reorder|duplicate|hide|unhide [args]");
        return;
    }
    match args[0].to_lowercase().as_str() {
        "list" => sheet_list(session),
        "switch" => {
            if args.len() < 2 {
                output::error("Usage: worksheet switch <name|index>");
                return;
            }
            sheet_select(args[1], engine, session);
        }
        "add" => {
            if args.len() < 2 {
                output::error("Usage: worksheet add <name>");
                return;
            }
            sheet_add(args[1], engine, session);
        }
        "delete" => {
            if args.len() < 2 {
                output::error("Usage: worksheet delete <name|index>");
                return;
            }
            sheet_delete(args[1], engine, session);
        }
        "rename" => {
            if args.len() < 3 {
                output::error("Usage: worksheet rename <old_name> <new_name>");
                return;
            }
            sheet_rename(args[1], args[2], engine, session);
        }
        "reorder" => {
            if args.len() < 2 {
                output::error("Usage: worksheet reorder <newPosition> (0-based)");
                return;
            }
            sheet_reorder(args[1], engine, session);
        }
        "duplicate" => sheet_duplicate(engine, session),
        "hide" => sheet_hide(args.get(1).copied(), engine, session),
        "unhide" => {
            if args.len() < 2 {
                output::error("Usage: worksheet unhide <name|index>");
                return;
            }
            sheet_unhide(args[1], engine, session);
        }
        other => output::error(&format!(
            "Unknown worksheet sub-command: '{}'. Use: list, switch, add, delete, rename, reorder, duplicate, hide, unhide",
            other
        )),
    }
}

fn sheet_list(session: &CliSession) {
    let doc_name = session.workbook_name.as_deref().unwrap_or("Workbook");
    output::line(&format!("Sheets in '{}':", doc_name), 0);
    for (i, name) in session.sheet_names.iter().enumerate() {
        let marker = if i == session.active_sheet_index {
            "  \u{2190} active"
        } else {
            ""
        };
        output::line(&format!("  [{}] {}{}", i, name, marker), 0);
    }
}

fn sheet_select(name_or_index: &str, engine: &EngineHandle, session: &mut CliSession) {
    if let Ok(idx) = name_or_index.parse::<usize>() {
        if idx >= session.sheet_names.len() {
            output::error(&format!(
                "Sheet index {} out of range. Workbook has {} sheet(s) (0-{}).",
                idx,
                session.sheet_names.len(),
                session.sheet_names.len() - 1
            ));
            return;
        }
        session.active_sheet_index = idx;
        session.active_sheet_name = Some(session.sheet_names[idx].clone());
        // Notify the engine about the active sheet change
        notify_engine_active_sheet(engine, session);
        output::success(&format!(
            "Active sheet: [{}] {}",
            idx,
            session.active_sheet_name.as_deref().unwrap()
        ));
    } else {
        let found = session
            .sheet_names
            .iter()
            .position(|s| s.eq_ignore_ascii_case(name_or_index));
        match found {
            Some(i) => {
                session.active_sheet_index = i;
                session.active_sheet_name = Some(session.sheet_names[i].clone());
                // Notify the engine about the active sheet change
                notify_engine_active_sheet(engine, session);
                output::success(&format!(
                    "Active sheet: [{}] {}",
                    i,
                    session.active_sheet_name.as_deref().unwrap()
                ));
            }
            None => {
                // eprintln!("[DEBUG sheet_select] looking for {:?}, available: {:?}", name_or_index, session.sheet_names);
                output::error(&format!(
                    "Sheet '{}' not found. Use 'worksheet list' to see available sheets.",
                    name_or_index
                ));
            }
        }
    }
}

fn sheet_add(name: &str, engine: &EngineHandle, session: &mut CliSession) {
    let rid = session.rid.clone().unwrap();
    let old_ids: Vec<String> = session.sheet_ids.clone();
    let request = rb::build_add_sheet(&rid);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            // Get new sheet's ID: try response first, fall back to diff
            let new_id = rp::extract_sheet_id_from_response(&resp).or_else(|| {
                refresh_sheet_list(engine, session);
                session.sheet_ids.iter()
                    .find(|id| !old_ids.contains(id))
                    .cloned()
            });
            if let Some(id) = new_id {
                let rename_req = rb::build_rename_sheet(&rid, &id, name);
                let _ = engine.process_request_json(&rename_req);
            }
            refresh_sheet_list(engine, session);
            let new_idx = session.sheet_names.iter()
                .position(|n| n.eq_ignore_ascii_case(name))
                .unwrap_or(session.sheet_names.len().saturating_sub(1));
            output::success(&format!(
                "Added sheet: '{}' at index [{}]",
                name, new_idx
            ));
        }
        Err(_) => output::error("Failed to add sheet."),
    }
}

fn resolve_sheet_id<'a>(
    name_or_index: &str,
    session: &'a CliSession,
) -> Option<(String, String)> {
    if let Ok(idx) = name_or_index.parse::<usize>() {
        if idx >= session.sheet_names.len() {
            output::error(&format!(
                "Sheet index {} out of range (0-{}).",
                idx,
                session.sheet_names.len() - 1
            ));
            return None;
        }
        let name = session.sheet_names[idx].clone();
        let id = session
            .sheet_ids
            .get(idx)
            .cloned()
            .unwrap_or_else(|| idx.to_string());
        Some((id, name))
    } else {
        let found = session
            .sheet_names
            .iter()
            .position(|s| s.eq_ignore_ascii_case(name_or_index));
        match found {
            Some(i) => {
                let name = session.sheet_names[i].clone();
                let id = session
                    .sheet_ids
                    .get(i)
                    .cloned()
                    .unwrap_or_else(|| i.to_string());
                Some((id, name))
            }
            None => {
                output::error(&format!(
                    "Sheet '{}' not found. Use 'worksheet list' to see available sheets.",
                    name_or_index
                ));
                None
            }
        }
    }
}

fn sheet_delete(name_or_index: &str, engine: &EngineHandle, session: &mut CliSession) {
    let (sheet_id, resolved_name) = match resolve_sheet_id(name_or_index, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.clone().unwrap();
    let request = rb::build_delete_sheet(&rid, &sheet_id);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                refresh_sheet_list(engine, session);
                output::success(&format!("Deleted sheet: '{}'.", resolved_name));
            } else {
                output::error(&format!(
                    "Failed to delete sheet: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn sheet_rename(old_name: &str, new_name: &str, engine: &EngineHandle, session: &mut CliSession) {
    let rid = session.rid.clone().unwrap();
    let (sheet_id, _) = match resolve_sheet_id(old_name, session) {
        Some(v) => v,
        None => return,
        };
   
    let request = rb::build_rename_sheet(&rid, &sheet_id, new_name);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                refresh_sheet_list(engine, session);
                output::success(&format!("Renamed '{}' \u{2192} '{}'.", old_name, new_name));
            } else {
                output::error(&format!(
                    "Failed to rename sheet: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn sheet_reorder(pos_str: &str, engine: &EngineHandle, session: &mut CliSession) {
    let new_pos: i32 = match pos_str.parse() {
        Ok(p) if p >= 0 => p,
        _ => {
            output::error("Position must be a non-negative integer (0-based).");
            return;
        }
    };
    let rid = session.rid.clone().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();
    let sheet_name = session
        .active_sheet_name
        .clone()
        .unwrap_or_else(|| format!("Sheet{}", session.active_sheet_index));

    let request = rb::build_reorder_sheet(&rid, &sheet_id, new_pos);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                session.active_sheet_index = new_pos as usize;
                refresh_sheet_list(engine, session);
                output::success(&format!("Moved '{}' to position [{}].", sheet_name, new_pos));
            } else {
                output::error(&format!(
                    "Failed to reorder sheet: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn sheet_duplicate(engine: &EngineHandle, session: &mut CliSession) {
    let rid = session.rid.clone().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();
    let sheet_name = session
        .active_sheet_name
        .clone()
        .unwrap_or_else(|| format!("Sheet{}", session.active_sheet_index));

    let request = rb::build_duplicate_sheet(&rid, &sheet_id);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                refresh_sheet_list(engine, session);
                output::success(&format!("Duplicated '{}'.", sheet_name));
            } else {
                output::error(&format!(
                    "Failed to duplicate sheet: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn sheet_hide(name_or_index: Option<&str>, engine: &EngineHandle, session: &mut CliSession) {
    let (sheet_id, sheet_name) = match name_or_index {
        Some(n) => match resolve_sheet_id(n, session) {
            Some(v) => v,
            None => return,
        },
        None => (
            session.get_active_sheet_id_or_default(),
            session
                .active_sheet_name
                .clone()
                .unwrap_or_else(|| format!("Sheet{}", session.active_sheet_index)),
        ),
    };

    let rid = session.rid.clone().unwrap();
    let request = rb::build_hide_sheet(&rid, &[sheet_id.as_str()]);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                refresh_sheet_list(engine, session);
                output::success(&format!("Hidden sheet: '{}'.", sheet_name));
            } else {
                output::error(&format!(
                    "Failed to hide sheet: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn sheet_unhide(name_or_index: &str, engine: &EngineHandle, session: &mut CliSession) {
    let (sheet_id, resolved_name) = match resolve_sheet_id(name_or_index, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.clone().unwrap();
    let request = rb::build_unhide_sheet(&rid, &[sheet_id.as_str()]);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                refresh_sheet_list(engine, session);
                output::success(&format!("Unhidden sheet: '{}'.", resolved_name));
            } else {
                output::error(&format!(
                    "Failed to unhide sheet: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── Row ─────────────────────────────────────────────────────────────────────

fn cmd_row(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if !session.is_active() {
        output::error("No workbook open. Use 'open' first.");
        return;
    }
    if args.len() < 2 {
        output::error("Usage: row insert|delete|hide|unhide|resize <rowNum> [options]");
        return;
    }
    let rid = session.rid.clone().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();
    match args[0].to_lowercase().as_str() {
        "insert" => row_insert(&rid, &sheet_id, args, engine, session),
        "delete" => row_delete(&rid, &sheet_id, args, engine, session),
        "hide" => row_hide(&rid, &sheet_id, args, engine, session),
        "unhide" => row_unhide(&rid, &sheet_id, args, engine, session),
        "resize" => row_resize(&rid, &sheet_id, args, engine, session),
        other => output::error(&format!(
            "Unknown row sub-command: '{}'. Use: insert, delete, hide, unhide, resize",
            other
        )),
    }
}

fn row_insert(
    rid: &str,
    sheet_id: &str,
    args: &[&str],
    engine: &EngineHandle,
    session: &mut CliSession,
) {
    let row_num: i32 = match args[1].parse() {
        Ok(n) if n >= 1 => n,
        _ => {
            output::error("Row number must be a positive integer (1-based).");
            return;
        }
    };
    let count: i32 = args
        .get(2)
        .and_then(|s| s.parse().ok())
        .filter(|&c: &i32| c > 0)
        .unwrap_or(1);
    if count > MAX_ROW_COL_COUNT {
        output::error(&format!("Count exceeds maximum allowed ({}).", MAX_ROW_COL_COUNT));
        return;
    }
    let start_row = row_num - 1;
    let end_row = start_row + count - 1;
    let request = rb::build_insert_row(rid, sheet_id, start_row, 0, end_row, 0, true);
    exec_status_cmd(engine, &request, session, &format!("Inserted {} row(s) at row {}.", count, row_num));
}

fn row_delete(
    rid: &str,
    sheet_id: &str,
    args: &[&str],
    engine: &EngineHandle,
    session: &mut CliSession,
) {
    let row_num: i32 = match args[1].parse() {
        Ok(n) if n >= 1 => n,
        _ => {
            output::error("Row number must be a positive integer (1-based).");
            return;
        }
    };
    let count: i32 = args.get(2).and_then(|s| s.parse().ok()).filter(|&c: &i32| c > 0).unwrap_or(1);
    if count > MAX_ROW_COL_COUNT {
        output::error(&format!("Count exceeds maximum allowed ({}).", MAX_ROW_COL_COUNT));
        return;
    }
    let start_row = row_num - 1;
    let end_row = start_row + count - 1;
    let request = rb::build_delete_row(rid, sheet_id, start_row, 0, end_row, 0);
    exec_status_cmd(engine, &request, session, &format!("Deleted {} row(s) starting at row {}.", count, row_num));
}

fn row_hide(
    rid: &str,
    sheet_id: &str,
    args: &[&str],
    engine: &EngineHandle,
    session: &mut CliSession,
) {
    let start: i32 = match args[1].parse() {
        Ok(n) if n >= 1 => n,
        _ => { output::error("Row number must be a positive integer (1-based)."); return; }
    };
    let end: i32 = args.get(2).and_then(|s| s.parse().ok()).filter(|&e: &i32| e >= start).unwrap_or(start);
    let request = rb::build_hide_row(rid, sheet_id, start - 1, end - 1, 0, 0);
    let label = if end != start { format!("{}-{}", start, end) } else { start.to_string() };
    exec_status_cmd(engine, &request, session, &format!("Hidden row(s) {}.", label));
}

fn row_unhide(
    rid: &str,
    sheet_id: &str,
    args: &[&str],
    engine: &EngineHandle,
    session: &mut CliSession,
) {
    let start: i32 = match args[1].parse() {
        Ok(n) if n >= 1 => n,
        _ => { output::error("Row number must be a positive integer (1-based)."); return; }
    };
    let end: i32 = args.get(2).and_then(|s| s.parse().ok()).filter(|&e: &i32| e >= start).unwrap_or(start);
    let request = rb::build_unhide_row(rid, sheet_id, start - 1, end - 1, 0, 0);
    let label = if end != start { format!("{}-{}", start, end) } else { start.to_string() };
    exec_status_cmd(engine, &request, session, &format!("Unhidden row(s) {}.", label));
}

fn row_resize(
    rid: &str,
    sheet_id: &str,
    args: &[&str],
    engine: &EngineHandle,
    session: &mut CliSession,
) {
    let row_num: i32 = match args[1].parse() {
        Ok(n) if n >= 1 => n,
        _ => { output::error("Row number must be a positive integer (1-based)."); return; }
    };
    if args.len() < 3 {
        output::error("Usage: row resize <rowNum> <height> | row resize <rowNum> --auto");
        return;
    }
    let (auto_fit, height) = if args[2].eq_ignore_ascii_case("--auto") {
        (true, 0)
    } else {
        match args[2].parse::<i32>() {
            Ok(h) if h >= 1 => (false, h),
            _ => { output::error("Height must be a positive integer or use --auto."); return; }
        }
    };
    let request = rb::build_resize_row(rid, sheet_id, row_num - 1, row_num - 1, height, auto_fit, 0, 0);
    let desc = if auto_fit { "auto-fit".to_string() } else { format!("{}px", height) };
    exec_status_cmd(engine, &request, session, &format!("Resized row {} to {}.", row_num, desc));
}

// ─── Col ─────────────────────────────────────────────────────────────────────

fn cmd_col(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if !session.is_active() {
        output::error("No workbook open. Use 'open' first.");
        return;
    }
    if args.len() < 2 {
        output::error("Usage: col insert|delete|hide|unhide|resize <colLetter> [options]");
        return;
    }
    let rid = session.rid.clone().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();
    match args[0].to_lowercase().as_str() {
        "insert" => col_insert(&rid, &sheet_id, args, engine, session),
        "delete" => col_delete(&rid, &sheet_id, args, engine, session),
        "hide" => col_hide(&rid, &sheet_id, args, engine, session),
        "unhide" => col_unhide(&rid, &sheet_id, args, engine, session),
        "resize" => col_resize(&rid, &sheet_id, args, engine, session),
        other => output::error(&format!(
            "Unknown col sub-command: '{}'. Use: insert, delete, hide, unhide, resize",
            other
        )),
    }
}

fn col_insert(rid: &str, sheet_id: &str, args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let col_idx = match cell_ref::try_parse_col_letter(args[1]) {
        Some(c) => c,
        None => { output::error("Column must be a letter (e.g., A, B, AA)."); return; }
    };
    let count: i32 = args.get(2).and_then(|s| s.parse().ok()).filter(|&c: &i32| c > 0).unwrap_or(1);
    if count > MAX_ROW_COL_COUNT {
        output::error(&format!("Count exceeds maximum allowed ({}).", MAX_ROW_COL_COUNT));
        return;
    }
    let end_col = col_idx + count - 1;
    let request = rb::build_insert_column(rid, sheet_id, 0, col_idx, 0, end_col, true);
    exec_status_cmd(engine, &request, session, &format!("Inserted {} column(s) at column {}.", count, cell_ref::col_to_letter(col_idx)));
}

fn col_delete(rid: &str, sheet_id: &str, args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let col_idx = match cell_ref::try_parse_col_letter(args[1]) {
        Some(c) => c,
        None => { output::error("Column must be a letter (e.g., A, B, AA)."); return; }
    };
    let count: i32 = args.get(2).and_then(|s| s.parse().ok()).filter(|&c: &i32| c > 0).unwrap_or(1);
    if count > MAX_ROW_COL_COUNT {
        output::error(&format!("Count exceeds maximum allowed ({}).", MAX_ROW_COL_COUNT));
        return;
    }
    let end_col = col_idx + count - 1;
    let request = rb::build_delete_column(rid, sheet_id, 0, col_idx, 0, end_col);
    exec_status_cmd(engine, &request, session, &format!("Deleted {} column(s) starting at column {}.", count, cell_ref::col_to_letter(col_idx)));
}

fn col_hide(rid: &str, sheet_id: &str, args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let start = match cell_ref::try_parse_col_letter(args[1]) {
        Some(c) => c,
        None => { output::error("Column must be a letter (e.g., A, B, AA)."); return; }
    };
    let end = args.get(2).and_then(|s| cell_ref::try_parse_col_letter(s)).unwrap_or(start);
    let request = rb::build_hide_column(rid, sheet_id, start, end, 0, 0);
    let label = if end != start {
        format!("{}-{}", cell_ref::col_to_letter(start), cell_ref::col_to_letter(end))
    } else {
        cell_ref::col_to_letter(start)
    };
    exec_status_cmd(engine, &request, session, &format!("Hidden column(s) {}.", label));
}

fn col_unhide(rid: &str, sheet_id: &str, args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let start = match cell_ref::try_parse_col_letter(args[1]) {
        Some(c) => c,
        None => { output::error("Column must be a letter (e.g., A, B, AA)."); return; }
    };
    let end = args.get(2).and_then(|s| cell_ref::try_parse_col_letter(s)).unwrap_or(start);
    let request = rb::build_unhide_column(rid, sheet_id, start, end, 0, 0);
    let label = if end != start {
        format!("{}-{}", cell_ref::col_to_letter(start), cell_ref::col_to_letter(end))
    } else {
        cell_ref::col_to_letter(start)
    };
    exec_status_cmd(engine, &request, session, &format!("Unhidden column(s) {}.", label));
}

fn col_resize(rid: &str, sheet_id: &str, args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let col_idx = match cell_ref::try_parse_col_letter(args[1]) {
        Some(c) => c,
        None => { output::error("Column must be a letter (e.g., A, B, AA)."); return; }
    };
    if args.len() < 3 {
        output::error("Usage: col resize <colLetter> <width> | col resize <colLetter> --auto");
        return;
    }
    let (auto_fit, width) = if args[2].eq_ignore_ascii_case("--auto") {
        (true, 0)
    } else {
        match args[2].parse::<i32>() {
            Ok(w) if w >= 1 => (false, w),
            _ => { output::error("Width must be a positive integer or use --auto."); return; }
        }
    };
    let request = rb::build_resize_column(rid, sheet_id, col_idx, col_idx, width, auto_fit, 0, 0);
    let desc = if auto_fit { "auto-fit".to_string() } else { format!("{}px", width) };
    exec_status_cmd(engine, &request, session, &format!("Resized column {} to {}.", cell_ref::col_to_letter(col_idx), desc));
}

// ─── Copy / Move ─────────────────────────────────────────────────────────────

fn cmd_copy(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.len() < 2 {
        output::error("Usage: copy <source_range> <dest_range> [--values|--format]");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let (dest_sc, dest_sr, dest_ec, dest_er) = match cell_ref::try_parse_range(args[1]) {
        Some(r) => r,
        None => { output::error(&format!("Invalid destination: '{}'. Use A1 or A1:C5 format.", args[1])); return; }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let src_rows = er - sr + 1;
    let src_cols = ec - sc + 1;
    let dest_rows = dest_er - dest_sr + 1;
    let dest_cols = dest_ec - dest_sc + 1;

    // Read all source cell values
    let mut src_grid: Vec<Vec<(String, String)>> = vec![vec![(String::new(), String::new()); src_cols as usize]; src_rows as usize];
    for r in sr..=er {
        for c in sc..=ec {
            let fetch_req = rb::build_cell_fetch(rid, &sid, r, c);
            if let Ok(resp) = engine.fetch_json(&fetch_req) {
                if let Some(result) = rp::parse_cell_fetch(&resp) {
                    let ri = (r - sr) as usize;
                    let ci = (c - sc) as usize;
                    src_grid[ri][ci] = (result.raw_value, result.formula);
                }
            }
        }
    }

    // Write values to destination, tiling the source as needed
    for dr in 0..dest_rows {
        for dc in 0..dest_cols {
            let ri = (dr % src_rows) as usize;
            let ci = (dc % src_cols) as usize;
            let (ref raw, ref formula) = src_grid[ri][ci];
            let (is_formula, value) = if !formula.is_empty() && formula != "null" {
                (true, formula.as_str())
            } else {
                (false, raw.as_str())
            };
            let set_req = rb::build_set_cell_value(rid, &sid, dest_sr + dr, dest_sc + dc, value, is_formula);
            match engine.process_request_json(&set_req) {
                Ok(resp) => {
                    if let Some(r) = rp::parse_set_cell_value(&resp) {
                        if !rp::is_success(r.status_code) {
                            output::error(&format!("Failed to write {}: {}",
                                cell_ref::to_ref(dest_sc + dc, dest_sr + dr),
                                r.status_message.unwrap_or_else(|| "engine error".into())));
                            return;
                        }
                    }
                }
                Err(e) => { output::error(&format!("Engine error: {}", e)); return; }
            }
        }
    }

    let dest_display = if dest_rows == 1 && dest_cols == 1 {
        cell_ref::to_ref(dest_sc, dest_sr)
    } else {
        format!("{}:{}", cell_ref::to_ref(dest_sc, dest_sr), cell_ref::to_ref(dest_ec, dest_er))
    };
    output::success(&format!("Copied {} to {}.", args[0].to_uppercase(), dest_display));
    session.is_dirty = true;
}

fn cmd_move(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.len() < 2 {
        output::error("Usage: move <source_range> <dest_range>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let (dest_sc, dest_sr, dest_ec, dest_er) = match cell_ref::try_parse_range(args[1]) {
        Some(r) => r,
        None => { output::error(&format!("Invalid destination: '{}'. Use A1 or A1:C5 format.", args[1])); return; }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let src_rows = er - sr + 1;
    let src_cols = ec - sc + 1;
    let dest_rows = dest_er - dest_sr + 1;
    let dest_cols = dest_ec - dest_sc + 1;

    // Read all source cell values first
    let mut src_grid: Vec<Vec<(String, String)>> = vec![vec![(String::new(), String::new()); src_cols as usize]; src_rows as usize];
    for r in sr..=er {
        for c in sc..=ec {
            let fetch_req = rb::build_cell_fetch(rid, &sid, r, c);
            if let Ok(resp) = engine.fetch_json(&fetch_req) {
                if let Some(result) = rp::parse_cell_fetch(&resp) {
                    let ri = (r - sr) as usize;
                    let ci = (c - sc) as usize;
                    src_grid[ri][ci] = (result.raw_value, result.formula);
                }
            }
        }
    }

    // Clear the source range
    let clear_req = rb::build_clear(rid, &sid, rb::ACTION_CLEAR_CONTENT, sr, sc, er, ec);
    match engine.process_request_json(&clear_req) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if !rp::is_success(status.status_code) {
                output::error(&format!("Failed to clear source range: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())));
                return;
            }
        }
        Err(e) => { output::error(&format!("Engine error: {}", e)); return; }
    }

    // Write values to destination, tiling the source as needed
    for dr in 0..dest_rows {
        for dc in 0..dest_cols {
            let ri = (dr % src_rows) as usize;
            let ci = (dc % src_cols) as usize;
            let (ref raw, ref formula) = src_grid[ri][ci];
            let (is_formula, value) = if !formula.is_empty() && formula != "null" {
                (true, formula.as_str())
            } else {
                (false, raw.as_str())
            };
            let set_req = rb::build_set_cell_value(rid, &sid, dest_sr + dr, dest_sc + dc, value, is_formula);
            match engine.process_request_json(&set_req) {
                Ok(resp) => {
                    if let Some(r) = rp::parse_set_cell_value(&resp) {
                        if !rp::is_success(r.status_code) {
                            output::error(&format!("Failed to write {}: {}",
                                cell_ref::to_ref(dest_sc + dc, dest_sr + dr),
                                r.status_message.unwrap_or_else(|| "engine error".into())));
                            return;
                        }
                    }
                }
                Err(e) => { output::error(&format!("Engine error: {}", e)); return; }
            }
        }
    }

    let dest_display = if dest_rows == 1 && dest_cols == 1 {
        cell_ref::to_ref(dest_sc, dest_sr)
    } else {
        format!("{}:{}", cell_ref::to_ref(dest_sc, dest_sr), cell_ref::to_ref(dest_ec, dest_er))
    };
    output::success(&format!("Moved {} to {}.", args[0].to_uppercase(), dest_display));
    session.is_dirty = true;
}

fn cmd_clipboard_copy(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: clipboardcopy <range> (e.g., clipboardcopy A1 or clipboardcopy A1:C5)");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let num_rows = (er - sr + 1) as usize;
    let num_cols = (ec - sc + 1) as usize;
    const MAX_CLIPBOARD_CELLS: usize = 10_000;
    if num_rows * num_cols > MAX_CLIPBOARD_CELLS {
        output::error(&format!(
            "Range too large ({} cells). Maximum allowed for clipboard is {}.",
            num_rows * num_cols,
            MAX_CLIPBOARD_CELLS
        ));
        return;
    }

    // Fetch each cell individually using the proven single-cell path
    let mut grid = vec![vec![String::new(); num_cols]; num_rows];
    for r in sr..=er {
        for c in sc..=ec {
            let fetch_req = rb::build_cell_fetch(rid, &sid, r, c);
            if let Ok(resp) = engine.fetch_json(&fetch_req) {
                if let Some(result) = rp::parse_cell_fetch(&resp) {
                    let ri = (r - sr) as usize;
                    let ci = (c - sc) as usize;
                    grid[ri][ci] = result.display_value;
                }
            }
        }
    }

    let tsv: String = grid
        .iter()
        .map(|row| row.join("\t"))
        .collect::<Vec<_>>()
        .join("\n");

    match arboard::Clipboard::new() {
        Ok(mut clipboard) => {
            if let Err(e) = clipboard.set_text(&tsv) {
                output::error(&format!("Failed to set system clipboard: {}", e));
            } else {
                output::success(&format!("Copied {} to system clipboard.", args[0].to_uppercase()));
            }
        }
        Err(e) => output::error(&format!("Could not access system clipboard: {}", e)),
    }
}

// ─── Find / Replace ──────────────────────────────────────────────────────────

fn cmd_find(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: find <text> [--exact] [--case]");
        return;
    }
    let search_text = args[0];
    let mut is_exact = false;
    let mut is_case = false;
    for a in args.iter().skip(1) {
        if a.eq_ignore_ascii_case("--exact") { is_exact = true; }
        if a.eq_ignore_ascii_case("--case") { is_case = true; }
    }
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_find(rid, &sid, search_text, 0, 0, is_exact, is_case, true);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_find_replace(&resp);
            if rp::is_success(result.status_code) {
                if result.match_count > 0 {
                    output::success(&format!("Found {} match(es) for '{}'.", result.match_count, search_text));
                    if result.found_row >= 0 && result.found_col >= 0 {
                        output::key_value("First match", &cell_ref::to_ref(result.found_col, result.found_row), 2);
                    }
                } else {
                    output::info(&format!("No matches found for '{}'.", search_text));
                }
            } else {
                output::error(&format!("Find failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn cmd_replace(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.len() < 2 {
        output::error("Usage: replace <search> <replacement> [--all] [--case] [--exact]");
        return;
    }
    let search_text = args[0];
    let replace_text = args[1];
    let mut replace_all = false;
    let mut is_case = false;
    let mut is_exact = false;
    for a in args.iter().skip(2) {
        if a.eq_ignore_ascii_case("--all") { replace_all = true; }
        if a.eq_ignore_ascii_case("--case") { is_case = true; }
        if a.eq_ignore_ascii_case("--exact") { is_exact = true; }
    }
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_replace(rid, &sid, search_text, replace_text, 0, 0, is_exact, is_case, replace_all);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_find_replace(&resp);
            if rp::is_success(result.status_code) {
                let mode = if replace_all { "Replaced all" } else { "Replaced" };
                if result.match_count > 0 {
                    output::success(&format!("{}: '{}' \u{2192} '{}' ({} match(es)).", mode, search_text, replace_text, result.match_count));
                    session.is_dirty = true;
                } else if !replace_all && result.has_meta {
                    output::success(&format!("Replaced: '{}' \u{2192} '{}' (last occurrence).", search_text, replace_text));
                    session.is_dirty = true;
                } else {
                    output::info(&format!("No matches found for '{}'.", search_text));
                }
            } else {
                output::error(&format!("Replace failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── Sort / Filter ───────────────────────────────────────────────────────────

fn cmd_sort(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.len() < 2 {
        output::error("Usage: sort <range> <colLetter> [--desc] [--header]");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let sort_col = match cell_ref::try_parse_col_letter(args[1]) {
        Some(c) => c,
        None => { output::error(&format!("Invalid column letter: '{}'.", args[1])); return; }
    };
    let mut is_asc = true;
    let mut has_header = false;
    for a in args.iter().skip(2) {
        if a.eq_ignore_ascii_case("--desc") { is_asc = false; }
        if a.eq_ignore_ascii_case("--header") { has_header = true; }
    }
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_sort(rid, &sid, sr, sc, er, ec, sort_col, is_asc, has_header);
    let dir = if is_asc { "ascending" } else { "descending" };
    exec_status_cmd(engine, &request, session, &format!("Sorted {} by column {} ({}).", args[0].to_uppercase(), args[1].to_uppercase(), dir));
}

fn cmd_filter(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: filter create <range> | filter remove");
        return;
    }
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    match args[0].to_lowercase().as_str() {
        "create" => {
            if args.len() < 2 {
                output::error("Usage: filter create <range> (e.g., filter create A1:D10)");
                return;
            }
            let (sc, sr, ec, er) = parse_range_arg!(args[1]);
            let request = rb::build_create_filter(rid, &sid, sr, sc, er, ec);
            exec_status_cmd(engine, &request, session, &format!("Auto-filter created on {}.", args[1].to_uppercase()));
        }
        "remove" => {
            let request = rb::build_remove_filter(rid, &sid);
            exec_status_cmd(engine, &request, session, "Auto-filter removed.");
        }
        other => output::error(&format!("Unknown filter sub-command: '{}'. Use: create, remove", other)),
    }
}

// ─── Merge / Clear ───────────────────────────────────────────────────────────

fn cmd_merge(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: merge <range> | merge undo <range>");
        return;
    }
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    if args[0].eq_ignore_ascii_case("undo") {
        if args.len() < 2 {
            output::error("Usage: merge undo <range>");
            return;
        }
        let (sc, sr, ec, er) = parse_range_arg!(args[1]);
        let request = rb::build_unmerge_cells(rid, &sid, sr, sc, er, ec);
        exec_status_cmd(engine, &request, session, &format!("Unmerged cells {}.", args[1].to_uppercase()));
    } else {
        let (sc, sr, ec, er) = parse_range_arg!(args[0]);
        let request = rb::build_merge_cells(rid, &sid, sr, sc, er, ec);
        exec_status_cmd(engine, &request, session, &format!("Merged cells {}.", args[0].to_uppercase()));
    }
}

fn cmd_clear(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: clear <range> [--content|--format]");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let mut action_id = rb::ACTION_CLEAR_ALL;
    let mut mode_label = "all";
    for a in args.iter().skip(1) {
        if a.eq_ignore_ascii_case("--content") { action_id = rb::ACTION_CLEAR_CONTENT; mode_label = "content"; }
        if a.eq_ignore_ascii_case("--format") { action_id = rb::ACTION_CLEAR_FORMAT; mode_label = "formatting"; }
    }
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_clear(rid, &sid, action_id, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Cleared {} in {}.", mode_label, args[0].to_uppercase()));
}

// ─── Undo / Redo ─────────────────────────────────────────────────────────────

fn cmd_undo(engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    let request = rb::build_undo(session.rid.as_deref().unwrap());
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success("Undo successful.");
                session.is_dirty = true;
            } else {
                output::warning(&format!("Undo: {}", status.status_message.unwrap_or_else(|| "nothing to undo".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn cmd_redo(engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    let request = rb::build_redo(session.rid.as_deref().unwrap());
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success("Redo successful.");
                session.is_dirty = true;
            } else {
                output::warning(&format!("Redo: {}", status.status_message.unwrap_or_else(|| "nothing to redo".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── Freeze / Unfreeze ───────────────────────────────────────────────────────

fn cmd_freeze(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: freeze <cellRef> (e.g., freeze B2)");
        return;
    }
    let (col, row) = match cell_ref::try_parse(args[0]) {
        Some(p) => p,
        None => { output::error(&format!("Invalid cell reference: '{}'", args[0])); return; }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_freeze(rid, &sid, row, col);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success(&format!("Panes frozen at {}.", cell_ref::to_ref(col, row)));
            } else {
                output::error(&format!("Failed to freeze: {}", status.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn cmd_unfreeze(engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_unfreeze(rid, &sid);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success("Panes unfrozen.");
            } else {
                output::error(&format!("Failed to unfreeze: {}", status.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── Named Ranges ────────────────────────────────────────────────────────────

fn cmd_name(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: name add <name> <expression> | name delete <name> | name list");
        return;
    }
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    match args[0].to_lowercase().as_str() {
        "add" => {
            if args.len() < 3 {
                output::error("Usage: name add <name> <expression> [comment]");
                return;
            }
            let comment = args.get(3).unwrap_or(&"");
            let request = rb::build_add_defined_name(rid, &sid, args[1], args[2], comment);
            match engine.process_request_json(&request) {
                Ok(resp) => {
                    let status = rp::parse_status_response(&resp);
                    if rp::is_success(status.status_code) {
                        output::success(&format!("Defined name '{}' added \u{2192} {}", args[1], args[2]));
                    } else {
                        output::error(&format!("Failed to add name: {}", status.status_message.unwrap_or_else(|| "engine error".into())));
                    }
                }
                Err(e) => output::error(&format!("Engine error: {}", e)),
            }
        }
        "delete" => {
            if args.len() < 2 {
                output::error("Usage: name delete <name>");
                return;
            }
            let request = rb::build_delete_defined_name(rid, &sid, args[1]);
            match engine.process_request_json(&request) {
                Ok(resp) => {
                    let status = rp::parse_status_response(&resp);
                    if rp::is_success(status.status_code) {
                        output::success(&format!("Defined name '{}' deleted.", args[1]));
                    } else {
                        output::error(&format!("Failed to delete name: {}", status.status_message.unwrap_or_else(|| "engine error".into())));
                    }
                }
                Err(e) => output::error(&format!("Engine error: {}", e)),
            }
        }
        "list" => {
            let request = rb::build_manage_defined_names(rid);
            match engine.process_request_json(&request) {
                Ok(resp) => {
                    let status = rp::parse_status_response(&resp);
                    if rp::is_success(status.status_code) {
                        output::success("Defined names retrieved successfully.");
                    } else {
                        output::error(&format!("Failed to list names: {}", status.status_message.unwrap_or_else(|| "engine error".into())));
                    }
                }
                Err(e) => output::error(&format!("Engine error: {}", e)),
            }
        }
        other => output::error(&format!("Unknown name sub-command: '{}'. Use: add, delete, list", other)),
    }
}

// ─── Table ────────────────────────────────────────────────────────────────────

/// Resolves a table identifier (either a table ID or table name) to the actual table_id.
/// Returns `Some(table_id)` on success, or `None` if no matching table is found.
fn resolve_table_id(identifier: &str, engine: &EngineHandle, session: &CliSession) -> Option<String> {
    let rid = session.rid.as_deref()?;
    let sheet_id = session.get_active_sheet_id_or_default();
    let fetch_req = rb::build_table_list_fetch(rid, &sheet_id);
    let resp = engine.fetch_json(&fetch_req).ok()?;
    let tables = rp::parse_table_list(&resp);

    // First, check if the identifier matches a table_id directly
    for t in &tables {
        if t.table_id == identifier {
            return Some(identifier.to_string());
        }
    }

    // Otherwise, treat it as a table name and search by calling manage on each
    for t in &tables {
        let manage_req = rb::build_manage_table(rid, &t.table_id);
        if let Ok(manage_resp) = engine.process_request_json(&manage_req) {
            if let Some(info) = rp::parse_manage_table(&manage_resp) {
                if info.table_name.eq_ignore_ascii_case(identifier) {
                    return Some(t.table_id.clone());
                }
            }
        }
    }

    None
}

fn cmd_table(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: table list|create|select|delete|rename|options|source|style|defaultstyle|insertrow|insertcol|deleterow|deletecol|manage [args]");
        return;
    }
    match args[0].to_lowercase().as_str() {
        "list" => table_list(engine, session),
        "create" => table_create(&args[1..], engine, session),
        "select" => table_select(&args[1..], engine, session),
        "delete" => table_delete(&args[1..], engine, session),
        "rename" => table_rename(&args[1..], engine, session),
        "options" => table_options(&args[1..], engine, session),
        "source" => table_source(&args[1..], engine, session),
        "style" => table_style(&args[1..], engine, session),
        "defaultstyle" => table_default_style(&args[1..], engine, session),
        "insertrow" => table_insert_row(&args[1..], engine, session),
        "insertcol" => table_insert_col(&args[1..], engine, session),
        "deleterow" => table_delete_row(&args[1..], engine, session),
        "deletecol" => table_delete_col(&args[1..], engine, session),
        "manage" => table_manage(&args[1..], engine, session),
        other => output::error(&format!(
            "Unknown table sub-command: '{}'. Use: list, create, select, delete, rename, options, source, style, defaultstyle, insertrow, insertcol, deleterow, deletecol, manage",
            other
        )),
    }
}

fn table_list(engine: &EngineHandle, session: &CliSession) {
    let rid = match &session.rid {
        Some(r) => r.clone(),
        None => return,
    };
    let sheet_id = session.get_active_sheet_id_or_default();
    let fetch_req = rb::build_table_list_fetch(&rid, &sheet_id);
    match engine.fetch_json(&fetch_req) {
        Ok(resp) => {
            let tables = rp::parse_table_list(&resp);
            if tables.is_empty() {
                output::info("No tables found in the active sheet.");
            } else {
                let sheet_name = session.active_sheet_name.as_deref().unwrap_or("Sheet");
                output::line(&format!("Tables in '{}':", sheet_name), 0);
                for (i, t) in tables.iter().enumerate() {
                    let start = format!("{}{}", cell_ref::col_to_letter(t.start_col), t.start_row + 1);
                    let end = format!("{}{}", cell_ref::col_to_letter(t.end_col), t.end_row + 1);
                    // Try to fetch the table name via manage
                    let name = {
                        let manage_req = rb::build_manage_table(&rid, &t.table_id);
                        engine.process_request_json(&manage_req).ok()
                            .and_then(|resp| rp::parse_manage_table(&resp))
                            .map(|info| info.table_name)
                    };
                    if let Some(ref name) = name {
                        output::line(&format!("  [{}] {} ({})  ({}:{})", i, t.table_id, name, start, end), 0);
                    } else {
                        output::line(&format!("  [{}] {}  ({}:{})", i, t.table_id, start, end), 0);
                    }
                }
            }
        }
        Err(e) => output::error(&format!("Failed to fetch table info: {}", e)),
    }
}

fn table_create(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: table create <range> [--headers]");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let has_headers = args.iter().skip(1).any(|a| a.eq_ignore_ascii_case("--headers"));
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_insert_table(rid, &sid, sr, sc, er, ec, has_headers);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let (status_code, status_message, table_id) = rp::parse_insert_table(&resp);
            if rp::is_success(status_code) {
                output::success(&format!(
                    "Table created on {}.",
                    args[0].to_uppercase()
                ));
                if let Some(id) = table_id {
                    output::key_value("Table ID", &id, 2);
                }
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to create table: {}",
                    status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn table_select(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: table select <range>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_select_table_range(rid, &sid, sr, sc, er, ec);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let (status_code, status_message, has_headers, range) =
                rp::parse_select_table_range(&resp);
            if rp::is_success(status_code) {
                output::success("Table range selected.");
                output::key_value("Has headers", if has_headers { "yes" } else { "no" }, 2);
                if let Some((sr, sc, er, ec)) = range {
                    output::key_value(
                        "Range",
                        &format!(
                            "{}:{}",
                            cell_ref::to_ref(sc, sr),
                            cell_ref::to_ref(ec, er)
                        ),
                        2,
                    );
                }
            } else {
                output::error(&format!(
                    "Failed to select table range: {}",
                    status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn table_delete(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: table delete <tableId|tableName> [--keep-format]");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let keep_format = args.iter().skip(1).any(|a| a.eq_ignore_ascii_case("--keep-format"));
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_delete_table(rid, &table_id, keep_format);
    exec_status_cmd(engine, &request, session, &format!("Table '{}' deleted.", table_id));
}

fn table_rename(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table rename <tableId|tableName> <newName>");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let new_name = args[1];
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_change_table_name(rid, &sid, &table_id, new_name);
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table renamed to '{}'.", new_name),
    );
}

fn table_options(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: table options <tableId|tableName> <settingType> <true|false>");
        output::info("  Setting types: 0=header_row, 1=total_row, 2=banded_row, 3=banded_column, 4=first_column, 5=last_column, 6=filter_button");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let setting_type: i32 = match args[1].parse() {
        Ok(n) if (0..=6).contains(&n) => n,
        _ => {
            output::error("Setting type must be 0-6.");
            return;
        }
    };
    let is_enabled = args[2].eq_ignore_ascii_case("true");
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_change_table_options(rid, &table_id, &sid, setting_type, is_enabled);
    let setting_names = [
        "header_row", "total_row", "banded_row", "banded_column",
        "first_column", "last_column", "filter_button",
    ];
    let label = setting_names.get(setting_type as usize).unwrap_or(&"unknown");
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table option '{}' set to {}.", label, is_enabled),
    );
}

fn table_source(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table source <tableId|tableName> <range>");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_change_table_source(rid, &table_id, &sid, sr, sc, er, ec);
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table source changed to {}.", args[1].to_uppercase()),
    );
}

fn table_style(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table style <tableId|tableName> <stylePattern> [--keep-format]");
        output::info("  Style patterns: 0=none, 1-3=light, 4-8=medium, 9=dark");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let pattern: i32 = match args[1].parse() {
        Ok(n) if (0..=9).contains(&n) => n,
        _ => {
            output::error("Style pattern must be 0-9.");
            return;
        }
    };
    let keep_format = args.iter().skip(2).any(|a| a.eq_ignore_ascii_case("--keep-format"));
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_change_table_style_pattern(rid, &table_id, pattern, keep_format);
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table style changed to pattern {}.", pattern),
    );
}

fn table_default_style(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: table defaultstyle <stylePattern>");
        output::info("  Style patterns: 0=none, 1-3=light, 4-8=medium, 9=dark");
        return;
    }
    let pattern: i32 = match args[0].parse() {
        Ok(n) if (0..=9).contains(&n) => n,
        _ => {
            output::error("Style pattern must be 0-9.");
            return;
        }
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_set_default_table_style(rid, pattern);
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Default table style set to pattern {}.", pattern),
    );
}

fn table_insert_row(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table insertrow <tableId|tableName> <range> [--above]");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let is_below = !args.iter().skip(2).any(|a| a.eq_ignore_ascii_case("--above"));
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_insert_table_row(rid, &table_id, &sid, sr, sc, er, ec, is_below);
    let pos = if is_below { "below" } else { "above" };
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table row(s) inserted {}.", pos),
    );
}

fn table_insert_col(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table insertcol <tableId|tableName> <range> [--after]");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let is_after = args.iter().skip(2).any(|a| a.eq_ignore_ascii_case("--after"));
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_insert_table_column(rid, &table_id, &sid, sr, sc, er, ec, is_after);
    let pos = if is_after { "after" } else { "before" };
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table column(s) inserted {}.", pos),
    );
}

fn table_delete_row(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table deleterow <tableId|tableName> <range>");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_delete_table_row(rid, &table_id, &sid, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, "Table row(s) deleted.");
}

fn table_delete_col(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table deletecol <tableId|tableName> <range>");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_delete_table_column(rid, &table_id, &sid, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, "Table column(s) deleted.");
}

fn table_manage(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: table manage <tableId|tableName>");
        return;
    }
    let table_id = match resolve_table_id(args[0], engine, session) {
        Some(id) => id,
        None => {
            output::error(&format!("Table '{}' not found. Use 'table list' to see available tables.", args[0]));
            return;
        }
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_manage_table(rid, &table_id);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            if let Some(info) = rp::parse_manage_table(&resp) {
                if !rp::is_success(info.status_code) {
                    output::error(&format!(
                        "Failed to get table info: {}",
                        info.status_message.unwrap_or_else(|| "engine error".into())
                    ));
                    return;
                }
                let range_display = format!(
                    "{}:{}",
                    cell_ref::to_ref(info.source_start_col, info.source_start_row),
                    cell_ref::to_ref(info.source_end_col, info.source_end_row)
                );
                output::line(&format!("Table: {}", info.table_name), 0);
                output::key_value("Table ID", &table_id, 2);
                output::key_value("Source", &range_display, 2);
                output::key_value("Style", &format!("{} ({})", info.table_style_type, info.table_color_pattern), 2);
                output::line("  Options:", 0);
                output::key_value("  Header Row", if info.is_header_row { "yes" } else { "no" }, 2);
                output::key_value("  Total Row", if info.is_total_row { "yes" } else { "no" }, 2);
                output::key_value("  Banded Rows", if info.is_banded_row { "yes" } else { "no" }, 2);
                output::key_value("  Banded Columns", if info.is_banded_column { "yes" } else { "no" }, 2);
                output::key_value("  First Column", if info.is_first_column { "yes" } else { "no" }, 2);
                output::key_value("  Last Column", if info.is_last_column { "yes" } else { "no" }, 2);
                output::key_value("  Filter Button", if info.is_show_filter_button { "yes" } else { "no" }, 2);
                if !info.column_headers.is_empty() {
                    output::key_value("Columns", &info.column_headers.join(", "), 2);
                }
            } else {
                output::error("Failed to parse table info from engine response.");
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── Pivot Table ─────────────────────────────────────────────────────────────

fn cmd_pivot(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: pivot list|create|delete|info|fields|refresh|rename|move|copy|selectfield|changefield|filter|removefilter|filterinfo|sort|removesort|group|dategroup|removegroup|removefield|properties|aggregation|showdataas|changesource|cellinfo|refreshonload [args]");
        return;
    }
    match args[0].to_lowercase().as_str() {
        "list" => pivot_list(engine, session),
        "create" => pivot_create(&args[1..], engine, session),
        "delete" => pivot_delete(&args[1..], engine, session),
        "info" => pivot_info(&args[1..], engine, session),
        "fields" => pivot_fields(&args[1..], engine, session),
        "refresh" => pivot_refresh(&args[1..], engine, session),
        "rename" => pivot_rename(&args[1..], engine, session),
        "move" => pivot_move(&args[1..], engine, session),
        "copy" => pivot_copy(&args[1..], engine, session),
        "selectfield" => pivot_select_field(&args[1..], engine, session),
        "changefield" => pivot_change_field(&args[1..], engine, session),
        "filter" => pivot_apply_filter(&args[1..], engine, session),
        "removefilter" => pivot_remove_filter(&args[1..], engine, session),
        "filterinfo" => pivot_filter_info(&args[1..], engine, session),
        "sort" => pivot_apply_sort(&args[1..], engine, session),
        "removesort" => pivot_remove_sort(&args[1..], engine, session),
        "group" => pivot_apply_grouping(&args[1..], engine, session),
        "dategroup" => pivot_apply_date_grouping(&args[1..], engine, session),
        "removegroup" => pivot_remove_group(&args[1..], engine, session),
        "removefield" => pivot_remove_field(&args[1..], engine, session),
        "properties" => pivot_modify_properties(&args[1..], engine, session),
        "aggregation" => pivot_modify_aggregation(&args[1..], engine, session),
        "showdataas" => pivot_modify_show_data_as(&args[1..], engine, session),
        "changesource" => pivot_change_source(&args[1..], engine, session),
        "cellinfo" => pivot_cell_info(&args[1..], engine, session),
        "refreshonload" => pivot_refresh_on_load(&args[1..], engine, session),
        other => output::error(&format!(
            "Unknown pivot sub-command: '{}'. Use: list, create, delete, info, fields, refresh, rename, move, copy, selectfield, changefield, filter, removefilter, filterinfo, sort, removesort, group, dategroup, removegroup, removefield, properties, aggregation, showdataas, changesource, cellinfo, refreshonload",
            other
        )),
    }
}

fn pivot_list(engine: &EngineHandle, session: &CliSession) {
    let rid = match &session.rid {
        Some(r) => r.clone(),
        None => return,
    };

    // Scan all sheets in the workbook for pivot tables
    let mut total_count = 0usize;
    let mut idx = 0usize;
    let mut seen_ids: Vec<String> = Vec::new();

    for (_si, sheet_id) in session.sheet_ids.iter().enumerate() {
        let fetch_req = rb::build_pivot_list_fetch(&rid, sheet_id);
        let pivots = match engine.fetch_json(&fetch_req) {
            Ok(resp) => rp::parse_pivot_list(&resp),
            Err(_) => Vec::new(),
        };
        if pivots.is_empty() {
            continue;
        }
        for p in &pivots {
            if seen_ids.contains(&p.pivot_id) {
                continue;
            }
            seen_ids.push(p.pivot_id.clone());
            if total_count == 0 {
                output::line("Pivot tables:", 0);
            }
            // Try to get the name and actual sheet where the pivot TABLE lives
            let (name, actual_sheet) = match get_pivot_name_and_sheet(engine, &rid, &session.sheet_ids, &p.pivot_id) {
                Some((n, sid)) => {
                    let sname = session.sheet_ids.iter().position(|s| s == &sid)
                        .and_then(|i| session.sheet_names.get(i))
                        .map(|s| s.as_str())
                        .unwrap_or("?");
                    (if n.is_empty() { None } else { Some(n) }, sname.to_string())
                }
                None => (None, "?".to_string()),
            };
            let start = format!("{}{}", cell_ref::col_to_letter(p.start_col), p.start_row + 1);
            let end = format!("{}{}", cell_ref::col_to_letter(p.end_col), p.end_row + 1);
            let empty_marker = if p.is_empty { " (empty)" } else { "" };
            if let Some(n) = name {
                output::line(&format!("  [{}] {} ({})  {}:{}  [{}]{}", idx, n, p.pivot_id, start, end, actual_sheet, empty_marker), 0);
            } else {
                output::line(&format!("  [{}] {}  {}:{}  [{}]{}", idx, p.pivot_id, start, end, actual_sheet, empty_marker), 0);
            }
            idx += 1;
            total_count += 1;
        }
    }

    if total_count == 0 {
        output::info("No pivot tables found in this workbook.");
    }
}

/// Attempts to get the pivot name by calling pivot_table_info on all known sheets.
/// Returns (name, actual_sheet_id) if found.
fn get_pivot_name_and_sheet(
    engine: &EngineHandle,
    rid: &str,
    sheet_ids: &[String],
    pivot_id: &str,
) -> Option<(String, String)> {
    for sid in sheet_ids {
        let request = rb::build_pivot_table_info(rid, sid, pivot_id);
        if let Ok(resp) = engine.process_request_json(&request) {
            if let Some(info) = rp::parse_pivot_table_info(&resp) {
                if rp::is_success(info.status_code) {
                    let name = if info.pivot_name.is_empty() {
                        None
                    } else {
                        Some(info.pivot_name)
                    };
                    return Some((name.unwrap_or_default(), sid.clone()));
                }
            }
        }
    }
    None
}

/// Resolves a pivot name or ID to the actual pivot_id and the sheet_id where it lives.
/// Tries pivot_table_info on each known sheet to find the correct one.
/// Returns (pivot_id, sheet_id) or None with an error message.
fn resolve_pivot_id(
    name_or_id: &str,
    engine: &EngineHandle,
    session: &CliSession,
) -> Option<(String, String)> {
    let rid = session.rid.as_deref()?;

    // First, try pivot_table_info directly with the input as an ID on each sheet
    for sid in &session.sheet_ids {
        let request = rb::build_pivot_table_info(rid, sid, name_or_id);
        if let Ok(resp) = engine.process_request_json(&request) {
            // eprintln!("[DEBUG resolve] pivot_table_info on sheet {:?} => {}", sid, &resp[..resp.len().min(500)]);
            if let Some(info) = rp::parse_pivot_table_info(&resp) {
                // eprintln!("[DEBUG resolve] parsed: status={}, pivot_id={:?}, name={:?}", info.status_code, info.pivot_id, info.pivot_name);
                if rp::is_success(info.status_code) {
                    // Engine may not echo back pivot_id; use input if empty
                    let resolved_id = if info.pivot_id.is_empty() {
                        name_or_id.to_string()
                    } else {
                        info.pivot_id
                    };
                    return Some((resolved_id, sid.clone()));
                }
            }
        }
    }

    // If that failed, search by name: fetch pivot lists, get names, match
    for sid in &session.sheet_ids {
        let fetch_req = rb::build_pivot_list_fetch(rid, sid);
        let pivots = match engine.fetch_json(&fetch_req) {
            Ok(resp) => rp::parse_pivot_list(&resp),
            Err(_) => continue,
        };
        for p in &pivots {
            // Try to get the name for this pivot by querying all sheets
            for target_sid in &session.sheet_ids {
                let req = rb::build_pivot_table_info(rid, target_sid, &p.pivot_id);
                if let Ok(resp) = engine.process_request_json(&req) {
                    if let Some(info) = rp::parse_pivot_table_info(&resp) {
                        if rp::is_success(info.status_code) {
                            if info.pivot_name.eq_ignore_ascii_case(name_or_id) {
                                let resolved_id = if info.pivot_id.is_empty() {
                                    p.pivot_id.clone()
                                } else {
                                    info.pivot_id
                                };
                                return Some((resolved_id, target_sid.clone()));
                            }
                            break; // Found the sheet for this pivot, name didn't match
                        }
                    }
                }
            }
        }
    }

    // If nothing matched, output an error
    output::error(&format!(
        "Pivot table '{}' not found. Use 'pivot list' to see available pivot tables.",
        name_or_id
    ));
    None
}

fn pivot_create(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: pivot create <range|Sheet!Range> [--newsheet | --dest <destCell>] [--name <name>]");
        return;
    }
    let rid = session.rid.as_deref().unwrap();
    let dest_sid = session.get_active_sheet_id_or_default();

    // Support cross-sheet references: 'Sheet Name'!A1:Q51 or SheetName!A1:Q51
    let (source_sid, sc, sr, ec, er) = if let Some((sheet_part, range_part)) = parse_sheet_range_ref(args[0]) {
        // Resolve the sheet name to an ID
        let found = session.sheet_names.iter().position(|s| s.eq_ignore_ascii_case(&sheet_part));
        match found {
            Some(i) => {
                let id = session.sheet_ids.get(i).cloned().unwrap_or_else(|| i.to_string());
                match cell_ref::try_parse_range(range_part) {
                    Some((sc, sr, ec, er)) => (id, sc, sr, ec, er),
                    None => {
                        output::error(&format!("Invalid range: '{}'. Use A1:C5 format.", range_part));
                        return;
                    }
                }
            }
            None => {
                output::error(&format!("Sheet '{}' not found.", sheet_part));
                return;
            }
        }
    } else {
        let (sc, sr, ec, er) = parse_range_arg!(args[0]);
        (dest_sid.clone(), sc, sr, ec, er)
    };

    let has_newsheet = args.iter().any(|a| a.eq_ignore_ascii_case("--newsheet"));
    let dest_pos = args.iter().position(|a| a.eq_ignore_ascii_case("--dest"));

    let request = if let Some(pos) = dest_pos {
        if pos + 1 >= args.len() {
            output::error("Usage: pivot create <range> --dest <destCell>");
            return;
        }
        let dest_cell = args[pos + 1];
        let (dest_col, dest_row) = match cell_ref::try_parse(dest_cell) {
            Some(v) => v,
            None => {
                output::error(&format!("Invalid destination cell: '{}'. Use A1 format.", dest_cell));
                return;
            }
        };
        rb::build_create_pivot_table_at_dest(rid, &source_sid, sr, sc, er, ec, &dest_sid, dest_row, dest_col)
    } else {
        rb::build_create_pivot_table_new_sheet(rid, &source_sid, sr, sc, er, ec)
    };

    match engine.process_request_json(&request) {
        Ok(resp) => {
            // eprintln!("[DEBUG pivot_create] request={}", request);
            // eprintln!("[DEBUG pivot_create] response={}", &resp[..resp.len().min(500)]);
            let (status_code, status_message, pivot_id) = rp::parse_create_pivot_table(&resp);
            if rp::is_success(status_code) {
                output::success("Pivot table created.");
                if let Some(ref id) = pivot_id {
                    output::key_value("Pivot ID", id, 2);
                }
                session.is_dirty = true;
                // Refresh sheet list to pick up any newly created sheets
                let old_count = session.sheet_names.len();
                refresh_sheet_list(engine, session);
                if session.sheet_names.len() > old_count {
                    let new_name = &session.sheet_names[session.sheet_names.len() - 1];
                    output::key_value("New Sheet", new_name, 2);
                }
                // --newsheet causes the engine to activate the new pivot sheet server-side.
                // Re-anchor the session to the source sheet so subsequent bare-range commands
                // still target the original data sheet, not the newly created pivot sheet.
                if has_newsheet {
                    if let Some(src_idx) = session.sheet_ids.iter().position(|id| id == &source_sid) {
                        session.active_sheet_index = src_idx;
                        session.active_sheet_name = Some(session.sheet_names[src_idx].clone());
                    }
                }
                // Auto-rename if --name was provided
                if let Some(name_pos) = args.iter().position(|a| a.eq_ignore_ascii_case("--name")) {
                    if name_pos + 1 < args.len() {
                        if let Some(ref id) = pivot_id {
                            // Re-borrow rid since refresh_sheet_list may have invalidated the old one
                            let rid2 = session.rid.as_deref().unwrap().to_string();
                            for sid in &session.sheet_ids {
                                let rename_req = rb::build_edit_pivot_name(&rid2, sid, id, args[name_pos + 1]);
                                if let Ok(rename_resp) = engine.process_request_json(&rename_req) {
                                    let result = rp::parse_status_response(&rename_resp);
                                    if rp::is_success(result.status_code) {
                                        output::key_value("Name", args[name_pos + 1], 2);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            } else {
                output::error(&format!(
                    "Failed to create pivot table: {}",
                    status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_delete(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: pivot delete <pivotId|pivotName>");
        return;
    }
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_delete_pivot_table(rid, &pivot_sid, &pivot_id);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Pivot table '{}' deleted.", pivot_id));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to delete pivot table: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_info(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: pivot info <pivotId|pivotName>");
        return;
    }
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_pivot_table_info(rid, &pivot_sid, &pivot_id);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            if let Some(info) = rp::parse_pivot_table_info(&resp) {
                if rp::is_success(info.status_code) {
                    // Use resolved values since engine doesn't echo them back
                    let sheet_name = session.sheet_ids.iter().position(|s| s == &pivot_sid)
                        .and_then(|i| session.sheet_names.get(i))
                        .map(|s| s.as_str())
                        .unwrap_or("?");
                    output::success("Pivot table info:");
                    output::key_value("Name", &info.pivot_name, 2);
                    output::key_value("Pivot ID", &pivot_id, 2);
                    output::key_value("Sheet", sheet_name, 2);
                    if !info.source_range.is_empty() {
                        output::key_value("Source Range", &info.source_range, 2);
                    }
                    if !info.headers.is_empty() {
                        output::line(&format!("  Headers ({}):", info.headers.len()), 0);
                        for (i, (name, dtype)) in info.headers.iter().enumerate() {
                            output::line(&format!("    [{}] {} ({})", i, name, dtype.to_lowercase()), 0);
                        }
                    }
                } else {
                    output::error(&format!(
                        "Failed to get pivot info: {}",
                        info.status_message.unwrap_or_else(|| "engine error".into())
                    ));
                }
            } else {
                output::error("Failed to parse pivot table info from engine response.");
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_fields(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: pivot fields <pivotId|pivotName>");
        return;
    }
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_pivot_table_info(rid, &pivot_sid, &pivot_id);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            if let Some(info) = rp::parse_pivot_table_info(&resp) {
                if rp::is_success(info.status_code) {
                    if info.headers.is_empty() {
                        output::info("No fields found in this pivot table.");
                    } else {
                        output::line(&format!("Fields in pivot '{}' ({} total):", if info.pivot_name.is_empty() { &pivot_id } else { &info.pivot_name }, info.headers.len()), 0);
                        output::line("", 0);
                        output::line("  Idx  Name                      Type", 0);
                        output::line("  ---  ------------------------  ----------", 0);
                        for (i, (name, dtype)) in info.headers.iter().enumerate() {
                            output::line(&format!("  {:>3}  {:<24}  {}", i, name, dtype.to_lowercase()), 0);
                        }
                        output::line("", 0);
                        output::info("Use <idx> with selectfield, changefield, filter, sort, group, removefield, aggregation, showdataas.");
                    }
                } else {
                    output::error(&format!(
                        "Failed to get pivot fields: {}",
                        info.status_message.unwrap_or_else(|| "engine error".into())
                    ));
                }
            } else {
                output::error("Failed to parse pivot table info from engine response.");
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_refresh(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: pivot refresh <pivotId|pivotName>");
        return;
    }
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_refresh_pivot_table(rid, &pivot_sid, &pivot_id);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Pivot table '{}' refreshed.", pivot_id));
            } else {
                output::error(&format!(
                    "Failed to refresh pivot table: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_rename(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: pivot rename <pivotId|pivotName> <newName>");
        return;
    }
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let new_name = args[1];
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_edit_pivot_name(rid, &pivot_sid, &pivot_id, new_name);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Pivot table renamed to '{}'.", new_name));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to rename pivot table: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_move(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: pivot move <pivotId|pivotName> <destCell> [--sheet <sheetName>]");
        return;
    }
    let dest_cell = args[1];
    let (dest_col, dest_row) = match cell_ref::try_parse(dest_cell) {
        Some(v) => v,
        None => {
            output::error(&format!("Invalid destination cell: '{}'. Use A1 format.", dest_cell));
            return;
        }
    };

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let dest_sheet_id = if let Some(pos) = args.iter().position(|a| a.eq_ignore_ascii_case("--sheet")) {
        if pos + 1 >= args.len() {
            output::error("Usage: pivot move <pivotId|pivotName> <destCell> --sheet <sheetName>");
            return;
        }
        match resolve_sheet_id(args[pos + 1], session) {
            Some((id, _name)) => id,
            None => return,
        }
    } else {
        sid.clone()
    };

    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };

    let request = rb::build_move_pivot_table(rid, &pivot_sid, &pivot_id, &dest_sheet_id, dest_row, dest_col);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Pivot table '{}' moved to {}.", pivot_id, dest_cell.to_uppercase()));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to move pivot table: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_copy(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: pivot copy <pivotId|pivotName> <destCell> [--sheet <sheetName>]");
        return;
    }
    let dest_cell = args[1];
    let (dest_col, dest_row) = match cell_ref::try_parse(dest_cell) {
        Some(v) => v,
        None => {
            output::error(&format!("Invalid destination cell: '{}'. Use A1 format.", dest_cell));
            return;
        }
    };

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let dest_sheet_id = if let Some(pos) = args.iter().position(|a| a.eq_ignore_ascii_case("--sheet")) {
        if pos + 1 >= args.len() {
            output::error("Usage: pivot copy <pivotId|pivotName> <destCell> --sheet <sheetName>");
            return;
        }
        match resolve_sheet_id(args[pos + 1], session) {
            Some((id, _name)) => id,
            None => return,
        }
    } else {
        sid.clone()
    };

    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };

    let request = rb::build_copy_pivot_table(rid, &pivot_sid, &pivot_id, &dest_sheet_id, dest_row, dest_col);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Pivot table '{}' copied to {}.", pivot_id, dest_cell.to_uppercase()));
                session.is_dirty = true;
            } else {
                let message = result.status_message.unwrap_or_else(|| "engine error".into());
                output::error(&format!(
                    "Failed to copy pivot table: {}",
                    message
                ));
                if message.to_ascii_lowercase().contains("overrides data") {
                    output::info("Hint: destination overlaps existing data or another pivot output range.");
                    output::info("Use an empty top-left destination cell (for example, A20) or a different sheet.");
                }
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_select_field(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: pivot selectfield <pivot> <headerIdx> <area> [fieldIdx]");
        output::info("  <area>: row, column, value, filter, none (or 0-4)");
        return;
    }
    let header_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => {
            output::error("Invalid headerIdx: must be a number.");
            return;
        }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => {
            output::error("Invalid area: use row, column, value, filter, none (or 0-4).");
            return;
        }
    };
    let field_index: i32 = if args.len() > 3 {
        match args[3].parse() {
            Ok(v) => v,
            Err(_) => {
                output::error("Invalid fieldIdx: must be a number.");
                return;
            }
        }
    } else {
        0
    };

    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };

    let rid = session.rid.as_deref().unwrap();
    // eprintln!("[DEBUG selectfield] pivot_id={:?}, sheet_id={:?}, rid={:?}", pivot_id, pivot_sid, rid);
    let request = rb::build_select_pivot_field(rid, &pivot_sid, &pivot_id, header_index, field_type, field_index);
    // eprintln!("[DEBUG selectfield] request={}", request);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            // eprintln!("[DEBUG selectfield] response={}", resp);
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot field selected.");
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to select pivot field: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_change_field(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 5 {
        output::error("Usage: pivot changefield <pivot> <fieldIdx> <fromArea> <destIdx> <toArea>");
        output::info("  Move a field from one area/position to another.");
        output::info("  Areas: row, column, value, filter, none (or 0-4)");
        output::info("  Example: pivot changefield MyPivot 0 row 0 column");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => {
            output::error("Invalid fieldIdx: must be a number.");
            return;
        }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => {
            output::error("Invalid fromArea: use row, column, value, filter, none (or 0-4).");
            return;
        }
    };
    let dest_index: i32 = match args[3].parse() {
        Ok(v) => v,
        Err(_) => {
            output::error("Invalid destIdx: must be a number.");
            return;
        }
    };
    let dest_type: i32 = match parse_pivot_area(args[4]) {
        Some(v) => v,
        None => {
            output::error("Invalid toArea: use row, column, value, filter, none (or 0-4).");
            return;
        }
    };

    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };

    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_change_pivot_field_type(rid, &pivot_sid, &pivot_id, field_index, field_type, dest_index, dest_type);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot field type changed.");
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to change pivot field type: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── New pivot subcommands ───────────────────────────────────────────────────

fn pivot_response_has_failure_hint(resp: &str, status_message: Option<&str>) -> bool {
    let mut combined = String::new();
    combined.push_str(resp);
    if let Some(msg) = status_message {
        if !msg.is_empty() {
            combined.push(' ');
            combined.push_str(msg);
        }
    }
    let hay = combined.to_ascii_lowercase();
    hay.contains("parser error")
        || hay.contains("syntax error")
        || hay.contains("duckdbfailure")
        || hay.contains("error in creating filteredoutputtable")
        || hay.contains("error in preparing statement")
}

fn pivot_apply_filter(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    // Two modes: condition-based or selection-based
    // Condition: pivot filter <pivot> <fieldIdx> <area> --condition <operator> <val1> [val2] [--valuefield <idx>]
    // Selection: pivot filter <pivot> <fieldIdx> <area> --selection <idx1,idx2,...>
    if args.len() < 4 {
        output::error("Usage: pivot filter <pivot> <fieldIdx> <area> --condition <operator> <val1> [val2] [--valuefield <idx>]");
        output::info("   or: pivot filter <pivot> <fieldIdx> <area> --selection <idx1,idx2,...>");
        output::info("  Operators: equals, notequals, greaterthan, gte, lessthan, lte, between, notbetween");
        output::info("             top10, bottom10, top10percent, bottom10percent, top10sum, bottom10sum");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid area."); return; }
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();

    let request = if args.len() > 3 && args[3].eq_ignore_ascii_case("--selection") {
        if args.len() < 5 {
            output::error("Usage: pivot filter <pivot> <fieldIdx> <area> --selection <idx1,idx2,...>");
            return;
        }
        let indices: Vec<i32> = args[4].split(',')
            .filter_map(|s| s.trim().parse::<i32>().ok())
            .collect();
        if indices.is_empty() {
            output::error("Invalid selection indices.");
            return;
        }
        rb::build_apply_pivot_filter_selection(rid, &pivot_sid, &pivot_id, field_index, field_type, indices)
    } else if args.len() > 3 && args[3].eq_ignore_ascii_case("--condition") {
        if args.len() < 6 {
            output::error("Usage: pivot filter <pivot> <fieldIdx> <area> --condition <operator> <val1> [val2] [--valuefield <idx>]");
            return;
        }
        let operator = args[4];
        let (criteria_id, sub_criteria_id) = match parse_filter_condition(operator) {
            Some(v) => v,
            None => {
                output::error(&format!("Unknown operator '{}'. Number: equals, notequals, gt, gte, lt, lte, between, top, bottom. Text: contains, beginswith, endswith, ... Date: after, before, ...", operator));
                return;
            }
        };
        // Engine currently emits malformed SQL for date criteria on row/column/value
        // filter targets (DuckDB parser error). Block these combinations early.
        if criteria_id == 2 && matches!(field_type, 0 | 1 | 2) {
            output::error("Date conditions are currently supported only for pivot filter-area fields. Use area 'filter' for date conditions.");
            output::info("Hint: move the date field into pivot filter area first, then apply the date condition.");
            output::info("Example: pivot selectfield <pivot> <dateFieldHeaderIdx> filter");
            output::info("         pivot filter <pivot> 0 filter --condition onorbefore 2024-07-01");
            return;
        }
        let val1 = args[5];
        let needs_two = filter_condition_needs_two_values(operator);
        let (val2, remaining_start) = if needs_two {
            if args.len() < 7 {
                output::error(&format!("'{}' requires two values: <val1> <val2>", operator));
                return;
            }
            (args[6], 7)
        } else {
            ("", 6)
        };
        let value_field_index: i32 = if let Some(pos) = args[remaining_start..].iter().position(|a| a.eq_ignore_ascii_case("--valuefield")) {
            let idx_pos = remaining_start + pos + 1;
            if idx_pos >= args.len() {
                output::error("--valuefield requires a field index.");
                return;
            }
            match args[idx_pos].parse() {
                Ok(v) => v,
                Err(_) => { output::error("Invalid --valuefield index."); return; }
            }
        } else {
            0
        };
        rb::build_apply_pivot_filter_condition(rid, &pivot_sid, &pivot_id, field_index, field_type, criteria_id, sub_criteria_id, val1, val2, value_field_index)
    } else {
        output::error("Use --condition or --selection flag.");
        return;
    };

    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            let status_message = result.status_message.clone();
            let has_failure_hint = pivot_response_has_failure_hint(&resp, status_message.as_deref());
            if rp::is_success(result.status_code) && !has_failure_hint {
                output::success("Pivot filter applied.");
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to apply pivot filter: {}",
                    status_message.unwrap_or_else(|| {
                        if has_failure_hint {
                            "engine reported SQL/parser failure".into()
                        } else {
                            "engine error".into()
                        }
                    })
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_remove_filter(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: pivot removefilter <pivot> <fieldIdx> <area>");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid area."); return; }
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_remove_pivot_filter(rid, &pivot_sid, &pivot_id, field_index, field_type);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot filter removed.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to remove pivot filter: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_filter_info(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: pivot filterinfo <cell>");
        return;
    }
    let (col, row) = match cell_ref::try_parse(args[0]) {
        Some(v) => v,
        None => { output::error(&format!("Invalid cell: '{}'. Use A1 format.", args[0])); return; }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_pivot_filter_info(rid, &sid, row, col);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            match rp::parse_pivot_filter_info(&resp) {
                Some(info) => {
                    if !rp::is_success(info.status_code) {
                        output::error(&format!(
                            "Failed to get pivot filter info: {}",
                            info.status_message.unwrap_or_else(|| "engine error".into())
                        ));
                        return;
                    }

                    output::success("Pivot filter info:");
                    if !info.pivot_id.is_empty() {
                        output::key_value("Pivot ID", &info.pivot_id, 2);
                    }

                    let active_type = if info.active_filter_type.is_empty() {
                        "UNKNOWN"
                    } else {
                        info.active_filter_type.as_str()
                    };
                    output::key_value("Active Type", active_type, 2);

                    if !info.label_field_name.is_empty() {
                        output::key_value(
                            "Label Field",
                            &format!(
                                "{} (idx: {}, area: {})",
                                info.label_field_name,
                                info.label_field_index,
                                if info.label_field_type.is_empty() { "?" } else { info.label_field_type.as_str() }
                            ),
                            2,
                        );
                    }

                    if !info.value_field_info_list.is_empty() {
                        output::line("  Value Fields:", 0);
                        for (i, name) in info.value_field_info_list.iter().enumerate() {
                            output::line(&format!("    [{}] {}", i, name), 0);
                        }
                    }

                    if !info.column_data.is_empty() {
                        output::line(&format!("  Items ({}):", info.column_data.len()), 0);
                        for (i, item) in info.column_data.iter().enumerate() {
                            let mark = info.check_mark_vector.get(i).copied().unwrap_or(0);
                            let checked = if mark != 0 { "x" } else { " " };
                            output::line(&format!("    [{}] [{}] {}", i, checked, item), 0);
                        }
                    }

                    if let Some(cond) = info.condition {
                        output::line("  Condition:", 0);
                        output::line(
                            &format!(
                                "    criteria_id={} sub_criteria_id={} val1='{}'{}",
                                cond.criteria_id,
                                cond.sub_criteria_id,
                                cond.val1,
                                if cond.val2.is_empty() {
                                    "".to_string()
                                } else {
                                    format!(" val2='{}'", cond.val2)
                                }
                            ),
                            0,
                        );

                        if info.custom_filter_value_field_index >= 0 {
                            output::key_value(
                                "Condition Value Field Index",
                                &info.custom_filter_value_field_index.to_string(),
                                2,
                            );
                        }
                    }
                }
                None => output::error("Failed to parse pivot filter info response."),
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_apply_sort(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 4 {
        output::error("Usage: pivot sort <pivot> <fieldIdx> <area> <asc|desc> [sortAggIdx]");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid area."); return; }
    };
    let is_asc = match args[3].to_lowercase().as_str() {
        "asc" | "ascending" | "true" | "1" => true,
        "desc" | "descending" | "false" | "0" => false,
        _ => { output::error("Invalid sort order: use 'asc' or 'desc'."); return; }
    };
    let sort_agg_idx: i32 = if args.len() > 4 {
        match args[4].parse() { Ok(v) => v, Err(_) => 0 }
    } else { 0 };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_apply_pivot_sort(rid, &pivot_sid, &pivot_id, field_index, field_type, is_asc, sort_agg_idx);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot sort applied.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to apply pivot sort: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_remove_sort(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: pivot removesort <pivot> <fieldIdx> <area>");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid area."); return; }
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_remove_pivot_sort(rid, &pivot_sid, &pivot_id, field_index, field_type);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot sort removed.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to remove pivot sort: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_apply_grouping(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 6 {
        output::error("Usage: pivot group <pivot> <fieldIdx> <area> <min> <max> <range> [--mindefault] [--maxdefault]");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid area."); return; }
    };
    let minimum: f64 = match args[3].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid min value."); return; }
    };
    let maximum: f64 = match args[4].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid max value."); return; }
    };
    let range: f64 = match args[5].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid range value."); return; }
    };
    let is_min_default = args.iter().any(|a| a.eq_ignore_ascii_case("--mindefault"));
    let is_max_default = args.iter().any(|a| a.eq_ignore_ascii_case("--maxdefault"));

    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_apply_pivot_grouping(rid, &pivot_sid, &pivot_id, field_index, field_type, minimum, maximum, range, is_min_default, is_max_default);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot grouping applied.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to apply pivot grouping: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_apply_date_grouping(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 7 {
        output::error("Usage: pivot dategroup <pivot> <fieldIdx> <area> <types> <min> <max> <days> [--mindefault] [--maxdefault]");
        output::info("  <types>: comma-separated values from year,quarter,month,day,hour,minute,second");
        output::info("           examples: month   |   year,month   |   year,quarter,month");
        output::info("  <min>/<max>: date bounds in YYYY-MM-DD (or date serial number)");
        output::info("  <days>: day interval used when 'day' is included in <types>");
        output::info("  Areas: row|column|value|filter|none (or 0-4)");
        output::info("  Example: pivot dategroup SalesByRegion 0 row year,month 2025-01-01 2025-12-31 1");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid area."); return; }
    };
    let date_grouping_types: Vec<i32> = args[3].split(',')
        .filter_map(|s| parse_date_grouping_type(s.trim()))
        .collect();
    if date_grouping_types.is_empty() {
        output::error("Invalid date grouping types. Use comma-separated: year,month,day (or 0,1,2,...)");
        return;
    }
    let minimum: f64 = match date_serial::parse_date_or_number(args[4]) {
        Some(v) => v,
        None => { output::error("Invalid min value. Use YYYY-MM-DD or a serial number."); return; }
    };
    let maximum: f64 = match date_serial::parse_date_or_number(args[5]) {
        Some(v) => v,
        None => { output::error("Invalid max value. Use YYYY-MM-DD or a serial number."); return; }
    };
    let no_of_days: i32 = match args[6].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid days value."); return; }
    };
    let is_min_default = args.iter().any(|a| a.eq_ignore_ascii_case("--mindefault"));
    let is_max_default = args.iter().any(|a| a.eq_ignore_ascii_case("--maxdefault"));

    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_apply_pivot_date_grouping(rid, &pivot_sid, &pivot_id, field_index, field_type, date_grouping_types, minimum, maximum, is_min_default, is_max_default, no_of_days);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot date grouping applied.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to apply pivot date grouping: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_remove_group(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: pivot removegroup <pivot> <fieldIdx> <area>");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid area."); return; }
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_remove_group(rid, &pivot_sid, &pivot_id, field_index, field_type);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot grouping removed.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to remove pivot grouping: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_remove_field(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: pivot removefield <pivot> <fieldIdx> <area>");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let field_type: i32 = match parse_pivot_area(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid area."); return; }
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_remove_pivot_field(rid, &pivot_sid, &pivot_id, field_index, field_type);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot field removed.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to remove pivot field: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_modify_properties(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: pivot properties <pivot> <property> <true|false>");
        output::info("  Properties: subtotal, rowtotal, coltotal, repeat, hideerrors");
        return;
    }
    let pivot_property: i32 = match parse_pivot_property(args[1]) {
        Some(v) => v,
        None => { output::error("Invalid property. Use: subtotal, rowtotal, coltotal, repeat, hideerrors (or 0-4)."); return; }
    };
    let is_enabled = match args[2].to_lowercase().as_str() {
        "true" | "1" | "yes" | "on" => true,
        "false" | "0" | "no" | "off" => false,
        _ => { output::error("Invalid value: use true/false."); return; }
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_modify_pivot_properties(rid, &pivot_sid, &pivot_id, pivot_property, is_enabled);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot property modified.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to modify pivot property: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_modify_aggregation(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: pivot aggregation <pivot> <fieldIdx> <type>");
        output::info("  Types: sum, count, countnums, distinct, avg, min, max, median, product, stdev, stdevp, var, varp");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let summarise_by: i32 = match parse_aggregation_type(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid aggregation type. Use: sum, count, countnums, distinct, avg, min, max, median, product, stdev, stdevp, var, varp (or 0-12)."); return; }
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_modify_value_aggregation_type(rid, &pivot_sid, &pivot_id, field_index, summarise_by);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot aggregation type modified.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to modify aggregation type: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_modify_show_data_as(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: pivot showdataas <pivot> <fieldIdx> <type>");
        output::info("  Types: nochange, percent_row, percent_col, percent_total");
        return;
    }
    let field_index: i32 = match args[1].parse() {
        Ok(v) => v,
        Err(_) => { output::error("Invalid fieldIdx."); return; }
    };
    let show_data_as: i32 = match parse_show_data_as(args[2]) {
        Some(v) => v,
        None => { output::error("Invalid show_data_as type. Use: nochange, percent_row, percent_col, percent_total (or 0-3)."); return; }
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_modify_value_show_data_as(rid, &pivot_sid, &pivot_id, field_index, show_data_as);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot show data as modified.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to modify show data as: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_change_source(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: pivot changesource <pivot> <range> [--sheet <sheetName>]");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let dest_sheet_id = if let Some(pos) = args.iter().position(|a| a.eq_ignore_ascii_case("--sheet")) {
        if pos + 1 >= args.len() {
            output::error("Usage: pivot changesource <pivot> <range> --sheet <sheetName>");
            return;
        }
        match resolve_sheet_id(args[pos + 1], session) {
            Some((id, _name)) => id,
            None => return,
        }
    } else {
        session.get_active_sheet_id_or_default()
    };
    let (pivot_id, pivot_sid) = match resolve_pivot_id(args[0], engine, session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_change_pivot_table_source(rid, &pivot_sid, &pivot_id, &dest_sheet_id, sr, er, sc, ec);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success("Pivot table source changed.");
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed to change pivot source: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_cell_info(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 1 {
        output::error("Usage: pivot cellinfo <cell>");
        return;
    }
    let (col, row) = match cell_ref::try_parse(args[0]) {
        Some(v) => v,
        None => { output::error(&format!("Invalid cell: '{}'. Use A1 format.", args[0])); return; }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_pivot_cell_info(rid, &sid, row, col);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            match rp::parse_pivot_cell_info(&resp) {
                Some(info) => {
                    if !rp::is_success(info.status_code) {
                        output::error(&format!(
                            "Failed to get pivot cell info: {}",
                            info.status_message.unwrap_or_else(|| "engine error".into())
                        ));
                        return;
                    }
                    output::success("Pivot cell info:");
                    output::key_value("Cell", &args[0].to_uppercase(), 2);
                    if !info.pivot_id.is_empty() {
                        output::key_value("Pivot ID", &info.pivot_id, 2);
                    } else {
                        output::key_value("Pivot ID", "(not returned)", 2);
                    }
                }
                None => output::error("Failed to parse pivot cell info response."),
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn pivot_refresh_on_load(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    // Accept explicit state: true/false. If omitted, default to true for backward compat.
    let enable = if args.is_empty() {
        true
    } else {
        match args[0].to_lowercase().as_str() {
            "true" | "1" | "yes" | "on" => true,
            "false" | "0" | "no" | "off" => false,
            _ => {
                output::error("Usage: pivot refreshonload <true|false>");
                return;
            }
        }
    };

    if !enable {
        output::success("Pivot refresh-on-load disabled (no pivots will be auto-refreshed on open).");
        return;
    }

    let rid = session.rid.as_deref().unwrap().to_string();
    let mut refreshed = 0usize;
    let mut seen_ids: Vec<String> = Vec::new();

    for sheet_id in session.sheet_ids.clone().iter() {
        let fetch_req = rb::build_pivot_list_fetch(&rid, sheet_id);
        let pivots = match engine.fetch_json(&fetch_req) {
            Ok(resp) => rp::parse_pivot_list(&resp),
            Err(_) => Vec::new(),
        };
        for p in &pivots {
            if seen_ids.contains(&p.pivot_id) {
                continue;
            }
            seen_ids.push(p.pivot_id.clone());
            let request = rb::build_refresh_pivot_table(&rid, sheet_id, &p.pivot_id);
            if let Ok(resp) = engine.process_request_json(&request) {
                let result = rp::parse_status_response(&resp);
                if rp::is_success(result.status_code) {
                    refreshed += 1;
                }
            }
        }
    }

    if refreshed > 0 {
        output::success(&format!("Refreshed {} pivot table(s) (refresh-on-load enabled).", refreshed));
    } else {
        output::info("No pivot tables to refresh (refresh-on-load enabled).");
    }
    session.is_dirty = true;
}

// ─── Helper macros & functions ───────────────────────────────────────────────

/// Parses a chart type from a string — accepts names or numeric values (0-16).
fn parse_chart_type(s: &str) -> Option<i32> {
    match s.to_ascii_lowercase().as_str() {
        "bar" => Some(0),
        "column" | "col" => Some(1),
        "line" => Some(2),
        "pie" => Some(3),
        "area" => Some(4),
        "scatter" | "xy" => Some(5),
        "race" => Some(6),
        "waterfall" => Some(7),
        "bullet" => Some(8),
        "funnel" => Some(9),
        "pareto" => Some(10),
        "histogram" | "hist" => Some(11),
        "stock" => Some(12),
        "radar" => Some(13),
        "wordcloud" => Some(14),
        "combo" => Some(15),
        "boxplot" | "box" => Some(16),
        _ => s.parse::<i32>().ok().filter(|&v| (0..=16).contains(&v)),
    }
}

/// Parses a combined chart type_subtype string (e.g. "bar_stacked") and returns (type, subtype).
/// For chart types without variants, subtype is omitted (None).
fn parse_chart_type_subtype(s: &str) -> Option<(i32, Option<i32>)> {
    let input = s.trim().to_ascii_lowercase();
    let tuple_input = input
        .strip_prefix('(')
        .and_then(|v| v.strip_suffix(')'))
        .unwrap_or(input.as_str())
        .trim();

    // Accept numeric forms: "7", "(7)", "7,1", "(7,1)".
    if let Some((chart_type, chart_sub_type)) = tuple_input.split_once(',') {
        let chart_type = chart_type.trim().parse::<i32>().ok()?;
        let chart_sub_type = chart_sub_type.trim().parse::<i32>().ok()?;
        if (0..=16).contains(&chart_type) {
            return Some((chart_type, Some(chart_sub_type)));
        }
        return None;
    }

    if let Ok(chart_type) = tuple_input.parse::<i32>() {
        if (0..=16).contains(&chart_type) {
            // For one-sized tuple/numeric inputs like "(7)" or "7", omit subtype in request payload.
            return Some((chart_type, None));
        }
        return None;
    }

    match input.as_str() {
        // BAR (type=0)
        "bar" | "bar_default" => Some((0, Some(0))),
        "bar_stacked" => Some((0, Some(1))),
        "bar_stacked_100" | "bar_stacked_100_percent" => Some((0, Some(2))),
        "bar_grouped" => Some((0, Some(3))),
        // COLUMN (type=1)
        "column" | "col" | "column_default" | "col_default" => Some((1, Some(0))),
        "column_stacked" | "col_stacked" => Some((1, Some(1))),
        "column_stacked_100" | "col_stacked_100" | "column_stacked_100_percent" => Some((1, Some(2))),
        "column_grouped" | "col_grouped" => Some((1, Some(3))),
        // LINE (type=2)
        "line" | "line_default" => Some((2, Some(0))),
        "line_spline" | "spline" => Some((2, Some(1))),
        "line_step" | "step" => Some((2, Some(2))),
        "line_timeline" | "timeline" => Some((2, Some(3))),
        // PIE (type=3)
        "pie" | "pie_default" => Some((3, Some(0))),
        "pie_semi" | "semipie" => Some((3, Some(1))),
        "pie_doughnut" | "doughnut" => Some((3, Some(2))),
        "pie_semi_doughnut" | "semi_doughnut" => Some((3, Some(3))),
        "pie_parliament" => Some((3, Some(4))),
        "doughnut_parliament" => Some((3, Some(5))),
        // AREA (type=4)
        "area" | "area_default" => Some((4, Some(0))),
        "area_stacked" => Some((4, Some(1))),
        "area_stacked_100" | "area_stacked_100_percent" => Some((4, Some(2))),
        "area_time" | "timearea" => Some((4, Some(3))),
        // SCATTER (type=5)
        "scatter" | "xy" | "scatter_default" => Some((5, Some(0))),
        "scatter_line" => Some((5, Some(1))),
        "scatter_line_markers" => Some((5, Some(2))),
        "scatter_bubble" | "bubble" => Some((5, Some(3))),
        // RACE (type=6) — no subtypes
        "race" => Some((6, None)),
        // WATERFALL (type=7) — no subtypes
        "waterfall" => Some((7, None)),
        // BULLET (type=8)
        "bullet" | "bullet_horizontal" => Some((8, Some(0))),
        "bullet_vertical" => Some((8, Some(1))),
        // FUNNEL (type=9)
        "funnel" | "funnel_default" => Some((9, Some(0))),
        "funnel_weighted" => Some((9, Some(1))),
        // PARETO (type=10) — no subtypes
        "pareto" => Some((10, None)),
        // HISTOGRAM (type=11) — no subtypes
        "histogram" | "hist" => Some((11, None)),
        // STOCK (type=12)
        "stock" | "stock_candlestick" | "candlestick" => Some((12, Some(0))),
        "stock_ohlc" | "ohlc" => Some((12, Some(1))),
        // RADAR (type=13)
        "radar" | "radar_polar" | "polar" => Some((13, Some(0))),
        "radar_spiderweb" | "spiderweb" => Some((13, Some(1))),
        // WORDCLOUD (type=14) — no subtypes
        "wordcloud" => Some((14, None)),
        // COMBO (type=15) — no subtypes
        "combo" => Some((15, None)),
        // BOXPLOT (type=16)
        "boxplot" | "box" | "boxplot_horizontal" => Some((16, Some(0))),
        "boxplot_grouped_horizontal" => Some((16, Some(1))),
        "boxplot_vertical" => Some((16, Some(2))),
        "boxplot_grouped_vertical" => Some((16, Some(3))),
        _ => None,
    }
}

/// Parses a pivot area type from a string — accepts names (row, column, value, filter, none)
/// or numeric values (0-4). Returns None if invalid.
fn parse_pivot_area(s: &str) -> Option<i32> {
    match s.to_ascii_lowercase().as_str() {
        "row" | "r" => Some(0),
        "column" | "col" | "c" => Some(1),
        "value" | "val" | "v" => Some(2),
        "filter" | "f" => Some(3),
        "none" | "n" | "remove" => Some(4),
        _ => s.parse::<i32>().ok().filter(|&v| (0..=4).contains(&v)),
    }
}

fn parse_aggregation_type(s: &str) -> Option<i32> {
    match s.to_ascii_lowercase().as_str() {
        "sum" => Some(0),
        "count" => Some(1),
        "countnums" | "count_nums" | "count_of_numbers" => Some(2),
        "distinct" | "distinct_count" | "distinctcount" => Some(3),
        "avg" | "average" => Some(4),
        "min" => Some(5),
        "max" => Some(6),
        "median" => Some(7),
        "product" => Some(8),
        "stdev" => Some(9),
        "stdevp" | "stdepv" => Some(10),
        "var" => Some(11),
        "varp" => Some(12),
        _ => s.parse::<i32>().ok().filter(|&v| (0..=12).contains(&v)),
    }
}

fn parse_show_data_as(s: &str) -> Option<i32> {
    match s.to_ascii_lowercase().as_str() {
        "nochange" | "normal" | "none" => Some(0),
        "percent_row" | "pct_row" | "percentage_of_row" => Some(1),
        "percent_col" | "pct_col" | "percentage_of_column" => Some(2),
        "percent_total" | "pct_total" | "percentage_of_grand_total" => Some(3),
        _ => s.parse::<i32>().ok().filter(|&v| (0..=3).contains(&v)),
    }
}

fn parse_pivot_property(s: &str) -> Option<i32> {
    match s.to_ascii_lowercase().as_str() {
        "subtotal" | "sub_total" => Some(0),
        "row_grand_total" | "rowgrandtotal" | "rowtotal" => Some(1),
        "col_grand_total" | "colgrandtotal" | "coltotal" => Some(2),
        "repeat_labels" | "repeatlabels" | "repeat" => Some(3),
        "hide_errors" | "hideerrors" => Some(4),
        _ => s.parse::<i32>().ok().filter(|&v| (0..=4).contains(&v)),
    }
}

fn parse_date_grouping_type(s: &str) -> Option<i32> {
    match s.to_ascii_lowercase().as_str() {
        "year" | "y" => Some(0),
        "quarter" | "q" => Some(1),
        "month" | "m" => Some(2),
        "day" | "d" => Some(3),
        "hour" | "h" => Some(4),
        "minute" | "min" => Some(5),
        "second" | "sec" | "s" => Some(6),
        _ => s.parse::<i32>().ok().filter(|&v| (0..=6).contains(&v)),
    }
}

/// Parses a filter condition name to (criteria_id, sub_criteria_id).
/// Criteria: 0=Number, 1=Text, 2=Date
fn parse_filter_condition(s: &str) -> Option<(i32, i32)> {
    match s.to_ascii_lowercase().as_str() {
        // Number criteria (criteria_id = 0)
        "equals" | "eq" => Some((0, 0)),
        "notequals" | "neq" | "ne" => Some((0, 1)),
        "greaterthan" | "gt" => Some((0, 2)),
        "greaterthanorequal" | "gte" | "ge" => Some((0, 3)),
        "lessthan" | "lt" => Some((0, 4)),
        "lessthanorequal" | "lte" | "le" => Some((0, 5)),
        "between" => Some((0, 6)),
        "top" | "top10" | "topn" => Some((0, 7)),
        "bottom" | "bottom10" | "bottomn" => Some((0, 8)),
        // Text criteria (criteria_id = 1)
        "equalstring" | "texteq" => Some((1, 0)),
        "notequalstring" | "textneq" => Some((1, 1)),
        "beginswith" | "startswith" => Some((1, 2)),
        "notbeginswith" | "notstartswith" => Some((1, 3)),
        "endswith" => Some((1, 4)),
        "notendswith" => Some((1, 5)),
        "contains" => Some((1, 6)),
        "notcontains" | "doesnotcontain" => Some((1, 7)),
        "matchlabel" => Some((1, 8)),
        "notmatchlabel" => Some((1, 9)),
        // Date criteria (criteria_id = 2)
        "equaldate" | "dateeq" => Some((2, 0)),
        "notequaldate" | "dateneq" => Some((2, 1)),
        "afterdate" | "after" => Some((2, 2)),
        "onorafter" => Some((2, 3)),
        "beforedate" | "before" => Some((2, 4)),
        "onorbefore" => Some((2, 5)),
        "betweendate" | "datebetween" => Some((2, 6)),
        _ => None,
    }
}

/// Returns true if the condition requires two values (val1 and val2).
fn filter_condition_needs_two_values(s: &str) -> bool {
    matches!(s.to_ascii_lowercase().as_str(), "between" | "betweendate" | "datebetween")
}

/// Parses a cross-sheet range reference like `'Sheet Name'!A1:Q51` or `SheetName!A1:Q51`.
/// Returns (sheet_name, range_part) if the format matches, or None for a plain range.
fn parse_sheet_range_ref(input: &str) -> Option<(String, &str)> {
    if input.starts_with('\'') {
        // 'Sheet Name'!A1:Q51
        if let Some(end_quote) = input[1..].find('\'') {
            let sheet_name = &input[1..1 + end_quote];
            let rest = &input[1 + end_quote + 1..];
            if rest.starts_with('!') {
                return Some((sheet_name.to_string(), &rest[1..]));
            }
        }
    } else if let Some(bang) = input.find('!') {
        // SheetName!A1:Q51
        let sheet_name = &input[..bang];
        let range_part = &input[bang + 1..];
        if !sheet_name.is_empty() && !range_part.is_empty() {
            return Some((sheet_name.to_string(), range_part));
        }
    }
    None
}

/// Macro to check if session is active, printing error and returning early if not.
macro_rules! require_active {
    ($session:expr) => {
        if !$session.is_active() {
            output::error("No workbook open. Use 'open' first.");
            return;
        }
    };
}
use require_active;

/// Macro to parse a range argument, printing error and returning early on failure.
macro_rules! parse_range_arg {
    ($arg:expr) => {
        match cell_ref::try_parse_range($arg) {
            Some(r) => r,
            None => {
                output::error(&format!("Invalid range: '{}'. Use A1:C5 format.", $arg));
                return;
            }
        }
    };
}
use parse_range_arg;

// ─── Format (Font) ───────────────────────────────────────────────────────────

fn cmd_format(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: format <bold|italic|underline|doubleunderline|strikethrough|superscript|subscript|fontsize|fontcolor|halign|valign|textwrap|rotate|indent|fillcolor|border|numformat|decimal|numpreview|numinfo|nummanage|customformat|default> <range> ...");
        return;
    }
    let sub = args[0].to_lowercase();
    let rest = &args[1..];
    match sub.as_str() {
        "bold" => cmd_format_bool_toggle(rest, engine, session, "bold"),
        "italic" => cmd_format_bool_toggle(rest, engine, session, "italic"),
        "underline" => cmd_format_bool_toggle(rest, engine, session, "underline"),
        "doubleunderline" => cmd_format_bool_toggle(rest, engine, session, "doubleunderline"),
        "strikethrough" => cmd_format_bool_toggle(rest, engine, session, "strikethrough"),
        "superscript" => cmd_format_bool_toggle(rest, engine, session, "superscript"),
        "subscript" => cmd_format_bool_toggle(rest, engine, session, "subscript"),
        "fontsize" => cmd_format_font_size(rest, engine, session),
        "fontcolor" => cmd_format_font_color(rest, engine, session),
        "halign" => cmd_format_halign(rest, engine, session),
        "valign" => cmd_format_valign(rest, engine, session),
        "textwrap" => cmd_format_wrap(rest, engine, session),
        "rotate" => cmd_format_rotate(rest, engine, session),
        "indent" => cmd_format_indent(rest, engine, session),
        "fillcolor" => cmd_format_fill_color(rest, engine, session),
        "border" => cmd_format_border(rest, engine, session),
        "numformat" => {
            // Handle --list-custom as the canonical way to list saved custom formats
            if rest.first().map(|a| *a == "--list-custom").unwrap_or(false) {
                cmd_format_list_custom(engine, session);
            } else if rest.first().map(|a| *a == "--list-currency").unwrap_or(false) {
                cmd_format_list_currency();
            } else {
                cmd_format_numformat(rest, engine, session);
            }
        }
        "decimal" => cmd_format_decimal(rest, engine, session),
        "numpreview" => cmd_format_numpreview(rest, engine, session),
        "numinfo" => cmd_format_numinfo(rest, engine, session),
        "nummanage" => cmd_format_nummanage(engine, session),
        "customformat" => {
            output::info("Warning: 'format customformat' is deprecated. Use 'format numformat --list-custom' instead.");
            cmd_format_list_custom(engine, session);
        }
        "default" => cmd_format_default(rest, engine, session),
        _ => {
            output::error(&format!("Unknown format sub-command: '{}'. Use: bold, italic, underline, doubleunderline, strikethrough, superscript, subscript, fontsize, fontcolor, halign, valign, textwrap, rotate, indent, fillcolor, border, numformat, decimal, numpreview, numinfo, nummanage, customformat, default.", sub));
        }
    }
}

fn cmd_format_bool_toggle(
    args: &[&str],
    engine: &EngineHandle,
    session: &mut CliSession,
    prop: &str,
) {
    if args.len() < 2 {
        output::error(&format!("Usage: format {} <range> <true|false>", prop));
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let value = match args[1].to_lowercase().as_str() {
        "true" | "1" | "on" | "yes" => true,
        "false" | "0" | "off" | "no" => false,
        _ => {
            output::error(&format!("Invalid value '{}'. Use true/false.", args[1]));
            return;
        }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = match prop {
        "bold" => rb::build_set_bold(rid, &sid, value, sr, sc, er, ec),
        "italic" => rb::build_set_italic(rid, &sid, value, sr, sc, er, ec),
        "underline" => rb::build_set_underline(rid, &sid, value, sr, sc, er, ec),
        "doubleunderline" => rb::build_set_double_underline(rid, &sid, value, sr, sc, er, ec),
        "strikethrough" => rb::build_strike_through(rid, &sid, value, sr, sc, er, ec),
        "superscript" => rb::build_set_superscript(rid, &sid, value, sr, sc, er, ec),
        "subscript" => rb::build_set_subscript(rid, &sid, value, sr, sc, er, ec),
        _ => unreachable!(),
    };
    let label = if value { "enabled" } else { "disabled" };
    exec_status_cmd(engine, &request, session, &format!("{} {} on {}.", prop, label, args[0].to_uppercase()));
}

fn cmd_format_font_size(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format fontsize <range> <size>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let size: i32 = match args[1].parse() {
        Ok(v) if v > 0 => v,
        _ => {
            output::error(&format!("Invalid font size '{}'. Must be a positive integer.", args[1]));
            return;
        }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_set_font_size(rid, &sid, size, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Font size set to {} on {}.", size, args[0].to_uppercase()));
}

fn cmd_format_font_color(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: format fontcolor <range> <r> <g> <b>  OR  format fontcolor <range> --auto");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    if args.len() >= 2 && args[1].eq_ignore_ascii_case("--auto") {
        let request = rb::build_set_font_color_auto(rid, &sid, sr, sc, er, ec);
        exec_status_cmd(engine, &request, session, &format!("Font color set to automatic on {}.", args[0].to_uppercase()));
        return;
    }

    if args.len() < 4 {
        output::error("Usage: format fontcolor <range> <r> <g> <b>  OR  format fontcolor <range> --auto");
        return;
    }
    let parse_channel = |s: &str, name: &str| -> Result<i32, ()> {
        match s.parse::<i32>() {
            Ok(v) if (0..=255).contains(&v) => Ok(v),
            _ => {
                output::error(&format!("Invalid {} value '{}'. Must be 0-255.", name, s));
                Err(())
            }
        }
    };
    let r = match parse_channel(args[1], "red") { Ok(v) => v, Err(_) => return };
    let g = match parse_channel(args[2], "green") { Ok(v) => v, Err(_) => return };
    let b = match parse_channel(args[3], "blue") { Ok(v) => v, Err(_) => return };
    let request = rb::build_set_font_color_rgb(rid, &sid, r, g, b, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Font color set to ({},{},{}) on {}.", r, g, b, args[0].to_uppercase()));
}

fn cmd_format_halign(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format halign <range> <general|left|center|right|fill|justify|centeracross|distributed>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let alignment_type = match args[1].to_lowercase().as_str() {
        "general" => 1,
        "left" => 2,
        "center" => 3,
        "right" => 4,
        "fill" => 5,
        "justify" => 6,
        "centeracross" => 7,
        "distributed" => 8,
        _ => {
            output::error(&format!("Invalid alignment '{}'. Use: general, left, center, right, fill, justify, centeracross, distributed.", args[1]));
            return;
        }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_horizontal_alignment(rid, &sid, alignment_type, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Horizontal alignment set to '{}' on {}.", args[1].to_lowercase(), args[0].to_uppercase()));
}

fn cmd_format_valign(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format valign <range> <top|center|bottom|justify|distributed>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let alignment_type = match args[1].to_lowercase().as_str() {
        "top" => 1,
        "center" => 2,
        "bottom" => 3,
        "justify" => 4,
        "distributed" => 5,
        _ => {
            output::error(&format!("Invalid alignment '{}'. Use: top, center, bottom, justify, distributed.", args[1]));
            return;
        }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_vertical_alignment(rid, &sid, alignment_type, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Vertical alignment set to '{}' on {}.", args[1].to_lowercase(), args[0].to_uppercase()));
}

fn cmd_format_wrap(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format textwrap <range> <overflow|clip|wrap|shrink>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let wrap_type = match args[1].to_lowercase().as_str() {
        "overflow" => 1,
        "clip" => 2,
        "wrap" => 3,
        "shrink" => 4,
        _ => {
            output::error(&format!("Invalid wrap type '{}'. Use: overflow, clip, wrap, shrink.", args[1]));
            return;
        }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_wrap_text(rid, &sid, wrap_type, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Wrap text set to '{}' on {}.", args[1].to_lowercase(), args[0].to_uppercase()));
}

fn cmd_format_rotate(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format rotate <range> <angle>  (angle: -90 to 90)");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let angle = match args[1].parse::<i32>() {
        Ok(v) if (-90..=90).contains(&v) => v,
        _ => {
            output::error(&format!("Invalid angle '{}'. Must be between -90 and 90.", args[1]));
            return;
        }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_text_rotation(rid, &sid, angle, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Text rotation set to {}° on {}.", angle, args[0].to_uppercase()));
}

fn cmd_format_indent(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format indent <range> <increase|decrease>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = match args[1].to_lowercase().as_str() {
        "increase" | "inc" | "+" => rb::build_increase_indent(rid, &sid, sr, sc, er, ec),
        "decrease" | "dec" | "-" => rb::build_decrease_indent(rid, &sid, sr, sc, er, ec),
        _ => {
            output::error(&format!("Invalid indent direction '{}'. Use: increase, decrease.", args[1]));
            return;
        }
    };
    exec_status_cmd(engine, &request, session, &format!("Indent {} on {}.", args[1].to_lowercase(), args[0].to_uppercase()));
}

fn cmd_format_fill_color(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: format fillcolor <range> <r> <g> <b>  OR  format fillcolor <range> --none");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    if args.len() >= 2 && args[1].eq_ignore_ascii_case("--none") {
        let request = rb::build_fill_color_none(rid, &sid, sr, sc, er, ec);
        exec_status_cmd(engine, &request, session, &format!("Fill color removed on {}.", args[0].to_uppercase()));
        return;
    }

    if args.len() < 4 {
        output::error("Usage: format fillcolor <range> <r> <g> <b>  OR  format fillcolor <range> --none");
        return;
    }
    let parse_channel = |s: &str, name: &str| -> Result<i32, ()> {
        match s.parse::<i32>() {
            Ok(v) if (0..=255).contains(&v) => Ok(v),
            _ => {
                output::error(&format!("Invalid {} value '{}'. Must be 0-255.", name, s));
                Err(())
            }
        }
    };
    let r = match parse_channel(args[1], "red") { Ok(v) => v, Err(_) => return };
    let g = match parse_channel(args[2], "green") { Ok(v) => v, Err(_) => return };
    let b = match parse_channel(args[3], "blue") { Ok(v) => v, Err(_) => return };
    let request = rb::build_fill_color_rgb(rid, &sid, r, g, b, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Fill color set to ({},{},{}) on {}.", r, g, b, args[0].to_uppercase()));
}

fn cmd_format_border(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: format border <range> <type> <style> [r g b]\n  type: all|outer|inner|left|right|top|bottom|horizontal|vertical|diagonal\n  style: none|thin|medium|dashed|dotted|thick|double|hair|mediumdashed|dashdot|mediumdashdot|dashdotdot|mediumdashdotdot|slantdashdot");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let border_type = match args[1].to_lowercase().as_str() {
        "all" => 101,
        "outer" => 102,
        "inner" => 103,
        "left" => 104,
        "right" => 105,
        "top" => 106,
        "bottom" => 107,
        "horizontal" => 108,
        "vertical" => 109,
        "diagonal" => 110,
        _ => {
            output::error(&format!("Invalid border type '{}'. Use: all, outer, inner, left, right, top, bottom, horizontal, vertical, diagonal.", args[1]));
            return;
        }
    };
    let border_line_style = match args[2].to_lowercase().as_str() {
        "none" => 1,
        "thin" => 2,
        "medium" => 3,
        "dashed" => 4,
        "dotted" => 5,
        "thick" => 6,
        "double" => 7,
        "hair" => 8,
        "mediumdashed" => 9,
        "dashdot" => 10,
        "mediumdashdot" => 11,
        "dashdotdot" => 12,
        "mediumdashdotdot" => 13,
        "slantdashdot" => 14,
        _ => {
            output::error(&format!("Invalid border style '{}'. Use: none, thin, medium, dashed, dotted, thick, double, hair, mediumdashed, dashdot, mediumdashdot, dashdotdot, mediumdashdotdot, slantdashdot.", args[2]));
            return;
        }
    };
    // Default to black if no color specified
    let (r, g, b) = if args.len() >= 6 {
        let parse_channel = |s: &str, name: &str| -> Result<i32, ()> {
            match s.parse::<i32>() {
                Ok(v) if (0..=255).contains(&v) => Ok(v),
                _ => {
                    output::error(&format!("Invalid {} value '{}'. Must be 0-255.", name, s));
                    Err(())
                }
            }
        };
        let r = match parse_channel(args[3], "red") { Ok(v) => v, Err(_) => return };
        let g = match parse_channel(args[4], "green") { Ok(v) => v, Err(_) => return };
        let b = match parse_channel(args[5], "blue") { Ok(v) => v, Err(_) => return };
        (r, g, b)
    } else {
        (0, 0, 0)
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_set_border(rid, &sid, border_type, border_line_style, r, g, b, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Border '{}' with style '{}' set on {}.", args[1].to_lowercase(), args[2].to_lowercase(), args[0].to_uppercase()));
}

// ─── Number formatting commands ──────────────────────────────────────────────

fn cmd_format_numformat(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format numformat <range> <type> [format_text | --flags...]\n  type: general|number|currency|accounting|date|time|duration|percentage|scientific|fraction|text|custom\n  Use --decimals, --leading-zeros, --negative, --currency, --digits etc. for parameterized patterns.\n  Use 'format numformat --list-custom' to list saved custom formats.");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let type_str = args[1].to_lowercase();
    let number_format_type = match type_str.as_str() {
        "general" => 1,
        "number" => 2,
        "currency" => 3,
        "accounting" => 4,
        "date" => 5,
        "time" => 6,
        "duration" => 7,
        "percentage" => 8,
        "scientific" => 9,
        "fraction" => 10,
        "text" => 11,
        "custom" => 13,
        _ => {
            output::error(&format!("Invalid number format type '{}'. Use: general, number, currency, accounting, date, time, duration, percentage, scientific, fraction, text, custom.", args[1]));
            return;
        }
    };

    let format_text = resolve_numformat_text(&type_str, &args[2..]);
    let format_text = match format_text {
        Ok(f) => f,
        Err(e) => { output::error(&e); return; }
    };

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_apply_number_format(rid, &sid, &format_text, number_format_type, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, &format!("Number format '{}' ({}) applied to {}.", format_text, type_str, args[0].to_uppercase()));
}

/// Shared logic to resolve format_text from args — supports flags, numbered shortcuts, and raw patterns.
fn resolve_numformat_text(type_str: &str, rest: &[&str]) -> Result<String, String> {
    // Types that need no format_text
    match type_str {
        "general" => return Ok(String::from("General")),
        "text" => return Ok(String::from("@")),
        _ => {}
    }

    // Detect if flags and a positional shortcut are both present (ambiguous mix)
    let has_flags = rest.iter().any(|a| a.starts_with("--"));
    let has_positional_shortcut = rest.first()
        .map(|a| !a.starts_with("--") && a.parse::<u32>().is_ok())
        .unwrap_or(false);
    if has_flags && has_positional_shortcut {
        return Err(String::from(
            "Cannot combine shortcut index with flags. Use one or the other.\n  Example (shortcut): format numformat A1 number 2\n  Example (flags):    format numformat A1 number --decimals 2"
        ));
    }

    // Check if flags are present (parameterized mode)
    if let Some(params) = numformat::parse_flags(rest) {
        return numformat::generate_pattern(type_str, &params);
    }

    // Accounting default when no args
    if type_str == "accounting" && rest.is_empty() {
        return Ok(String::from("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"));
    }

    // Require format_text for other types
    if rest.is_empty() {
        return Err(format!("Type '{}' requires a format_text argument or --flags. Use 'help' to see options.", type_str));
    }

    let raw = rest.join(" ");
    // Currency/accounting: resolve locale key to actual format pattern
    // The engine does NOT resolve locale keys — it treats format_text as a literal pattern.
    if type_str == "currency" && !raw.starts_with("--") {
        return Ok(numformat::resolve_currency_locale(raw.trim()));
    }
    if type_str == "accounting" && !raw.starts_with("--") {
        return Ok(numformat::resolve_accounting_locale(raw.trim()));
    }
    let raw = raw;
    // Resolve numbered shortcuts for common types
    let resolved = match type_str {
        "number" => match raw.as_str() {
            "1" => String::from("#,##0"),
            "2" => String::from("#,##0.00"),
            "3" => String::from("#0"),
            "4" => String::from("#0.00"),
            _ => raw,
        },
        "date" => match raw.as_str() {
            "1" => String::from("dddd, d mmmm, yyyy"),
            "2" => String::from("d mmmm yyyy"),
            "3" => String::from("dd-mmm-yyyy"),
            "4" => String::from("dd/mm/yy"),
            _ => raw,
        },
        "time" => match raw.as_str() {
            "1" => String::from("h:mm:ss"),
            "2" => String::from("h:mm:ss AM/PM"),
            _ => raw,
        },
        "duration" => match raw.as_str() {
            "1" => String::from("[hh]:mm:ss"),
            "2" => String::from("[hh]:mm"),
            "3" => String::from("[hh]"),
            "4" => String::from("[mm]"),
            "5" => String::from("[ss]"),
            _ => raw,
        },
        "percentage" => match raw.as_str() {
            "1" => String::from("0%"),
            "2" => String::from("0.00%"),
            _ => raw,
        },
        "scientific" => match raw.as_str() {
            "1" => String::from("0.00E+00"),
            "2" => String::from("0.0E+00"),
            _ => raw,
        },
        "fraction" => match raw.as_str() {
            "1" => String::from("# ?/?"),
            "2" => String::from("# ??/??"),
            _ => raw,
        },
        _ => raw,
    };
    Ok(resolved)
}

fn cmd_format_decimal(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format decimal <range> <increase|decrease>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    match args[1].to_lowercase().as_str() {
        "increase" => {
            let request = rb::build_increase_decimal(rid, &sid, sr, sc, er, ec);
            exec_status_cmd(engine, &request, session, &format!("Decimal places increased on {}.", args[0].to_uppercase()));
        }
        "decrease" => {
            let request = rb::build_decrease_decimal(rid, &sid, sr, sc, er, ec);
            exec_status_cmd(engine, &request, session, &format!("Decimal places decreased on {}.", args[0].to_uppercase()));
        }
        _ => {
            output::error(&format!("Invalid decimal direction '{}'. Use: increase, decrease.", args[1]));
        }
    }
}

fn cmd_format_numpreview(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: format numpreview <cellRef|range> <type> [format_text | --flags...]\n  type: general|number|currency|accounting|date|time|duration|percentage|scientific|fraction|text|custom\n  Accepts a single cell (A1) or range (A1:B5); preview uses the first cell.");
        return;
    }
    // Accept both single cell refs and ranges; use the top-left cell for preview
    let (col, row) = if let Some(p) = cell_ref::try_parse(args[0]) {
        p
    } else if let Some((sc, sr, _ec, _er)) = cell_ref::try_parse_range(args[0]) {
        (sc, sr)
    } else {
        output::error(&format!("Invalid cell reference or range: '{}'", args[0]));
        return;
    };
    let type_str = args[1].to_lowercase();
    let number_format_type = match type_str.as_str() {
        "general" => 1,
        "number" => 2,
        "currency" => 3,
        "accounting" => 4,
        "date" => 5,
        "time" => 6,
        "duration" => 7,
        "percentage" => 8,
        "scientific" => 9,
        "fraction" => 10,
        "text" => 11,
        "custom" => 13,
        _ => {
            output::error(&format!("Invalid number format type '{}'. Use: general, number, currency, accounting, date, time, duration, percentage, scientific, fraction, text, custom.", args[1]));
            return;
        }
    };
    let format_text = resolve_numformat_text(&type_str, &args[2..]);
    let format_text = match format_text {
        Ok(f) => f,
        Err(e) => { output::error(&e); return; }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_preview_number_format(rid, &sid, &format_text, number_format_type, row, col);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let v: serde_json::Value = match serde_json::from_str(&resp) {
                Ok(v) => v,
                Err(e) => { output::error(&format!("Failed to parse response: {}", e)); return; }
            };
            if let Some(response) = v.get("response") {
                let preview = response.get("preview_for_selected_format")
                    .and_then(|v| v.as_str())
                    .unwrap_or("(none)");
                let valid = response.get("is_pattern_valid")
                    .and_then(|v| v.as_bool())
                    .unwrap_or(false);
                let resolved_format = response.get("format_text")
                    .and_then(|v| v.as_str())
                    .unwrap_or("");
                output::success("Number format preview:");
                output::key_value("Format", resolved_format, 2);
                output::key_value("Valid", if valid { "yes" } else { "no" }, 2);
                output::key_value("Preview", preview, 2);
            } else {
                let status = rp::parse_status_response(&resp);
                output::error(&format!("Preview failed: {}", status.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn cmd_format_numinfo(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: format numinfo <cellRef>");
        return;
    }
    let (col, row) = match cell_ref::try_parse(args[0]) {
        Some(p) => p,
        None => { output::error(&format!("Invalid cell reference: '{}'", args[0])); return; }
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_get_number_format_info(rid, &sid, row, col);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let v: serde_json::Value = match serde_json::from_str(&resp) {
                Ok(v) => v,
                Err(e) => { output::error(&format!("Failed to parse response: {}", e)); return; }
            };
            if let Some(response) = v.get("response") {
                let format_text = response.get("format_text")
                    .and_then(|v| v.as_str())
                    .unwrap_or("(none)");
                let format_type = response.get("number_format_type")
                    .and_then(|v| v.as_str())
                    .unwrap_or("UNKNOWN");
                let decimal_places = response.get("decimal_places")
                    .and_then(|v| v.as_i64())
                    .map(|v| v.to_string())
                    .unwrap_or_else(|| "N/A".to_string());
                let leading_zeroes = response.get("leading_zeroes")
                    .and_then(|v| v.as_i64())
                    .map(|v| v.to_string())
                    .unwrap_or_else(|| "N/A".to_string());
                output::success(&format!("Number format info for {}:", args[0].to_uppercase()));
                output::key_value("Format text", format_text, 2);
                output::key_value("Format type", format_type, 2);
                output::key_value("Decimal places", &decimal_places, 2);
                output::key_value("Leading zeroes", &leading_zeroes, 2);
            } else {
                let status = rp::parse_status_response(&resp);
                output::error(&format!("Get number format info failed: {}", status.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn cmd_format_nummanage(_engine: &EngineHandle, _session: &mut CliSession) {
    output::success("Built-in number format types and shortcuts:");
    output::info("");
    output::info("  general       (no shortcuts) — Displays as entered");
    output::info("  number        1: #,##0    2: #,##0.00    3: #0    4: #0.00");
    output::info("                Flags: --decimals, --noseparator, --leading-zeros, --negative, --prefix, --suffix");
    output::info("  currency      Locale key, e.g.: en-US  en-IN  en-GB  en-JP");
    output::info("                Flags: --currency, --decimals, --negative");
    output::info("  accounting    Default: _(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)");
    output::info("                Flags: --currency, --decimals, --negative");
    output::info("  date          1: dddd, d mmmm, yyyy   2: d mmmm yyyy");
    output::info("                3: dd-mmm-yyyy           4: dd/mm/yy");
    output::info("                Flags: --date");
    output::info("  time          1: h:mm:ss              2: h:mm:ss AM/PM");
    output::info("                Flags: --time");
    output::info("  duration      1: [hh]:mm:ss  2: [hh]:mm  3: [hh]  4: [mm]  5: [ss]");
    output::info("  percentage    1: 0%          2: 0.00%");
    output::info("                Flags: --decimals");
    output::info("  scientific    1: 0.00E+00    2: 0.0E+00");
    output::info("                Flags: --decimals");
    output::info("  fraction      1: # ?/?       2: # ??/??");
    output::info("                Flags: --digits");
    output::info("  text          (no shortcuts) — Displays as text (@)");
    output::info("  custom        <raw_pattern>  — Any format string");
    output::info("                Flags: --date, --time, --prefix, --suffix");
    output::info("");
    output::info("  Use 'format numformat --list-custom' to see saved custom formats.");
    output::info("  Use 'format numformat --list-currency' to see supported currency codes.");
}

fn cmd_format_list_currency() {
    output::success("Supported currency country codes:");
    output::info("  Usage: format numformat <range> currency <code>");
    output::info("");
    output::info("  en-US       US Dollar ($)");
    output::info("  en-GB       British Pound (£)");
    output::info("  en-IN       Indian Rupee (₹)");
    output::info("  en-CA       Canadian Dollar (C$)");
    output::info("  en-AU       Australian Dollar (A$)");
    output::info("  de-DE       Euro (€)");
    output::info("  fr-FR       Euro (€)");
    output::info("  ja-JP       Japanese Yen (¥)");
    output::info("  zh-CN       Chinese Yuan (¥)");
    output::info("  ko-KR       Korean Won (₩)");
    output::info("  pt-BR       Brazilian Real (R$)");
    output::info("  es-MX       Mexican Peso (MX$)");
    output::info("  ru-RU       Russian Ruble (₽)");
    output::info("  ar-SA       Saudi Riyal (﷼)");
    output::info("  en-ZA       South African Rand (R)");
}

fn cmd_format_list_custom(engine: &EngineHandle, session: &mut CliSession) {
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_manage_custom_format(rid);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let v: serde_json::Value = match serde_json::from_str(&resp) {
                Ok(v) => v,
                Err(e) => { output::error(&format!("Failed to parse response: {}", e)); return; }
            };
            if let Some(response) = v.get("response") {
                output::success("Custom number formats:");
                if let Some(user_formats) = response.get("user_level_custom_format").and_then(|v| v.as_array()) {
                    if user_formats.is_empty() {
                        output::key_value("User formats", "(none)", 2);
                    } else {
                        output::key_value("User formats", &format!("{}", user_formats.len()), 2);
                        for f in user_formats {
                            if let Some(s) = f.as_str() {
                                output::info(&format!("    {}", s));
                            }
                        }
                    }
                }
                if let Some(doc_formats) = response.get("document_level_custom_format").and_then(|v| v.as_array()) {
                    if doc_formats.is_empty() {
                        output::key_value("Document formats", "(none)", 2);
                    } else {
                        output::key_value("Document formats", &format!("{}", doc_formats.len()), 2);
                        for f in doc_formats {
                            if let Some(s) = f.as_str() {
                                output::info(&format!("    {}", s));
                            }
                        }
                    }
                }
            } else {
                let status = rp::parse_status_response(&resp);
                output::error(&format!("Manage custom format failed: {}", status.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn cmd_format_default(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() || (args.len() == 1 && args[0] == "--help") {
        output::info("Usage: format default <range> [--flags...]");
        output::info("  Sets the default cell format for the specified range.");
        output::info("  Does not affect existing cell-level overrides.");
        output::info("");
        output::info("  Flags:");
        output::info("    --font-name NAME           Font name (e.g. Arial, Calibri)");
        output::info("    --font-size N              Font size in points");
        output::info("    --bold true|false          Bold text");
        output::info("    --italic true|false        Italic text");
        output::info("    --underline true|false     Underline text");
        output::info("    --font-color R G B         Font color (RGB 0-255)");
        output::info("    --fill-color R G B         Fill/background color (RGB 0-255)");
        output::info("    --halign TYPE              Horizontal alignment: left|center|right|justify");
        output::info("    --valign TYPE              Vertical alignment: top|center|bottom");
        output::info("    --wrap true|false          Text wrap");
        output::info("");
        output::info("  Examples:");
        output::info("    format default A1:Z100 --font-name Arial --font-size 11 --bold false");
        output::info("    format default A1:Z100 --fill-color 255 255 255 --halign center");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rest = &args[1..];

    // Backward compatibility: detect raw JSON input
    let rest_joined = rest.join(" ");
    let trimmed = rest_joined.trim();
    if trimmed.starts_with('{') {
        output::info("Warning: JSON input for 'format default' is deprecated. Use structured flags instead. See 'format default --help'.");
        let mut format_json: serde_json::Value = match serde_json::from_str(trimmed) {
            Ok(v) => v,
            Err(e) => {
                output::error(&format!("Invalid JSON: {}.", e));
                return;
            }
        };
        let rid = session.rid.as_deref().unwrap();
        let sid = session.get_active_sheet_id_or_default();
        format_json["active_info"] = serde_json::json!({
            "active_sheet_id": sid,
            "active_cell": { "active_row": sr, "active_column": sc },
            "active_range_list": [{ "start_row": sr, "end_row": er, "start_column": sc, "end_column": ec }]
        });
        let request = rb::build_default_format(rid, format_json);
        exec_status_cmd(engine, &request, session, &format!("Default format applied to {}.", args[0].to_uppercase()));
        return;
    }

    // Parse structured flags
    if rest.is_empty() {
        output::error("No flags provided. Run 'format default --help' to see supported flags.");
        return;
    }

    let mut font_obj = serde_json::Map::new();
    let mut alignment_obj = serde_json::Map::new();
    let mut fill_obj = serde_json::Map::new();
    let mut has_font = false;
    let mut has_alignment = false;
    let mut has_fill = false;

    let mut i = 0;
    while i < rest.len() {
        match rest[i] {
            "--font-name" => {
                i += 1;
                if i >= rest.len() { output::error("--font-name requires a value."); return; }
                font_obj.insert("font_name".to_string(), serde_json::Value::String(rest[i].to_string()));
                has_font = true;
            }
            "--font-size" => {
                i += 1;
                if i >= rest.len() { output::error("--font-size requires a value."); return; }
                let size: i64 = match rest[i].parse() {
                    Ok(v) => v,
                    Err(_) => { output::error(&format!("Invalid font size: '{}'", rest[i])); return; }
                };
                font_obj.insert("font_size".to_string(), serde_json::json!(size));
                has_font = true;
            }
            "--bold" => {
                i += 1;
                if i >= rest.len() { output::error("--bold requires true|false."); return; }
                let v = parse_bool_flag(rest[i], "--bold");
                match v { Ok(b) => { font_obj.insert("is_bold".to_string(), serde_json::json!(b)); has_font = true; }, Err(e) => { output::error(&e); return; } }
            }
            "--italic" => {
                i += 1;
                if i >= rest.len() { output::error("--italic requires true|false."); return; }
                let v = parse_bool_flag(rest[i], "--italic");
                match v { Ok(b) => { font_obj.insert("is_italic".to_string(), serde_json::json!(b)); has_font = true; }, Err(e) => { output::error(&e); return; } }
            }
            "--underline" => {
                i += 1;
                if i >= rest.len() { output::error("--underline requires true|false."); return; }
                let v = parse_bool_flag(rest[i], "--underline");
                match v { Ok(b) => { font_obj.insert("is_underline".to_string(), serde_json::json!(b)); has_font = true; }, Err(e) => { output::error(&e); return; } }
            }
            "--font-color" => {
                if i + 3 >= rest.len() { output::error("--font-color requires R G B values (0-255)."); return; }
                let r: i64 = rest[i+1].parse().unwrap_or(-1);
                let g: i64 = rest[i+2].parse().unwrap_or(-1);
                let b: i64 = rest[i+3].parse().unwrap_or(-1);
                if r < 0 || r > 255 || g < 0 || g > 255 || b < 0 || b > 255 {
                    output::error("--font-color RGB values must be 0-255."); return;
                }
                font_obj.insert("font_color".to_string(), serde_json::json!({"r": r, "g": g, "b": b}));
                has_font = true;
                i += 3;
            }
            "--fill-color" => {
                if i + 3 >= rest.len() { output::error("--fill-color requires R G B values (0-255)."); return; }
                let r: i64 = rest[i+1].parse().unwrap_or(-1);
                let g: i64 = rest[i+2].parse().unwrap_or(-1);
                let b: i64 = rest[i+3].parse().unwrap_or(-1);
                if r < 0 || r > 255 || g < 0 || g > 255 || b < 0 || b > 255 {
                    output::error("--fill-color RGB values must be 0-255."); return;
                }
                fill_obj.insert("bg_color".to_string(), serde_json::json!({"r": r, "g": g, "b": b}));
                has_fill = true;
                i += 3;
            }
            "--halign" => {
                i += 1;
                if i >= rest.len() { output::error("--halign requires a value."); return; }
                let align_type: i32 = match rest[i].to_lowercase().as_str() {
                    "left" => 2,
                    "center" => 3,
                    "right" => 4,
                    "justify" => 6,
                    _ => { output::error(&format!("Invalid --halign value '{}'. Use: left, center, right, justify.", rest[i])); return; }
                };
                alignment_obj.insert("horizontal_alignment_type".to_string(), serde_json::json!(align_type));
                has_alignment = true;
            }
            "--valign" => {
                i += 1;
                if i >= rest.len() { output::error("--valign requires a value."); return; }
                let align_type: i32 = match rest[i].to_lowercase().as_str() {
                    "top" => 1,
                    "center" => 2,
                    "bottom" => 3,
                    _ => { output::error(&format!("Invalid --valign value '{}'. Use: top, center, bottom.", rest[i])); return; }
                };
                alignment_obj.insert("vertical_alignment_type".to_string(), serde_json::json!(align_type));
                has_alignment = true;
            }
            "--wrap" => {
                i += 1;
                if i >= rest.len() { output::error("--wrap requires true|false."); return; }
                let v = parse_bool_flag(rest[i], "--wrap");
                match v { Ok(b) => { alignment_obj.insert("is_text_wrap".to_string(), serde_json::json!(b)); has_alignment = true; }, Err(e) => { output::error(&e); return; } }
            }
            other => {
                output::error(&format!("Unknown flag '{}' for format default. Run 'format default --help' to see supported flags.", other));
                return;
            }
        }
        i += 1;
    }

    let mut format_json = serde_json::json!({});
    if has_font {
        format_json["font"] = serde_json::Value::Object(font_obj);
    }
    if has_alignment {
        format_json["alignment"] = serde_json::Value::Object(alignment_obj);
    }
    if has_fill {
        format_json["fill"] = serde_json::Value::Object(fill_obj);
    }

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    format_json["active_info"] = serde_json::json!({
        "active_sheet_id": sid,
        "active_cell": { "active_row": sr, "active_column": sc },
        "active_range_list": [{ "start_row": sr, "end_row": er, "start_column": sc, "end_column": ec }]
    });
    let request = rb::build_default_format(rid, format_json);
    exec_status_cmd(engine, &request, session, &format!("Default format applied to {}.", args[0].to_uppercase()));
}

fn parse_bool_flag(value: &str, flag_name: &str) -> Result<bool, String> {
    match value.to_lowercase().as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(format!("Invalid value '{}' for {}. Use: true, false.", value, flag_name)),
    }
}

fn parse_on_off_flag(value: &str, flag_name: &str) -> Result<bool, String> {
    match value.to_lowercase().as_str() {
        "on" => Ok(true),
        "off" => Ok(false),
        _ => Err(format!("Invalid value '{}' for {}. Use: on, off.", value, flag_name)),
    }
}

/// Execute a ProcessRequestJson call, check status, and print success/error.
fn exec_status_cmd(
    engine: &EngineHandle,
    request: &str,
    session: &mut CliSession,
    success_msg: &str,
) {
    match engine.process_request_json(request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success(success_msg);
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Operation failed: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── Chart ───────────────────────────────────────────────────────────────────

fn cmd_chart(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: chart list|insert|delete|move|clone|rename|position|type|info|get|source|manage|recommend|customize|style|series|datalabel|axis|gridline|autoexpand [args]");
        return;
    }
    match args[0].to_lowercase().as_str() {
        "list" => chart_list(&args[1..], engine, session),
        "insert" => chart_insert(&args[1..], engine, session),
        "delete" => chart_delete(&args[1..], engine, session),
        "move" => chart_move(&args[1..], engine, session),
        "clone" => chart_clone(&args[1..], engine, session),
        "rename" => chart_rename(&args[1..], engine, session),
        "position" => chart_position(&args[1..], engine, session),
        "type" => chart_type(&args[1..], engine, session),
        "info" => chart_info(&args[1..], engine, session),
        "get" => chart_manage(&args[1..], engine, session),
        "source" => chart_source(&args[1..], engine, session),
        "manage" => chart_manage(&args[1..], engine, session),
        "recommend" => chart_recommend(&args[1..], engine, session),
        "customize" => chart_customize(&args[1..], engine, session),
        "style" => chart_style(&args[1..], engine, session),
        "series" => chart_series(&args[1..], engine, session),
        "datalabel" => chart_datalabel(&args[1..], engine, session),
        "axis" => chart_axis(&args[1..], engine, session),
        "gridline" => chart_gridline(&args[1..], engine, session),
        "autoexpand" => chart_autoexpand(&args[1..], engine, session),
        // Deprecated aliases — will be removed next major version
        "property" => chart_property_deprecated(&args[1..], engine, session),
        "charttype" => chart_charttype_deprecated(&args[1..], engine, session),
        "property2" => chart_property2_deprecated(&args[1..], engine, session),
        other => output::error(&format!(
            "Unknown chart sub-command: '{}'. Use: list, insert, delete, move, clone, rename, position, type, info, get, source, manage, recommend, customize, style, series, datalabel, axis, gridline, autoexpand",
            other
        )),
    }
}

/// Fetches all charts on the active sheet and returns them.
fn fetch_all_charts(engine: &EngineHandle, session: &mut CliSession) -> Option<Vec<rp::ChartInfo>> {
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let range_list = vec![serde_json::json!({
        "start_row": 0,
        "start_column": 0,
        "end_row": 1048576,
        "end_column": 16384
    })];
    let request = rb::build_manage_chart_with_range(rid, &sid, range_list);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let (status_code, _status_message, charts) = rp::parse_manage_chart(&resp);
            if rp::is_success(status_code) {
                Some(charts)
            } else {
                None
            }
        }
        Err(_) => None,
    }
}

/// Resolves a chart name or ID to the actual chart_id.
/// Accepts: full chart_id, user-assigned name (from chart rename cache), or chart title.
/// Parses key=value pairs into a nested JSON object.
/// Supports dot notation for nesting (e.g. "color.red=255" → {"color":{"red":255}}).
/// Values are auto-detected: integers, floats, booleans, or strings.
fn parse_kv_to_json(args: &[&str]) -> Option<serde_json::Value> {
    let mut root = serde_json::Map::new();
    for &arg in args {
        let (key, val) = match arg.split_once('=') {
            Some(kv) => kv,
            None => {
                output::error(&format!("Invalid property format: '{}'. Use key=value.", arg));
                return None;
            }
        };
        let json_val = if val == "true" {
            serde_json::Value::Bool(true)
        } else if val == "false" {
            serde_json::Value::Bool(false)
        } else if let Ok(i) = val.parse::<i64>() {
            serde_json::Value::Number(serde_json::Number::from(i))
        } else if let Ok(f) = val.parse::<f64>() {
            serde_json::Number::from_f64(f)
                .map(serde_json::Value::Number)
                .unwrap_or_else(|| serde_json::Value::String(val.to_string()))
        } else {
            serde_json::Value::String(val.to_string())
        };

        // Support dot notation for nested keys
        let parts: Vec<&str> = key.split('.').collect();
        if parts.len() == 1 {
            root.insert(key.to_string(), json_val);
        } else {
            // Walk/create nested maps
            let mut current = &mut root;
            for &part in &parts[..parts.len() - 1] {
                current = current
                    .entry(part.to_string())
                    .or_insert_with(|| serde_json::Value::Object(serde_json::Map::new()))
                    .as_object_mut()
                    .unwrap();
            }
            current.insert(parts.last().unwrap().to_string(), json_val);
        }
    }
    Some(serde_json::Value::Object(root))
}

/// When multiple charts share the same title (e.g. default "Chart Title"),
/// the last one (most recently inserted) is used — this supports the pattern
/// of inserting a chart and immediately renaming it.
fn resolve_chart_id(name_or_id: &str, engine: &EngineHandle, session: &mut CliSession) -> Option<String> {
    let charts = fetch_all_charts(engine, session)?;
    // First try exact ID match
    if let Some(c) = charts.iter().find(|c| c.chart_id == name_or_id) {
        return Some(c.chart_id.clone());
    }
    // Check the session's chart name cache (populated by chart rename)
    let lower = name_or_id.to_lowercase();
    if let Some(cached_id) = session.chart_name_cache.get(&lower) {
        // Verify the chart still exists on this sheet
        if charts.iter().any(|c| c.chart_id == *cached_id) {
            return Some(cached_id.clone());
        }
    }
    // Then try case-insensitive title match from engine
    let matches: Vec<&rp::ChartInfo> = charts.iter()
        .filter(|c| c.chart_title.to_lowercase() == lower)
        .collect();
    if matches.len() == 1 {
        return Some(matches[0].chart_id.clone());
    }
    if matches.len() > 1 {
        // Multiple charts with same title: use the last one (most recently inserted)
        return Some(matches.last().unwrap().chart_id.clone());
    }
    // No match found
    output::error(&format!("Chart '{}' not found. Use 'chart list' to see available charts.", name_or_id));
    None
}

/// Returns the best display name for a chart: the user-assigned name from the
/// rename cache if available, otherwise the engine-returned title.
fn resolve_display_name<'a>(chart_id: &str, engine_title: &'a str, session: &'a CliSession) -> &'a str {
    // Check if there's a cached name pointing to this chart_id
    for (name, cid) in &session.chart_name_cache {
        if cid == chart_id {
            return name.as_str();
        }
    }
    if engine_title.is_empty() { "(untitled)" } else { engine_title }
}

fn chart_list(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let all_sheets = args.iter().any(|a| a.eq_ignore_ascii_case("--all"));

    if all_sheets {
        // List charts across all sheets (like pivot list)
        let rid = session.rid.as_deref().unwrap();
        let mut total = 0usize;
        for (i, sheet_id) in session.sheet_ids.iter().enumerate() {
            let range_list = vec![serde_json::json!({
                "start_row": 0,
                "start_column": 0,
                "end_row": 1048576,
                "end_column": 16384
            })];
            let request = rb::build_manage_chart_with_range(rid, sheet_id, range_list);
            let charts = match engine.process_request_json(&request) {
                Ok(resp) => {
                    let (status_code, _, charts) = rp::parse_manage_chart(&resp);
                    if rp::is_success(status_code) { charts } else { Vec::new() }
                }
                Err(_) => Vec::new(),
            };
            if charts.is_empty() {
                continue;
            }
            let sheet_name = session.sheet_names.get(i).map(|s| s.as_str()).unwrap_or("?");
            if total == 0 {
                output::line("Charts across all sheets:", 0);
                output::line("", 0);
            }
            for c in &charts {
                let title_display = resolve_display_name(&c.chart_id, &c.chart_title, session);
                output::line(&format!(
                    "  \"{}\"  [ID: {}]  type: {} sub: {}  [{}]",
                    title_display, c.chart_id, c.chart_type, c.chart_sub_type, sheet_name
                ), 0);
                total += 1;
            }
        }
        if total == 0 {
            output::info("No charts found in this workbook.");
        }
    } else {
        match fetch_all_charts(engine, session) {
            Some(charts) => {
                if charts.is_empty() {
                    output::info("No charts found on the active sheet. Use 'chart list --all' for all sheets.");
                } else {
                    output::line(&format!("Charts on active sheet ({} found):", charts.len()), 0);
                    output::line("", 0);
                    for (i, c) in charts.iter().enumerate() {
                        let title_display = resolve_display_name(&c.chart_id, &c.chart_title, session);
                        output::line(&format!(
                            "  {}. \"{}\"  [ID: {}]  type: {} sub: {}",
                            i + 1, title_display, c.chart_id, c.chart_type, c.chart_sub_type
                        ), 0);
                    }
                }
            }
            None => {
                output::error("Failed to fetch charts from the active sheet.");
            }
        }
    }
}

fn chart_info(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: chart info <chartName|chartId>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    // Fetch full chart details via manage by id
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let ids = vec![chart_id.clone()];
    let request = rb::build_manage_chart_with_id(rid, &sid, ids);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let (status_code, status_message, charts) = rp::parse_manage_chart(&resp);
            if rp::is_success(status_code) {
                if let Some(c) = charts.first() {
                    let title_display = if c.chart_title.is_empty() { "(untitled)" } else { &c.chart_title };
                    let type_names = ["bar","column","line","pie","area","scatter","race",
                        "waterfall","bullet","funnel","pareto","histogram","stock",
                        "radar","wordcloud","combo","boxplot"];
                    let type_name = type_names.get(c.chart_type as usize).unwrap_or(&"unknown");
                    output::success("Chart info:");
                    output::key_value("Title", title_display, 2);
                    output::key_value("Chart ID", &c.chart_id, 2);
                    output::key_value("Type", &format!("{} ({})", type_name, c.chart_type), 2);
                    output::key_value("Sub-type", &c.chart_sub_type.to_string(), 2);
                    output::key_value("Position", &format!("({},{},{},{})", c.start_x, c.start_y, c.end_x, c.end_y), 2);
                    output::key_value("Position Type", &c.position_type.to_string(), 2);
                } else {
                    output::error("Chart not found in response.");
                }
            } else {
                output::error(&format!(
                    "Failed to get chart info: {}",
                    status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_source(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: chart source <chartName|chartId> <range>");
        output::info("  Updates the chart's data source range.");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let range_list = serde_json::json!([{
        "sheet_id": sid,
        "start_row": sr,
        "start_column": sc,
        "end_row": er,
        "end_column": ec
    }]);
    let props = serde_json::json!({
        "sheet_range_list": range_list
    });
    let request = rb::build_customize_chart_property_two(rid, &sid, &chart_id, props);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' data source updated to {}.", chart_id, args[1].to_uppercase()));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to update chart source: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_insert(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    // Usage: chart insert <range> <type> [--pos x1,y1,x2,y2] [--postype 0|1|2]
    if args.is_empty() {
        output::error("Usage: chart insert <range> <type> [--pos x1,y1,x2,y2] [--postype 0|1|2]");
        output::info("Type uses combined type_subtype format. Available types:");
        output::info("  BAR:       bar, bar_stacked, bar_stacked_100, bar_grouped");
        output::info("  COLUMN:    column, column_stacked, column_stacked_100, column_grouped");
        output::info("  LINE:      line, line_spline, line_step, line_timeline");
        output::info("  PIE:       pie, pie_semi, pie_doughnut, pie_semi_doughnut, pie_parliament, doughnut_parliament");
        output::info("  AREA:      area, area_stacked, area_stacked_100, area_time");
        output::info("  SCATTER:   scatter, scatter_line, scatter_line_markers, scatter_bubble");
        output::info("  RACE:      race");
        output::info("  WATERFALL: waterfall");
        output::info("  BULLET:    bullet, bullet_vertical");
        output::info("  FUNNEL:    funnel, funnel_weighted");
        output::info("  PARETO:    pareto");
        output::info("  HISTOGRAM: histogram");
        output::info("  STOCK:     stock, stock_ohlc");
        output::info("  RADAR:     radar, radar_spiderweb");
        output::info("  WORDCLOUD: wordcloud");
        output::info("  COMBO:     combo");
        output::info("  BOXPLOT:   boxplot, boxplot_grouped_horizontal, boxplot_vertical, boxplot_grouped_vertical");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    if args.len() < 2 {
        output::error("Missing chart type. Use: bar, column_stacked, line_spline, pie_doughnut, etc.");
        return;
    }
    let (chart_type, chart_sub_type) = match parse_chart_type_subtype(args[1]) {
        Some(v) => v,
        None => { output::error("Invalid chart type. Use combined type_subtype format (e.g. bar_stacked, line_spline, pie_doughnut). Run 'chart insert' for full list."); return; }
    };

    // Parse --pos
    let (start_x, start_y, end_x, end_y) = if let Some(pos) = args.iter().position(|a| a.eq_ignore_ascii_case("--pos")) {
        if pos + 1 >= args.len() { output::error("--pos requires startX,startY,endX,endY"); return; }
        let parts: Vec<&str> = args[pos + 1].split(',').collect();
        if parts.len() != 4 { output::error("--pos format: startX,startY,endX,endY"); return; }
        let nums: Vec<i32> = match parts.iter().map(|p| p.parse::<i32>()).collect::<Result<Vec<_>, _>>() {
            Ok(v) => v,
            Err(_) => { output::error("--pos values must be integers."); return; }
        };
        (nums[0], nums[1], nums[2], nums[3])
    } else {
        (0, 0, 500, 300)
    };

    let position_type: i32 = if let Some(pos) = args.iter().position(|a| a.eq_ignore_ascii_case("--postype")) {
        if pos + 1 >= args.len() { output::error("--postype requires 0, 1, or 2"); return; }
        match args[pos + 1].parse() {
            Ok(v) => v,
            Err(_) => { output::error("Invalid position type."); return; }
        }
    } else {
        0
    };

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let range_list = vec![serde_json::json!({
        "sheet_id": sid,
        "start_row": sr,
        "start_column": sc,
        "end_row": er,
        "end_column": ec
    })];

    let request = rb::build_insert_chart(rid, &sid, chart_type, chart_sub_type, range_list, start_x, start_y, end_x, end_y, position_type, sr, sc);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let (status_code, status_message, chart_id) = rp::parse_insert_chart(&resp);
            if rp::is_success(status_code) {
                output::success("Chart inserted.");
                if let Some(id) = chart_id {
                    output::key_value("Chart ID", &id, 2);
                }
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to insert chart: {}",
                    status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_delete(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: chart delete <chartName|chartId>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_delete_chart(rid, &sid, &chart_id, 0, 0);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' deleted.", chart_id));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to delete chart: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_move(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: chart move <chartName|chartId> <destinationSheet>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let (dest_sheet_id, dest_name) = match resolve_sheet_id(args[1], session) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_move_chart(rid, &sid, &dest_sheet_id, &chart_id, 0, 0);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' moved to sheet '{}'.", chart_id, dest_name));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to move chart: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_clone(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: chart clone <chartName|chartId>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_clone_chart(rid, &sid, &chart_id, 0, 0);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let (status_code, status_message, new_chart_id) = rp::parse_clone_chart(&resp);
            if rp::is_success(status_code) {
                output::success(&format!("Chart '{}' cloned.", chart_id));
                if let Some(id) = new_chart_id {
                    output::key_value("New Chart ID", &id, 2);
                }
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to clone chart: {}",
                    status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_position(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: chart position <chartName|chartId> <startX,startY,endX,endY> [--postype 0|1|2]");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let parts: Vec<&str> = args[1].split(',').collect();
    if parts.len() != 4 {
        output::error("Position format: startX,startY,endX,endY");
        return;
    }
    let nums: Vec<i32> = match parts.iter().map(|p| p.parse::<i32>()).collect::<Result<Vec<_>, _>>() {
        Ok(v) => v,
        Err(_) => { output::error("Position values must be integers."); return; }
    };
    // Accept both positional and --postype flag for consistency with chart insert
    let position_type: i32 = if let Some(pos) = args.iter().position(|a| a.eq_ignore_ascii_case("--postype")) {
        if pos + 1 < args.len() {
            args[pos + 1].parse().unwrap_or(0)
        } else { 0 }
    } else if args.len() > 2 && !args[2].starts_with("--") {
        args[2].parse().unwrap_or(0)
    } else {
        0
    };

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_update_chart_position(rid, &sid, &chart_id, nums[0], nums[1], nums[2], nums[3], position_type);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' position updated.", chart_id));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to update chart position: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_type(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: chart type <chartName|chartId> <type>");
        output::info("Type uses combined type_subtype format (e.g. bar_stacked, line_spline). Run 'chart insert' for full list.");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let (chart_type, chart_sub_type) = match parse_chart_type_subtype(args[1]) {
        Some(v) => v,
        None => { output::error("Invalid chart type. Use combined type_subtype format (e.g. bar_stacked, line_spline). Run 'chart insert' for full list."); return; }
    };

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_update_chart_type(rid, &sid, &chart_id, chart_type, chart_sub_type, 0, 0);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' type updated.", chart_id));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to update chart type: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_manage(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: chart get <id|range|name> <value>  (alias: chart manage)");
        output::info("  chart get range <A1:C5>  — get charts in range");
        output::info("  chart get id <id1,id2,...>  — get charts by IDs");
        output::info("  chart get name <chartName>  — get chart by name/title");
        return;
    }
    let rid = session.rid.clone().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let request = match args[0].to_lowercase().as_str() {
        "range" => {
            if args.len() < 2 {
                output::error("Usage: chart get range <A1:C5>");
                return;
            }
            let (sc, sr, ec, er) = parse_range_arg!(args[1]);
            let range_list = vec![serde_json::json!({
                "start_row": sr,
                "start_column": sc,
                "end_row": er,
                "end_column": ec
            })];
            rb::build_manage_chart_with_range(&rid, &sid, range_list)
        }
        "id" => {
            if args.len() < 2 {
                output::error("Usage: chart get id <id1,id2,...>");
                return;
            }
            let ids: Vec<String> = args[1].split(',').map(|s| s.trim().to_string()).collect();
            rb::build_manage_chart_with_id(&rid, &sid, ids)
        }
        "name" => {
            if args.len() < 2 {
                output::error("Usage: chart get name <chartName>");
                return;
            }
            let resolved = resolve_chart_id(args[1], engine, session);
            match resolved {
                Some(id) => rb::build_manage_chart_with_id(&rid, &sid, vec![id]),
                None => return,
            }
        }
        other => {
            // Treat as a chart name/id directly — resolve and use id mode
            let resolved = resolve_chart_id(other, engine, session);
            match resolved {
                Some(id) => rb::build_manage_chart_with_id(&rid, &sid, vec![id]),
                None => return,
            }
        }
    };

    match engine.process_request_json(&request) {
        Ok(resp) => {
            let (status_code, status_message, charts) = rp::parse_manage_chart(&resp);
            if rp::is_success(status_code) {
                if charts.is_empty() {
                    output::info("No charts found.");
                } else {
                    output::line(&format!("Found {} chart(s):", charts.len()), 0);
                    for c in &charts {
                        let title_display = if c.chart_title.is_empty() { "(untitled)" } else { &c.chart_title };
                        output::line(&format!(
                            "  ID: {}  title: \"{}\"  type: {} sub: {}  pos: ({},{},{},{})  posType: {}",
                            c.chart_id, title_display, c.chart_type, c.chart_sub_type,
                            c.start_x, c.start_y, c.end_x, c.end_y, c.position_type
                        ), 0);
                    }
                }
            } else {
                output::error(&format!(
                    "Failed to manage charts: {}",
                    status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_recommend(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: chart recommend <range>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let range_list = vec![serde_json::json!({
        "sheet_id": sid,
        "start_row": sr,
        "start_column": sc,
        "end_row": er,
        "end_column": ec
    })];
    let request = rb::build_recommend_chart(rid, &sid, range_list);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let (status_code, status_message, recs) = rp::parse_recommend_chart(&resp);
            if rp::is_success(status_code) {
                if recs.is_empty() {
                    output::info("No chart recommendations.");
                } else {
                    let type_names = ["BAR","COLUMN","LINE","PIE","AREA","SCATTER","RACE",
                        "WATERFALL","BULLET","FUNNEL","PARETO","HISTOGRAM","STOCK",
                        "RADAR","WORDCLOUD","COMBO","BOXPLOT"];
                    output::line(&format!("Recommended charts ({}):", recs.len()), 0);
                    for r in &recs {
                        let name = type_names.get(r.chart_type as usize).unwrap_or(&"UNKNOWN");
                        output::line(&format!("  {} (type:{}, sub:{})", name, r.chart_type, r.chart_sub_type), 0);
                    }
                }
            } else {
                output::error(&format!(
                    "Failed to get recommendations: {}",
                    status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_rename(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: chart rename <chartName|chartId> <newName>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let new_name = args[1..].join(" ");

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    // Use action 6001 sub 122 (same as chart style title) — this updates the
    // title field that manage_chart returns.
    let request = rb::build_customize_chart_with_subaction(
        rid, &sid, &chart_id,
        rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 122,
        serde_json::json!({"title_string": new_name}),
    );
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart renamed to '{}'.", new_name));
                session.chart_name_cache.insert(new_name.to_lowercase(), chart_id);
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to rename chart: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_style(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: chart style <chartName|chartId> <property> <value>");
        output::info("Properties: title, titlestyle, titlealign, subtitle, subtitlestyle, bgcolor, border,");
        output::info("            font, transparency, animation, gradient, tooltip, spline,");
        output::info("            legend, legendstyle, invert, 3d, colorscheme");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let property = args[1].to_lowercase();
    let value = args[2..].join(" ");

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let (action_id, sub_action_id, chart_properties) = match property.as_str() {
        "title" => (
            rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 122,
            serde_json::json!({"title_string": value}),
        ),
        "titlestyle" => {
            let mut props = serde_json::Map::new();
            props.insert("is_bold".to_string(), serde_json::json!(false));
            props.insert("is_italic".to_string(), serde_json::json!(false));
            props.insert("is_default_color".to_string(), serde_json::json!(true));

            if let Some(custom_props) = parse_kv_to_json(&args[2..]) {
                if let Some(custom_map) = custom_props.as_object() {
                    for (k, v) in custom_map {
                        props.insert(k.to_string(), v.clone());
                    }
                }
            } else {
                return;
            }

            (
                rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 119,
                serde_json::Value::Object(props),
            )
        }
        "titlealign" => {
            let align = match value.to_lowercase().as_str() {
                "left" => 0,
                "center" | "centre" => 1,
                "right" => 2,
                _ => match value.parse::<i32>() {
                    Ok(v) if (0..=2).contains(&v) => v,
                    _ => {
                        output::error("titlealign must be 0|1|2 or left|center|right");
                        return;
                    }
                },
            };
            (
                rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 197,
                serde_json::json!({"chart_title_alignment": align}),
            )
        }
        "subtitle" => (
            rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 123,
            serde_json::json!({"title_string": value}),
        ),
        "subtitlestyle" => {
            let mut props = serde_json::Map::new();
            props.insert("is_bold".to_string(), serde_json::json!(false));
            props.insert("is_italic".to_string(), serde_json::json!(false));
            props.insert("is_default_color".to_string(), serde_json::json!(true));

            if let Some(custom_props) = parse_kv_to_json(&args[2..]) {
                if let Some(custom_map) = custom_props.as_object() {
                    for (k, v) in custom_map {
                        props.insert(k.to_string(), v.clone());
                    }
                }
            } else {
                return;
            }

            (
                rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 120,
                serde_json::Value::Object(props),
            )
        }
        "bgcolor" => {
            let parts: Vec<&str> = value.split(',').collect();
            if parts.len() != 3 {
                output::error("bgcolor requires r,g,b values (e.g. 255,0,0)");
                return;
            }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b),
                _ => { output::error("bgcolor values must be integers"); return; }
            };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 108,
             serde_json::json!({"color": {"red": r, "green": g, "blue": b}}))
        }
        "border" => {
            let parts: Vec<&str> = value.split(',').collect();
            if parts.len() != 3 {
                output::error("border requires r,g,b values (e.g. 0,0,0)");
                return;
            }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b),
                _ => { output::error("border values must be integers"); return; }
            };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 199,
             serde_json::json!({"color": {"red": r, "green": g, "blue": b}}))
        }
        "font" => (
            rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 114,
            serde_json::json!({"font_name": value}),
        ),
        "transparency" => {
            let val: i32 = match value.parse() {
                Ok(v) => v,
                Err(_) => { output::error("transparency must be an integer (0-100)"); return; }
            };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 109,
             serde_json::json!({"transparency": val}))
        }
        "animation" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("animation: use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 105,
             serde_json::json!({"is_animation_applied": on}))
        }
        "gradient" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("gradient: use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 190,
             serde_json::json!({"is_gradient_applied": on}))
        }
        "tooltip" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("tooltip: use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 106,
             serde_json::json!({"is_tool_tip_enabled": on}))
        }
        "spline" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("spline: use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 251,
             serde_json::json!({"is_spline_enabled": on}))
        }
        "legend" => {
            let pos: i32 = match value.to_lowercase().as_str() {
                "none" => 0,
                "top" => 1,
                "bottom" => 2,
                "left" => 3,
                "right" => 4,
                "top-right" | "topright" => 5,
                _ => match value.parse() {
                    Ok(v) if (0..=5).contains(&v) => v,
                    _ => { output::error("legend position must be 0-5 or one of: none, top, bottom, left, right, top-right"); return; }
                }
            };
            (rb::ACTION_UPDATE_CHART_TYPE, 101,
             serde_json::json!({"legend_position": pos}))
        }
        "legendstyle" => {
            let mut props = serde_json::Map::new();
            props.insert("is_bold".to_string(), serde_json::json!(false));
            props.insert("is_italic".to_string(), serde_json::json!(false));
            props.insert("is_default_color".to_string(), serde_json::json!(true));

            if let Some(custom_props) = parse_kv_to_json(&args[2..]) {
                if let Some(custom_map) = custom_props.as_object() {
                    for (k, v) in custom_map {
                        props.insert(k.to_string(), v.clone());
                    }
                }
            } else {
                return;
            }

            (
                rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 125,
                serde_json::Value::Object(props),
            )
        }
        "invert" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("invert: use on/off"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 155,
             serde_json::json!({"is_invert_chart": on, "series_index": 0}))
        }
        "3d" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("3d: use on/off"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 107,
             serde_json::json!({"is_3d_view_enabled": on}))
        }
        "colorscheme" => {
            let normalized = value.trim().to_lowercase();

            // Engine expects strict palette enum names, not UI labels.
            let palette = match normalized.as_str() {
                // Common UI alias used in logs/examples.
                "office" => Some((0, "CHART_COLOR_SCHEME_1".to_string())),

                // Quantitative palette aliases.
                "chart_color_scheme_1" | "scheme1" | "color_scheme_1" => Some((0, "CHART_COLOR_SCHEME_1".to_string())),
                "chart_color_scheme_2" | "scheme2" | "color_scheme_2" => Some((0, "CHART_COLOR_SCHEME_2".to_string())),
                "chart_color_scheme_3" | "scheme3" | "color_scheme_3" => Some((0, "CHART_COLOR_SCHEME_3".to_string())),
                "chart_color_scheme_4" | "scheme4" | "color_scheme_4" => Some((0, "CHART_COLOR_SCHEME_4".to_string())),

                // Sequential/monochromatic palette aliases.
                "chart_monochromatic_1" | "mono1" | "monochromatic1" => Some((1, "CHART_MONOCHROMATIC_1".to_string())),
                "chart_monochromatic_2" | "mono2" | "monochromatic2" => Some((1, "CHART_MONOCHROMATIC_2".to_string())),
                "chart_monochromatic_3" | "mono3" | "monochromatic3" => Some((1, "CHART_MONOCHROMATIC_3".to_string())),
                "chart_monochromatic_4" | "mono4" | "monochromatic4" => Some((1, "CHART_MONOCHROMATIC_4".to_string())),
                "chart_monochromatic_5" | "mono5" | "monochromatic5" => Some((1, "CHART_MONOCHROMATIC_5".to_string())),
                "chart_monochromatic_6" | "mono6" | "monochromatic6" => Some((1, "CHART_MONOCHROMATIC_6".to_string())),
                _ => None,
            };

            if let Some((palette_type, palette_name)) = palette {
                (
                    rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE,
                    133,
                    serde_json::json!({
                        "color_palette_type": palette_type,
                        "palette_name": palette_name
                    }),
                )
            } else {
                let parts: Vec<&str> = value.split(',').collect();
                if parts.len() == 3 {
                    let (r, g, b) = match (
                        parts[0].trim().parse::<i32>(),
                        parts[1].trim().parse::<i32>(),
                        parts[2].trim().parse::<i32>(),
                    ) {
                        (Ok(r), Ok(g), Ok(b)) => (r, g, b),
                        _ => {
                            output::error("colorscheme custom RGB must be integers: <r,g,b>");
                            return;
                        }
                    };

                    (
                        rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE,
                        133,
                        serde_json::json!({
                            "color_palette_type": 2,
                            "base_color": {"red": r, "green": g, "blue": b}
                        }),
                    )
                } else {
                    output::error("colorscheme must be one of: CHART_COLOR_SCHEME_1..4, CHART_MONOCHROMATIC_1..6, office, or custom RGB <r,g,b>");
                    return;
                }
            }
        }
        _ => {
            output::error(&format!("Unknown style property: '{}'. Use: title, titlestyle, titlealign, subtitle, subtitlestyle, bgcolor, border, font, transparency, animation, gradient, tooltip, spline, legend, legendstyle, invert, 3d, colorscheme", property));
            return;
        }
    };

    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, action_id, sub_action_id, chart_properties);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' style '{}' updated.", chart_id, property));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to update chart style: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_series(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: chart series <chart> <property> <value> [idx|--all]");
        output::info("Per-series properties (default --all when idx omitted):");
        output::info("  color <r,g,b> [idx|--all], transparency <0-100> [idx|--all],");
        output::info("  linestyle <0-10> [idx|--all], bordercolor <r,g,b> [idx|--all],");
        output::info("  marker <on|off> [idx|--all], markershape <0-4|name> [idx|--all],");
        output::info("  markersize <size> [idx|--all], markerfill <r,g,b> [idx|--all],");
        output::info("  markerborder <r,g,b> [idx|--all], combotype <0-5|name> [idx|--all]");
        output::info("Chart-wide properties (apply to entire chart, no idx):");
        output::info("  threshold <value|off>, thresholdcolor <r,g,b>, trendline <0-6|name>,");
        output::info("  trendlinepoly <degree>, trendlinemovingavg <period>,");
        output::info("  trendlinestyle <trendline_idx> <0-10>, trendlinecolor <trendline_idx> <r,g,b>,");
        output::info("  trendlinetransparency <trendline_idx> <0-100>, angle <on|off>,");
        output::info("  sort <on|off>, sortby <name|value>, sortorder <asc|desc>,");
        output::info("  startangle <deg>, endangle <deg>, sliceangle <deg>,");
        output::info("  racecount <n>, raceduration <seconds>, racecaption <on|off>,");
        output::info("  racecaptionstyle <k=v...>, raceseriesorder <top|bottom|0|1>, raceblank <auto|zero|last|0|1|2>,");
        output::info("  racecumulate <on|off>, racedecimals <n>");
        output::info("  Box-plot props: boxoutliers <on|off> [idx|--all], boxinnerpoints <on|off> [idx|--all],");
        output::info("  boxmeanliner <on|off> [idx|--all], boxmeanmarker <on|off> [idx|--all],");
        output::info("  boxoutliercolor <r,g,b> [idx|--all], boxmeancolor <r,g,b> [idx|--all],");
        output::info("  boxwhiskercolor <r,g,b> [idx|--all], boxmediancolor <r,g,b> [idx|--all],");
        output::info("  boxgroupheaders <on|off>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let property = args[1].to_lowercase();
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let (action_id, sub_action_id, chart_properties) = match property.as_str() {
        "color" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> color <r,g,b> [idx|--all]"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("color requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            let mut props = serde_json::json!({"color": {"red": r, "green": g, "blue": b}});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("series_index must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 112, props)
        }
        "transparency" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> transparency <0-100> [idx|--all]"); return; }
            let val: f64 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be number"); return; } };
            let mut props = serde_json::json!({"transparency": val});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("series_index must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 113, props)
        }
        "linestyle" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> linestyle <0-10> [idx|--all]"); return; }
            let style: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            let mut props = serde_json::json!({"line_style": style});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("series_index must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 183, props)
        }
        "bordercolor" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> bordercolor <r,g,b> [--all|series_index]"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("bordercolor requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            let mut props = serde_json::json!({"border_type": 2, "border_color": {"red": r, "green": g, "blue": b}});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("series_index must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 174, props)
        }
        "threshold" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> threshold <value|off>"); return; }
            let props = if args[2] == "off" {
                serde_json::json!({"is_enabled": false})
            } else {
                let val: f64 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be number or 'off'"); return; } };
                serde_json::json!({"is_enabled": true, "threshold_value": val})
            };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 195, props)
        }
        "thresholdcolor" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> thresholdcolor <r,g,b>"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("Requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 196,
             serde_json::json!({"threshold_color": {"red": r, "green": g, "blue": b}}))
        }
        "trendline" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> trendline <0-6|none|linear|power|exponential|logarithmic|polynomial|moving_average>"); return; }
            let t: i32 = match args[2].parse() {
                Ok(v) => v,
                Err(_) => match args[2].to_lowercase().as_str() {
                    "none" => 0,
                    "linear" => 1,
                    "power" => 2,
                    "exponential" | "exp" => 3,
                    "logarithmic" | "log" => 4,
                    "polynomial" | "poly" => 5,
                    "moving_average" | "movingavg" | "moving-average" => 6,
                    _ => { output::error("Use 0-6 or one of: none, linear, power, exponential, logarithmic, polynomial, moving_average"); return; }
                }
            };
            (rb::ACTION_UPDATE_CHART_TYPE, 137, serde_json::json!({"trendline_type": t}))
        }
        "trendlinepoly" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> trendlinepoly <degree>"); return; }
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 138, serde_json::json!({"trendline_value": val}))
        }
        "trendlinemovingavg" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> trendlinemovingavg <period>"); return; }
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 148, serde_json::json!({"trendline_value": val}))
        }
        "trendlinestyle" => {
            if args.len() < 4 { output::error("Usage: chart series <chart> trendlinestyle <trendline_idx> <0-10>"); return; }
            let trendline_index: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("trendline_idx must be int"); return; } };
            let style: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("line_style must be int"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 121, serde_json::json!({"trendline_index": trendline_index, "line_style": style}))
        }
        "trendlinecolor" => {
            if args.len() < 4 { output::error("Usage: chart series <chart> trendlinecolor <trendline_idx> <r,g,b>"); return; }
            let trendline_index: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("trendline_idx must be int"); return; } };
            let parts: Vec<&str> = args[3].split(',').collect();
            if parts.len() != 3 { output::error("color requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 110,
             serde_json::json!({"trendline_index": trendline_index, "color": {"red": r, "green": g, "blue": b}}))
        }
        "trendlinetransparency" => {
            if args.len() < 4 { output::error("Usage: chart series <chart> trendlinetransparency <trendline_idx> <0-100>"); return; }
            let trendline_index: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("trendline_idx must be int"); return; } };
            let val: f64 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("transparency must be number"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 111,
             serde_json::json!({"trendline_index": trendline_index, "transparency": val}))
        }
        "sort" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> sort <on|off>"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 173, serde_json::json!({"is_sort_enabled": on}))
        }
        "sortby" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> sortby <name|value>"); return; }
            let t = match args[2] { "name" => 0, "value" => 1, _ => { output::error("Use name/value"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 175, serde_json::json!({"sort_type": t}))
        }
        "sortorder" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> sortorder <asc|desc>"); return; }
            let o = match args[2] { "asc" => 0, "desc" => 1, _ => { output::error("Use asc/desc"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 176, serde_json::json!({"sort_order": o}))
        }
        "racecount" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> racecount <n>"); return; }
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 158, serde_json::json!({"display_count": val}))
        }
        "raceduration" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> raceduration <seconds>"); return; }
            let val: f64 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be number"); return; } };
            if !(0.0..=2.0).contains(&val) {
                output::error("Invalid animation duration. Must be between 0 and 2 seconds.");
                return;
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 157, serde_json::json!({"animation_duration": val}))
        }
        "racecaption" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> racecaption <on|off>"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 168, serde_json::json!({"is_caption_enabled": on}))
        }
        "racecaptionstyle" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> racecaptionstyle <k=v> ..."); return; }
            let custom_props = match parse_kv_to_json(&args[2..]) {
                Some(v) => v,
                None => return,
            };

            let mut props = serde_json::Map::new();
            props.insert("is_bold".to_string(), serde_json::json!(false));
            props.insert("is_italic".to_string(), serde_json::json!(false));
            props.insert("is_default_color".to_string(), serde_json::json!(true));

            if let Some(custom_map) = custom_props.as_object() {
                for (k, v) in custom_map {
                    props.insert(k.to_string(), v.clone());
                }
                // If user passes color fields, use custom color unless explicitly overridden.
                if custom_map.contains_key("color") && !custom_map.contains_key("is_default_color") {
                    props.insert("is_default_color".to_string(), serde_json::json!(false));
                }
            }

            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 163, serde_json::Value::Object(props))
        }
        "raceseriesorder" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> raceseriesorder <top|bottom|0|1>"); return; }
            let val: i32 = match args[2].to_lowercase().as_str() {
                "top" => 0,
                "bottom" => 1,
                _ => match args[2].parse() {
                    Ok(v) if (0..=1).contains(&v) => v,
                    _ => { output::error("Use top|bottom|0|1"); return; }
                }
            };
            (rb::ACTION_UPDATE_CHART_TYPE, 160, serde_json::json!({"series_order": val}))
        }
        "raceblank" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> raceblank <auto|zero|last|0|1|2>"); return; }
            let val: i32 = match args[2].to_lowercase().as_str() {
                "auto" | "interpolate" => 0,
                "zero" => 1,
                "last" | "last_valid_value" | "last-valid-value" => 2,
                _ => match args[2].parse() {
                    Ok(v) if (0..=2).contains(&v) => v,
                    _ => { output::error("Use auto|zero|last|0|1|2"); return; }
                }
            };
            (rb::ACTION_UPDATE_CHART_TYPE, 162, serde_json::json!({"blank_cell_value": val}))
        }
        "racecumulate" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> racecumulate <on|off>"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 167, serde_json::json!({"is_cumulate_values": on}))
        }
        "racedecimals" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> racedecimals <n>"); return; }
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 177, serde_json::json!({"decimal_places": val}))
        }
        "boxoutliers" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxoutliers <on|off> [idx|--all]"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            let mut props = serde_json::json!({"is_show_outliers": on});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 200, props)
        }
        "boxinnerpoints" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxinnerpoints <on|off> [idx|--all]"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            let mut props = serde_json::json!({"is_show_inner_points": on});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 202, props)
        }
        "boxmeanliner" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxmeanliner <on|off> [idx|--all]"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            let mut props = serde_json::json!({"is_show_mean_liner": on});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 203, props)
        }
        "boxmeanmarker" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxmeanmarker <on|off> [idx|--all]"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            let mut props = serde_json::json!({"is_show_mean_marker": on});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 208, props)
        }
        "boxoutliercolor" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxoutliercolor <r,g,b> [idx|--all]"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("boxoutliercolor requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            let mut props = serde_json::json!({"outliers_color": {"red": r, "green": g, "blue": b}});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 204, props)
        }
        "boxmeancolor" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxmeancolor <r,g,b> [idx|--all]"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("boxmeancolor requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            let mut props = serde_json::json!({"mean_color": {"red": r, "green": g, "blue": b}});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 205, props)
        }
        "boxwhiskercolor" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxwhiskercolor <r,g,b> [idx|--all]"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("boxwhiskercolor requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            let mut props = serde_json::json!({"whiskers_color": {"red": r, "green": g, "blue": b}});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 206, props)
        }
        "boxmediancolor" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxmediancolor <r,g,b> [idx|--all]"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("boxmediancolor requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            let mut props = serde_json::json!({"median_color": {"red": r, "green": g, "blue": b}});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 207, props)
        }
        "boxgroupheaders" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> boxgroupheaders <on|off>"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 209, serde_json::json!({"is_grouped_box_plot": on}))
        }
        "marker" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> marker <on|off> [idx|--all]"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            let mut props = serde_json::json!({"is_enabled": on});
            if args.len() > 3 && args[3] == "--all" { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            else if args.len() > 3 { let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } }; props["series_index"] = serde_json::json!(idx); }
            else { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 186, props)
        }
        "markershape" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> markershape <0-4|circle|square|diamond|triangle|triangle_down> [idx|--all]"); return; }
            let shape: i32 = match args[2].parse() {
                Ok(v) => v,
                Err(_) => match args[2].to_lowercase().as_str() {
                    "circle" => 0,
                    "square" => 1,
                    "diamond" => 2,
                    "triangle" => 3,
                    "triangle_down" | "triangledown" | "triangle-down" => 4,
                    _ => { output::error("Use 0-4 or one of: circle, square, diamond, triangle, triangle_down"); return; }
                }
            };
            let mut props = serde_json::json!({"marker_shape": shape});
            if args.len() > 3 && args[3] == "--all" { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            else if args.len() > 3 { let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } }; props["series_index"] = serde_json::json!(idx); }
            else { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 171, props)
        }
        "markersize" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> markersize <size> [idx|--all]"); return; }
            let size: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            let mut props = serde_json::json!({"marker_size": size});
            if args.len() > 3 && args[3] == "--all" { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            else if args.len() > 3 { let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } }; props["series_index"] = serde_json::json!(idx); }
            else { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 172, props)
        }
        "markerfill" | "markerfillcolor" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> markerfill <r,g,b> [idx|--all]"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("markerfill requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            let mut props = serde_json::json!({"fill_color": {"red": r, "green": g, "blue": b}});
            if args.len() > 3 && args[3] == "--all" { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            else if args.len() > 3 { let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } }; props["series_index"] = serde_json::json!(idx); }
            else { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 184, props)
        }
        "markerborder" | "markerbordercolor" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> markerborder <r,g,b> [idx|--all]"); return; }
            let parts: Vec<&str> = args[2].split(',').collect();
            if parts.len() != 3 { output::error("markerborder requires r,g,b"); return; }
            let (r, g, b) = match (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<i32>(), parts[2].trim().parse::<i32>()) {
                (Ok(r), Ok(g), Ok(b)) => (r, g, b), _ => { output::error("Invalid RGB"); return; }
            };
            let mut props = serde_json::json!({"border_color": {"red": r, "green": g, "blue": b}});
            if args.len() > 3 && args[3] == "--all" { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            else if args.len() > 3 { let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } }; props["series_index"] = serde_json::json!(idx); }
            else { props["is_apply_to_all_series"] = serde_json::json!(true); props["series_index"] = serde_json::json!(0); }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 185, props)
        }
        "combotype" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> combotype <0-5|bar|column|line|spline|stepline|area> [idx|--all]"); return; }
            let series_type: i32 = match args[2].parse() {
                Ok(v) => v,
                Err(_) => match args[2].to_lowercase().as_str() {
                    "bar" => 0,
                    "column" => 1,
                    "line" => 2,
                    "spline" => 3,
                    "stepline" | "step" => 4,
                    "area" => 5,
                    _ => { output::error("Use 0-5 or one of: bar, column, line, spline, stepline, area"); return; }
                }
            };
            let mut props = serde_json::json!({"series_type": series_type});
            if args.len() > 3 && args[3] == "--all" {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            } else if args.len() > 3 {
                let idx: i32 = match args[3].parse() { Ok(v) => v, Err(_) => { output::error("idx must be int"); return; } };
                props["series_index"] = serde_json::json!(idx);
            } else {
                props["is_apply_to_all_series"] = serde_json::json!(true);
                props["series_index"] = serde_json::json!(0);
            }
            (rb::ACTION_UPDATE_CHART_TYPE, 210, props)
        }
        "angle" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> angle <on|off>"); return; }
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 180, serde_json::json!({"is_angle_present": on}))
        }
        "startangle" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> startangle <degrees>"); return; }
            let a: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 178, serde_json::json!({"start_angle": a}))
        }
        "endangle" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> endangle <degrees>"); return; }
            let a: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 179, serde_json::json!({"end_angle": a}))
        }
        "sliceangle" => {
            if args.len() < 3 { output::error("Usage: chart series <chart> sliceangle <degrees>"); return; }
            let a: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 182, serde_json::json!({"slice_start_angle": a}))
        }
        _ => { output::error(&format!("Unknown series property: '{}'", property)); return; }
    };

    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, action_id, sub_action_id, chart_properties);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' series '{}' updated.", chart_id, property));
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_datalabel(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: chart datalabel <chart> <property> <value>");
        output::info("Properties: component <0-10>, position <0-5>, style <k=v>,");
        output::info("  total <on|off>, totalstyle <k=v>  (supported values depend on chart type)");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let property = args[1].to_lowercase();
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let (action_id, sub_action_id, chart_properties) = match property.as_str() {
        "component" => {
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 102, serde_json::json!({"data_label_component": val}))
        }
        "position" => {
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 170, serde_json::json!({"data_label_position": val}))
        }
        "style" => {
            let mut props = serde_json::Map::new();
            props.insert("is_bold".to_string(), serde_json::json!(false));
            props.insert("is_italic".to_string(), serde_json::json!(false));
            props.insert("is_default_color".to_string(), serde_json::json!(true));

            if let Some(custom_props) = parse_kv_to_json(&args[2..]) {
                if let Some(custom_map) = custom_props.as_object() {
                    for (k, v) in custom_map {
                        props.insert(k.to_string(), v.clone());
                    }
                }
            } else {
                return;
            }

            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 191, serde_json::Value::Object(props))
        }
        "total" => {
            let on = match args[2] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_UPDATE_CHART_TYPE, 193, serde_json::json!({"is_total_data_labels_enabled": on}))
        }
        "totalstyle" => {
            let mut props = serde_json::Map::new();
            props.insert("is_bold".to_string(), serde_json::json!(false));
            props.insert("is_italic".to_string(), serde_json::json!(false));
            props.insert("is_default_color".to_string(), serde_json::json!(true));

            if let Some(custom_props) = parse_kv_to_json(&args[2..]) {
                if let Some(custom_map) = custom_props.as_object() {
                    for (k, v) in custom_map {
                        props.insert(k.to_string(), v.clone());
                    }
                }
            } else {
                return;
            }

            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 194, serde_json::Value::Object(props))
        }
        _ => { output::error(&format!("Unknown datalabel property: '{}'", property)); return; }
    };

    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, action_id, sub_action_id, chart_properties);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' datalabel '{}' updated.", chart_id, property));
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_axis(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: chart axis <chart> <property> <value>");
        output::info("Properties: htitle <text>, vtitle <text>, htitlestyle <k=v...>, hlabelstyle <k=v...>,");
        output::info("            hreverse <on|off>, vreverse <on|off> (Y-axis log scale),");
        output::info("  multipleaxes <on|off>, slant <0-3>, stagger <0-2>, binning <interval>,");
        output::info("  vmin <number>, vmax <number>, vlogbase <int>, vlabelenabled <on|off>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let property = args[1].to_lowercase();
    let value = args[2..].join(" ");
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let (action_id, sub_action_id, chart_properties) = match property.as_str() {
        "htitle" => (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 115, serde_json::json!({"title_string": value})),
        "vtitle" => (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 116, serde_json::json!({"title_string": value, "series_index": 0})),
        "htitlestyle" => {
            let mut props = serde_json::Map::new();
            props.insert("is_bold".to_string(), serde_json::json!(false));
            props.insert("is_italic".to_string(), serde_json::json!(false));
            props.insert("is_default_color".to_string(), serde_json::json!(true));

            if let Some(custom_props) = parse_kv_to_json(&args[2..]) {
                if let Some(custom_map) = custom_props.as_object() {
                    for (k, v) in custom_map {
                        props.insert(k.to_string(), v.clone());
                    }
                }
            } else {
                return;
            }

            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 117, serde_json::Value::Object(props))
        }
        "hlabelstyle" => {
            let mut props = serde_json::Map::new();
            props.insert("is_bold".to_string(), serde_json::json!(false));
            props.insert("is_italic".to_string(), serde_json::json!(false));
            props.insert("is_default_color".to_string(), serde_json::json!(true));

            if let Some(custom_props) = parse_kv_to_json(&args[2..]) {
                if let Some(custom_map) = custom_props.as_object() {
                    for (k, v) in custom_map {
                        props.insert(k.to_string(), v.clone());
                    }
                }
            } else {
                return;
            }

            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 127, serde_json::Value::Object(props))
        }
        "hreverse" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 141, serde_json::json!({"is_horizontal_axis_reversed": on}))
        }
        "vreverse" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            let props = if on {
                // Native API requires a base value when enabling logarithmic scale.
                serde_json::json!({"is_scale_logarithmic": true, "is_y_axis": true, "base_value": 10})
            } else {
                serde_json::json!({"is_scale_logarithmic": false, "is_y_axis": true})
            };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 142, props)
        }
        "multipleaxes" | "multiyaxis" | "multipleyaxis" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 154, serde_json::json!({"is_multiple_y_axis_enabled": on}))
        }
        "slant" => {
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int 0-3"); return; } };
            if !(0..=3).contains(&val) {
                output::error("slant must be in range 0-3");
                return;
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 129, serde_json::json!({"slant_degree": val}))
        }
        "stagger" => {
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int 0-2"); return; } };
            if !(0..=2).contains(&val) {
                output::error("stagger must be in range 0-2");
                return;
            }
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 139, serde_json::json!({"stagger_lines": val}))
        }
        "binning" => {
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be int > 0"); return; } };
            if val <= 0 {
                output::error("binning interval must be > 0");
                return;
            }
            (rb::ACTION_UPDATE_CHART_TYPE, 151, serde_json::json!({"binning_interval": val}))
        }
        "vmin" => {
            let val: f64 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be a number"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 130, serde_json::json!({"is_y_axis": true, "axis_minimum_value": val}))
        }
        "vmax" => {
            let val: f64 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be a number"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 131, serde_json::json!({"is_y_axis": true, "axis_maximum_value": val}))
        }
        "vlogbase" => {
            let val: i32 = match args[2].parse() { Ok(v) => v, Err(_) => { output::error("Must be an integer (e.g., 2, 8, 10)"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 143, serde_json::json!({"is_y_axis": true, "base_value": val}))
        }
        "vlabelenabled" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            (rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 144, serde_json::json!({"is_y_axis": true, "is_label_enabled": on}))
        }
        _ => { output::error(&format!("Unknown axis property: '{}'. Use: htitle, vtitle, htitlestyle, hlabelstyle, hreverse, vreverse, multipleaxes, slant, stagger, binning, vmin, vmax, vlogbase, vlabelenabled", property)); return; }
    };

    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, action_id, sub_action_id, chart_properties);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' axis '{}' updated.", chart_id, property));
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_gridline(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 4 {
        output::error("Usage: chart gridline <chart> <x|y> <property> <value>");
        output::info("Properties: major <on|off>, minor <on|off>, majortype <0-10>, minortype <0-10>, majorcolor <default|r,g,b>, minorcolor <default|r,g,b>, counttype <0|1>, count <positive int>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let is_y_axis = match args[1] { "y" => true, "x" => false, _ => { output::error("Use x or y"); return; } };
    let property = args[2].to_lowercase();
    let value = args[3..].join(" ");
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let parse_gridline_color = |raw_value: &str| -> Option<serde_json::Value> {
        let normalized = raw_value.trim().to_lowercase();
        if normalized == "default" || normalized == "true" || normalized == "on" {
            return Some(serde_json::json!({"is_default_color": true}));
        }

        let parts: Vec<&str> = raw_value
            .split(|c: char| c == ',' || c.is_whitespace())
            .filter(|part| !part.is_empty())
            .collect();
        if parts.len() != 3 {
            return None;
        }

        let red: i32 = parts[0].parse().ok()?;
        let green: i32 = parts[1].parse().ok()?;
        let blue: i32 = parts[2].parse().ok()?;
        Some(serde_json::json!({
            "is_default_color": false,
            "red": red,
            "green": green,
            "blue": blue
        }))
    };

    let mut props = serde_json::json!({"is_y_axis": is_y_axis});
    match property.as_str() {
        "major" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            props["is_major_gridline_enabled"] = serde_json::json!(on);
        }
        "minor" => {
            let on = match value.as_str() { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
            props["is_minor_gridline_enabled"] = serde_json::json!(on);
        }
        "majortype" => {
            let val: i32 = match value.parse() { Ok(v) => v, Err(_) => { output::error("Must be int 0-10"); return; } };
            props["major_gridline_type"] = serde_json::json!(val);
        }
        "minortype" => {
            let val: i32 = match value.parse() { Ok(v) => v, Err(_) => { output::error("Must be int 0-10"); return; } };
            props["minor_gridline_type"] = serde_json::json!(val);
        }
        "majorcolor" | "major_gridline_color" => {
            let color = match parse_gridline_color(&value) {
                Some(v) => v,
                None => { output::error("Use majorcolor default|r,g,b"); return; }
            };
            props["major_gridline_color"] = color;
        }
        "minorcolor" | "minor_gridline_color" => {
            let color = match parse_gridline_color(&value) {
                Some(v) => v,
                None => { output::error("Use minorcolor default|r,g,b"); return; }
            };
            props["minor_gridline_color"] = color;
        }
        "counttype" | "grid_line_count_type" => {
            let val = match value.trim().to_lowercase().as_str() {
                "0" | "auto" => 0,
                "1" | "custom" => 1,
                _ => { output::error("Use counttype 0|1 (0=auto, 1=custom)"); return; }
            };
            props["grid_line_count_type"] = serde_json::json!(val);
        }
        "count" | "grid_line_count" => {
            let val: i32 = match value.parse() { Ok(v) => v, Err(_) => { output::error("Must be int > 0"); return; } };
            if val <= 0 {
                output::error("Gridline count must be > 0");
                return;
            }
            props["grid_line_count"] = serde_json::json!(val);
        }
        _ => { output::error(&format!("Unknown gridline property: '{}'", property)); return; }
    }

    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, 181, props);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' gridline updated.", chart_id));
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_autoexpand(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: chart autoexpand <chart> <on|off>");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let on = match args[1] { "on" | "true" => true, "off" | "false" => false, _ => { output::error("Use on/off"); return; } };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, rb::ACTION_UPDATE_CHART_TYPE, 145, serde_json::json!({"is_auto_expand": on}));
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' auto-expand set to {}.", chart_id, if on { "on" } else { "off" }));
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── Deprecated chart commands (will be removed next major version) ───────────

fn chart_property_deprecated(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    output::warning("DEPRECATED: 'chart property' is removed. Use the named equivalents instead:");
    output::warning("  bgcolor -> chart style bgcolor; series color -> chart series color <r,g,b> <idx>");
    output::warning("  title -> chart style title; font -> chart style font; etc.");
    output::warning("  See 'help chart' for the full list. This alias will be removed next major version.");
    if args.len() < 3 {
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let sub_action_id: i32 = match args[1].parse() { Ok(v) => v, Err(_) => { output::error("sub_action_id must be integer"); return; } };
    let chart_properties = match parse_kv_to_json(&args[2..]) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, rb::ACTION_CUSTOMIZE_CHART_PROPERTY_ONE, sub_action_id, chart_properties);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' property (sub:{}) updated.", chart_id, sub_action_id));
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_charttype_deprecated(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    output::warning("DEPRECATED: 'chart charttype' is removed. Use the named equivalents instead:");
    output::warning("  legend -> chart style legend; 3d -> chart style 3d; invert -> chart style invert;");
    output::warning("  trendline -> chart series trendline; autoexpand -> chart autoexpand");
    output::warning("  See 'help chart' for the full list. This alias will be removed next major version.");
    if args.len() < 3 {
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let sub_action_id: i32 = match args[1].parse() { Ok(v) => v, Err(_) => { output::error("sub_action_id must be integer"); return; } };
    let chart_properties = match parse_kv_to_json(&args[2..]) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, rb::ACTION_UPDATE_CHART_TYPE, sub_action_id, chart_properties);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' charttype (sub:{}) updated.", chart_id, sub_action_id));
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_property2_deprecated(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    output::warning("DEPRECATED: 'chart property2' is removed. Use 'chart customize' or 'chart source' instead.");
    output::warning("  See 'help chart' for the full list. This alias will be removed next major version.");
    if args.len() < 3 {
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let sub_action_id: i32 = match args[1].parse() { Ok(v) => v, Err(_) => { output::error("sub_action_id must be integer"); return; } };
    let chart_properties = match parse_kv_to_json(&args[2..]) {
        Some(v) => v,
        None => return,
    };
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_customize_chart_with_subaction(rid, &sid, &chart_id, rb::ACTION_CUSTOMIZE_CHART_PROPERTY_TWO, sub_action_id, chart_properties);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' property2 (sub:{}) updated.", chart_id, sub_action_id));
                session.is_dirty = true;
            } else {
                output::error(&format!("Failed: {}", result.status_message.unwrap_or_else(|| "engine error".into())));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn chart_customize(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: chart customize <chartName|chartId> [options]");
        output::info("Options: --range <A1:B10> (repeatable), --series-in <rows|columns>,");
        output::info("  --combine-horizontal <on|off>, --include-hidden <on|off>,");
        output::info("  --first-row-labels <on|off>, --first-col-labels <on|off>");
        output::info("Use 'chart style' for visual properties like bgcolor, font, animation, etc.");
        output::info("Deprecated: key=value form is still accepted for compatibility.");
        return;
    }
    let chart_id = match resolve_chart_id(args[0], engine, session) {
        Some(id) => id,
        None => return,
    };
    let mut props = serde_json::Map::new();
    let sid = session.get_active_sheet_id_or_default();
    let rest = &args[1..];
    let is_legacy_kv = rest.iter().any(|a| a.contains('=')) && rest.iter().all(|a| !a.starts_with("--"));

    if is_legacy_kv {
        output::warning("DEPRECATED: 'chart customize <key=value>' is deprecated; use flags like --range/--series-in/--include-hidden instead.");

        let mut i = 0usize;
        while i < rest.len() {
            let arg = rest[i];
            let (key, val, consumed) = if let Some((k, v)) = arg.split_once('=') {
                (k, v, 1usize)
            } else if i + 1 < rest.len() && !rest[i + 1].contains('=') {
                // Backward-compatible convenience: allow key value in addition to key=value.
                (arg, rest[i + 1], 2usize)
            } else {
                output::error(&format!("Invalid property format: '{}'. Use key=value.", arg));
                return;
            };

            if key == "sheet_range_list" {
                // Parse range references into proper sheet_range_list array
                // Accepts: ["Sheet1!A1:D6"] or just A1:D6
                let trimmed = val.trim_start_matches('[').trim_end_matches(']');
                let mut ranges = Vec::new();
                for entry in trimmed.split(',') {
                    let entry = entry.trim().trim_matches('"').trim_matches('\'');
                    if entry.is_empty() { continue; }
                    // Try to split on '!' for sheet reference
                    let (range_sheet_id, range_str) = if let Some(pos) = entry.find('!') {
                        (entry[..pos].to_string(), &entry[pos + 1..])
                    } else {
                        (sid.clone(), entry)
                    };
                    let (sc, sr, ec, er) = match cell_ref::try_parse_range(range_str) {
                        Some(v) => v,
                        None => { output::error(&format!("Invalid range: '{}'", entry)); return; }
                    };
                    ranges.push(serde_json::json!({
                        "sheet_id": range_sheet_id,
                        "start_row": sr,
                        "start_column": sc,
                        "end_row": er,
                        "end_column": ec
                    }));
                }
                props.insert(key.to_string(), serde_json::Value::Array(ranges));
            } else {
                match val {
                    "true" => { props.insert(key.to_string(), serde_json::Value::Bool(true)); }
                    "false" => { props.insert(key.to_string(), serde_json::Value::Bool(false)); }
                    _ => {
                        // Try parsing as JSON first (for numbers, arrays, objects)
                        if let Ok(json_val) = serde_json::from_str::<serde_json::Value>(val) {
                            props.insert(key.to_string(), json_val);
                        } else {
                            props.insert(key.to_string(), serde_json::Value::String(val.to_string()));
                        }
                    }
                }
            }

            i += consumed;
        }
    } else {
        let mut i = 0usize;
        let mut ranges: Vec<serde_json::Value> = Vec::new();
        while i < rest.len() {
            match rest[i] {
                "--range" => {
                    i += 1;
                    if i >= rest.len() {
                        output::error("--range requires a value like A1:B10");
                        return;
                    }
                    let range_arg = rest[i].trim();
                    if range_arg.starts_with('[')
                        || range_arg.starts_with('{')
                        || range_arg.contains('"')
                        || range_arg.contains('\'')
                    {
                        output::error("--range accepts a single A1:B10-style range per flag. JSON is not allowed.");
                        return;
                    }
                    let (range_sheet_id, range_str) = if let Some(pos) = range_arg.find('!') {
                        (range_arg[..pos].to_string(), &range_arg[pos + 1..])
                    } else {
                        (sid.clone(), range_arg)
                    };
                    let (sc, sr, ec, er) = match cell_ref::try_parse_range(range_str) {
                        Some(v) => v,
                        None => {
                            output::error(&format!("Invalid range: '{}'", range_arg));
                            return;
                        }
                    };
                    ranges.push(serde_json::json!({
                        "sheet_id": range_sheet_id,
                        "start_row": sr,
                        "start_column": sc,
                        "end_row": er,
                        "end_column": ec
                    }));
                }
                "--series-in" => {
                    i += 1;
                    if i >= rest.len() {
                        output::error("--series-in requires rows|columns");
                        return;
                    }
                    let is_rows = match rest[i].to_lowercase().as_str() {
                        "rows" => true,
                        "columns" => false,
                        _ => {
                            output::error("--series-in must be rows|columns");
                            return;
                        }
                    };
                    props.insert("is_series_in_rows".to_string(), serde_json::Value::Bool(is_rows));
                }
                "--combine-horizontal" => {
                    i += 1;
                    if i >= rest.len() {
                        output::error("--combine-horizontal requires on|off");
                        return;
                    }
                    let v = match parse_on_off_flag(rest[i], "--combine-horizontal") {
                        Ok(v) => v,
                        Err(e) => { output::error(&e); return; }
                    };
                    props.insert("is_combine_range_horizontally".to_string(), serde_json::Value::Bool(v));
                }
                "--include-hidden" => {
                    i += 1;
                    if i >= rest.len() {
                        output::error("--include-hidden requires on|off");
                        return;
                    }
                    let v = match parse_on_off_flag(rest[i], "--include-hidden") {
                        Ok(v) => v,
                        Err(e) => { output::error(&e); return; }
                    };
                    props.insert("is_include_hidden_cells".to_string(), serde_json::Value::Bool(v));
                }
                "--first-row-labels" => {
                    i += 1;
                    if i >= rest.len() {
                        output::error("--first-row-labels requires on|off");
                        return;
                    }
                    let v = match parse_on_off_flag(rest[i], "--first-row-labels") {
                        Ok(v) => v,
                        Err(e) => { output::error(&e); return; }
                    };
                    props.insert("is_first_row_label".to_string(), serde_json::Value::Bool(v));
                }
                "--first-col-labels" => {
                    i += 1;
                    if i >= rest.len() {
                        output::error("--first-col-labels requires on|off");
                        return;
                    }
                    let v = match parse_on_off_flag(rest[i], "--first-col-labels") {
                        Ok(v) => v,
                        Err(e) => { output::error(&e); return; }
                    };
                    props.insert("is_first_column_label".to_string(), serde_json::Value::Bool(v));
                }
                unknown if unknown.contains('=') => {
                    output::error("Do not mix key=value and --flags in one command. Use only one style.");
                    return;
                }
                unknown => {
                    output::error(&format!("Unknown option '{}'.", unknown));
                    output::info("Use: --range, --series-in, --combine-horizontal, --include-hidden, --first-row-labels, --first-col-labels");
                    return;
                }
            }
            i += 1;
        }

        if !ranges.is_empty() {
            props.insert("sheet_range_list".to_string(), serde_json::Value::Array(ranges));
        }
    }

    let rid = session.rid.as_deref().unwrap();
    // Internal request mapping retained for developers:
    // action_id=652 with rid/sheet_id/chart_id/chart_properties/active_info.
    let request = rb::build_customize_chart_property_two(rid, &sid, &chart_id, serde_json::Value::Object(props));
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let result = rp::parse_status_response(&resp);
            if rp::is_success(result.status_code) {
                output::success(&format!("Chart '{}' properties updated.", chart_id));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to customize chart: {}",
                    result.status_message.unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

// ─── Help ────────────────────────────────────────────────────────────────────

fn print_help(args: &[&str]) {
    let topic = args.first().map(|s| s.to_lowercase());

    match topic.as_deref() {
        None => print_help_overview(),
        Some("--all") | Some("all") => print_help_all(),
        Some("file") | Some("open") | Some("save") | Some("close") => print_help_file(),
        Some("worksheet") | Some("sheet") => print_help_worksheet(),
        Some("cell") | Some("cells") => print_help_cell(),
        Some("row") | Some("col") | Some("column") | Some("rows") | Some("columns") => print_help_rowcol(),
        Some("editing") | Some("copy") | Some("move") | Some("merge") | Some("clear")
        | Some("undo") | Some("redo") | Some("clipboardcopy") => print_help_editing(),
        Some("find") | Some("replace") | Some("sort") | Some("filter") | Some("search") => print_help_find_sort(),
        Some("view") | Some("freeze") | Some("unfreeze") | Some("name") => print_help_view_names(),
        Some("table") | Some("tables") => print_help_table(),
        Some("pivot") => print_help_pivot(),
        Some("chart") | Some("charts") => print_help_chart(),
        Some("format") | Some("formatting") => print_help_format(),
        Some(unknown) => {
            output::error(&format!("Unknown help topic: '{}'. Type 'help' for all topics.", unknown));
            println!();
            output::help_header("Available Help Topics");
            output::help_section("TOPICS");
            output::help_cmd("help file", "File operations (open, save, close)");
            output::help_cmd("help worksheet", "Worksheet management");
            output::help_cmd("help cell", "Cell get/set operations");
            output::help_cmd("help row", "Row & column operations");
            output::help_cmd("help editing", "Copy, move, merge, clear, undo/redo");
            output::help_cmd("help find", "Find, replace, sort & filter");
            output::help_cmd("help view", "Freeze panes & named ranges");
            output::help_cmd("help table", "Table operations");
            output::help_cmd("help pivot", "Pivot table operations");
            output::help_cmd("help chart", "Chart operations");
            output::help_cmd("help format", "Cell formatting & number formats");
            output::help_cmd("help --all", "Show all commands (full reference)");
            output::help_footer();
        }
    }
}

fn print_help_overview() {
    output::help_header("Zoho Sheet CLI \u{2014} Help");

    output::help_section("HELP TOPICS (type 'help <topic>' for details)");
    output::help_cmd("help file", "File operations (open, save, close)");
    output::help_cmd("help worksheet", "Worksheet management");
    output::help_cmd("help cell", "Cell get/set operations");
    output::help_cmd("help row", "Row & column operations");
    output::help_cmd("help editing", "Copy, move, merge, clear, undo/redo");
    output::help_cmd("help find", "Find, replace, sort & filter");
    output::help_cmd("help view", "Freeze panes & named ranges");
    output::help_cmd("help table", "Table operations");
    output::help_cmd("help pivot", "Pivot table operations");
    output::help_cmd("help chart", "Chart operations");
    output::help_cmd("help format", "Cell formatting & number formats");
    output::help_cmd("help --all", "Show all commands (full reference)");

    output::help_section("QUICK REFERENCE");
    output::help_cmd("open <filepath>", "Open a file (.xlsx, .csv, .tsv)");
    output::help_cmd("open --new <docname>", "Create a new blank workbook");
    output::help_cmd("save / save --as <path>", "Save workbook");
    output::help_cmd("close [--force]", "Close the current workbook");
    output::help_cmd("worksheet <sub>", "Manage worksheets");
    output::help_cmd("cell get|set <ref> ...", "Read/write cell values");
    output::help_cmd("row|col <sub> ...", "Row & column operations");
    output::help_cmd("copy|move <range> <dest>", "Copy/move ranges");
    output::help_cmd("find|replace ...", "Search & replace");
    output::help_cmd("sort <range> <col> ...", "Sort data");
    output::help_cmd("table <sub> ...", "Table operations");
    output::help_cmd("pivot list|create|... ", "Pivot table operations");
    output::help_cmd("chart <sub> ...", "Chart operations");
    output::help_cmd("format <prop> <range> ...", "Cell formatting");
    output::help_cmd("undo / redo", "Undo/redo actions");
    output::help_cmd("help / help <topic>", "Show help");
    output::help_cmd("exit / quit", "Exit the CLI");

    output::help_footer();
}

fn print_help_all() {
    output::help_header("Zoho Sheet CLI \u{2014} Full Command Reference");
    print_help_file_content();
    print_help_worksheet_content();
    print_help_cell_content();
    print_help_rowcol_content();
    print_help_editing_content();
    print_help_find_sort_content();
    print_help_view_names_content();
    print_help_table_content();
    print_help_pivot_content();
    print_help_chart_content();
    print_help_format_content();

    output::help_section("SESSION");
    output::help_cmd("help / help <topic>", "Show help");
    output::help_cmd("exit / quit", "Exit the CLI");
    output::help_footer();
}

fn print_help_file() {
    output::help_header("File Operations");
    print_help_file_content();
    output::help_footer();
}

fn print_help_file_content() {
    output::help_section("FILE");
    output::help_cmd("open <filepath>", "Open a local file (.xlsx, .csv, .tsv)");
    output::help_cmd("open --new <docname>", "Create a new blank workbook");
    output::help_cmd("save", "Save to original path");
    output::help_cmd("save --as <path>", "Save a copy / export (format from extension)");
    output::help_cmd("close [--force]", "Close the current workbook");
    output::help_detail("--force skips unsaved-changes prompt");
}

fn print_help_worksheet() {
    output::help_header("Worksheet Management");
    print_help_worksheet_content();
    output::help_footer();
}

fn print_help_worksheet_content() {
    output::help_section("WORKSHEETS");
    output::help_cmd("worksheet list", "List all sheets in the open workbook");
    output::help_cmd("worksheet switch <name|index>", "Switch the active sheet");
    output::help_cmd("worksheet add <name>", "Add a new sheet");
    output::help_cmd("worksheet delete <name|index>", "Delete a sheet");
    output::help_cmd("worksheet rename <old> <new>", "Rename a sheet");
    output::help_cmd("worksheet reorder <position>", "Move active sheet to position (0-based)");
    output::help_cmd("worksheet duplicate", "Duplicate the active sheet");
    output::help_cmd("worksheet hide [name|index]", "Hide a sheet");
    output::help_cmd("worksheet unhide <name|index>", "Unhide a sheet");
}

fn print_help_cell() {
    output::help_header("Cell Operations");
    print_help_cell_content();
    output::help_footer();
}

fn print_help_cell_content() {
    output::help_section("CELLS");
    output::help_cmd("cell get <ref>", "Get cell value (e.g., A1)");
    output::help_cmd("cell set <ref> <value>", "Set a cell value");
    output::help_cmd("cell set <ref> --formula <f>", "Set a formula in a cell");
}

fn print_help_rowcol() {
    output::help_header("Row & Column Operations");
    print_help_rowcol_content();
    output::help_footer();
}

fn print_help_rowcol_content() {
    output::help_section("ROWS & COLUMNS");
    output::help_cmd("row insert <row> [count]", "Insert rows (1-based)");
    output::help_cmd("row delete <row> [count]", "Delete rows");
    output::help_cmd("row hide <row> [endRow]", "Hide rows");
    output::help_cmd("row unhide <row> [endRow]", "Unhide rows");
    output::help_cmd("row resize <row> <height|--auto>", "Resize a row (or auto-fit)");
    output::help_cmd("col insert <col> [count]", "Insert columns (letter, e.g., B)");
    output::help_cmd("col delete <col> [count]", "Delete columns");
    output::help_cmd("col hide <col> [endCol]", "Hide columns");
    output::help_cmd("col unhide <col> [endCol]", "Unhide columns");
    output::help_cmd("col resize <col> <width|--auto>", "Resize a column (or auto-fit)");
}

fn print_help_editing() {
    output::help_header("Editing Commands");
    print_help_editing_content();
    output::help_footer();
}

fn print_help_editing_content() {
    output::help_section("EDITING");
    output::help_cmd("copy <range> <dest> [--values|--format]", "Copy range to destination");
    output::help_cmd("move <range> <dest>", "Move range to destination");
    output::help_cmd("clipboardcopy <range>", "Copy cell values to system clipboard");
    output::help_cmd("merge <range>", "Merge cells");
    output::help_cmd("merge undo <range>", "Unmerge cells");
    output::help_cmd("clear <range> [--content|--format]", "Clear cells");
    output::help_detail("Default (no flag): clears all content and formatting");
    output::help_cmd("undo", "Undo last action");
    output::help_cmd("redo", "Redo last undone action");
}

fn print_help_find_sort() {
    output::help_header("Find, Replace, Sort & Filter");
    print_help_find_sort_content();
    output::help_footer();
}

fn print_help_find_sort_content() {
    output::help_section("FIND & SORT");
    output::help_cmd("find <text> [--case] [--exact]", "Find in sheet");
    output::help_detail("--case: case-sensitive  --exact: whole-cell match");
    output::help_cmd("replace <old> <new> [--all] [--case] [--exact]", "Find and replace");
    output::help_detail("--all: replace all occurrences  --exact: whole-cell match");
    output::help_cmd("sort <range> <col> [--desc] [--header]", "Sort a range");
    output::help_detail("<col> is column letter; --header excludes first row from sort");
    output::help_cmd("filter create <range>", "Create auto-filter");
    output::help_cmd("filter remove", "Remove auto-filter");
}

fn print_help_view_names() {
    output::help_header("View & Named Ranges");
    print_help_view_names_content();
    output::help_footer();
}

fn print_help_view_names_content() {
    output::help_section("VIEW & NAMES");
    output::help_cmd("freeze <ref>", "Freeze panes at cell");
    output::help_cmd("unfreeze", "Unfreeze panes");
    output::help_cmd("name add <name> <expr> [comment]", "Add a named range");
    output::help_detail("<expr> is a range or formula; optional comment for documentation");
    output::help_cmd("name delete <name>", "Delete a named range");
    output::help_cmd("name list", "List all defined names");
}

fn print_help_table() {
    output::help_header("Table Operations");
    print_help_table_content();
    output::help_footer();
}

fn print_help_table_content() {
    output::help_section("TABLES");
    output::help_cmd("table list", "List all tables in active sheet");
    output::help_cmd("table create <range> [--headers]", "Create a table on range");
    output::help_detail("--headers: first row is treated as column headers");
    output::help_cmd("table select <range>", "Select table range");
    output::help_cmd("table delete <id|name> [--keep-format]", "Delete a table");
    output::help_detail("--keep-format: removes table but preserves cell formatting");
    output::help_cmd("table rename <id|name> <newName>", "Rename a table");
    output::help_cmd("table options <id|name> <type> <true|false>", "Toggle table option");
    output::help_detail("types: 0=Header Row  1=Total Row  2=Banded Rows");
    output::help_detail("       3=Banded Columns  4=First Column  5=Last Column  6=Filter Button");
    output::help_cmd("table source <id|name> <range>", "Change table source range");
    output::help_cmd("table style <id|name> <pattern> [--keep-format]", "Change table style (0-9)");
    output::help_detail("0=Light1 1=Light2 2=Light3 3=Light4 4=Light5");
    output::help_detail("5=Medium1 6=Medium2 7=Medium3 8=Dark1 9=Dark2");
    output::help_cmd("table defaultstyle <pattern>", "Set default table style");
    output::help_cmd("table insertrow <id|name> <range> [--above]", "Insert table row(s)");
    output::help_cmd("table insertcol <id|name> <range> [--after]", "Insert table column(s)");
    output::help_cmd("table deleterow <id|name> <range>", "Delete table row(s)");
    output::help_cmd("table deletecol <id|name> <range>", "Delete table column(s)");
    output::help_cmd("table manage <id|name>", "Get table info");
}

fn print_help_pivot() {
    output::help_header("Pivot Table Operations");
    print_help_pivot_content();
    output::help_footer();
}

fn print_help_pivot_content() {
    output::help_section("PIVOT TABLES (accept pivot name or pivot ID)");
    output::help_cmd("pivot list", "List all pivot tables across all sheets");
    output::help_cmd("pivot create <range> [--newsheet|--dest <cell>] [--name <n>]", "Create pivot table");
    output::help_cmd("pivot delete <pivot>", "Delete pivot table");
    output::help_cmd("pivot info <pivot>", "Get pivot table info (name, source, headers)");
    output::help_cmd("pivot fields <pivot>", "List all field indices and types");
    output::help_detail("Use field indices from here with selectfield, changefield, etc.");
    output::help_cmd("pivot refresh <pivot>", "Refresh pivot table");
    output::help_cmd("pivot rename <pivot> <newName>", "Rename pivot table");
    output::help_cmd("pivot move <pivot> <dest> [--sheet <n>]", "Move pivot table");
    output::help_cmd("pivot copy <pivot> <dest> [--sheet <n>]", "Copy pivot table");
    output::help_cmd("pivot selectfield <pivot> <headerIdx> <area> [fieldIdx]", "Select/add field");
    output::help_detail("areas: row(0) column(1) value(2) filter(3) none(4)");
    output::help_cmd("pivot changefield <pivot> <fieldIdx> <fromArea> <destIdx> <toArea>", "Move field between areas");
    output::help_detail("<destIdx> = destination position index within the target area");
    output::help_cmd("pivot filter <pivot> <fieldIdx> <area> --condition ...", "Condition filter on field");
    output::help_detail("Number ops: equals, notequals, gt, gte, lt, lte, between, top, bottom");
    output::help_detail("Text ops: contains, notcontains, beginswith, endswith, matchlabel");
    output::help_detail("Date ops: after, onorafter, before, onorbefore, betweendate");
    output::help_detail("'between'/'betweendate' require two values. Others need one.");
    output::help_cmd("pivot filter <pivot> <fieldIdx> <area> --selection ...", "Selection filter (indices)");
    output::help_cmd("pivot removefilter <pivot> <fieldIdx> <area>", "Remove filter from field");
    output::help_cmd("pivot filterinfo <cell>", "Get filter info for pivot cell");
    output::help_cmd("pivot sort <pivot> <fieldIdx> <area> <asc|desc> [aggIdx]", "Sort pivot field");
    output::help_cmd("pivot removesort <pivot> <fieldIdx> <area>", "Remove sort from field");
    output::help_cmd("pivot group <pivot> <fieldIdx> <area> <min> <max> <range>", "Numeric grouping");
    output::help_detail("Optional: [--mindefault] [--maxdefault]");
    output::help_cmd("pivot dategroup <pivot> <fieldIdx> <area> <types> <min> <max> <days>", "Date grouping");
    output::help_detail("  <types> is comma-separated: year,quarter,month,day,hour,minute,second");
    output::help_detail("  Example: pivot dategroup SalesByRegion 0 row year,month 2025-01-01 2025-12-31 1");
    output::help_detail("  Optional: [--mindefault] [--maxdefault]");
    output::help_cmd("pivot removegroup <pivot> <fieldIdx> <area>", "Remove grouping from field");
    output::help_cmd("pivot removefield <pivot> <fieldIdx> <area>", "Remove field from pivot");
    output::help_cmd("pivot properties <pivot> <prop> <true|false>", "Modify pivot property");
    output::help_detail("Props: subtotal, rowtotal, coltotal, repeat, hideerrors");
    output::help_cmd("pivot aggregation <pivot> <fieldIdx> <type>", "Change value aggregation");
    output::help_detail("Types: sum, count, countnums, distinct, avg, min, max, median, product, stdev, stdevp, var, varp");
    output::help_cmd("pivot showdataas <pivot> <fieldIdx> <type>", "Change show-data-as");
    output::help_detail("Types: nochange, percent_row, percent_col, percent_total");
    output::help_cmd("pivot changesource <pivot> <range> [--sheet <n>]", "Change pivot source range");
    output::help_cmd("pivot cellinfo <cell>", "Get info for a pivot cell");
    output::help_cmd("pivot refreshonload <true|false>", "Enable/disable refresh on file load");
}

fn print_help_chart() {
    output::help_header("Chart Operations");
    print_help_chart_content();
    output::help_footer();
}

fn print_help_chart_content() {
    output::help_section("CHARTS (accept chart name or chart ID)");
    output::help_cmd("chart list [--all]", "List charts on active sheet (--all for all sheets)");
    output::help_cmd("chart insert <range> <type> [--pos x1,y1,x2,y2] [--postype 0|1|2]", "Insert a chart");
    output::help_detail("  <range>   Data range, e.g. A1:D10");
    output::help_detail("  <type>    Combined type_subtype name:");
    output::help_detail("    BAR:       bar, bar_stacked, bar_stacked_100, bar_grouped");
    output::help_detail("    COLUMN:    column, column_stacked, column_stacked_100, column_grouped");
    output::help_detail("    LINE:      line, line_spline, line_step, line_timeline");
    output::help_detail("    PIE:       pie, pie_semi, pie_doughnut, pie_semi_doughnut, pie_parliament, doughnut_parliament");
    output::help_detail("    AREA:      area, area_stacked, area_stacked_100, area_time");
    output::help_detail("    SCATTER:   scatter, scatter_line, scatter_line_markers, scatter_bubble");
    output::help_detail("    RACE:      race");
    output::help_detail("    WATERFALL: waterfall");
    output::help_detail("    BULLET:    bullet, bullet_vertical");
    output::help_detail("    FUNNEL:    funnel, funnel_weighted");
    output::help_detail("    PARETO:    pareto");
    output::help_detail("    HISTOGRAM: histogram");
    output::help_detail("    STOCK:     stock, stock_ohlc");
    output::help_detail("    RADAR:     radar, radar_spiderweb");
    output::help_detail("    WORDCLOUD: wordcloud");
    output::help_detail("    COMBO:     combo");
    output::help_detail("    BOXPLOT:   boxplot, boxplot_grouped_horizontal, boxplot_vertical, boxplot_grouped_vertical");
    output::help_detail("  --pos     Pixel coordinates: startX,startY,endX,endY");
    output::help_detail("  --postype 0=absolute pixels  1=one-cell anchor  2=two-cell anchor");
    output::help_detail("  Example: chart insert A1:D10 line_spline --pos 50,50,500,300");
    output::help_cmd("chart delete <chart>", "Delete a chart");
    output::help_cmd("chart move <chart> <sheet>", "Move chart to another sheet");
    output::help_detail("  <sheet>   Target sheet name or ID");
    output::help_cmd("chart clone <chart>", "Clone a chart on the same sheet");
    output::help_cmd("chart rename <chart> <newName>", "Rename a chart");
    output::help_cmd("chart info <chart>", "Get chart details (type, position, title)");
    output::help_cmd("chart source <chart> <range>", "Change chart data source range");
    output::help_detail("  <range>   New data range, e.g. A1:E20");
    output::help_cmd("chart position <chart> <x1,y1,x2,y2> [--postype 0|1|2]", "Update chart position");
    output::help_detail("  Coordinates are startX,startY,endX,endY (comma-separated, no spaces)");
    output::help_cmd("chart type <chart> <type>", "Change chart type");
    output::help_detail("  <type>    Combined type_subtype name (see 'chart insert' for full list).");
    output::help_detail("  Example: chart type MyChart pie_doughnut");
    output::help_detail("  Note: not every type/subtype is valid for every existing chart or engine build.");
    output::help_detail("        Some conversions require specific source data shapes (for example timeline/stock). ");
    output::help_cmd("chart get range <range>", "Get charts overlapping a range");
    output::help_cmd("chart get id <id1,id2,...>", "Get charts by IDs (comma-separated)");
    output::help_cmd("chart get name <chartName>", "Get chart by name/title");
    output::help_detail("  (\"chart manage\" is an alias for \"chart get\")");
    output::help_cmd("chart recommend <range>", "Get chart type recommendations for data");
    output::help_cmd("chart customize <chart> [options]", "Set chart data-source properties");
    output::help_detail("  --range <A1:B10>            Data range; repeat for multi-range");
    output::help_detail("                              e.g. --range A1:B10 --range D1:E10");
    output::help_detail("  --series-in <rows|columns>  Where each series lives (default: columns)");
    output::help_detail("  --combine-horizontal <on|off>  Join multiple ranges horizontally");
    output::help_detail("  --include-hidden <on|off>   Include hidden cells (default: off)");
    output::help_detail("  --first-row-labels <on|off> Treat first row as labels");
    output::help_detail("  --first-col-labels <on|off> Treat first column as labels");
    output::help_detail("  (Use 'chart style' for visual appearance properties.)");
    output::help_cmd("chart style <chart> <property> <value>", "Style chart appearance");
    output::help_detail("  title <text>          Set chart title");
    output::help_detail("  titlestyle <k=v...>   Set title text style (sub-action 119)");
    output::help_detail("    keys: font_size=<int> is_italic=<bool> is_bold=<bool>");
    output::help_detail("          is_default_color=<bool> color.red=<0-255> color.green=<0-255> color.blue=<0-255>");
    output::help_detail("    Example: chart style MyChart titlestyle font_size=12 is_bold=true color.red=110 color.green=100 color.blue=90 is_default_color=false");
    output::help_detail("  titlealign <0|1|2|left|center|right>  Set title alignment (0=left 1=center 2=right)");
    output::help_detail("  subtitle <text>       Set chart subtitle");
    output::help_detail("  subtitlestyle <k=v...> Set subtitle text style (sub-action 120)");
    output::help_detail("    keys: font_size=<int> is_italic=<bool> is_bold=<bool>");
    output::help_detail("          is_default_color=<bool> color.red=<0-255> color.green=<0-255> color.blue=<0-255>");
    output::help_detail("    Example: chart style MyChart subtitlestyle font_size=10 is_italic=true color.red=110 color.green=100 color.blue=90 is_default_color=false");
    output::help_detail("  bgcolor <r,g,b>       Background color (0-255 each)");
    output::help_detail("  border <r,g,b>        Border color");
    output::help_detail("  font <name>           Font family (e.g. Arial, Roboto, Open Sans)");
    output::help_detail("  transparency <0-100>  Chart transparency percentage");
    output::help_detail("  animation <on|off>    Toggle animation");
    output::help_detail("  gradient <on|off>     Toggle gradient fill");
    output::help_detail("  tooltip <on|off>      Toggle hover tooltips");
    output::help_detail("  spline <on|off>       Smooth line curves");
    output::help_detail("  legend <0-5|name>     Legend position (0=none 1=top 2=bottom 3=left 4=right 5=top-right)");
    output::help_detail("  legendstyle <k=v...>  Set legend text style (sub-action 125)");
    output::help_detail("    keys: font_size=<int> is_italic=<bool> is_bold=<bool>");
    output::help_detail("          is_default_color=<bool> color.red=<0-255> color.green=<0-255> color.blue=<0-255>");
    output::help_detail("  invert <on|off>       Swap category/value axes (transpose the plot);");
    output::help_detail("                        distinct from axis hreverse (x direction) and vreverse (y log scale)");
    output::help_detail("  3d <on|off>           Toggle 3D view");
    output::help_detail("  colorscheme <name|r,g,b>  Color palette or custom base color");
    output::help_detail("    Names: CHART_COLOR_SCHEME_1..4, CHART_MONOCHROMATIC_1..6, office(alias of CHART_COLOR_SCHEME_1)");
    output::help_detail("    Custom: chart style MyChart colorscheme 110,100,90  (uses color_palette_type=2)");
    output::help_detail("  Example: chart style MyChart bgcolor 66,133,244");
    output::help_cmd("chart series <chart> <prop> <val> [idx|--all]", "Customize data series");
    output::help_detail("  Per-series props (omit idx to apply to all series):");
    output::help_detail("  color <r,g,b> [idx|--all]         Series fill color");
    output::help_detail("  transparency <0-100> [idx|--all]  Series transparency");
    output::help_detail("  linestyle <0-10> [idx|--all]      Line dash style");
    output::help_detail("  bordercolor <r,g,b> [idx|--all]   Series border color");
    output::help_detail("  marker <on|off> [idx|--all]       Show data point markers");
    output::help_detail("  markershape <0-4|name> [idx|--all] Marker shape");
    output::help_detail("    0=circle 1=square 2=diamond 3=triangle 4=triangle_down");
    output::help_detail("  markersize <size> [idx|--all]     Marker pixel size");
    output::help_detail("  markerfill <r,g,b> [idx|--all]    Marker fill color");
    output::help_detail("  markerborder <r,g,b> [idx|--all]  Marker border color");
    output::help_detail("  combotype <0-5|name> [idx|--all]  Combo chart series type");
    output::help_detail("    0=bar 1=column 2=line 3=spline 4=stepline 5=area");
    output::help_detail("  Chart-wide props (no idx — apply to entire chart):");
    output::help_detail("  threshold <value|off>       Threshold line value");
    output::help_detail("  thresholdcolor <r,g,b>      Threshold line color");
    output::help_detail("  trendline <0-6|name>        0=none 1=linear 2=power 3=exp 4=log 5=poly 6=moving");
    output::help_detail("  trendlinepoly <degree>      Polynomial trendline degree");
    output::help_detail("  trendlinemovingavg <period> Moving average trendline period");
    output::help_detail("  trendlinestyle <idx> <0-10> Trendline line style by trendline index");
    output::help_detail("  trendlinecolor <idx> <r,g,b> Trendline color by trendline index");
    output::help_detail("  trendlinetransparency <idx> <0-100> Trendline transparency by trendline index");
    output::help_detail("  angle <on|off>              Enable/disable start-end angle mode");
    output::help_detail("  sort <on|off>               Enable series sorting (chart-level)");
    output::help_detail("  sortby <name|value>         Sort criterion");
    output::help_detail("  sortorder <asc|desc>        Sort direction");
    output::help_detail("  startangle <deg>            Pie parliament/donut start angle (chart-level)");
    output::help_detail("  endangle <deg>              Pie parliament/donut end angle (chart-level)");
    output::help_detail("  sliceangle <deg>            Pie parliament/donut slice start angle (chart-level)");
    output::help_detail("  racecount <n>               Race chart: top/bottom value count");
    output::help_detail("  raceduration <seconds>      Race chart: animation duration");
    output::help_detail("  racecaption <on|off>        Race chart: show/hide caption");
    output::help_detail("  racecaptionstyle <k=v...>   Race chart caption style (font_size,is_italic,is_bold,color.red/green/blue)");
    output::help_detail("  raceseriesorder <top|bottom|0|1> Race chart ordering");
    output::help_detail("  raceblank <auto|zero|last|0|1|2> Race chart blank-cell handling");
    output::help_detail("  racecumulate <on|off>       Race chart cumulative values");
    output::help_detail("  racedecimals <n>            Race chart decimal precision");
    output::help_detail("  boxoutliers <on|off> [idx|--all]     Box plot: show outliers");
    output::help_detail("  boxinnerpoints <on|off> [idx|--all]  Box plot: show inner points");
    output::help_detail("  boxmeanliner <on|off> [idx|--all]    Box plot: show mean liner");
    output::help_detail("  boxmeanmarker <on|off> [idx|--all]   Box plot: show mean marker");
    output::help_detail("  boxoutliercolor <r,g,b> [idx|--all]  Box plot: outlier color");
    output::help_detail("  boxmeancolor <r,g,b> [idx|--all]     Box plot: mean color");
    output::help_detail("  boxwhiskercolor <r,g,b> [idx|--all]  Box plot: whisker color");
    output::help_detail("  boxmediancolor <r,g,b> [idx|--all]   Box plot: median color");
    output::help_detail("  boxgroupheaders <on|off>             Box plot: grouped/ungrouped headers");
    output::help_detail("  Example: chart series MyChart color 255,0,0 0");
    output::help_cmd("chart datalabel <chart> <prop> <value>", "Customize data labels");
    output::help_detail("  component <0-10>      Which parts to show (supported subset varies by chart type):");
    output::help_detail("    0=none 1=value 2=percentage 3=category 4=series name");
    output::help_detail("    5=val+% 6=val+cat 7=cat+% 8=series+val 9=series+% 10=all");
    output::help_detail("  position <0-5>        Label placement (supported subset varies by chart type):");
    output::help_detail("    0=auto 1=center 2=inside-end 3=outside-end 4=best-fit 5=above");
    output::help_detail("  style <k=v ...>       Font styling: font_size=12 is_bold=true is_italic=false");
    output::help_detail("                        is_default_color=<bool> color.red=<0-255> color.green=<0-255> color.blue=<0-255>");
    output::help_detail("  total <on|off>        Show total labels (only on supported chart types)");
    output::help_detail("  totalstyle <k=v ...>  Style for total labels (same keys as style)");
    output::help_cmd("chart axis <chart> <prop> <value>", "Customize axis");
    output::help_detail("  htitle <text>         Horizontal (category) axis title");
    output::help_detail("  vtitle <text>         Vertical (value) axis title");
    output::help_detail("  htitlestyle <k=v...>  Horizontal axis title style (sub-action 117)");
    output::help_detail("  hlabelstyle <k=v...>  Horizontal axis label style (sub-action 127)");
    output::help_detail("    keys: font_size=<int> is_italic=<bool> is_bold=<bool>");
    output::help_detail("          is_default_color=<bool> color.red=<0-255> color.green=<0-255> color.blue=<0-255>");
    output::help_detail("  hreverse <on|off>     Reverse horizontal axis direction");
    output::help_detail("  vreverse <on|off>     Toggle Y-axis logarithmic scale (base=10 when enabled)");
    output::help_detail("  multipleaxes <on|off> Enable/disable multiple vertical axes");
    output::help_detail("  slant <0-3>           Axis label slant (0=none 1=45° 2=90° 3=auto)");
    output::help_detail("  stagger <0-2>         Stagger label lines (0=none)");
    output::help_detail("  binning <interval>    Histogram bin width (integer > 0)");
    output::help_detail("  vmin <number>         Vertical axis minimum value");
    output::help_detail("  vmax <number>         Vertical axis maximum value");
    output::help_detail("  vlogbase <int>        Logarithmic scale base for Y-axis (e.g., 2, 8, 10)");
    output::help_detail("  vlabelenabled <on|off> Enable/disable Y-axis labels");
    output::help_detail("  Example: chart axis MyChart vtitle \"Revenue ($)\"");
    output::help_cmd("chart gridline <chart> <x|y> <prop> <val>", "Customize gridlines");
    output::help_detail("  <x|y>              Axis: x=horizontal (category), y=vertical (value)");
    output::help_detail("                     (corresponds to h/v in 'chart axis' commands)");
    output::help_detail("  major <on|off>     Show/hide major gridlines");
    output::help_detail("  minor <on|off>     Show/hide minor gridlines");
    output::help_detail("  majortype <0-10>   Major gridline dash style");
    output::help_detail("  minortype <0-10>   Minor gridline dash style");
    output::help_detail("  majorcolor <default|r,g,b> Major gridline color");
    output::help_detail("  minorcolor <default|r,g,b> Minor gridline color");
    output::help_detail("  counttype <0|1>    Gridline count mode (0=auto, 1=custom)");
    output::help_detail("  count <int>        Custom gridline count (> 0)");
    output::help_detail("  Example: chart gridline MyChart y major on");
    output::help_cmd("chart autoexpand <chart> <on|off>", "Auto-expand data range when new data is added");
}

fn print_help_format() {
    output::help_header("Formatting");
    print_help_format_content();
    output::help_footer();
}

fn print_help_format_content() {
    output::help_section("FORMATTING");
    output::help_cmd("format bold <range> <true|false>", "Toggle bold");
    output::help_cmd("format italic <range> <true|false>", "Toggle italic");
    output::help_cmd("format underline <range> <true|false>", "Toggle underline");
    output::help_cmd("format doubleunderline <range> <bool>", "Toggle double underline");
    output::help_cmd("format strikethrough <range> <bool>", "Toggle strikethrough");
    output::help_cmd("format superscript <range> <bool>", "Toggle superscript");
    output::help_cmd("format subscript <range> <bool>", "Toggle subscript");
    output::help_cmd("format fontsize <range> <size>", "Set font size");
    output::help_cmd("format fontcolor <range> <r> <g> <b>", "Set font color (RGB 0-255)");
    output::help_cmd("format fontcolor <range> --auto", "Set automatic font color");
    output::help_cmd("format halign <range> <type>", "Horizontal alignment");
    output::help_detail("types: general|left|center|right|fill|justify|centeracross|distributed");
    output::help_cmd("format valign <range> <type>", "Vertical alignment");
    output::help_detail("types: top|center|bottom|justify|distributed");
    output::help_cmd("format textwrap <range> <mode>", "Set text wrapping");
    output::help_detail("modes: overflow|clip|wrap|shrink");
    output::help_cmd("format rotate <range> <angle>", "Set text rotation (-90 to 90)");
    output::help_cmd("format indent <range> <increase|decrease>", "Adjust indent level");
    output::help_cmd("format fillcolor <range> <r> <g> <b>", "Set fill color (RGB 0-255)");
    output::help_cmd("format fillcolor <range> --none", "Remove fill color");
    output::help_cmd("format border <range> <type> <style> [r g b]", "Set border");
    output::help_detail("types: all|outer|inner|left|right|top|bottom|horizontal|vertical|diagonal");
    output::help_detail("styles: none|thin|medium|dashed|dotted|thick|double|hair|");
    output::help_detail("  mediumdashed|dashdot|mediumdashdot|dashdotdot|mediumdashdotdot|slantdashdot");

    output::help_section("NUMBER FORMATTING");
    output::help_cmd("format numformat <range> <type> [shortcut|--flags]", "Apply number format");
    output::help_detail("Types: general, number, currency, accounting, date, time,");
    output::help_detail("  duration, percentage, scientific, fraction, text, custom");
    output::help_detail("Number shortcuts: 1=#,##0  2=#,##0.00  3=#0  4=#0.00");
    output::help_detail("Date shortcuts: 1=dddd,d mmmm,yyyy  2=d mmmm yyyy  3=dd-mmm-yyyy  4=dd/mm/yy");
    output::help_detail("Time shortcuts: 1=h:mm:ss  2=h:mm:ss AM/PM");
    output::help_detail("Duration shortcuts: 1=[hh]:mm:ss  2=[hh]:mm  3=[hh]  4=[mm]  5=[ss]");
    output::help_detail("Percentage shortcuts: 1=0%  2=0.00%");
    output::help_detail("Scientific shortcuts: 1=0.00E+00  2=0.0E+00");
    output::help_detail("Fraction shortcuts: 1=# ?/?  2=# ??/??");
    output::help_detail("Flags: --decimals N, --noseparator, --leading-zeros N, --negative STYLE");
    output::help_detail("  --currency KEY, --prefix TEXT, --suffix TEXT, --digits N");
    output::help_detail("  --date PATTERN, --time PATTERN");
    output::help_detail("--negative styles: minus, red, red-minus, parens, red-parens");
    output::help_detail("Combine --date and --time with type 'custom'");
    output::help_cmd("format numformat --list-custom", "List saved custom formats");
    output::help_cmd("format numformat --list-currency", "List supported currency codes");
    output::help_cmd("format decimal <range> <increase|decrease>", "Adjust decimal places");
    output::help_cmd("format numpreview <ref|range> <type> <fmt>", "Preview number format on cell");
    output::help_detail("<fmt> is a raw pattern string or a shortcut index");
    output::help_cmd("format numinfo <cellRef>", "Get number format info for cell");
    output::help_cmd("format nummanage", "List all built-in format types and shortcuts");
    output::help_cmd("format customformat", "(Deprecated) Alias for numformat --list-custom");
    output::help_cmd("format default <range> [--flags...]", "Set default cell format");
    output::help_detail("Does not affect existing cell-level overrides");
    output::help_detail("Flags: --font-name, --font-size, --bold, --italic, --underline,");
    output::help_detail("  --font-color, --fill-color, --halign, --valign, --wrap");
}
