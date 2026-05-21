/// All CLI command implementations — one per verb.
///
/// Instead of the C# DI-based ICliCommand pattern, Rust uses a dispatch function
/// that takes shared references to the engine, session, and parser.
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

use crate::engine::ffi::EngineHandle;
use crate::engine::request_builder as rb;
use crate::engine::response_parser as rp;
use crate::output;
use crate::session::CliSession;
use crate::util::cell_ref;

/// Maximum allowed single argument length to guard the native engine.
const MAX_INPUT_ARG_LEN: usize = 32_768;
/// Max rows / columns for batch insert/delete.
const MAX_ROW_COL_COUNT: i32 = 10_000;
/// Supported import extensions.
const SUPPORTED_EXTENSIONS: &[&str] = &[".xlsx", ".csv", ".ods", ".xls", ".tsv"];

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
        "format" => cmd_format(&args, engine, session),
        "help" => { print_help(); }
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
            "Unsupported file type: '{}'. Supported: .xlsx, .csv, .ods, .xls, .tsv",
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

    let request = rb::build_open_workbook(&working_copy.to_string_lossy());
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

    let supported: Vec<&str> = vec!["xlsx", "csv", "ods", "pdf","tsv"];
    if !supported.contains(&fmt.as_str()) {
        output::error(&format!(
            "Unsupported format: '{}'. Supported: xlsx, csv, ods, pdf",
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
        "ods" => 3,
        "pdf" => 4,
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
            sheet_select(args[1], session);
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

fn sheet_select(name_or_index: &str, session: &mut CliSession) {
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
                output::success(&format!(
                    "Active sheet: [{}] {}",
                    i,
                    session.active_sheet_name.as_deref().unwrap()
                ));
            }
            None => output::error(&format!(
                "Sheet '{}' not found. Use 'worksheet list' to see available sheets.",
                name_or_index
            )),
        }
    }
}

fn sheet_add(name: &str, engine: &EngineHandle, session: &mut CliSession) {
    let rid = session.rid.clone().unwrap();
    let request = rb::build_add_sheet(&rid);
    if let Ok(_resp) = engine.process_request_json(&request) {
        refresh_sheet_list(engine, session);
        // Rename the newly created sheet
        let new_idx = session.sheet_names.len().saturating_sub(1);
        if new_idx < session.sheet_ids.len() {
            let new_id = session.sheet_ids[new_idx].clone();
            let rename_req = rb::build_rename_sheet(&rid, &new_id, name);
            let _ = engine.process_request_json(&rename_req);
            refresh_sheet_list(engine, session);
        }
        output::success(&format!(
            "Added sheet: '{}' at index [{}]",
            name,
            session.sheet_names.len().saturating_sub(1)
        ));
    } else {
        output::error("Failed to add sheet.");
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
                    output::line(&format!("  [{}] {}  ({}:{})", i, t.table_id, start, end), 0);
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
        output::error("Usage: table delete <tableId> [--keep-format]");
        return;
    }
    let table_id = args[0];
    let keep_format = args.iter().skip(1).any(|a| a.eq_ignore_ascii_case("--keep-format"));
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_delete_table(rid, table_id, keep_format);
    exec_status_cmd(engine, &request, session, &format!("Table '{}' deleted.", table_id));
}

fn table_rename(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table rename <tableId> <newName>");
        return;
    }
    let table_id = args[0];
    let new_name = args[1];
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_change_table_name(rid, &sid, table_id, new_name);
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table renamed to '{}'.", new_name),
    );
}

fn table_options(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 3 {
        output::error("Usage: table options <tableId> <settingType> <true|false>");
        output::info("  Setting types: 0=header_row, 1=total_row, 2=banded_row, 3=banded_column, 4=first_column, 5=last_column, 6=filter_button");
        return;
    }
    let table_id = args[0];
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
    let request = rb::build_change_table_options(rid, table_id, &sid, setting_type, is_enabled);
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
        output::error("Usage: table source <tableId> <range>");
        return;
    }
    let table_id = args[0];
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_change_table_source(rid, table_id, &sid, sr, sc, er, ec);
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table source changed to {}.", args[1].to_uppercase()),
    );
}

fn table_style(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table style <tableId> <stylePattern> [--keep-format]");
        output::info("  Style patterns: 0=none, 1-3=light, 4-8=medium, 9=dark");
        return;
    }
    let table_id = args[0];
    let pattern: i32 = match args[1].parse() {
        Ok(n) if (0..=9).contains(&n) => n,
        _ => {
            output::error("Style pattern must be 0-9.");
            return;
        }
    };
    let keep_format = args.iter().skip(2).any(|a| a.eq_ignore_ascii_case("--keep-format"));
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_change_table_style_pattern(rid, table_id, pattern, keep_format);
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
        output::error("Usage: table insertrow <tableId> <range> [--above]");
        return;
    }
    let table_id = args[0];
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let is_below = !args.iter().skip(2).any(|a| a.eq_ignore_ascii_case("--above"));
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_insert_table_row(rid, table_id, &sid, sr, sc, er, ec, is_below);
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
        output::error("Usage: table insertcol <tableId> <range> [--after]");
        return;
    }
    let table_id = args[0];
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let is_before = !args.iter().skip(2).any(|a| a.eq_ignore_ascii_case("--after"));
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_insert_table_column(rid, table_id, &sid, sr, sc, er, ec, is_before);
    let pos = if is_before { "before" } else { "after" };
    exec_status_cmd(
        engine,
        &request,
        session,
        &format!("Table column(s) inserted {}.", pos),
    );
}

fn table_delete_row(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table deleterow <tableId> <range>");
        return;
    }
    let table_id = args[0];
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_delete_table_row(rid, table_id, &sid, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, "Table row(s) deleted.");
}

fn table_delete_col(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.len() < 2 {
        output::error("Usage: table deletecol <tableId> <range>");
        return;
    }
    let table_id = args[0];
    let (sc, sr, ec, er) = parse_range_arg!(args[1]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let request = rb::build_delete_table_column(rid, table_id, &sid, sr, sc, er, ec);
    exec_status_cmd(engine, &request, session, "Table column(s) deleted.");
}

fn table_manage(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: table manage <tableId>");
        return;
    }
    let table_id = args[0];
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_manage_table(rid, table_id);
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
                output::key_value("Table ID", table_id, 2);
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

// ─── Helper macros & functions ───────────────────────────────────────────────

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
        output::error("Usage: format <bold|italic|underline|doubleunderline|strikethrough|superscript|subscript|fontsize|fontcolor> <range> ...");
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
        _ => {
            output::error(&format!("Unknown format sub-command: '{}'. Use: bold, italic, underline, doubleunderline, strikethrough, superscript, subscript, fontsize, fontcolor.", sub));
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

// ─── Help ────────────────────────────────────────────────────────────────────

fn print_help() {
    output::info("Available commands:");
    output::info("");
    output::info("  FILE");
    output::info("  open <filepath>              Open a local file (.xlsx, .csv, .tsv)");
    output::info("  open --new <docname>         Create a new blank workbook");
    output::info("  save                         Save to original path");
    output::info("  save --as <path>             Save a copy / export (format from extension)");
    output::info("  close                        Close the current workbook");
    output::info("");
    output::info("  WORKSHEETS");
    output::info("  worksheet list                   List all sheets in the open workbook");
    output::info("  worksheet switch <name|index>    Switch the active sheet");
    output::info("  worksheet add <name>             Add a new sheet");
    output::info("  worksheet delete <name|index>    Delete a sheet");
    output::info("  worksheet rename <name>          Rename the active sheet");
    output::info("  worksheet reorder <position>     Move active sheet to position (0-based)");
    output::info("  worksheet duplicate              Duplicate the active sheet");
    output::info("  worksheet hide [name|index]      Hide a sheet");
    output::info("  worksheet unhide <name|index>    Unhide a sheet");
    output::info("");
    output::info("  CELLS");
    output::info("  cell get <ref>               Get cell value (e.g., A1)");
    output::info("  cell set <ref> <value>       Set a cell value");
    output::info("  cell set <ref> --formula <f> Set a formula in a cell");
    output::info("");
    output::info("  ROWS & COLUMNS");
    output::info("  row insert <row> [count]     Insert rows (1-based)");
    output::info("  row delete <row> [count]     Delete rows");
    output::info("  row hide <row> [endRow]      Hide rows");
    output::info("  row unhide <row> [endRow]    Unhide rows");
    output::info("  row resize <row> <height>    Resize a row");
    output::info("  col insert <col> [count]     Insert columns (letter, e.g., B)");
    output::info("  col delete <col> [count]     Delete columns");
    output::info("  col hide <col> [endCol]      Hide columns");
    output::info("  col unhide <col> [endCol]    Unhide columns");
    output::info("  col resize <col> <width>     Resize a column");
    output::info("");
    output::info("  EDITING");
    output::info("  copy <range> <dest> [--values|--format]   Copy range to destination");
    output::info("  move <range> <dest> [--values|--format]   Move range to destination");
    output::info("  clipboardcopy <range>            Copy cell values to system clipboard");
    output::info("  merge <range>                Merge cells");
    output::info("  merge undo <range>           Unmerge cells");
    output::info("  clear <range> [--content|--format]  Clear cells");
    output::info("  undo                         Undo last action");
    output::info("  redo                         Redo last undone action");
    output::info("");
    output::info("  FIND & SORT");
    output::info("  find <text> [--case] [--exact]           Find in sheet");
    output::info("  replace <old> <new> [--all] [--case]     Find and replace");
    output::info("  sort <range> <col> [--desc] [--header]   Sort a range");
    output::info("  filter create <range>        Create auto-filter");
    output::info("  filter remove                Remove auto-filter");
    output::info("");
    output::info("  VIEW & NAMES");
    output::info("  freeze <ref>                 Freeze panes at cell");
    output::info("  unfreeze                     Unfreeze panes");
    output::info("  name add <name> <range>      Add a named range");
    output::info("  name delete <name>           Delete a named range");
    output::info("  name list                    List all defined names");
    output::info("");
    output::info("  TABLES");
    output::info("  table list                                   List all tables in active sheet");
    output::info("  table create <range> [--headers]             Create a table on range");
    output::info("  table select <range>                         Select table range");
    output::info("  table delete <tableId> [--keep-format]       Delete a table");
    output::info("  table rename <tableId> <name>                Rename a table");
    output::info("  table options <tableId> <type> <true|false>  Toggle table option");
    output::info("         types: 0=Header Row    1=Total Row      2=Banded Rows");
    output::info("                3=Banded Columns 4=First Column   5=Last Column");
    output::info("                6=Filter Button");
    output::info("  table source <tableId> <range>               Change table source range");
    output::info("  table style <tableId> <pattern>              Change table style (0-9)");
    output::info("         patterns: 0=Light1  1=Light2  2=Light3  3=Light4  4=Light5");
    output::info("                   5=Medium1 6=Medium2 7=Medium3 8=Dark1   9=Dark2");
    output::info("  table defaultstyle <pattern>                 Set default table style");
    output::info("  table insertrow <tableId> <range> [--above]  Insert table row(s)");
    output::info("  table insertcol <tableId> <range> [--after]  Insert table column(s)");
    output::info("  table deleterow <tableId> <range>            Delete table row(s)");
    output::info("  table deletecol <tableId> <range>            Delete table column(s)");
    output::info("  table manage <tableId>                       Get table info");
    output::info("");
    output::info("");
    output::info("  FORMATTING");
    output::info("  format bold <range> <true|false>              Toggle bold");
    output::info("  format italic <range> <true|false>            Toggle italic");
    output::info("  format underline <range> <true|false>         Toggle underline");
    output::info("  format doubleunderline <range> <true|false>   Toggle double underline");
    output::info("  format strikethrough <range> <true|false>     Toggle strikethrough");
    output::info("  format superscript <range> <true|false>       Toggle superscript");
    output::info("  format subscript <range> <true|false>         Toggle subscript");
    output::info("  format fontsize <range> <size>                Set font size");
    output::info("  format fontcolor <range> <r> <g> <b>          Set font color (RGB 0-255)");
    output::info("  format fontcolor <range> --auto               Set automatic font color");
    output::info("");
    output::info("  help                         Show this help");
    output::info("  exit / quit                  Exit the CLI");
}
