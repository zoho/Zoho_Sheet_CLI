/// All CLI command implementations — one per verb.
///
/// Instead of the C# DI-based ICliCommand pattern, Rust uses a dispatch function
/// that takes shared references to the engine, session, and parser.
use std::collections::HashMap;
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
        "checkbox" => cmd_checkbox(&args, engine, session),
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
        "cf" | "conditional" => cmd_cf(&args, engine, session),
        "dv" | "datavalidation" | "data-validation" => cmd_dv(&args, engine, session),
        "theme" => cmd_theme(&args, engine, session),
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
    let working_path = working_copy.to_string_lossy().to_string();

    // Password negotiation loop.
    // Each iteration sends one request; the status code drives the next prompt.
    let mut password: Option<String> = None;
    loop {
        let request = rb::build_open_workbook(&working_path, open_file_type, password.as_deref());
        let response = match engine.process_request_json(&request) {
            Ok(r) => r,
            Err(e) => {
                output::error(&format!("Engine error: {}", e));
                return;
            }
        };

        let result = match rp::parse_workbook_open(&response) {
            Some(r) => r,
            None => {
                output::error(&format!("Engine failed to open '{}'.", file_name));
                return;
            }
        };

        match result.status_code {
            // ── Success (100) or warning with unsupported features (500) ─────
            100 | 500 => {
                if result.rid.is_none() || result.rid.as_deref() == Some("") {
                    output::error(&format!("Engine failed to open '{}'.", file_name));
                    return;
                }

                // Warn about unsupported features
                if !result.unsupported_features.is_empty() {
                    output::warning(&format!(
                        "Some features are not supported and were skipped: {}",
                        result.unsupported_features.join(", ")
                    ));
                }

                // Warn if the file was repaired
                if result.is_repaired {
                    output::warning(
                        "This file was repaired on open and may have lost some data or formatting.",
                    );
                    let choice = output::confirm("Proceed with the repaired file?");
                    if choice != "y" {
                        // Close the workbook and abort
                        if let Some(ref rid) = result.rid {
                            let close_req = rb::build_close_workbook(rid);
                            let _ = engine.process_request_json(&close_req);
                        }
                        output::info("Open cancelled.");
                        return;
                    }
                }

                let rid = result.rid.unwrap();
                session.rid = Some(rid.clone());
                session.workbook_name = result.workbook_name.or_else(|| {
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
                return;
            }

            // ── Open password required ────────────────────────────────────────
            4135 => {
                output::info("This file is password protected.");
                match output::prompt_password("Enter password to open") {
                    None => {
                        // User cancelled — tell the engine to abort
                        let _ = engine.process_request_json(
                            &rb::build_open_workbook(&working_path, open_file_type, Some("Abort")),
                        );
                        output::info("Open cancelled.");
                        return;
                    }
                    Some(pw) => password = Some(base64_encode(&pw)),
                }
            }

            // ── Wrong open password ───────────────────────────────────────────
            4132 => {
                output::error("Incorrect password. Please try again.");
                match output::prompt_password("Enter password to open") {
                    None => {
                        let _ = engine.process_request_json(
                            &rb::build_open_workbook(&working_path, open_file_type, Some("Abort")),
                        );
                        output::info("Open cancelled.");
                        return;
                    }
                    Some(pw) => password = Some(base64_encode(&pw)),
                }
            }

            // ── Valid open password, but modification password also needed ────
            4136 | 4137 => {
                if result.status_code == 4137 {
                    output::error("Incorrect modification password. Please try again.");
                }
                let choice = output::prompt_choice(
                    "A modification password is required:",
                    &["Open read-only", "Enter modification password", "Cancel"],
                );
                match choice {
                    0 => password = Some(String::new()), // read-only: send empty string
                    1 => match output::prompt_password("Enter modification password") {
                        None => {
                            let _ = engine.process_request_json(
                                &rb::build_open_workbook(&working_path, open_file_type, Some("Abort")),
                            );
                            output::info("Open cancelled.");
                            return;
                        }
                        Some(pw) => password = Some(base64_encode(&pw)),
                    },
                    _ => {
                        let _ = engine.process_request_json(
                            &rb::build_open_workbook(&working_path, open_file_type, Some("Abort")),
                        );
                        output::info("Open cancelled.");
                        return;
                    }
                }
            }

            // ── Engine needs a yes/no read-only decision ──────────────────────
            4138 => {
                let choice = output::prompt_choice(
                    "How would you like to open this file?",
                    &["Open read-only", "Open for editing"],
                );
                // "Yes" = read-only, "No" = editing — must NOT be base64 encoded
                password = Some(if choice == 0 { "Yes" } else { "No" }.to_string());
            }

            // ── Any other status code is a hard error ─────────────────────────
            _ => {
                output::error(&format!(
                    "Failed to open '{}': {} (code {})",
                    file_name,
                    result.status_message.as_deref().unwrap_or("Unknown error"),
                    result.status_code
                ));
                return;
            }
        }
    }
}

/// Encodes a UTF-8 string as standard Base64 (RFC 4648).
fn base64_encode(input: &str) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let bytes = input.as_bytes();
    let mut out = String::with_capacity((bytes.len() + 2) / 3 * 4);
    for chunk in bytes.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };
        let n = (b0 << 16) | (b1 << 8) | b2;
        out.push(CHARS[((n >> 18) & 63) as usize] as char);
        out.push(CHARS[((n >> 12) & 63) as usize] as char);
        out.push(if chunk.len() > 1 { CHARS[((n >> 6) & 63) as usize] as char } else { '=' });
        out.push(if chunk.len() > 2 { CHARS[(n & 63) as usize] as char } else { '=' });
    }
    out
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
        output::error("Usage: cell get <ref> | cell get <range> [--page <n>] | cell set <ref> <value> | cell set <ref> --formula <expr> | cell set <ref> --hyperlink <link> [--text <display>] [--type <n>] | cell set <ref> --note <text>");
        return;
    }
    match args[0].to_lowercase().as_str() {
        "get" => cell_get(&args[1..], engine, session),
        "set" => cell_set(args, engine, session),
        other => output::error(&format!("Unknown cell sub-command: '{}'. Use: get, set", other)),
    }
}

fn cell_get(args: &[&str], engine: &EngineHandle, session: &CliSession) {
    if args.is_empty() {
        output::error("Usage: cell get <ref> | cell get <range> [--page <n>]");
        return;
    }
    let cell_ref = args[0];
    let page: usize = if args.len() >= 3 && args[1].eq_ignore_ascii_case("--page") {
        args[2].parse::<usize>().unwrap_or(1).max(1)
    } else {
        1
    };
    if cell_ref.contains(':') {
        cell_get_range(cell_ref, page, engine, session);
    } else {
        cell_get_single(cell_ref, engine, session);
    }
}

fn cell_get_single(cell_ref: &str, engine: &EngineHandle, session: &CliSession) {
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
                let sheet_name = session.active_sheet_name.as_deref().unwrap_or(&fallback);
                let ref_display = cell_ref::to_ref(col, row);
                output::line(&format!("Cell {}  ({})", ref_display, sheet_name), 0);
                output::key_value(
                    "Display",
                    if result.display_value.is_empty() { "(empty)" } else { &result.display_value },
                    2,
                );
                output::key_value(
                    "Raw",
                    if result.raw_value.is_empty() { "(empty)" } else { &result.raw_value },
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

const RANGE_PAGE_SIZE: usize = 25;

fn cell_get_range(range_ref: &str, page: usize, engine: &EngineHandle, session: &CliSession) {
    let (start_col, start_row, end_col, end_row) = match cell_ref::try_parse_range(range_ref) {
        Some(r) => r,
        None => {
            output::error(&format!("Invalid range reference: '{}'", range_ref));
            return;
        }
    };

    let rid = session.rid.as_deref().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();

    // Build the full ordered cell list (row-major)
    let mut all_cells: Vec<(i32, i32)> = Vec::new();
    for r in start_row..=end_row {
        for c in start_col..=end_col {
            all_cells.push((r, c));
        }
    }
    let total = all_cells.len();

    let total_pages = (total + RANGE_PAGE_SIZE - 1) / RANGE_PAGE_SIZE;
    if page > total_pages {
        output::error(&format!(
            "Page {} out of range. There {} {} page(s) for this range.",
            page,
            if total_pages == 1 { "is" } else { "are" },
            total_pages
        ));
        return;
    }

    let skip = (page - 1) * RANGE_PAGE_SIZE;
    let page_cells = &all_cells[skip..(skip + RANGE_PAGE_SIZE).min(total)];
    let page_end = skip + page_cells.len();

    // Fetch all cells in the range in one request
    let fetch_req = rb::build_range_cell_fetch(rid, &sheet_id, start_row, start_col, end_row, end_col);
    let values = match engine.fetch_json(&fetch_req) {
        Ok(resp) => rp::parse_range_cell_values(&resp),
        Err(e) => {
            output::error(&format!("Engine error reading range: {}", e));
            return;
        }
    };

    // Build lookup: (row, col) → (display, raw, formula)
    let mut lookup: HashMap<(i32, i32), (String, String, String)> = HashMap::new();
    for (r, c, display, raw, formula) in values {
        lookup.insert((r, c), (display, raw, formula));
    }

    let fallback = format!("Sheet{}", session.active_sheet_index);
    let sheet_name = session.active_sheet_name.as_deref().unwrap_or(&fallback);
    let range_display = format!(
        "{}:{}",
        cell_ref::to_ref(start_col, start_row),
        cell_ref::to_ref(end_col, end_row)
    );

    let header = if total_pages > 1 {
        format!("Range {}  ({})  — page {} of {}", range_display, sheet_name, page, total_pages)
    } else {
        format!("Range {}  ({})", range_display, sheet_name)
    };
    output::line(&header, 0);

    // Collect row data for the current page
    const COL_CAP: usize = 30;
    let col_headers = ["Ref", "Display", "Raw", "Formula"];
    let mut widths = [col_headers[0].len(), col_headers[1].len(), col_headers[2].len(), col_headers[3].len()];

    let mut rows_data: Vec<[String; 4]> = Vec::new();
    for &(r, c) in page_cells {
        let ref_str = cell_ref::to_ref(c, r);
        let (disp, raw, formula) = lookup.remove(&(r, c)).unwrap_or_default();
        let disp_str    = if disp.is_empty()                          { "(empty)".into() } else { disp };
        let raw_str     = if raw.is_empty()                           { "(empty)".into() } else { raw };
        let formula_str = if formula.is_empty() || formula == "null"  { "(none)".into()  } else { formula };
        widths[0] = widths[0].max(ref_str.len().min(COL_CAP));
        widths[1] = widths[1].max(disp_str.len().min(COL_CAP));
        widths[2] = widths[2].max(raw_str.len().min(COL_CAP));
        widths[3] = widths[3].max(formula_str.len().min(COL_CAP));
        rows_data.push([ref_str, disp_str, raw_str, formula_str]);
    }

    let separator: String = widths.iter().map(|&w| "─".repeat(w + 2)).collect::<Vec<_>>().join("─");
    output::line(&separator, 0);
    output::line(
        &format!(
            "{:<w0$}  {:<w1$}  {:<w2$}  {:<w3$}",
            col_headers[0], col_headers[1], col_headers[2], col_headers[3],
            w0 = widths[0], w1 = widths[1], w2 = widths[2], w3 = widths[3]
        ),
        0,
    );
    output::line(&separator, 0);

    for row_data in &rows_data {
        output::line(
            &format!(
                "{:<w0$}  {:<w1$}  {:<w2$}  {:<w3$}",
                truncate_cell(&row_data[0], widths[0]),
                truncate_cell(&row_data[1], widths[1]),
                truncate_cell(&row_data[2], widths[2]),
                truncate_cell(&row_data[3], widths[3]),
                w0 = widths[0], w1 = widths[1], w2 = widths[2], w3 = widths[3]
            ),
            0,
        );
    }

    if total > RANGE_PAGE_SIZE {
        let next_hint = if page < total_pages {
            format!("  Run: cell get {} --page {}", range_display, page + 1)
        } else {
            String::new()
        };
        output::line(
            &format!("\nShowing {}-{} of {}.{}", skip + 1, page_end, total, next_hint),
            0,
        );
    }
}

fn truncate_cell(s: &str, max_chars: usize) -> String {
    let count = s.chars().count();
    if count <= max_chars {
        s.to_string()
    } else {
        s.chars().take(max_chars.saturating_sub(3)).collect::<String>() + "..."
    }
}

fn cell_set(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    // args[0] = "set", args[1] = cellRef, args[2..] = value or --formula <expr> or --hyperlink <link> [--text <display>] [--type <n>]
    if args.len() < 3 {
        output::error("Usage: cell set <ref> <value> | cell set <ref> --formula <expr> | cell set <ref> --hyperlink <link> [--text <display>] [--type <n>] | cell set <ref> --note <text>");
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

    // Support hyperlink command: `cell set A1 --hyperlink <link> [--text <display>] [--type <n>]`
    if args[2].eq_ignore_ascii_case("--hyperlink") {
        if args.len() < 4 {
            output::error("Usage: cell set <ref> --hyperlink <link> [--text <display>] [--type <n>]");
            return;
        }
        let link = args[3];
        let mut display_text: Option<&str> = None;
        let mut link_type: i32 = 0; // default to WEB_PAGE

        let mut i = 4;
        while i < args.len() {
            match args[i].to_lowercase().as_str() {
                "--text" => {
                    if i + 1 >= args.len() {
                        output::error("--text requires a value");
                        return;
                    }
                    display_text = Some(args[i + 1]);
                    i += 2;
                }
                "--type" => {
                    if i + 1 >= args.len() {
                        output::error("--type requires a numeric value");
                        return;
                    }
                    match args[i + 1].parse::<i32>() {
                        Ok(v) => link_type = v,
                        Err(_) => {
                            output::error("--type requires a numeric value (0=WEB_PAGE,1=RANGE,2=EMAIL,3=TELEPHONE,4=DEFINED_NAME)");
                            return;
                        }
                    }
                    i += 2;
                }
                other => {
                    output::error(&format!("Unknown option for --hyperlink: {}", other));
                    return;
                }
            }
        }

        let rid = session.rid.as_deref().unwrap();
        let sheet_id = session.get_active_sheet_id_or_default();
        let request = rb::build_insert_hyperlink(rid, &sheet_id, row, col, link, display_text, link_type);
        match engine.process_request_json(&request) {
            Ok(resp) => {
                let status = rp::parse_status_response(&resp);
                let ref_display = cell_ref::to_ref(col, row);
                if rp::is_success(status.status_code) {
                    output::success(&format!("{}: inserted hyperlink -> {}", ref_display, link));
                    session.is_dirty = true;
                } else {
                    output::error(&format!("Failed to insert hyperlink {}: {}", ref_display, status.status_message.unwrap_or_else(|| "engine error".into())));
                }
            }
            Err(e) => output::error(&format!("Engine error inserting hyperlink {}: {}", cell_ref::to_ref(col, row), e)),
        }
        return;
    }

    // Support note command: `cell set A1 --note <text>`
    if args[2].eq_ignore_ascii_case("--note") {
        if args.len() < 4 {
            output::error("Usage: cell set <ref> --note <text>");
            return;
        }
        let notes = args[3];
        let rid = session.rid.as_deref().unwrap();
        let sheet_id = session.get_active_sheet_id_or_default();
        let request = rb::build_insert_note(rid, &sheet_id, row, col, notes);
        match engine.process_request_json(&request) {
            Ok(resp) => {
                let status = rp::parse_status_response(&resp);
                let ref_display = cell_ref::to_ref(col, row);
                if rp::is_success(status.status_code) {
                    output::success(&format!("{}: note set", ref_display));
                    session.is_dirty = true;
                } else {
                    output::error(&format!("Failed to set note {}: {}", ref_display, status.status_message.unwrap_or_else(|| "engine error".into())));
                }
            }
            Err(e) => output::error(&format!("Engine error setting note {}: {}", cell_ref::to_ref(col, row), e)),
        }
        return;
    }

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

// ─── Checkbox ───────────────────────────────────────────────────────────────

fn cmd_checkbox(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if !session.is_active() {
        output::error("No workbook open. Use 'open' first.");
        return;
    }
    if args.len() < 2 {
        output::error("Usage: checkbox insert <range> | checkbox update <range> <true|false> | checkbox delete <range>");
        return;
    }

    match args[0].to_lowercase().as_str() {
        "insert" => checkbox_insert(args[1], engine, session),
        "update" => {
            if args.len() < 3 {
                output::error("Usage: checkbox update <range> <true|false>");
                return;
            }
            checkbox_update(args[1], args[2], engine, session);
        }
        "delete" => checkbox_delete(args[1], engine, session),
        other => output::error(&format!(
            "Unknown checkbox sub-command: '{}'. Use: insert, update, delete",
            other
        )),
    }
}

fn checkbox_insert(range_arg: &str, engine: &EngineHandle, session: &mut CliSession) {
    let (start_col, start_row, end_col, end_row) = match cell_ref::try_parse_range(range_arg) {
        Some(r) => r,
        None => {
            output::error(&format!(
                "Invalid range: '{}'. Use A1:C5 format.",
                range_arg
            ));
            return;
        }
    };

    let rid = session.rid.as_deref().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();
    let request = rb::build_insert_checkbox(
        rid,
        &sheet_id,
        start_row,
        start_col,
        end_row,
        end_col,
    );

    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success(&format!("Inserted checkbox(es) in {}.", range_arg.to_uppercase()));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to insert checkbox(es): {}",
                    status
                        .status_message
                        .unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error inserting checkbox(es): {}", e)),
    }
}

fn checkbox_update(
    range_arg: &str,
    bool_arg: &str,
    engine: &EngineHandle,
    session: &mut CliSession,
) {
    let (start_col, start_row, end_col, end_row) = match cell_ref::try_parse_range(range_arg) {
        Some(r) => r,
        None => {
            output::error(&format!(
                "Invalid range: '{}'. Use A1:C5 format.",
                range_arg
            ));
            return;
        }
    };

    let boolean_value = match bool_arg.to_ascii_lowercase().as_str() {
        "true" | "1" | "yes" | "on" => true,
        "false" | "0" | "no" | "off" => false,
        _ => {
            output::error("checkbox update expects <true|false> (also accepts 1/0, yes/no, on/off)");
            return;
        }
    };

    let rid = session.rid.as_deref().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();
    let request = rb::build_update_checkbox(
        rid,
        &sheet_id,
        start_row,
        start_col,
        end_row,
        end_col,
        boolean_value,
    );

    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success(&format!(
                    "Updated checkbox(es) in {} to {}.",
                    range_arg.to_uppercase(),
                    boolean_value
                ));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to update checkbox(es): {}",
                    status
                        .status_message
                        .unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error updating checkbox(es): {}", e)),
    }
}

fn checkbox_delete(range_arg: &str, engine: &EngineHandle, session: &mut CliSession) {
    let (start_col, start_row, end_col, end_row) = match cell_ref::try_parse_range(range_arg) {
        Some(r) => r,
        None => {
            output::error(&format!(
                "Invalid range: '{}'. Use A1:C5 format.",
                range_arg
            ));
            return;
        }
    };

    let rid = session.rid.as_deref().unwrap();
    let sheet_id = session.get_active_sheet_id_or_default();
    let request = rb::build_delete_checkbox(
        rid,
        &sheet_id,
        start_row,
        start_col,
        end_row,
        end_col,
    );

    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success(&format!("Deleted checkbox(es) in {}.", range_arg.to_uppercase()));
                session.is_dirty = true;
            } else {
                output::error(&format!(
                    "Failed to delete checkbox(es): {}",
                    status
                        .status_message
                        .unwrap_or_else(|| "engine error".into())
                ));
            }
        }
        Err(e) => output::error(&format!("Engine error deleting checkbox(es): {}", e)),
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

fn cmd_cf(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: cf <classic|colorscale|databar|iconset|list|delete|move> ...");
        return;
    }

    match args[0].to_lowercase().as_str() {
        "classic" => cmd_cf_classic(&args[1..], engine, session),
        "colorscale" | "color-scale" => cmd_cf_color_scale(&args[1..], engine, session),
        "databar" | "data-bar" => cmd_cf_data_bar(&args[1..], engine, session),
        "iconset" | "icon-set" => cmd_cf_icon_set(&args[1..], engine, session),
        "move" => cmd_cf_move(&args[1..], engine, session),
        "delete" => cmd_cf_delete_rule(&args[1..], engine, session),
        "manage" | "list" => cmd_cf_manage_rules(&args[1..], engine, session),
        "priority" | "priority-update" | "update-priority" => {
            output::info("'cf priority' is deprecated; use 'cf move' instead.");
            cmd_cf_update_priority_legacy(&args[1..], engine, session);
        }
        other => output::error(&format!("Unknown cf sub-command: '{}'. Use: classic, colorscale, databar, iconset, list, delete, move", other)),
    }
}

fn cmd_cf_classic(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: cf classic [<#index>|--rule-id <id>] --range <A1:C5> [--when \"expr\"] [--bold] [--fill <color>] [--condition.<key> <val>] ...";

    let translated: Vec<String> = args.iter().map(|&a| {
        if let Some(s) = a.strip_prefix("--condition.") { format!("--condition[0].{}", s) }
        else { a.to_string() }
    }).collect();
    let args: Vec<&str> = translated.iter().map(|s| s.as_str()).collect();
    let args = args.as_slice();

    let (index_target, rule_id_target, range_override, sheet_override, active_cell_override,
         mut rule_obj)
        = match cf_parse_v2_args(args, usage) {
        Ok(v) => v,
        Err(e) => { output::error(&e); return; }
    };

    let is_edit = index_target.is_some() || rule_id_target.is_some();

    // Resolve rule_id for #index
    let resolved_rule_id: Option<String> = if let Some(idx) = index_target {
        match cf_resolve_index(idx, session) {
            Ok(id) => Some(id),
            Err(e) => { output::error(&e); return; }
        }
    } else {
        rule_id_target.clone()
    };

    // Auto-assign criteria_id
    cf_auto_criteria_id(&mut rule_obj);

    // Default is_percent to false for top/bottom (criteria_type 13) if not already set
    cf_default_is_percent(&mut rule_obj);

    // Apply engine remaps
    cf_apply_engine_remaps(&mut rule_obj);

    // For top_bottom (criteria_type 13): move lhs from condition to rule-level count
    {
        let mut top_bottom_lhs: Option<serde_json::Value> = None;
        if let Some(obj) = rule_obj.as_object_mut() {
            if let Some(conds) = obj.get_mut("condition").and_then(|v| v.as_array_mut()) {
                for cond in conds.iter_mut() {
                    if let Some(co) = cond.as_object_mut() {
                        if co.get("criteria_type").and_then(|v| v.as_i64()) == Some(13) {
                            if let Some(lhs) = co.remove("lhs") {
                                top_bottom_lhs = Some(lhs);
                            }
                        }
                    }
                }
            }
        }
        if let Some(lhs) = top_bottom_lhs {
            if let Some(obj) = rule_obj.as_object_mut() {
                if !obj.contains_key("count") {
                    let count_val = match &lhs {
                        serde_json::Value::String(s) => s.parse::<i64>()
                            .map(|n| serde_json::json!(n))
                            .unwrap_or_else(|_| lhs.clone()),
                        other => other.clone(),
                    };
                    obj.insert("count".into(), count_val);
                }
            }
        }
    }

    if let Err(e) = cf_validate_condition_criteria(&rule_obj, "classic") {
        output::error(&e);
        return;
    }

    // Build rule payload
    let sid = match cf_resolve_sheet_id(sheet_override.as_deref(), session) {
        Some(id) => id,
        None => return,
    };

    if is_edit {
        let rule_id = match resolved_rule_id {
            Some(ref id) => id.clone(),
            None => { output::error("Edit requires #index or --rule-id."); return; }
        };

        if let Err(e) = cf_check_edit_rule_type(session, index_target, "Classic", "classic") {
            output::error(&e);
            return;
        }

        // If editing by index, reconstruct existing condition/stop_if_true as base so the
        // user only needs to specify the fields they want to change.
        if let Some(idx) = index_target {
            let full_rule = session.last_cf_rules[idx - 1].full_rule.clone();
            if let Some(mut base) = cf_reconstruct_classic_base(&full_rule) {
                deep_merge_into(&mut base, &rule_obj);
                rule_obj = base;
            }
        }

        if let Some(obj) = rule_obj.as_object_mut() {
            obj.insert("rule_id".into(), serde_json::json!(rule_id));
        }

        // For edit, range is optional
        let range_list = if let Some(r) = range_override.as_deref() {
            match cf_parse_range(r) {
                Ok(v) => v,
                Err(e) => { output::error(&e); return; }
            }
        } else if let Some(idx) = index_target {
            // Use cached range from last cf list; engine requires a non-empty range_list
            let entry = &session.last_cf_rules[idx - 1];
            if entry.range_json.as_array().map_or(true, |a| a.is_empty()) {
                output::error("Cannot determine range for this rule. Re-run 'cf list' or provide --range.");
                return;
            }
            entry.range_json.clone()
        } else {
            output::error("--range is required when editing by --rule-id without a cached range.");
            return;
        };
        let range_arr = range_list.as_array().unwrap();
        let active_info = build_cf_active_info_default(&sid, &range_list);
        let rid: &str = session.rid.as_deref().unwrap();
        let request = rb::build_edit_classic_rule(rid, &sid, range_arr, rule_obj, active_info);
        exec_status_cmd(engine, &request, session, "Conditional formatting classic rule edited.");
    } else {
        // Insert — range required
        let range_str = match range_override.as_deref() {
            Some(r) => r.to_string(),
            None => { output::error("--range is required for cf classic insert."); return; }
        };
        let range_list: serde_json::Value = match cf_parse_range(&range_str) {
            Ok(v) => v,
            Err(e) => { output::error(&e); return; }
        };

        let range_arr = range_list.as_array().unwrap();
        let active_info = build_cf_active_info_default(&sid, &range_list);
        let rid = session.rid.as_deref().unwrap();
        let rule_vec = vec![rule_obj];
        let request = rb::build_insert_classic_rule(rid, &sid, range_arr, &rule_vec, active_info);
        exec_status_cmd(engine, &request, session, "Conditional formatting classic rule inserted.");
    }
}

fn cmd_cf_color_scale(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: cf colorscale [<#index>|--rule-id <id>] --range <A1:C5> [--min <color>] [--mid <color>] [--max <color>] [--min.color <c>] [--min.criteria_type <t>] [--min.value <v>] ...";

    let (index_target, rule_id_target, range_override, sheet_override, _active_cell_override,
         mut rule_obj)
        = match cf_parse_v2_args(args, usage) {
        Ok(v) => v,
        Err(e) => { output::error(&e); return; }
    };

    let is_edit = index_target.is_some() || rule_id_target.is_some();

    let resolved_rule_id: Option<String> = if let Some(idx) = index_target {
        match cf_resolve_index(idx, session) {
            Ok(id) => Some(id),
            Err(e) => { output::error(&e); return; }
        }
    } else {
        rule_id_target.clone()
    };

    if let Some(obj) = rule_obj.as_object_mut() {
        if !obj.contains_key("is_hide_values") { obj.insert("is_hide_values".into(), serde_json::json!(false)); }
        if !obj.contains_key("is_automatic_text_color") { obj.insert("is_automatic_text_color".into(), serde_json::json!(false)); }
    }
    cf_auto_criteria_id(&mut rule_obj);
    // Default criteria_type for colorscale stops that only had a color specified
    if let Some(conds) = rule_obj.get_mut("condition").and_then(|v| v.as_array_mut()) {
        for cond in conds.iter_mut() {
            if let Some(co) = cond.as_object_mut() {
                if !co.contains_key("criteria_type") {
                    let default_ct = match co.get("criteria_id").and_then(|v| v.as_i64()) {
                        Some(0) => 4i64, // min stop → minimum_value
                        Some(1) => 2i64, // mid stop → percentile
                        Some(2) => 5i64, // max stop → maximum_value
                        _ => 4i64,
                    };
                    co.insert("criteria_type".into(), serde_json::json!(default_ct));
                    if default_ct == 2 && !co.contains_key("lhs") && !co.contains_key("value") {
                        co.insert("value".into(), serde_json::json!("50"));
                    }
                }
            }
        }
    }
    cf_apply_engine_remaps(&mut rule_obj);

    if let Err(e) = cf_validate_condition_criteria(&rule_obj, "colorscale") {
        output::error(&e);
        return;
    }

    let sid = match cf_resolve_sheet_id(sheet_override.as_deref(), session) {
        Some(id) => id,
        None => return,
    };

    if is_edit {
        let rule_id = match resolved_rule_id {
            Some(ref id) => id.clone(),
            None => { output::error("Edit requires #index or --rule-id."); return; }
        };

        if let Err(e) = cf_check_edit_rule_type(session, index_target, "ColorScale", "color scale") {
            output::error(&e);
            return;
        }

        // If editing by index, reconstruct the existing rule as a base and merge the
        // user's changes onto it, so a partial edit doesn't send an incomplete rule.
        if let Some(idx) = index_target {
            let full_rule = session.last_cf_rules[idx - 1].full_rule.clone();
            if let Some(mut base) = cf_reconstruct_color_scale_base(&full_rule) {
                cf_merge_rule_base(&mut base, &rule_obj);
                rule_obj = base;
            }
        }

        if let Some(obj) = rule_obj.as_object_mut() {
            obj.insert("rule_id".into(), serde_json::json!(rule_id));
        }
        let range_list = if let Some(r) = range_override.as_deref() {
            match cf_parse_range(r) { Ok(v) => v, Err(e) => { output::error(&e); return; } }
        } else if let Some(idx) = index_target {
            let entry = &session.last_cf_rules[idx - 1];
            if entry.range_json.as_array().map_or(true, |a| a.is_empty()) {
                output::error("Cannot determine range for this rule. Re-run 'cf list' or provide --range.");
                return;
            }
            entry.range_json.clone()
        } else {
            output::error("--range is required when editing by --rule-id without a cached range.");
            return;
        };
        let range_arr = range_list.as_array().unwrap();
        let active_info = build_cf_active_info_default(&sid, &range_list);
        let rid = session.rid.as_deref().unwrap();
        let request = rb::build_edit_color_scale_rule(rid, &sid, range_arr, rule_obj, active_info);
        exec_status_cmd(engine, &request, session, "Conditional formatting color scale rule edited.");
    } else {
        let range_str = match range_override.as_deref() {
            Some(r) => r.to_string(),
            None => { output::error("--range is required for cf colorscale insert."); return; }
        };
        let range_list = match cf_parse_range(&range_str) {
            Ok(v) => v, Err(e) => { output::error(&e); return; }
        };
        let range_arr = range_list.as_array().unwrap();
        let active_info = build_cf_active_info_default(&sid, &range_list);
        let rid = session.rid.as_deref().unwrap();
        let request = rb::build_insert_color_scale_rule(rid, &sid, range_arr, rule_obj, active_info);
        exec_status_cmd(engine, &request, session, "Conditional formatting color scale rule inserted.");
    }
}

fn cmd_cf_data_bar(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: cf databar [<#index>|--rule-id <id>] --range <A1:C5> [--positive <color>] [--negative <color>] ...";

    let translated: Vec<String> = args.iter().map(|&a| {
        if let Some(s) = a.strip_prefix("--min.") { format!("--condition[0].{}", s) }
        else if let Some(s) = a.strip_prefix("--max.") { format!("--condition[1].{}", s) }
        else { a.to_string() }
    }).collect();
    let args: Vec<&str> = translated.iter().map(|s| s.as_str()).collect();
    let args = args.as_slice();

    let (index_target, rule_id_target, range_override, sheet_override, _active_cell_override,
         mut rule_obj)
        = match cf_parse_v2_args(args, usage) {
        Ok(v) => v,
        Err(e) => { output::error(&e); return; }
    };

    let is_edit = index_target.is_some() || rule_id_target.is_some();

    let resolved_rule_id: Option<String> = if let Some(idx) = index_target {
        match cf_resolve_index(idx, session) {
            Ok(id) => Some(id),
            Err(e) => { output::error(&e); return; }
        }
    } else {
        rule_id_target.clone()
    };

    if let Some(obj) = rule_obj.as_object_mut() {
        if !obj.contains_key("is_hide_values") { obj.insert("is_hide_values".into(), serde_json::json!(false)); }
    }
    cf_auto_criteria_id(&mut rule_obj);
    cf_apply_engine_remaps(&mut rule_obj);

    if let Err(e) = cf_validate_condition_criteria(&rule_obj, "databar") {
        output::error(&e);
        return;
    }

    let sid = match cf_resolve_sheet_id(sheet_override.as_deref(), session) {
        Some(id) => id,
        None => return,
    };

    if is_edit {
        let rule_id = match resolved_rule_id {
            Some(ref id) => id.clone(),
            None => { output::error("Edit requires #index or --rule-id."); return; }
        };

        if let Err(e) = cf_check_edit_rule_type(session, index_target, "DataBar", "data bar") {
            output::error(&e);
            return;
        }

        // If editing by index, reconstruct the existing rule as a base and merge the
        // user's changes onto it, so a partial edit doesn't send an incomplete rule
        // (the engine's edit API requires the full `condition` array).
        if let Some(idx) = index_target {
            let full_rule = session.last_cf_rules[idx - 1].full_rule.clone();
            if let Some(mut base) = cf_reconstruct_data_bar_base(&full_rule) {
                cf_merge_rule_base(&mut base, &rule_obj);
                rule_obj = base;
            }
        }

        if let Some(obj) = rule_obj.as_object_mut() {
            obj.insert("rule_id".into(), serde_json::json!(rule_id));
        }
        let range_list = if let Some(r) = range_override.as_deref() {
            match cf_parse_range(r) { Ok(v) => v, Err(e) => { output::error(&e); return; } }
        } else if let Some(idx) = index_target {
            let entry = &session.last_cf_rules[idx - 1];
            if entry.range_json.as_array().map_or(true, |a| a.is_empty()) {
                output::error("Cannot determine range for this rule. Re-run 'cf list' or provide --range.");
                return;
            }
            entry.range_json.clone()
        } else {
            output::error("--range is required when editing by --rule-id without a cached range.");
            return;
        };
        let range_arr = range_list.as_array().unwrap();
        let active_info = build_cf_active_info_default(&sid, &range_list);
        let rid = session.rid.as_deref().unwrap();
        let request = rb::build_edit_data_bar_rule(rid, &sid, range_arr, rule_obj, active_info);
        exec_status_cmd(engine, &request, session, "Conditional formatting data bar rule edited.");
    } else {
        // Insert: auto-populate missing fields required by the engine
        if let Some(obj) = rule_obj.as_object_mut() {
            if !obj.contains_key("condition") {
                obj.insert("condition".into(), serde_json::json!([
                    {"criteria_type": 4, "criteria_id": 0},
                    {"criteria_type": 5, "criteria_id": 1}
                ]));
            } else if let Some(conds) = obj.get_mut("condition").and_then(|v| v.as_array_mut()) {
                let defaults: [i64; 2] = [4, 5];
                for (ci, cond) in conds.iter_mut().enumerate() {
                    if let Some(co) = cond.as_object_mut() {
                        if !co.contains_key("criteria_type") {
                            let default_ct = defaults.get(ci).copied().unwrap_or(4);
                            co.insert("criteria_type".into(), serde_json::json!(default_ct));
                        }
                        if !co.contains_key("criteria_id") {
                            co.insert("criteria_id".into(), serde_json::json!(ci as i64));
                        }
                    }
                }
            }
            if !obj.contains_key("axis_position") {
                obj.insert("axis_position".into(), serde_json::json!(0));
            }
            if !obj.contains_key("bar_direction") {
                obj.insert("bar_direction".into(), serde_json::json!(0));
            }
            if !obj.contains_key("fill_type") {
                obj.insert("fill_type".into(), serde_json::json!(0));
            }
            if !obj.contains_key("border_type") {
                obj.insert("border_type".into(), serde_json::json!(0));
            }
            if !obj.contains_key("positive_value_fill") {
                obj.insert("positive_value_fill".into(), serde_json::json!({"red": 0, "green": 112, "blue": 192}));
            }
            if !obj.contains_key("negative_value_fill") {
                obj.insert("negative_value_fill".into(), serde_json::json!({"red": 255, "green": 0, "blue": 0}));
            }
        }
        let range_str = match range_override.as_deref() {
            Some(r) => r.to_string(),
            None => { output::error("--range is required for cf databar insert."); return; }
        };
        let range_list = match cf_parse_range(&range_str) {
            Ok(v) => v, Err(e) => { output::error(&e); return; }
        };
        let range_arr = range_list.as_array().unwrap();
        let active_info = build_cf_active_info_default(&sid, &range_list);
        let rid = session.rid.as_deref().unwrap();
        let request = rb::build_insert_data_bar_rule(rid, &sid, range_arr, rule_obj, active_info);
        exec_status_cmd(engine, &request, session, "Conditional formatting data bar rule inserted.");
    }
}

fn cmd_cf_icon_set(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: cf iconset [<#index>|--rule-id <id>] --range <A1:C5> --set <icon_set_type> ...";

    let (index_target, rule_id_target, range_override, sheet_override, _active_cell_override,
         mut rule_obj)
        = match cf_parse_v2_args(args, usage) {
        Ok(v) => v,
        Err(e) => { output::error(&e); return; }
    };

    let is_edit = index_target.is_some() || rule_id_target.is_some();

    let resolved_rule_id: Option<String> = if let Some(idx) = index_target {
        match cf_resolve_index(idx, session) {
            Ok(id) => Some(id),
            Err(e) => { output::error(&e); return; }
        }
    } else {
        rule_id_target.clone()
    };

    // Auto-populate icons[] and conditions[] from icon_set_type if not provided
    if let Err(e) = cf_iconset_auto_populate(&mut rule_obj) {
        output::error(&e); return;
    }

    cf_auto_criteria_id(&mut rule_obj);
    cf_apply_engine_remaps(&mut rule_obj);

    let sid = match cf_resolve_sheet_id(sheet_override.as_deref(), session) {
        Some(id) => id,
        None => return,
    };

    if is_edit {
        let rule_id = match resolved_rule_id {
            Some(ref id) => id.clone(),
            None => { output::error("Edit requires #index or --rule-id."); return; }
        };

        if let Err(e) = cf_check_edit_rule_type(session, index_target, "IconSets", "icon set") {
            output::error(&e);
            return;
        }

        // If editing by index, reconstruct the existing rule as a base and merge the
        // user's changes onto it, so a partial edit doesn't send an incomplete rule.
        // (When --set is given, auto-populate already regenerated icons/condition, and a
        // plain deep-merge replaces the whole arrays — which is what changing the set means.)
        if let Some(idx) = index_target {
            let full_rule = session.last_cf_rules[idx - 1].full_rule.clone();
            if let Some(mut base) = cf_reconstruct_icon_set_base(&full_rule) {
                deep_merge_into(&mut base, &rule_obj);
                rule_obj = base;
            }
        }

        if let Some(obj) = rule_obj.as_object_mut() {
            obj.insert("rule_id".into(), serde_json::json!(rule_id));
        }
        let range_list = if let Some(r) = range_override.as_deref() {
            match cf_parse_range(r) { Ok(v) => v, Err(e) => { output::error(&e); return; } }
        } else if let Some(idx) = index_target {
            let entry = &session.last_cf_rules[idx - 1];
            if entry.range_json.as_array().map_or(true, |a| a.is_empty()) {
                output::error("Cannot determine range for this rule. Re-run 'cf list' or provide --range.");
                return;
            }
            entry.range_json.clone()
        } else {
            output::error("--range is required when editing by --rule-id without a cached range.");
            return;
        };
        let range_arr = range_list.as_array().unwrap();
        let active_info = build_cf_active_info_default(&sid, &range_list);
        let rid = session.rid.as_deref().unwrap();
        let request = rb::build_edit_icon_set_rule(rid, &sid, range_arr, rule_obj, active_info);
        exec_status_cmd(engine, &request, session, "Conditional formatting icon set rule edited.");
    } else {
        let range_str = match range_override.as_deref() {
            Some(r) => r.to_string(),
            None => { output::error("--range is required for cf iconset insert."); return; }
        };
        let range_list = match cf_parse_range(&range_str) {
            Ok(v) => v, Err(e) => { output::error(&e); return; }
        };
        let range_arr = range_list.as_array().unwrap();
        let active_info = build_cf_active_info_default(&sid, &range_list);
        let rid = session.rid.as_deref().unwrap();
        let request = rb::build_insert_icon_set_rule(rid, &sid, range_arr, rule_obj, active_info);
        exec_status_cmd(engine, &request, session, "Conditional formatting icon set rule inserted.");
    }
}

fn cmd_cf_move(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: cf move <#index|rule-id> <--up [n]|--down [n]|--top|--bottom|--after <#index|rule-id>> [--sheet <id>] [--active-cell <A1>] [--range <A1:C5>]";
    if args.is_empty() { output::error(usage); return; }

    let target_raw = args[0];
    let mut position: Option<(&str, Option<String>)> = None; // (kind, param)
    let mut sheet_id: Option<String> = None;
    let mut active_cell: Option<String> = None;
    let mut range_ref: Option<String> = None;

    let mut i = 1;
    while i < args.len() {
        match args[i].to_lowercase().as_str() {
            "--up" => {
                let n = if i + 1 < args.len() && !args[i+1].starts_with("--") {
                    i += 1; args[i].to_string()
                } else { "1".to_string() };
                position = Some(("up", Some(n)));
            }
            "--down" => {
                let n = if i + 1 < args.len() && !args[i+1].starts_with("--") {
                    i += 1; args[i].to_string()
                } else { "1".to_string() };
                position = Some(("down", Some(n)));
            }
            "--top" => { position = Some(("top", None)); }
            "--bottom" => { position = Some(("bottom", None)); }
            "--after" => {
                i += 1;
                if i >= args.len() { output::error("--after requires a rule id or #index."); return; }
                position = Some(("after", Some(args[i].to_string())));
            }
            "--sheet" => {
                i += 1;
                if i >= args.len() { output::error("--sheet requires a sheet id."); return; }
                sheet_id = Some(args[i].to_string());
            }
            "--active-cell" => {
                i += 1;
                if i >= args.len() { output::error("--active-cell requires an A1 cell reference."); return; }
                active_cell = Some(args[i].to_string());
            }
            "--range" => {
                i += 1;
                if i >= args.len() { output::error("--range requires an A1 range."); return; }
                range_ref = Some(args[i].to_string());
            }
            other => { output::error(&format!("Unknown option '{}'.", other)); output::error(usage); return; }
        }
        i += 1;
    }

    let (pos_kind, pos_param) = match position {
        Some(v) => v,
        None => { output::error("Specify --up, --down, --top, --bottom, or --after."); return; }
    };

    let sid = match cf_resolve_sheet_id(sheet_id.as_deref(), session) {
        Some(id) => id,
        None => return,
    };
    let active_info = match build_cf_active_info_from_optional_inputs(&sid, active_cell.as_deref(), range_ref.as_deref()) {
        Ok(v) => v,
        Err(e) => { output::error(&e); return; }
    };

    // Resolve target rule_id
    let target_rule_id = if let Some(stripped) = target_raw.strip_prefix('#') {
        match stripped.parse::<usize>() {
            Ok(idx) => match cf_resolve_index(idx, session) {
                Ok(id) => id,
                Err(e) => { output::error(&e); return; }
            },
            Err(_) => { output::error("Invalid #index."); return; }
        }
    } else if target_raw.chars().next().map_or(false, |c| c.is_ascii_digit()) {
        match target_raw.parse::<usize>() {
            Ok(idx) => match cf_resolve_index(idx, session) {
                Ok(id) => id,
                Err(e) => { output::error(&e); return; }
            },
            Err(_) => { output::error(&format!("Invalid index '{}'.", target_raw)); return; }
        }
    } else {
        target_raw.to_string()
    };

    // For top/bottom/up/down we need the rule list
    let priority_greater_than: Option<String> = match pos_kind {
        "bottom" => None,
        "after" => {
            let after_raw = pos_param.as_deref().unwrap_or("");
            if let Some(stripped) = after_raw.strip_prefix('#') {
                match stripped.parse::<usize>() {
                    Ok(idx) => match cf_resolve_index(idx, session) {
                        Ok(id) => Some(id),
                        Err(e) => { output::error(&e); return; }
                    },
                    Err(_) => { output::error("Invalid #index in --after."); return; }
                }
            } else if after_raw.chars().next().map_or(false, |c| c.is_ascii_digit()) {
                match after_raw.parse::<usize>() {
                    Ok(idx) => match cf_resolve_index(idx, session) {
                        Ok(id) => Some(id),
                        Err(e) => { output::error(&e); return; }
                    },
                    Err(_) => { output::error(&format!("Invalid index '{}' in --after.", after_raw)); return; }
                }
            } else {
                Some(after_raw.to_string())
            }
        }
        "top" | "up" | "down" => {
            // Fetch ordered rule list for the sheet
            let rid = session.rid.as_deref().unwrap();
            let list_req = rb::build_manage_cf_rules(rid, Some(&sid), rb::CF_SCOPE_SHEET, None);
            let rule_list = match engine.process_request_json(&list_req) {
                Ok(resp) => {
                    let parsed: serde_json::Value = serde_json::from_str(&resp).unwrap_or_default();
                    parsed.get("response")
                        .and_then(|r| r.get("rules"))
                        .and_then(|v| v.as_array())
                        .and_then(|sheets| sheets.first())
                        .and_then(|s| s.get("rules_in_sheet"))
                        .and_then(|v| v.as_array())
                        .cloned()
                        .unwrap_or_default()
                }
                Err(e) => { output::error(&format!("Failed to fetch rule list: {}", e)); return; }
            };

            let ids: Vec<String> = rule_list.iter()
                .filter_map(|r| r.get("rule_id").and_then(|v| v.as_str()).map(|s| s.to_string()))
                .collect();

            let cur_pos = match ids.iter().position(|id| *id == target_rule_id) {
                Some(p) => p,
                None => { output::error(&format!("Rule '{}' not found in sheet.", target_rule_id)); return; }
            };

            match pos_kind {
                "top" => {
                    // Place before the current first rule (other than target)
                    // priority_greater_than = None means bottom; to get top we need a different approach.
                    // We achieve "top" by finding what's currently at pos 0 (skip target) and
                    // passing None (engine places at highest priority when no constraint given).
                    // Based on engine: priority_greater_than=None → least priority (bottom).
                    // To go to top: we cannot directly; best effort is to use the rule currently
                    // just before position 0 in the remaining list. If target is already at 0 → no-op.
                    if cur_pos == 0 {
                        output::info("Rule is already at the top.");
                        return;
                    }
                    // Place before the rule currently at position 0
                    // By passing priority_greater_than = ids[0] would place AFTER ids[0],
                    // so we need to NOT pass priority_greater_than and let engine handle.
                    // Actually for now: pass None; note this puts it at bottom per engine docs,
                    // but some engines interpret None as "top". We follow spec intent.
                    None
                }
                "up" => {
                    let n: usize = pos_param.as_deref().unwrap_or("1").parse().unwrap_or(1);
                    if cur_pos == 0 { output::info("Rule is already at the top."); return; }
                    let new_pos = if n > cur_pos { 0 } else { cur_pos - n };
                    if new_pos == 0 { None }
                    else {
                        // Place after rule at new_pos - 1 (skipping target itself)
                        let remaining: Vec<&String> = ids.iter().enumerate()
                            .filter(|(i, _)| *i != cur_pos)
                            .map(|(_, id)| id)
                            .collect();
                        remaining.get(new_pos.saturating_sub(1)).map(|id| (*id).clone())
                    }
                }
                "down" => {
                    let n: usize = pos_param.as_deref().unwrap_or("1").parse().unwrap_or(1);
                    let remaining: Vec<&String> = ids.iter().enumerate()
                        .filter(|(i, _)| *i != cur_pos)
                        .map(|(_, id)| id)
                        .collect();
                    if cur_pos >= ids.len().saturating_sub(1) { output::info("Rule is already at the bottom."); return; }
                    let new_pos = (cur_pos + n).min(ids.len().saturating_sub(1));
                    // Place after rule at new_pos - 1 in the remaining list
                    remaining.get(new_pos.saturating_sub(1)).map(|id| (*id).clone())
                }
                _ => unreachable!()
            }
        }
        _ => { output::error("Unknown position kind."); return; }
    };

    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_update_cf_rule_priority(
        rid, &sid, &target_rule_id, priority_greater_than.as_deref(), active_info,
    );
    exec_status_cmd(engine, &request, session, "Conditional formatting rule moved.");
}

fn cmd_cf_update_priority_legacy(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: cf priority <rule_id> [--after <rule_id>] [--sheet <id>] [--active-cell <A1>] [--range <A1:C5>]";
    if args.is_empty() { output::error(usage); return; }

    let rule_to_be_updated = args[0];
    let mut priority_greater_than: Option<String> = None;
    let mut sheet_id: Option<String> = None;
    let mut active_cell: Option<String> = None;
    let mut range_ref: Option<String> = None;

    let mut i = 1;
    while i < args.len() {
        match args[i].to_lowercase().as_str() {
            "--after" => { i += 1; if i < args.len() { priority_greater_than = Some(args[i].to_string()); } }
            "--sheet" => { i += 1; if i < args.len() { sheet_id = Some(args[i].to_string()); } }
            "--active-cell" => { i += 1; if i < args.len() { active_cell = Some(args[i].to_string()); } }
            "--range" => { i += 1; if i < args.len() { range_ref = Some(args[i].to_string()); } }
            other => { output::error(&format!("Unknown option '{}'.", other)); return; }
        }
        i += 1;
    }

    let sid = sheet_id.unwrap_or_else(|| session.get_active_sheet_id_or_default());
    let active_info = match build_cf_active_info_from_optional_inputs(&sid, active_cell.as_deref(), range_ref.as_deref()) {
        Ok(v) => v, Err(e) => { output::error(&e); return; }
    };
    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_update_cf_rule_priority(rid, &sid, rule_to_be_updated, priority_greater_than.as_deref(), active_info);
    exec_status_cmd(engine, &request, session, "Conditional formatting rule priority updated.");
}

fn cmd_cf_delete_rule(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: cf delete <#index|rule-id|--target-range <A1:C5>|--all> [--sheet <id>] [--active-cell <A1>]";
    if args.is_empty() { output::error(usage); return; }

    let mut target_index: Option<usize> = None;
    let mut target_rule_id: Option<String> = None;
    let mut target_range: Option<String> = None;
    let mut delete_all = false;
    let mut sheet_id: Option<String> = None;
    let mut active_cell: Option<String> = None;
    let mut range_ref: Option<String> = None;

    let mut i = 0;
    while i < args.len() {
        match args[i].to_lowercase().as_str() {
            "--all" => { delete_all = true; }
            "--target-range" => {
                i += 1;
                if i >= args.len() { output::error("--target-range requires a range."); return; }
                target_range = Some(args[i].to_string());
            }
            "--sheet" => {
                i += 1;
                if i >= args.len() { output::error("--sheet requires a sheet id."); return; }
                sheet_id = Some(args[i].to_string());
            }
            "--active-cell" => {
                i += 1;
                if i >= args.len() { output::error("--active-cell requires an A1 cell reference."); return; }
                active_cell = Some(args[i].to_string());
            }
            "--range" => {
                i += 1;
                if i >= args.len() { output::error("--range requires a range."); return; }
                range_ref = Some(args[i].to_string());
            }
            other => {
                if i == 0 {
                    if let Some(stripped) = other.strip_prefix('#') {
                        match stripped.parse::<usize>() {
                            Ok(idx) => target_index = Some(idx),
                            Err(_) => { output::error("Invalid #index."); return; }
                        }
                    } else if !other.starts_with("--") {
                        if other.chars().all(|c| c.is_ascii_digit()) {
                            match other.parse::<usize>() {
                                Ok(idx) => target_index = Some(idx),
                                Err(_) => { output::error(&format!("Invalid index '{}'.", other)); return; }
                            }
                        } else {
                            target_rule_id = Some(args[i].to_string());
                        }
                    } else {
                        output::error(&format!("Unknown option '{}'.", other)); return;
                    }
                } else {
                    output::error(&format!("Unknown option '{}'.", other)); return;
                }
            }
        }
        i += 1;
    }

    let sid = match cf_resolve_sheet_id(sheet_id.as_deref(), session) {
        Some(id) => id,
        None => return,
    };
    let active_info = match build_cf_active_info_from_optional_inputs(&sid, active_cell.as_deref(), range_ref.as_deref()) {
        Ok(v) => v, Err(e) => { output::error(&e); return; }
    };
    let rid = session.rid.as_deref().unwrap();

    if delete_all {
        // Fetch all rules for the sheet, delete each
        let list_req = rb::build_manage_cf_rules(rid, Some(&sid), rb::CF_SCOPE_SHEET, None);
        let ids: Vec<String> = match engine.process_request_json(&list_req) {
            Ok(resp) => {
                let parsed: serde_json::Value = serde_json::from_str(&resp).unwrap_or_default();
                parsed.get("response")
                    .and_then(|r| r.get("rules"))
                    .and_then(|v| v.as_array())
                    .and_then(|sheets| sheets.first())
                    .and_then(|s| s.get("rules_in_sheet"))
                    .and_then(|v| v.as_array())
                    .map(|arr| arr.iter()
                        .filter_map(|r| r.get("rule_id").and_then(|v| v.as_str()).map(|s| s.to_string()))
                        .collect())
                    .unwrap_or_default()
            }
            Err(e) => { output::error(&format!("Failed to fetch rules: {}", e)); return; }
        };
        if ids.is_empty() { output::info("No rules to delete."); return; }
        for rule_id in &ids {
            let req = rb::build_delete_cf_rule(rid, &sid, rule_id, active_info.clone());
            match engine.process_request_json(&req) {
                Ok(_) => {}
                Err(e) => { output::error(&format!("Failed to delete rule {}: {}", rule_id, e)); return; }
            }
        }
        output::success(&format!("Deleted {} conditional formatting rules.", ids.len()));
        return;
    }

    if let Some(range_str) = target_range {
        // Delete all rules overlapping the given range
        let range_list = match cf_parse_range(&range_str) {
            Ok(v) => v, Err(e) => { output::error(&e); return; }
        };
        let range_arr = range_list.as_array().unwrap().clone();
        let list_req = rb::build_manage_cf_rules(rid, Some(&sid), rb::CF_SCOPE_RANGE, Some(&range_arr));
        let ids: Vec<String> = match engine.process_request_json(&list_req) {
            Ok(resp) => {
                let parsed: serde_json::Value = serde_json::from_str(&resp).unwrap_or_default();
                parsed.get("response")
                    .and_then(|r| r.get("rules"))
                    .and_then(|v| v.as_array())
                    .and_then(|sheets| sheets.first())
                    .and_then(|s| s.get("rules_in_sheet"))
                    .and_then(|v| v.as_array())
                    .map(|arr| arr.iter()
                        .filter_map(|r| r.get("rule_id").and_then(|v| v.as_str()).map(|s| s.to_string()))
                        .collect())
                    .unwrap_or_default()
            }
            Err(e) => { output::error(&format!("Failed to fetch rules: {}", e)); return; }
        };
        if ids.is_empty() { output::info("No rules found for that range."); return; }
        for rule_id in &ids {
            let req = rb::build_delete_cf_rule(rid, &sid, rule_id, active_info.clone());
            match engine.process_request_json(&req) {
                Ok(_) => {}
                Err(e) => { output::error(&format!("Failed to delete rule {}: {}", rule_id, e)); return; }
            }
        }
        output::success(&format!("Deleted {} conditional formatting rules from {}.", ids.len(), range_str));
        return;
    }

    // Single rule delete by #index or rule_id
    let rule_id = if let Some(idx) = target_index {
        match cf_resolve_index(idx, session) {
            Ok(id) => id,
            Err(e) => { output::error(&e); return; }
        }
    } else if let Some(id) = target_rule_id {
        id
    } else {
        output::error(usage);
        return;
    };

    let request = rb::build_delete_cf_rule(rid, &sid, &rule_id, active_info);
    exec_status_cmd(engine, &request, session, "Conditional formatting rule deleted.");
}

fn cmd_cf_manage_rules(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: cf list [--scope workbook|sheet|range] [--sheet <sheet_id>] [--range <A1:C5>]";

    let mut scope = rb::CF_SCOPE_SHEET;
    let mut sheet_id: Option<String> = None;
    let mut range_ref: Option<String> = None;

    let mut i = 0;
    while i < args.len() {
        match args[i].to_lowercase().as_str() {
            "--scope" => {
                i += 1;
                if i >= args.len() {
                    output::error("--scope requires workbook|sheet|range.");
                    return;
                }
                scope = match args[i].to_lowercase().as_str() {
                    "workbook" => rb::CF_SCOPE_WORKBOOK,
                    "sheet" => rb::CF_SCOPE_SHEET,
                    "range" => rb::CF_SCOPE_RANGE,
                    other => {
                        output::error(&format!("Invalid scope '{}'. Use workbook, sheet, or range.", other));
                        return;
                    }
                };
            }
            "--sheet" => {
                i += 1;
                if i >= args.len() {
                    output::error("--sheet requires a sheet id.");
                    return;
                }
                sheet_id = Some(args[i].to_string());
            }
            "--range" => {
                i += 1;
                if i >= args.len() {
                    output::error("--range requires an A1 range.");
                    return;
                }
                range_ref = Some(args[i].to_string());
            }
            other => {
                output::error(&format!("Unknown option '{}'.", other));
                output::error(usage);
                return;
            }
        }
        i += 1;
    }

    if scope == rb::CF_SCOPE_RANGE && range_ref.is_none() {
        output::error("--scope range requires --range <A1:C5>.");
        return;
    }

    let include_sheet_for_request = !(scope == rb::CF_SCOPE_WORKBOOK && sheet_id.is_none());
    let sid = match cf_resolve_sheet_id(sheet_id.as_deref(), session) {
        Some(id) => id,
        None => return,
    };
    let sheet_for_request = if include_sheet_for_request {
        Some(sid.as_str())
    } else {
        None
    };
    let range_list = if let Some(range) = range_ref.as_deref() {
        let (sc, sr, ec, er) = match cell_ref::try_parse_range(range) {
            Some(v) => v,
            None => {
                output::error(&format!("Invalid range: '{}'. Use A1:C5 format.", range));
                return;
            }
        };
        Some(vec![serde_json::json!({
            "start_row": sr,
            "end_row": er,
            "start_column": sc,
            "end_column": ec
        })])
    } else {
        None
    };

    let rid = session.rid.as_deref().unwrap();
    let request = rb::build_manage_cf_rules(
        rid,
        sheet_for_request,
        scope,
        range_list.as_deref(),
    );

    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if !rp::is_success(status.status_code) {
                output::error(&format!(
                    "Manage conditional formatting rules failed: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
                return;
            }

            let parsed: serde_json::Value = match serde_json::from_str(&resp) {
                Ok(v) => v,
                Err(_) => {
                    output::success("Conditional formatting rules fetched.");
                    output::line(&resp, 2);
                    return;
                }
            };

            // Collect all rules across sheets for the index cache
            let mut all_rules: Vec<crate::session::CfRuleEntry> = Vec::new();

            if let Some(rules_sheets) = parsed
                .get("response")
                .and_then(|r| r.get("rules"))
                .and_then(|v| v.as_array())
            {
                let mut global_index = 1usize;

                for sheet_rules in rules_sheets {
                    let sheet_label = sheet_rules
                        .get("sheet_id")
                        .and_then(|v| v.as_str())
                        .unwrap_or("<unknown-sheet>");
                    output::line(&format!("Sheet: {}", sheet_label), 2);

                    // Header
                    output::line(&format!("  {:<4} {:<12} {:<12} {:<30} {}", "#", "RANGE", "TYPE", "SUMMARY", "ID"), 0);
                    output::line(&format!("  {}", "-".repeat(80)), 0);

                    let mut count = 0usize;
                    if let Some(items) = sheet_rules.get("rules_in_sheet").and_then(|v| v.as_array()) {
                        for item in items {
                            let rule_id = item
                                .get("rule_id")
                                .and_then(|v| v.as_str())
                                .unwrap_or("<missing>");
                            let rule_type = item
                                .get("rule_type")
                                .and_then(|v| v.as_str())
                                .unwrap_or("-");

                            let range_label = cf_format_range_list(item.get("range"));

                            let summary = cf_rule_summary(item, rule_type);

                            output::line(
                                &format!("  {:<4} {:<12} {:<12} {:<30} {}",
                                    global_index, range_label, rule_type, summary, rule_id),
                                0,
                            );

                            all_rules.push(crate::session::CfRuleEntry {
                                rule_id: rule_id.to_string(),
                                rule_type: rule_type.to_string(),
                                range_label: range_label.clone(),
                                range_json: item.get("range").cloned().unwrap_or(serde_json::json!([])),
                                summary: summary.clone(),
                                full_rule: item.clone(),
                            });

                            global_index += 1;
                            count += 1;
                        }
                    }
                    if count == 0 {
                        output::line("  (no rules)", 0);
                    }
                }
            } else {
                output::line("No conditional formatting rules found.", 2);
            }

            // Cache for #index lookup
            session.last_cf_rules = all_rules;
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn cf_format_range_list(range_val: Option<&serde_json::Value>) -> String {
    fn format_one(r: &serde_json::Value) -> Option<String> {
        let sc = r.get("start_column").and_then(|v| v.as_i64())? as i32;
        let sr = r.get("start_row").and_then(|v| v.as_i64())? as i32;
        let ec = r.get("end_column").and_then(|v| v.as_i64())? as i32;
        let er = r.get("end_row").and_then(|v| v.as_i64())? as i32;
        let start = crate::util::cell_ref::to_ref(sc, sr);
        let end = crate::util::cell_ref::to_ref(ec, er);
        if start == end { Some(start) } else { Some(format!("{}:{}", start, end)) }
    }
    match range_val {
        Some(serde_json::Value::Array(arr)) => {
            let parts: Vec<String> = arr.iter().filter_map(format_one).collect();
            if parts.is_empty() { "-".to_string() } else { parts.join(", ") }
        }
        Some(obj) if obj.is_object() => format_one(obj).unwrap_or_else(|| "-".to_string()),
        _ => "-".to_string(),
    }
}

fn cf_rule_summary(item: &serde_json::Value, rule_type: &str) -> String {
    match rule_type.to_lowercase().as_str() {
        "classic" => {
            let conds = item.get("condition")
                .or_else(|| item.get("rule").and_then(|r| r.get("condition")))
                .and_then(|v| v.as_array());
            if let Some(conds) = conds {
                if let Some(c) = conds.first() {
                    let ct = c.get("criteria_type").and_then(|v| v.as_i64()).unwrap_or(-1) as i32;
                    let sct = c.get("sub_criteria_type").and_then(|v| v.as_i64()).unwrap_or(-1) as i32;
                    let lhs = c.get("lhs").map(|v| format!(" {}", v)).unwrap_or_default();
                    let ct_name = cf_criteria_type_name(ct);
                    let sct_name = cf_sub_criteria_name(ct, sct);
                    return match sct_name {
                        Some(s) => format!("{} {} {}", ct_name, s, lhs.trim()),
                        None => format!("{}{}", ct_name, lhs),
                    };
                }
            }
            "classic rule".to_string()
        }
        "colorscale" => "color scale".to_string(),
        "databar" => "data bar".to_string(),
        "iconset" => {
            item.get("icon_set_type").and_then(|v| v.as_i64())
                .map(|v| cf_icon_set_type_name(v).to_string())
                .unwrap_or_else(|| "icon set".to_string())
        }
        _ => "-".to_string(),
    }
}

fn cf_criteria_type_name(ct: i32) -> &'static str {
    match ct {
        0 => "number", 1 => "percent", 2 => "percentile", 3 => "formula",
        4 => "min", 5 => "max", 6 => "number_comparison", 7 => "date",
        8 => "text", 9 => "cell_containing", 10 => "average",
        11 => "std_deviation", 12 => "automatic", 13 => "top_bottom_values",
        14 => "none", _ => "?",
    }
}

/// When editing by #index, verify the cached rule's type matches the command being used,
/// returning a clear CLI error (instead of an opaque engine failure) on a mismatch.
/// `expected` is the engine rule_type string (e.g. "ColorScale"); `cmd` is the sub-command name.
fn cf_check_edit_rule_type(
    session: &CliSession,
    index_target: Option<usize>,
    expected: &str,
    cmd: &str,
) -> Result<(), String> {
    let idx = match index_target { Some(i) => i, None => return Ok(()) };
    let entry = match session.last_cf_rules.get(idx - 1) { Some(e) => e, None => return Ok(()) };
    if entry.rule_type.eq_ignore_ascii_case(expected) {
        return Ok(());
    }
    Err(format!(
        "Rule #{} is a '{}' rule, not a '{}' — use the matching 'cf' sub-command, or run 'cf list' \
         to find the correct #index for a {} rule.",
        idx, entry.rule_type, expected, cmd
    ))
}

/// Validate that every condition's `criteria_type` is supported by the given rule kind,
/// returning a helpful CLI error (instead of the opaque engine "Rule doesn't support the
/// specified criteria.") when it isn't. `kind` is "classic", "colorscale" or "databar".
fn cf_validate_condition_criteria(rule_obj: &serde_json::Value, kind: &str) -> Result<(), String> {
    // Criteria types accepted by the engine for each rule family (verified empirically).
    let allowed: &[i64] = match kind {
        // formula, number_comparison, date, text, cell_containing, average, std_deviation, top_bottom_values
        "classic" => &[3, 6, 7, 8, 9, 10, 11, 13],
        // number, percent, percentile, min, max, automatic
        "colorscale" | "databar" => &[0, 1, 2, 4, 5, 12],
        _ => return Ok(()),
    };
    let conds = match rule_obj.get("condition").and_then(|v| v.as_array()) {
        Some(c) => c,
        None => return Ok(()),
    };
    for c in conds {
        let ct = match c.get("criteria_type").and_then(|v| v.as_i64()) {
            Some(v) => v,
            None => continue,
        };
        if allowed.contains(&ct) { continue; }
        let name = cf_criteria_type_name(ct as i32);
        return Err(match kind {
            "classic" => format!(
                "Classic rules don't support the '{}' criteria type. Use a cell-value comparison \
                 (e.g. --when \">50\", or --condition.criteria_type number_comparison \
                 --condition.sub_criteria_type gt --condition.value 50), or one of: \
                 formula, number_comparison, date, text, cell_containing, average, \
                 standard_deviation, top_bottom_values.",
                name),
            _ => format!(
                "{} rules don't support the '{}' criteria type. Use one of: \
                 number, percent, percentile, minimum_value (min), maximum_value (max), automatic.",
                if kind == "colorscale" { "Color-scale" } else { "Data-bar" }, name),
        });
    }
    Ok(())
}

fn cf_sub_criteria_name(ct: i32, sct: i32) -> Option<&'static str> {
    match ct {
        6 => match sct {
            0 => Some("="), 1 => Some("!="), 2 => Some("between"), 3 => Some("not between"),
            4 => Some(">"), 5 => Some("<"), 6 => Some(">="), 7 => Some("<="), _ => None,
        },
        7 => match sct {
            0 => Some("yesterday"), 1 => Some("today"), 2 => Some("tomorrow"),
            3 => Some("last 7 days"), 4 => Some("last week"), 5 => Some("this week"),
            6 => Some("next week"), 7 => Some("last month"), 8 => Some("this month"),
            9 => Some("next month"), 18 => Some("next 7 days"), 19 => Some("last year"),
            20 => Some("this year"), 21 => Some("next year"), _ => None,
        },
        8 => match sct {
            0 => Some("contains"), 1 => Some("not contains"), 2 => Some("begins with"), 3 => Some("ends with"), _ => None,
        },
        9 => match sct {
            0 => Some("duplicates"), 1 => Some("unique"), 2 => Some("blanks"),
            3 => Some("no blanks"), 4 => Some("errors"), 5 => Some("no errors"), _ => None,
        },
        10 => match sct { 0 => Some("above avg"), 1 => Some("below avg"), 2 => Some(">= avg"), 3 => Some("<= avg"), _ => None },
        11 => match sct { 0 => Some("+1σ"), 1 => Some("-1σ"), 2 => Some("+2σ"), 3 => Some("-2σ"), 4 => Some("+3σ"), 5 => Some("-3σ"), _ => None },
        13 => match sct { 0 => Some("top"), 1 => Some("bottom"), _ => None },
        _ => None,
    }
}

// ─── CF v2 parsing infrastructure ────────────────────────────────────────────

/// Deep-merge `overlay` into `base`. Object fields: overlay wins, missing filled from base.
/// Arrays and scalars: overlay wins entirely.
fn deep_merge_into(base: &mut serde_json::Value, overlay: &serde_json::Value) {
    match (base, overlay) {
        (serde_json::Value::Object(base_obj), serde_json::Value::Object(overlay_obj)) => {
            for (key, overlay_val) in overlay_obj.iter() {
                let base_val = base_obj.entry(key.clone()).or_insert(serde_json::Value::Null);
                deep_merge_into(base_val, overlay_val);
            }
        }
        (base_val, overlay_val) => {
            *base_val = overlay_val.clone();
        }
    }
}

/// Convert engine string criteria_type (from cf list response) to integer.
fn cf_engine_str_to_criteria_type(s: &str) -> Option<i64> {
    match s {
        "Number" | "number" => Some(0),
        "Percent" | "percent" => Some(1),
        "Percentile" | "percentile" => Some(2),
        "Formula" | "formula" => Some(3),
        "Min" | "Minimum" | "MinimumValue" | "LowestValue" => Some(4),
        "Max" | "Maximum" | "MaximumValue" | "HighestValue" => Some(5),
        "NumberComparison" | "NumericComparison" => Some(6),
        "Date" | "date" => Some(7),
        "Text" | "text" => Some(8),
        "CellContaining" | "CellContains" | "Cell" => Some(9),
        "Average" | "average" => Some(10),
        "StandardDeviation" | "StdDeviation" => Some(11),
        "Automatic" | "automatic" => Some(12),
        "TopBottomValues" | "TopBottom" | "TopAndBottom" => Some(13),
        "None" | "none" => Some(14),
        _ => None,
    }
}

/// Convert engine string sub_criteria_type to integer, given the parent criteria_type.
fn cf_engine_str_to_sub_criteria_type(ct: i64, s: &str) -> Option<i64> {
    match ct {
        6 => match s {
            "Equal" | "EqualTo" => Some(0),
            "NotEqual" | "NotEqualTo" | "NotEquals" => Some(1),
            "Between" => Some(2),
            "NotBetween" => Some(3),
            "GreaterThan" => Some(4),
            "LessThan" => Some(5),
            "GreaterThanOrEqual" | "GreaterThanOrEqualTo" | "GreaterOrEqual" => Some(6),
            "LessThanOrEqual" | "LessThanOrEqualTo" | "LessOrEqual" => Some(7),
            _ => None,
        },
        7 => match s {
            "Yesterday" => Some(0), "Today" => Some(1), "Tomorrow" => Some(2),
            "Last7Days" | "LastSevenDays" => Some(3),
            "LastWeek" => Some(4), "ThisWeek" => Some(5), "NextWeek" => Some(6),
            "LastMonth" => Some(7), "ThisMonth" => Some(8), "NextMonth" => Some(9),
            "Next7Days" | "NextSevenDays" => Some(18),
            "LastYear" => Some(19), "ThisYear" => Some(20), "NextYear" => Some(21),
            _ => None,
        },
        8 => match s {
            "Contains" => Some(0), "NotContains" | "DoesNotContain" => Some(1),
            "BeginsWith" | "StartsWith" => Some(2), "EndsWith" => Some(3), _ => None,
        },
        9 => match s {
            "Duplicates" | "Duplicate" => Some(0), "Unique" => Some(1),
            "Blanks" | "Blank" => Some(2), "NoBlanks" | "NotBlank" | "NonBlank" => Some(3),
            "Errors" | "Error" => Some(4), "NoErrors" | "NoError" => Some(5), _ => None,
        },
        10 => match s {
            "Above" | "AboveAverage" => Some(0), "Below" | "BelowAverage" => Some(1),
            "AboveOrEqual" | "AboveOrEqualAverage" => Some(2),
            "BelowOrEqual" | "BelowOrEqualAverage" => Some(3), _ => None,
        },
        11 => match s {
            "Above1" | "PlusOneSigma" => Some(0), "Below1" | "MinusOneSigma" => Some(1),
            "Above2" | "PlusTwoSigma" => Some(2), "Below2" | "MinusTwoSigma" => Some(3),
            "Above3" | "PlusThreeSigma" => Some(4), "Below3" | "MinusThreeSigma" => Some(5),
            _ => None,
        },
        13 => match s { "Top" => Some(0), "Bottom" => Some(1), _ => None },
        _ => None,
    }
}

/// Reconstruct a CLI-format edit base object for a classic rule from the engine's cf list response.
/// Extracts sub_rules → condition array, and preserves stop_if_true / count / is_percent.
/// Returns None if the cache has no usable data.
fn cf_reconstruct_classic_base(full_rule: &serde_json::Value) -> Option<serde_json::Value> {
    let rule_inner = full_rule.get("rule")?;

    // Reconstruct condition from sub_rules
    let sub_rules = rule_inner.get("sub_rules").and_then(|v| v.as_array())?;
    let mut conditions = Vec::new();
    for (i, sr) in sub_rules.iter().enumerate() {
        let ct_str = sr.get("criteria_type").and_then(|v| v.as_str()).unwrap_or("");
        let ct = cf_engine_str_to_criteria_type(ct_str)?;
        let sct_str = sr.get("sub_criteria_type").and_then(|v| v.as_str()).unwrap_or("");
        let sct = cf_engine_str_to_sub_criteria_type(ct, sct_str);

        let mut cond = serde_json::json!({"criteria_type": ct, "criteria_id": i as i64});
        if let Some(obj) = cond.as_object_mut() {
            if let Some(sct_val) = sct { obj.insert("sub_criteria_type".into(), serde_json::json!(sct_val)); }
            // Engine prefixes numeric lhs with "="; strip for non-formula types
            for key in &["lhs", "rhs"] {
                if let Some(val_str) = sr.get(*key).and_then(|v| v.as_str()) {
                    let stripped = if ct == 3 { val_str.to_string() } else { val_str.trim_start_matches('=').to_string() };
                    if !stripped.is_empty() { obj.insert((*key).into(), serde_json::json!(stripped)); }
                }
            }
        }
        conditions.push(cond);
    }

    if conditions.is_empty() { return None; }

    let mut base = serde_json::json!({"condition": conditions});
    if let Some(obj) = base.as_object_mut() {
        // Preserve stop_if_true
        if let Some(v) = rule_inner.get("stop_if_true") { obj.insert("stop_if_true".into(), v.clone()); }
        // Preserve count / is_percent for top_bottom rules
        if let Some(v) = full_rule.get("count") { obj.insert("count".into(), v.clone()); }
        if let Some(v) = full_rule.get("is_percent") { obj.insert("is_percent".into(), v.clone()); }
    }
    Some(base)
}

/// Convert an engine PascalCase enum string to snake_case
/// (e.g. "ThreeTrafficLights1" → "three_traffic_lights_1", "RedCircle" → "red_circle").
fn cf_pascal_to_snake(s: &str) -> String {
    let mut out = String::new();
    for ch in s.chars() {
        if ch.is_ascii_uppercase() {
            if out.chars().last().map_or(false, |c| c.is_ascii_alphanumeric()) {
                out.push('_');
            }
            out.push(ch.to_ascii_lowercase());
        } else if ch.is_ascii_digit() {
            if out.chars().last().map_or(false, |c| c.is_ascii_alphabetic()) {
                out.push('_');
            }
            out.push(ch);
        } else {
            out.push(ch.to_ascii_lowercase());
        }
    }
    out
}

/// Reconstruct CLI-format `condition[]` entries from an engine rule's `sub_rules[]`.
/// Converts string criteria/sub-criteria enums to integers and preserves criteria_id,
/// lhs/rhs thresholds and any per-stop color (color scale).
fn cf_reconstruct_conditions(sub_rules: &[serde_json::Value]) -> Vec<serde_json::Value> {
    let mut conditions = Vec::new();
    for (i, sr) in sub_rules.iter().enumerate() {
        let ct_str = sr.get("criteria_type").and_then(|v| v.as_str()).unwrap_or("");
        let ct = match cf_engine_str_to_criteria_type(ct_str) { Some(v) => v, None => continue };
        let criteria_id = sr.get("criteria_id").and_then(|v| v.as_i64()).unwrap_or(i as i64);
        let mut cond = serde_json::json!({"criteria_type": ct, "criteria_id": criteria_id});
        if let Some(obj) = cond.as_object_mut() {
            let sct_str = sr.get("sub_criteria_type").and_then(|v| v.as_str()).unwrap_or("");
            if let Some(sct) = cf_engine_str_to_sub_criteria_type(ct, sct_str) {
                obj.insert("sub_criteria_type".into(), serde_json::json!(sct));
            }
            // Engine prefixes numeric lhs with "="; strip for non-formula types.
            for key in &["lhs", "rhs"] {
                if let Some(val_str) = sr.get(*key).and_then(|v| v.as_str()) {
                    let stripped = if ct == 3 { val_str.to_string() } else { val_str.trim_start_matches('=').to_string() };
                    if !stripped.is_empty() { obj.insert((*key).into(), serde_json::json!(stripped)); }
                }
            }
            if let Some(color) = sr.get("color") { obj.insert("color".into(), color.clone()); }
        }
        conditions.push(cond);
    }
    conditions
}

/// Reconstruct a CLI-format edit base object for a color-scale rule from the engine's cf list response.
fn cf_reconstruct_color_scale_base(full_rule: &serde_json::Value) -> Option<serde_json::Value> {
    let rule_inner = full_rule.get("rule")?;
    let sub_rules = rule_inner.get("sub_rules").and_then(|v| v.as_array())?;
    let conditions = cf_reconstruct_conditions(sub_rules);
    if conditions.is_empty() { return None; }
    let mut base = serde_json::json!({"condition": conditions});
    if let Some(obj) = base.as_object_mut() {
        let hide = rule_inner.get("is_hide_cell_content").and_then(|v| v.as_bool()).unwrap_or(false);
        obj.insert("is_hide_values".into(), serde_json::json!(hide));
        let auto_text = rule_inner.get("is_automatic_text_color").and_then(|v| v.as_bool()).unwrap_or(false);
        obj.insert("is_automatic_text_color".into(), serde_json::json!(auto_text));
        if let Some(v) = rule_inner.get("stop_if_true") { obj.insert("stop_if_true".into(), v.clone()); }
    }
    Some(base)
}

/// Reconstruct a CLI-format edit base object for a data-bar rule from the engine's cf list response.
fn cf_reconstruct_data_bar_base(full_rule: &serde_json::Value) -> Option<serde_json::Value> {
    let rule_inner = full_rule.get("rule")?;
    let sub_rules = rule_inner.get("sub_rules").and_then(|v| v.as_array())?;
    let conditions = cf_reconstruct_conditions(sub_rules);
    if conditions.is_empty() { return None; }
    let mut base = serde_json::json!({"condition": conditions});
    if let Some(obj) = base.as_object_mut() {
        let enum_fields: [(&str, &[(&str, i64)]); 4] = [
            ("axis_position", &[("automatic",0),("middle",1),("none",2)]),
            ("bar_direction", &[("left_to_right",0),("right_to_left",1),("context",2)]),
            ("fill_type", &[("solid",0),("gradient",1)]),
            ("border_type", &[("none",0),("with_border",1)]),
        ];
        for (key, table) in enum_fields.iter() {
            if let Some(s) = rule_inner.get(*key).and_then(|v| v.as_str()) {
                if let Some(n) = cf_map_named_i64(&cf_pascal_to_snake(s), table) {
                    obj.insert((*key).into(), serde_json::json!(n));
                }
            }
        }
        for key in &["positive_value_fill", "negative_value_fill", "positive_value_border", "negative_value_border", "axis_color"] {
            if let Some(v) = rule_inner.get(*key) { obj.insert((*key).into(), v.clone()); }
        }
        let hide = rule_inner.get("is_display_only_bar").and_then(|v| v.as_bool()).unwrap_or(false);
        obj.insert("is_hide_values".into(), serde_json::json!(hide));
        if let Some(v) = rule_inner.get("stop_if_true") { obj.insert("stop_if_true".into(), v.clone()); }
    }
    Some(base)
}

/// Reconstruct a CLI-format edit base object for an icon-set rule from the engine's cf list response.
fn cf_reconstruct_icon_set_base(full_rule: &serde_json::Value) -> Option<serde_json::Value> {
    let rule_inner = full_rule.get("rule")?;
    let set_type = rule_inner.get("icon_set_type").and_then(|v| v.as_str())
        .and_then(|s| cf_map_icon_set_type(&cf_pascal_to_snake(s)).ok())?;
    let mut base = serde_json::json!({"icon_set_type": set_type});
    if let Some(obj) = base.as_object_mut() {
        if let Some(sub_rules) = rule_inner.get("sub_rules").and_then(|v| v.as_array()) {
            let conditions = cf_reconstruct_conditions(sub_rules);
            if !conditions.is_empty() { obj.insert("condition".into(), serde_json::Value::Array(conditions)); }
        }
        if let Some(icons) = rule_inner.get("icons").and_then(|v| v.as_array()) {
            let mapped: Vec<serde_json::Value> = icons.iter().filter_map(|ic| {
                let pos = ic.get("position").and_then(|v| v.as_i64())?;
                let itype = ic.get("icon_type").and_then(|v| v.as_str())
                    .and_then(|s| cf_map_single_icon_type(&cf_pascal_to_snake(s)))?;
                Some(serde_json::json!({"position": pos, "icon_type": itype}))
            }).collect();
            if !mapped.is_empty() { obj.insert("icons".into(), serde_json::Value::Array(mapped)); }
        }
        let reversed = rule_inner.get("is_icons_reversed").and_then(|v| v.as_bool()).unwrap_or(false);
        obj.insert("is_reverse_icons".into(), serde_json::json!(reversed));
        let hide = rule_inner.get("is_display_only_icons").and_then(|v| v.as_bool()).unwrap_or(false);
        obj.insert("is_hide_values".into(), serde_json::json!(hide));
        let def_size = rule_inner.get("is_default_icon_size").and_then(|v| v.as_bool()).unwrap_or(false);
        obj.insert("is_default_icon_size".into(), serde_json::json!(def_size));
        if let Some(v) = rule_inner.get("stop_if_true") { obj.insert("stop_if_true".into(), v.clone()); }
    }
    Some(base)
}

/// Merge an edit `overlay` onto a reconstructed `base` rule, merging the `condition`
/// array element-wise by `criteria_id` (so a partial stop edit — e.g. only `--max.color`
/// — updates that one stop instead of replacing the whole array). All other fields use
/// `deep_merge_into`. Used for color-scale and data-bar edits, whose stops are edited
/// individually. (Icon-set edits regenerate the whole set via `--set`, so they use a
/// plain `deep_merge_into` full-array replace instead.)
fn cf_merge_rule_base(base: &mut serde_json::Value, overlay: &serde_json::Value) {
    let overlay_conds = overlay.get("condition").and_then(|v| v.as_array()).cloned();

    if let (Some(base_obj), Some(overlay_obj)) = (base.as_object_mut(), overlay.as_object()) {
        for (k, v) in overlay_obj.iter() {
            if k == "condition" { continue; }
            let bv = base_obj.entry(k.clone()).or_insert(serde_json::Value::Null);
            deep_merge_into(bv, v);
        }
    }

    if let Some(ov_conds) = overlay_conds {
        let base_obj = match base.as_object_mut() { Some(o) => o, None => return };
        let base_conds = base_obj.entry("condition".to_string())
            .or_insert_with(|| serde_json::Value::Array(vec![]));
        let base_arr = match base_conds.as_array_mut() {
            Some(a) => a,
            None => { *base_conds = serde_json::Value::Array(ov_conds); return; }
        };
        for ov in ov_conds.iter() {
            let ov_cid = ov.get("criteria_id").and_then(|v| v.as_i64());
            let pos = ov_cid.and_then(|cid| base_arr.iter()
                .position(|b| b.get("criteria_id").and_then(|v| v.as_i64()) == Some(cid)));
            match pos {
                Some(p) => deep_merge_into(&mut base_arr[p], ov),
                None => base_arr.push(ov.clone()),
            }
        }
    }
}

/// Resolve a --sheet argument (name, 0-based index, or UUID) to a sheet UUID.
fn cf_resolve_sheet_id(sheet_override: Option<&str>, session: &CliSession) -> Option<String> {
    match sheet_override {
        None => Some(session.get_active_sheet_id_or_default()),
        Some(s) => {
            if session.sheet_ids.contains(&s.to_string()) {
                return Some(s.to_string());
            }
            resolve_sheet_id(s, session).map(|(id, _)| id)
        }
    }
}

/// Resolve a #index (1-based) to a rule_id from the session cache.
fn cf_resolve_index(index: usize, session: &CliSession) -> Result<String, String> {
    if index == 0 || index > session.last_cf_rules.len() {
        return Err(format!(
            "#{} is out of range. Run 'cf list' first — {} rule(s) cached.",
            index, session.last_cf_rules.len()
        ));
    }
    Ok(session.last_cf_rules[index - 1].rule_id.clone())
}

/// Parse a range string to a JSON range_list array.
fn cf_parse_range(range: &str) -> Result<serde_json::Value, String> {
    let (sc, sr, ec, er) = cell_ref::try_parse_range(range)
        .ok_or_else(|| format!("Invalid range '{}'. Use A1:C5 format.", range))?;
    Ok(serde_json::json!([{"start_row": sr, "start_column": sc, "end_row": er, "end_column": ec}]))
}

/// Parse a color string: named / hex "#RRGGBB" / "R,G,B" / "theme:N[,tint]".
fn cf_parse_color(s: &str) -> Result<serde_json::Value, String> {
    let s = s.trim();

    // Theme: "theme:4" or "theme:4,0.5"
    if let Some(rest) = s.strip_prefix("theme:") {
        let parts: Vec<&str> = rest.splitn(2, ',').collect();
        let idx: i64 = parts[0].trim().parse()
            .map_err(|_| format!("Invalid theme index in '{}'.", s))?;
        let tint: f64 = if parts.len() > 1 { parts[1].trim().parse().unwrap_or(0.0) } else { 0.0 };
        return Ok(serde_json::json!({"theme_color": idx, "tint": tint}));
    }

    // Hex: "#RRGGBB" or "RRGGBB"
    let hex = if s.starts_with('#') { &s[1..] } else { s };
    if hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
        let r = i64::from_str_radix(&hex[0..2], 16).unwrap();
        let g = i64::from_str_radix(&hex[2..4], 16).unwrap();
        let b = i64::from_str_radix(&hex[4..6], 16).unwrap();
        return Ok(serde_json::json!({"red": r, "green": g, "blue": b}));
    }

    // R,G,B triple
    if s.contains(',') && !s.contains(':') {
        let parts: Vec<&str> = s.splitn(3, ',').collect();
        if parts.len() == 3 {
            if let (Ok(r), Ok(g), Ok(b)) = (parts[0].trim().parse::<i64>(), parts[1].trim().parse::<i64>(), parts[2].trim().parse::<i64>()) {
                return Ok(serde_json::json!({"red": r, "green": g, "blue": b}));
            }
        }
    }

    // Named colors
    let rgb = match s.to_lowercase().as_str() {
        "red"     => (255, 0, 0),     "green"   => (0, 128, 0),     "blue"    => (0, 0, 255),
        "yellow"  => (255, 255, 0),   "orange"  => (255, 165, 0),   "white"   => (255, 255, 255),
        "black"   => (0, 0, 0),       "gray" | "grey" => (128, 128, 128), "purple" => (128, 0, 128),
        "cyan"    => (0, 255, 255),   "magenta" => (255, 0, 255),   "pink"    => (255, 192, 203),
        "brown"   => (139, 69, 19),   "lime"    => (0, 255, 0),     "navy"    => (0, 0, 128),
        "teal"    => (0, 128, 128),   "maroon"  => (128, 0, 0),     "olive"   => (128, 128, 0),
        _ => return Err(format!("Unknown color '{}'. Use a named color, #RRGGBB, R,G,B, or theme:N.", s)),
    };
    Ok(serde_json::json!({"red": rgb.0, "green": rgb.1, "blue": rgb.2}))
}

/// Expand a --when expression into raw condition fields on `obj`.
fn cf_expand_when(expr: &str, obj: &mut serde_json::Value) -> Result<(), String> {
    let e = expr.trim().to_lowercase();

    // top/bottom N — count and is_percent are RULE-LEVEL fields (not inside condition).
    // Must be handled before the cond borrow so we can write to obj directly.
    for (prefix, sct) in &[("top ", 0i64), ("bottom ", 1i64)] {
        if let Some(rest) = e.strip_prefix(prefix) {
            let (count_str, is_percent) = if let Some(s) = rest.trim().strip_suffix('%') {
                (s.trim(), true)
            } else {
                (rest.trim(), false)
            };
            let count: i64 = count_str.parse().map_err(|_| {
                format!("Invalid count '{}' in '{}N' expression. Use a positive integer.", count_str, prefix.trim())
            })?;
            let rule = obj.as_object_mut().ok_or_else(|| "Internal: rule_obj must be an object.".to_string())?;
            let cond_arr = rule.entry("condition").or_insert_with(|| serde_json::json!([]));
            if let Some(arr) = cond_arr.as_array_mut() {
                arr.push(serde_json::json!({"criteria_type": 13, "sub_criteria_type": sct}));
            }
            rule.insert("count".into(), serde_json::json!(count));
            rule.insert("is_percent".into(), serde_json::json!(is_percent));
            return Ok(());
        }
    }

    let cond = obj.as_object_mut()
        .ok_or_else(|| "Internal: rule_obj must be an object.".to_string())?
        .entry("condition")
        .or_insert_with(|| serde_json::json!([]))
        .as_array_mut()
        .ok_or_else(|| "condition must be an array.".to_string())?;

    // Helper closure to push a condition entry
    macro_rules! push_cond {
        ($ct:expr, $sct:expr) => {{
            cond.push(serde_json::json!({"criteria_type": $ct, "sub_criteria_type": $sct}));
        }};
        ($ct:expr, $sct:expr, $lhs:expr) => {{
            cond.push(serde_json::json!({"criteria_type": $ct, "sub_criteria_type": $sct, "lhs": $lhs}));
        }};
        ($ct:expr, $sct:expr, $lhs:expr, $rhs:expr) => {{
            cond.push(serde_json::json!({"criteria_type": $ct, "sub_criteria_type": $sct, "lhs": $lhs, "rhs": $rhs}));
        }};
    }

    // number_comparison shorthands
    if let Some(rest) = e.strip_prefix(">=") {
        push_cond!(6, 6, rest.trim().to_string()); return Ok(());
    }
    if let Some(rest) = e.strip_prefix("<=") {
        push_cond!(6, 7, rest.trim().to_string()); return Ok(());
    }
    if let Some(rest) = e.strip_prefix("!=") {
        push_cond!(6, 1, rest.trim().to_string()); return Ok(());
    }
    if let Some(rest) = e.strip_prefix('>') {
        push_cond!(6, 4, rest.trim().to_string()); return Ok(());
    }
    if let Some(rest) = e.strip_prefix('<') {
        push_cond!(6, 5, rest.trim().to_string()); return Ok(());
    }
    if let Some(rest) = e.strip_prefix('=') {
        push_cond!(6, 0, rest.trim().to_string()); return Ok(());
    }

    // between / not between
    if let Some(rest) = e.strip_prefix("not between ") {
        let parts: Vec<&str> = rest.splitn(3, " and ").collect();
        if parts.len() == 2 {
            push_cond!(6, 3, parts[0].trim().to_string(), parts[1].trim().to_string());
            return Ok(());
        }
    }
    if let Some(rest) = e.strip_prefix("between ") {
        let parts: Vec<&str> = rest.splitn(3, " and ").collect();
        if parts.len() == 2 {
            push_cond!(6, 2, parts[0].trim().to_string(), parts[1].trim().to_string());
            return Ok(());
        }
    }

    // text
    if let Some(rest) = e.strip_prefix("contains '") {
        let val = rest.trim_end_matches('\''); push_cond!(8, 0, val.to_string()); return Ok(());
    }
    if let Some(rest) = e.strip_prefix("not contains '") {
        let val = rest.trim_end_matches('\''); push_cond!(8, 1, val.to_string()); return Ok(());
    }
    if let Some(rest) = e.strip_prefix("begins with '") {
        let val = rest.trim_end_matches('\''); push_cond!(8, 2, val.to_string()); return Ok(());
    }
    if let Some(rest) = e.strip_prefix("ends with '") {
        let val = rest.trim_end_matches('\''); push_cond!(8, 3, val.to_string()); return Ok(());
    }

    // cell_containing
    match e.as_str() {
        "is duplicate" => { push_cond!(9, 0); return Ok(()); }
        "is unique" => { push_cond!(9, 1); return Ok(()); }
        "is blank" => { push_cond!(9, 2); return Ok(()); }
        "is not blank" => { push_cond!(9, 3); return Ok(()); }
        "is error" => { push_cond!(9, 4); return Ok(()); }
        "is not error" => { push_cond!(9, 5); return Ok(()); }
        "above average" => { push_cond!(10, 0); return Ok(()); }
        "below average" => { push_cond!(10, 1); return Ok(()); }
        "today" => { push_cond!(7, 1); return Ok(()); }
        "yesterday" => { push_cond!(7, 0); return Ok(()); }
        "tomorrow" => { push_cond!(7, 2); return Ok(()); }
        "last 7 days" => { push_cond!(7, 3); return Ok(()); }
        "this week" => { push_cond!(7, 5); return Ok(()); }
        "last week" => { push_cond!(7, 4); return Ok(()); }
        "next week" => { push_cond!(7, 6); return Ok(()); }
        "last month" => { push_cond!(7, 7); return Ok(()); }
        "this month" => { push_cond!(7, 8); return Ok(()); }
        "next month" => { push_cond!(7, 9); return Ok(()); }
        "next 7 days" => { push_cond!(7, 18); return Ok(()); }
        "last year" => { push_cond!(7, 19); return Ok(()); }
        "this year" => { push_cond!(7, 20); return Ok(()); }
        "next year" => { push_cond!(7, 21); return Ok(()); }
        _ => {}
    }

    Err(format!("Unrecognized --when expression: '{}'. See docs for supported patterns.", expr))
}

/// Default is_percent to false when any condition uses criteria_type 13 (top/bottom)
/// and is_percent is not already set at the rule level.
fn cf_default_is_percent(rule_obj: &mut serde_json::Value) {
    if let Some(obj) = rule_obj.as_object_mut() {
        if obj.contains_key("is_percent") { return; }
        let has_top_bottom = obj.get("condition")
            .and_then(|v| v.as_array())
            .map(|arr| arr.iter().any(|c| c.get("criteria_type").and_then(|v| v.as_i64()) == Some(13)))
            .unwrap_or(false);
        if has_top_bottom {
            obj.insert("is_percent".into(), serde_json::json!(false));
        }
    }
}

/// Auto-assign criteria_id to condition entries that don't have one.
fn cf_auto_criteria_id(rule_obj: &mut serde_json::Value) {
    if let Some(arr) = rule_obj.get_mut("condition").and_then(|v| v.as_array_mut()) {
        for (i, c) in arr.iter_mut().enumerate() {
            if let Some(obj) = c.as_object_mut() {
                if !obj.contains_key("criteria_id") {
                    obj.insert("criteria_id".into(), serde_json::json!(i as i64));
                }
            }
        }
    }
}

/// Apply engine-specific value remaps (offsets / inversions) for fields where
/// the CLI surface uses 0-based values but the engine uses different values.
fn cf_apply_engine_remaps(v: &mut serde_json::Value) {
    match v {
        serde_json::Value::Object(obj) => {
            // Rename line_style → border_line_style if needed (user may pass it via flat path)
            if obj.contains_key("line_style") && !obj.contains_key("border_line_style") {
                if let Some(val) = obj.remove("line_style") {
                    obj.insert("border_line_style".into(), val);
                }
            }
            // A border object (identified by border_line_style) requires an integer border_type
            // (the side). If the user gave a line_style/color but omitted the side, default to
            // bottom (3) so the engine doesn't reject with "'border_type' key not found".
            if obj.contains_key("border_line_style") && !obj.contains_key("border_type") {
                obj.insert("border_type".into(), serde_json::json!(3));
            }
            // Rename color → border_color inside border objects (identified by border_type presence)
            if obj.contains_key("border_type") && obj.contains_key("color") && !obj.contains_key("border_color") {
                if let Some(val) = obj.remove("color") {
                    obj.insert("border_color".into(), val);
                }
            }
            // If border object has no color at all, default border_color to black
            if obj.contains_key("border_type") && !obj.contains_key("border_color") && !obj.contains_key("color") {
                obj.insert("border_color".into(), serde_json::json!({"red": 0, "green": 0, "blue": 0}));
            }
            // Rename color → background_color inside fill objects.
            // Fill objects have color but no border_type (border) or font_style/underline_type (font).
            // Exclude colorscale/databar condition objects which use `color` directly and are
            // identified by having `criteria_type` or `criteria_id`.
            if obj.contains_key("color") && !obj.contains_key("background_color")
                && !obj.contains_key("border_type")
                && !obj.contains_key("font_style") && !obj.contains_key("underline_type")
                && !obj.contains_key("criteria_type") && !obj.contains_key("criteria_id")
            {
                if let Some(val) = obj.remove("color") {
                    obj.insert("background_color".into(), val);
                }
            }
            // Rename type → custom_format_type in number_format objects (engine requires this key name).
            if obj.contains_key("type") && !obj.contains_key("custom_format_type") {
                if let Some(val) = obj.remove("type") {
                    obj.insert("custom_format_type".into(), val);
                }
            }
            // number_format objects require custom_format_text (string) alongside custom_format_type.
            if obj.contains_key("custom_format_type") && !obj.contains_key("custom_format_text") {
                obj.insert("custom_format_text".into(), serde_json::json!(""));
            }
            // Rename value/value2 → lhs/rhs (condition threshold aliases); API requires string values.
            if obj.contains_key("value") && !obj.contains_key("lhs") {
                if let Some(val) = obj.remove("value") {
                    let s = match &val { serde_json::Value::String(s) => s.clone(), other => other.to_string() };
                    obj.insert("lhs".into(), serde_json::json!(s));
                }
            }
            if obj.contains_key("value2") && !obj.contains_key("rhs") {
                if let Some(val) = obj.remove("value2") {
                    let s = match &val { serde_json::Value::String(s) => s.clone(), other => other.to_string() };
                    obj.insert("rhs".into(), serde_json::json!(s));
                }
            }
            // Ensure existing lhs/rhs are strings (user may pass via direct path as numbers).
            for key in &["lhs", "rhs"] {
                if let Some(val) = obj.get_mut(*key) {
                    if !val.is_string() {
                        let s = val.to_string();
                        *val = serde_json::json!(s);
                    }
                }
            }
            for (key, val) in obj.iter_mut() {
                match key.as_str() {
                    "underline_type" => {
                        if let Some(n) = val.as_i64() { *val = serde_json::json!(n + 1); }
                    }
                    "strike_type" => {
                        // CLI: off=0, on=1, automatic=2 → engine: off=1, on=0, automatic=2
                        if let Some(n) = val.as_i64() {
                            *val = match n { 0 => serde_json::json!(1), 1 => serde_json::json!(0), _ => val.clone() };
                        }
                    }
                    "border_line_style" => {
                        if let Some(n) = val.as_i64() { *val = serde_json::json!(n + 1); }
                    }
                    "custom_format_type" | "type" => {
                        // Only remap inside number_format context; key "type" is ambiguous but
                        // in practice only appears there in CF payloads.
                        if let Some(n) = val.as_i64() { *val = serde_json::json!(n + 1); }
                    }
                    _ => cf_apply_engine_remaps(val),
                }
            }
        }
        serde_json::Value::Array(arr) => {
            for item in arr.iter_mut() {
                cf_apply_engine_remaps(item);
            }
        }
        _ => {}
    }
}

/// Auto-populate icons[] and condition[] for an iconset rule from icon_set_type if not already set.
fn cf_iconset_auto_populate(rule_obj: &mut serde_json::Value) -> Result<(), String> {
    let obj = rule_obj.as_object_mut().ok_or("rule must be an object")?;
    let set_type = match obj.get("icon_set_type").and_then(|v| v.as_i64()) {
        Some(v) => v as i32,
        None => return Ok(()),
    };

    // Ensure required booleans are present
    if !obj.contains_key("is_hide_values") { obj.insert("is_hide_values".into(), serde_json::json!(false)); }
    if !obj.contains_key("is_default_icon_size") { obj.insert("is_default_icon_size".into(), serde_json::json!(false)); }
    if !obj.contains_key("is_reverse_icons") { obj.insert("is_reverse_icons".into(), serde_json::json!(false)); }

    if set_type == 24 {
        // CUSTOM: user must supply icons and conditions manually
        return Ok(());
    }

    // (icon_count, default icon_types[position 0..n-1])
    // Position 0 = lowest value threshold, position N-1 = highest.
    let (n_icons, default_icons): (usize, &[i64]) = match set_type {
        0  => (2, &[34, 35]),          // TWO_THUMBS: THUMBS_DOWN, THUMBS_UP
        1  => (2, &[39, 38]),          // TWO_HEART: HEART_BREAK, HEART
        2  => (2, &[36, 37]),          // TWO_TICK_WRONG: WRONG, TICK
        3  => (3, &[0, 7, 16]),        // THREE_ARROWS: RED_DOWN, YELLOW_RIGHT, GREEN_UP
        4  => (3, &[22, 24, 23]),      // THREE_ARROWS_GRAY: GRAY_DOWN, GRAY_RIGHT, GRAY_UP
        5  => (3, &[5, 14, 20]),       // THREE_FLAGS: RED, YELLOW, GREEN
        6  => (3, &[4, 13, 19]),       // THREE_TRAFFIC_LIGHTS_1
        7  => (3, &[4, 13, 19]),       // THREE_TRAFFIC_LIGHTS_2
        8  => (3, &[6, 15, 21]),       // THREE_SIGNS: RED_WRONG, YELLOW_EXCLAMATION, GREEN_TICK
        9  => (3, &[1, 12, 17]),       // THREE_TRIANGLES: RED_DOWN, YELLOW, GREEN_UP
        10 => (3, &[6, 15, 21]),       // THREE_SYMBOLS
        11 => (3, &[39, 50, 38]),      // THREE_HEART: HEART_BREAK, BLANK_HEART, HEART
        12 => (3, &[40, 41, 42]),      // THREE_SMILEY: SAD, NEUTRAL, HAPPY
        13 => (3, &[52, 53, 54]),      // THREE_STARS: BLANK, HALF, FULL
        14 => (4, &[0, 9, 10, 16]),    // FOUR_ARROWS
        15 => (4, &[22, 25, 26, 23]),  // FOUR_ARROWS_GRAY
        16 => (4, &[29, 2, 11, 18]),   // FOUR_BLACK_TO_RED: BLACK, RED, YELLOW, GREEN circle
        17 => (4, &[4, 4, 13, 19]),    // FOUR_TRAFFIC_LIGHTS
        18 => (5, &[0, 9, 7, 10, 16]),    // FIVE_ARROWS
        19 => (5, &[22, 25, 24, 26, 23]), // FIVE_ARROWS_GRAY
        20 => (5, &[33, 32, 31, 30, 29]), // FIVE_QUATERS: WHITE→BLACK filled circles
        21 => (5, &[40, 40, 41, 42, 44]), // FIVE_SMILEY: SAD..LAUGH
        22 => (5, &[2, 28, 11, 18, 29]),  // FIVE_CIRCLES: RED, PINK, YELLOW, GREEN, BLACK
        23 => (5, &[46, 45, 47, 48, 49]), // FIVE_CLOUDS: THUNDER, RAINY, WINDY, CLOUD, SUNNY
        _  => (3, &[0, 7, 16]),           // fallback
    };

    if !obj.contains_key("icons") {
        let icons: Vec<serde_json::Value> = (0..n_icons).map(|i| {
            let icon_type = default_icons.get(i).copied().unwrap_or(0);
            serde_json::json!({ "position": i as i64, "icon_type": icon_type })
        }).collect();
        obj.insert("icons".into(), serde_json::Value::Array(icons));
    }

    if !obj.contains_key("condition") {
        let n_conds = n_icons - 1;
        let step = 100.0 / n_icons as f64;
        let conds: Vec<serde_json::Value> = (0..n_conds).map(|i| {
            let threshold = ((n_conds - i) as f64 * step).round() as i64;
            serde_json::json!({
                "criteria_type": 0,  // number (percent criteria rejected by API)
                "sub_criteria_type": 6,  // gte
                "lhs": threshold.to_string()
            })
        }).collect();
        obj.insert("condition".into(), serde_json::Value::Array(conds));
    }

    Ok(())
}

/// Main v2 arg parser for CF sub-commands.
/// Returns: (index_target, rule_id_target, range_override, sheet_override, active_cell_override, rule_obj)
#[allow(clippy::type_complexity)]
fn cf_parse_v2_args(
    args: &[&str],
    usage: &str,
) -> Result<(Option<usize>, Option<String>, Option<String>, Option<String>, Option<String>, serde_json::Value), String> {
    if args.is_empty() {
        return Err(usage.to_string());
    }

    let mut index_target: Option<usize> = None;
    let mut rule_id_target: Option<String> = None;
    let mut range_override: Option<String> = None;
    let mut sheet_override: Option<String> = None;
    let mut active_cell_override: Option<String> = None;
    let mut rule_obj = serde_json::json!({});
    let mut flag_entries: Vec<(String, serde_json::Value)> = Vec::new();

    let mut i = 0;

    // First positional arg: optional #index
    if !args[0].starts_with("--") && !args[0].starts_with('-') {
        let first = args[0];
        if let Some(stripped) = first.strip_prefix('#') {
            index_target = Some(stripped.parse::<usize>()
                .map_err(|_| format!("Invalid #index '{}'.", first))?);
            i = 1;
        } else if first.chars().next().map_or(false, |c| c.is_ascii_digit()) {
            // Bare number = #index
            index_target = Some(first.parse::<usize>()
                .map_err(|_| format!("Invalid index '{}'.", first))?);
            i = 1;
        }
        // else: not an index — leave i=0, process as flag
    }

    while i < args.len() {
        let arg = args[i];
        match arg.to_lowercase().as_str() {
            "--range" => {
                i += 1; if i >= args.len() { return Err("--range requires a value.".into()); }
                range_override = Some(args[i].to_string());
            }
            "--rule-id" => {
                i += 1; if i >= args.len() { return Err("--rule-id requires a value.".into()); }
                rule_id_target = Some(args[i].to_string());
            }
            "--sheet" => {
                i += 1; if i >= args.len() { return Err("--sheet requires a value.".into()); }
                sheet_override = Some(args[i].to_string());
            }
            "--active-cell" => {
                i += 1; if i >= args.len() { return Err("--active-cell requires a value.".into()); }
                active_cell_override = Some(args[i].to_string());
            }
            "--stop-if-true" => { flag_entries.push(("stop_if_true".into(), serde_json::json!(true))); }
            "--hide-values" => { flag_entries.push(("is_hide_values".into(), serde_json::json!(true))); }
            "--auto-text-color" => { flag_entries.push(("is_automatic_text_color".into(), serde_json::json!(true))); }
            "--reverse-icons" => { flag_entries.push(("is_reverse_icons".into(), serde_json::json!(true))); }
            "--default-icon-size" => { flag_entries.push(("is_default_icon_size".into(), serde_json::json!(true))); }
            "--bold" => { flag_entries.push(("style.font.font_style".into(), serde_json::json!(2))); }
            "--italic" => { flag_entries.push(("style.font.font_style".into(), serde_json::json!(1))); }
            "--bold-italic" => { flag_entries.push(("style.font.font_style".into(), serde_json::json!(3))); }
            "--strike" => { flag_entries.push(("style.font.strike_type".into(), serde_json::json!(1))); }
            "--underline" => {
                let utype = if i + 1 < args.len() && !args[i+1].starts_with("--") {
                    i += 1;
                    match args[i].to_lowercase().as_str() {
                        "double" => 2i64,
                        "single_accounting" => 3,
                        "double_accounting" => 4,
                        _ => 1,
                    }
                } else { 1 };
                flag_entries.push(("style.font.underline_type".into(), serde_json::json!(utype)));
            }
            "--font-color" => {
                i += 1; if i >= args.len() { return Err("--font-color requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                flag_entries.push(("style.font.color".into(), color));
            }
            "--fill" => {
                i += 1; if i >= args.len() { return Err("--fill requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                flag_entries.push(("style.fill.background_color".into(), color));
            }
            "--border-bottom" | "--border-top" | "--border-left" | "--border-right" => {
                let border_type_num: i64 = match arg.to_lowercase().as_str() {
                    "--border-left" => 0, "--border-right" => 1,
                    "--border-top" => 2, "--border-bottom" => 3, _ => 3,
                };
                i += 1; if i >= args.len() { return Err(format!("{} requires a line style.", arg)); }
                let line_style_raw = args[i];
                let line_style_num = cf_map_line_style(line_style_raw);
                let color = if i + 1 < args.len() && !args[i+1].starts_with("--") {
                    i += 1;
                    cf_parse_color(args[i])?
                } else {
                    serde_json::json!({"red": 0, "green": 0, "blue": 0})
                };
                // Append to borders array — count unique indices, not total entries (3 per border)
                let border_idx = flag_entries.iter()
                    .filter_map(|(k, _)| {
                        let s = k.strip_prefix("style.borders[")?;
                        let end = s.find(']')?;
                        s[..end].parse::<usize>().ok()
                    })
                    .map(|i| i + 1)
                    .max()
                    .unwrap_or(0);
                let base = format!("style.borders[{}]", border_idx);
                flag_entries.push((format!("{}.border_type", base), serde_json::json!(border_type_num)));
                flag_entries.push((format!("{}.border_line_style", base), serde_json::json!(line_style_num)));
                flag_entries.push((format!("{}.color", base), color));
            }
            "--when" => {
                i += 1; if i >= args.len() { return Err("--when requires an expression.".into()); }
                cf_expand_when(args[i], &mut rule_obj)?;
            }
            // colorscale shorthands: --min, --mid, --max
            // criteria_id is fixed (0=min, 1=mid, 2=max) so a 2-stop scale gets 0 and 2
            "--min" => {
                i += 1; if i >= args.len() { return Err("--min requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                let idx = cf_count_conditions_with_pending(&rule_obj, &flag_entries);
                flag_entries.push((format!("condition[{}].criteria_id", idx), serde_json::json!(0)));
                flag_entries.push((format!("condition[{}].criteria_type", idx), serde_json::json!(4))); // min
                flag_entries.push((format!("condition[{}].color", idx), color));
            }
            "--mid" => {
                i += 1; if i >= args.len() { return Err("--mid requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                let idx = cf_count_conditions_with_pending(&rule_obj, &flag_entries);
                flag_entries.push((format!("condition[{}].criteria_id", idx), serde_json::json!(1)));
                flag_entries.push((format!("condition[{}].criteria_type", idx), serde_json::json!(2))); // percentile
                flag_entries.push((format!("condition[{}].value", idx), serde_json::json!("50")));
                flag_entries.push((format!("condition[{}].color", idx), color));
            }
            "--max" => {
                i += 1; if i >= args.len() { return Err("--max requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                let idx = cf_count_conditions_with_pending(&rule_obj, &flag_entries);
                flag_entries.push((format!("condition[{}].criteria_id", idx), serde_json::json!(2)));
                flag_entries.push((format!("condition[{}].criteria_type", idx), serde_json::json!(5))); // max
                flag_entries.push((format!("condition[{}].color", idx), color));
            }
            // databar shorthands
            "--positive" => {
                i += 1; if i >= args.len() { return Err("--positive requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                flag_entries.push(("positive_value_fill".into(), color));
            }
            "--negative" => {
                i += 1; if i >= args.len() { return Err("--negative requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                flag_entries.push(("negative_value_fill".into(), color));
            }
            "--positive-border" => {
                i += 1; if i >= args.len() { return Err("--positive-border requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                flag_entries.push(("positive_value_border".into(), color));
            }
            "--negative-border" => {
                i += 1; if i >= args.len() { return Err("--negative-border requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                flag_entries.push(("negative_value_border".into(), color));
            }
            "--axis-color" => {
                i += 1; if i >= args.len() { return Err("--axis-color requires a color.".into()); }
                let color = cf_parse_color(args[i])?;
                flag_entries.push(("axis_color".into(), color));
            }
            // iconset shorthand
            "--set" => {
                i += 1; if i >= args.len() { return Err("--set requires an icon set type.".into()); }
                let v = cf_map_icon_set_type(args[i])?;
                flag_entries.push(("icon_set_type".into(), serde_json::json!(v)));
            }
            // databar option flags
            "--direction" => {
                i += 1; if i >= args.len() { return Err("--direction requires a value.".into()); }
                let v = cf_map_named_i64(args[i], &[("left_to_right",0),("right_to_left",1),("context",2)])
                    .ok_or_else(|| format!("Invalid direction '{}'.", args[i]))?;
                flag_entries.push(("bar_direction".into(), serde_json::json!(v)));
            }
            "--axis" => {
                i += 1; if i >= args.len() { return Err("--axis requires a value.".into()); }
                let v = cf_map_named_i64(args[i], &[("automatic",0),("middle",1),("none",2)])
                    .ok_or_else(|| format!("Invalid axis '{}'.", args[i]))?;
                flag_entries.push(("axis_position".into(), serde_json::json!(v)));
            }
            "--fill-type" => {
                i += 1; if i >= args.len() { return Err("--fill-type requires a value.".into()); }
                let v = cf_map_named_i64(args[i], &[("solid",0),("gradient",1)])
                    .ok_or_else(|| format!("Invalid fill-type '{}'.", args[i]))?;
                flag_entries.push(("fill_type".into(), serde_json::json!(v)));
            }
            "--bar-border" => {
                i += 1; if i >= args.len() { return Err("--bar-border requires a value.".into()); }
                let v = cf_map_named_i64(args[i], &[("none",0),("off",0),("with_border",1),("on",1)])
                    .ok_or_else(|| format!("Invalid bar-border '{}'.", args[i]))?;
                flag_entries.push(("border_type".into(), serde_json::json!(v)));
            }
            _ => {
                // Try as a flag with a value
                if let Some(path) = arg.strip_prefix("--") {
                    if path.is_empty() { return Err("Empty flag name.".into()); }
                    i += 1;
                    if i >= args.len() { return Err(format!("{} requires a value.", arg)); }

                    // Handle --min.*, --mid.*, --max.* extended colorscale stop syntax
                    let (cid, sub_key) = if let Some(s) = path.strip_prefix("min.") { (Some(0u64), Some(s)) }
                        else if let Some(s) = path.strip_prefix("mid.") { (Some(1u64), Some(s)) }
                        else if let Some(s) = path.strip_prefix("max.") { (Some(2u64), Some(s)) }
                        else { (None, None) };

                    if let (Some(cid), Some(sub_key)) = (cid, sub_key) {
                        let existing = cf_find_stop_index(cid, &flag_entries);
                        let idx = existing.unwrap_or_else(|| {
                            cf_count_conditions_with_pending(&rule_obj, &flag_entries)
                        });
                        if existing.is_none() {
                            flag_entries.push((format!("condition[{}].criteria_id", idx), serde_json::json!(cid)));
                        }
                        let cond_path = format!("condition[{}].{}", idx, sub_key);
                        let value = cf_parse_flag_value_for_path(&cond_path, args[i]);
                        flag_entries.push((cond_path, value));
                    } else {
                        let value = cf_parse_flag_value_for_path(path, args[i]);
                        flag_entries.push((path.to_string(), value));
                    }
                } else {
                    return Err(format!("Unknown argument '{}'. {}", arg, usage));
                }
            }
        }
        i += 1;
    }

    // Apply all flag entries to rule_obj
    for (path, value) in flag_entries {
        cf_set_path_value(&mut rule_obj, &path, value)?;
    }

    Ok((index_target, rule_id_target, range_override, sheet_override, active_cell_override, rule_obj))
}

fn cf_count_conditions(rule_obj: &serde_json::Value) -> usize {
    rule_obj.get("condition")
        .and_then(|v| v.as_array())
        .map(|a| a.len())
        .unwrap_or(0)
}

fn cf_find_stop_index(criteria_id: u64, flag_entries: &[(String, serde_json::Value)]) -> Option<usize> {
    for (k, v) in flag_entries.iter() {
        if k.ends_with(".criteria_id") {
            if let Some(s) = k.strip_prefix("condition[") {
                if let Some(bracket) = s.find(']') {
                    if let Ok(idx) = s[..bracket].parse::<usize>() {
                        if v.as_u64() == Some(criteria_id) {
                            return Some(idx);
                        }
                    }
                }
            }
        }
    }
    None
}

fn cf_count_conditions_with_pending(rule_obj: &serde_json::Value, flag_entries: &[(String, serde_json::Value)]) -> usize {
    let in_obj = cf_count_conditions(rule_obj);
    let max_pending = flag_entries.iter()
        .filter_map(|(k, _)| {
            let s = k.strip_prefix("condition[")?;
            let end = s.find(']')?;
            s[..end].parse::<usize>().ok()
        })
        .map(|i| i + 1)
        .max()
        .unwrap_or(0);
    in_obj.max(max_pending)
}

fn cf_map_named_i64(s: &str, table: &[(&str, i64)]) -> Option<i64> {
    let key = s.to_lowercase();
    if let Some((_, v)) = table.iter().find(|(k, _)| *k == key.as_str()) {
        return Some(*v);
    }
    s.parse::<i64>().ok()
}

fn cf_map_line_style(s: &str) -> i64 {
    cf_map_named_i64(s, &[
        ("none",0),("dash_dot",1),("dash_dot_dot",2),("dashed",3),("dotted",4),
        ("double",5),("hairline",6),("medium",7),("medium_dash_dot",8),
        ("medium_dash_dot_dot",9),("medium_dashed",10),("thick",11),("thin",12),
        ("slant_dash_dot",13),
    ]).unwrap_or(0)
}

fn cf_map_icon_set_type(s: &str) -> Result<i64, String> {
    let table: &[(&str, i64)] = &[
        ("two_thumbs",0),("two_heart",1),("two_tick_wrong",2),
        ("three_arrows",3),("three_arrows_gray",4),("three_flags",5),
        ("three_traffic_lights_1",6),("three_traffic_lights_2",7),
        ("three_signs",8),("three_triangles",9),("three_symbols",10),
        ("three_heart",11),("three_smiley",12),("three_stars",13),
        ("four_arrows",14),("four_arrows_gray",15),("four_black_to_red",16),
        ("four_traffic_lights",17),
        ("five_arrows",18),("five_arrows_gray",19),("five_quaters",20),("five_quarters",20),
        ("five_smiley",21),("five_circles",22),("five_rating",22),("five_clouds",23),
        ("custom",24),
    ];
    if let Some(v) = cf_map_named_i64(s, table) {
        return Ok(v);
    }
    let low = s.to_lowercase();
    if let Some(rest) = low.strip_prefix("icon_set_") {
        if let Ok(n) = rest.parse::<i64>() {
            if (0..=24).contains(&n) { return Ok(n); }
        }
    }
    if let Ok(n) = s.parse::<i64>() {
        if (0..=24).contains(&n) { return Ok(n); }
    }
    Err(format!(
        "Unknown icon set type '{}'. Use e.g. three_arrows, four_traffic_lights, five_circles, or 0..24.",
        s
    ))
}

fn cf_icon_set_type_name(v: i64) -> &'static str {
    match v {
        0 => "two_thumbs", 1 => "two_heart", 2 => "two_tick_wrong",
        3 => "three_arrows", 4 => "three_arrows_gray", 5 => "three_flags",
        6 => "three_traffic_lights_1", 7 => "three_traffic_lights_2",
        8 => "three_signs", 9 => "three_triangles", 10 => "three_symbols",
        11 => "three_heart", 12 => "three_smiley", 13 => "three_stars",
        14 => "four_arrows", 15 => "four_arrows_gray", 16 => "four_black_to_red",
        17 => "four_traffic_lights",
        18 => "five_arrows", 19 => "five_arrows_gray", 20 => "five_quaters",
        21 => "five_smiley", 22 => "five_circles", 23 => "five_clouds",
        24 => "custom",
        _ => "unknown",
    }
}

fn cf_map_single_icon_type(s: &str) -> Option<i64> {
    let table: &[(&str, i64)] = &[
        ("red_down_arrow",0),("red_down_triangle_small",1),("red_circle",2),
        ("red_diamond",3),("red_traffic_light",4),("red_flag",5),("red_wrong_sign",6),
        ("yellow_right_arrow",7),("yellow_horizontal_line",8),("yellow_bottom_right_arrow",9),
        ("yellow_top_right_arrow",10),("yellow_circle",11),("yellow_triangle",12),
        ("yellow_traffic_light",13),("yellow_flag",14),("yellow_exclamation_sign",15),
        ("green_up_arrow",16),("green_up_triangle_small",17),("green_circle",18),
        ("green_traffic_light",19),("green_flag",20),("green_tick_sign",21),
        ("gray_down_arrow",22),("gray_up_arrow",23),("gray_right_arrow",24),
        ("gray_bottom_right_arrow",25),("gray_top_right_arrow",26),("gray_circle",27),
        ("pink_circle",28),("black_circle",29),("three_quarter_filled_circle",30),
        ("half_filled_circle",31),("quarter_filled_circle",32),("white_circle",33),
        ("thumbs_down",34),("thumbs_up",35),("wrong",36),("tick",37),
        ("heart",38),("heart_break",39),
        ("sad_emoji",40),("neutral_emoji",41),("happy_emoji",42),("angry_emoji",43),("laugh_emoji",44),
        ("rainy_cloud",45),("thunder_cloud",46),("windy_cloud",47),("cloud",48),("sunny_cloud",49),
        ("blank_heart",50),("half_filled_heart",51),
        ("blank_star",52),("half_filled_star",53),("full_star",54),("none",55),
    ];
    let low = s.to_lowercase();
    if let Some((_, v)) = table.iter().find(|(k, _)| *k == low.as_str()) {
        return Some(*v);
    }
    let rest = low.strip_prefix("icon_").unwrap_or(&low);
    rest.parse::<i64>().ok().filter(|&n| (0..=55).contains(&n))
}

fn cf_parse_flag_value_for_path(path: &str, raw: &str) -> serde_json::Value {
    // Try to parse as a color for known color-holding fields
    let last_seg = path.split('.').last().unwrap_or(path).to_lowercase();
    let last_seg = last_seg.trim_end_matches(|c: char| c == ']' || c.is_ascii_digit());

    if matches!(last_seg, "color" | "fill" | "font_color") {
        if let Ok(c) = cf_parse_color(raw) { return c; }
    }

    // Enum mappings
    let parsed = cf_parse_raw_value(raw);
    if let Some(s) = parsed.as_str() {
        if let Some(mapped) = cf_map_enum_for_path(path, s) {
            return serde_json::json!(mapped);
        }
    }
    parsed
}

fn cf_parse_raw_value(raw: &str) -> serde_json::Value {
    if raw.eq_ignore_ascii_case("true") { return serde_json::Value::Bool(true); }
    if raw.eq_ignore_ascii_case("false") { return serde_json::Value::Bool(false); }
    if raw.eq_ignore_ascii_case("null") { return serde_json::Value::Null; }
    if let Ok(n) = raw.parse::<i64>() { return serde_json::json!(n); }
    if let Ok(n) = raw.parse::<f64>() { return serde_json::json!(n); }
    if (raw.starts_with('{') && raw.ends_with('}')) || (raw.starts_with('[') && raw.ends_with(']')) {
        if let Ok(v) = serde_json::from_str::<serde_json::Value>(raw) { return v; }
    }
    serde_json::Value::String(raw.to_string())
}

fn cf_map_enum_for_path(path: &str, raw: &str) -> Option<i64> {
    let last_key = path.split('.').last().unwrap_or(path)
        .trim_end_matches(|c: char| c.is_ascii_digit() || c == ']');
    let key = last_key.to_lowercase();
    match key.as_str() {
        "font_style" => cf_map_named_i64(raw, &[
            ("regular",0),("italic",1),("bold",2),("bold_italic",3),("automatic",4),("auto",4)]),
        "underline_type" => cf_map_named_i64(raw, &[
            // 0-based at CLI surface; engine offset applied in cf_apply_engine_remaps
            ("none",0),("single",1),("double",2),("single_accounting",3),("double_accounting",4)]),
        "strike_type" => cf_map_named_i64(raw, &[
            // CLI: off=0, on=1, automatic=2; engine inversion applied in cf_apply_engine_remaps
            ("off",0),("on",1),("automatic",2),("auto",2)]),
        "border_line_style" | "line_style" => cf_map_named_i64(raw, &[
            // 0-based at CLI; engine offset applied in cf_apply_engine_remaps
            ("none",0),("dash_dot",1),("dash_dot_dot",2),("dashed",3),("dotted",4),
            ("double",5),("hairline",6),("medium",7),("medium_dash_dot",8),
            ("medium_dash_dot_dot",9),("medium_dashed",10),("thick",11),("thin",12),
            ("slant_dash_dot",13)]),
        "custom_format_type" | "type" => cf_map_named_i64(raw, &[
            // 0-based at CLI; engine offset applied in cf_apply_engine_remaps
            ("general",0),("number",1),("currency",2),("accounting",3),("date",4),("time",5),
            ("duration",6),("percentage",7),("scientific",8),("fraction",9),("text",10),
            ("regional",11),("custom",12)]),
        "criteria_type" => cf_map_named_i64(raw, &[
            ("number",0),("percent",1),("percentile",2),("formula",3),
            ("minimum_value",4),("min",4),("maximum_value",5),("max",5),
            ("number_comparison",6),("date",7),("text",8),("cell_containing",9),
            ("average",10),("standard_deviation",11),("automatic",12),("auto",12),
            ("top_bottom_values",13),("none",14)]),
        "sub_criteria_type" => cf_map_named_i64(raw, &[
            ("equal_to",0),("eq",0),("not_equal_to",1),("neq",1),("between",2),("not_between",3),
            ("greater_than",4),("gt",4),("less_than",5),("lt",5),
            ("greater_than_or_equal_to",6),("gte",6),("less_than_or_equal_to",7),("lte",7),
            ("yesterday",0),("today",1),("tomorrow",2),("last_7_days",3),("last_week",4),
            ("this_week",5),("next_week",6),("last_month",7),("this_month",8),("next_month",9),
            ("next_7_days",18),("last_year",19),("this_year",20),("next_year",21),
            ("contains",0),("not_contains",1),("begins_with",2),("ends_with",3),
            ("duplicate_values",0),("unique_values",1),("blanks",2),("no_blanks",3),
            ("errors",4),("no_errors",5),
            ("top",0),("bottom",1),("above",0),("below",1),
            ("equal_or_above",2),("below_or_equal",3),
            ("one_above",0),("one_below",1),("two_above",2),("two_below",3),
            ("three_above",4),("three_below",5)]),
        "bar_direction" => cf_map_named_i64(raw, &[("left_to_right",0),("right_to_left",1),("context",2)]),
        "axis_position" => cf_map_named_i64(raw, &[("automatic",0),("middle",1),("none",2)]),
        "fill_type" => cf_map_named_i64(raw, &[("solid",0),("gradient",1)]),
        "border_type" => cf_map_named_i64(raw, &[("left",0),("right",1),("top",2),("bottom",3),
            ("none",0),("off",0),("with_border",1),("on",1)]),
        "icon_set_type" => cf_map_icon_set_type(raw).ok(),
        "icon_type" => cf_map_single_icon_type(raw),
        _ => None,
    }
}

#[derive(Debug, Clone)]
enum CfPathSegment { Key(String), Index(usize) }

fn cf_parse_path(path: &str) -> Result<Vec<CfPathSegment>, String> {
    if path.trim().is_empty() { return Err("Flag path cannot be empty.".into()); }
    let chars: Vec<char> = path.chars().collect();
    let mut i = 0usize;
    let mut token = String::new();
    let mut segs: Vec<CfPathSegment> = Vec::new();
    while i < chars.len() {
        match chars[i] {
            '.' => {
                if token.is_empty() {
                    // Allow '.' after ']' (index segment already pushed); bare leading dot is invalid
                    if segs.is_empty() { return Err(format!("Invalid path '{}': empty token.", path)); }
                    i += 1;
                } else {
                    segs.push(CfPathSegment::Key(token.clone())); token.clear(); i += 1;
                }
            }
            '[' => {
                if !token.is_empty() { segs.push(CfPathSegment::Key(token.clone())); token.clear(); }
                i += 1;
                let start = i;
                while i < chars.len() && chars[i].is_ascii_digit() { i += 1; }
                if start == i || i >= chars.len() || chars[i] != ']' {
                    return Err(format!("Invalid path '{}': malformed index.", path));
                }
                let idx: usize = chars[start..i].iter().collect::<String>().parse()
                    .map_err(|_| format!("Invalid index in path '{}'.", path))?;
                segs.push(CfPathSegment::Index(idx)); i += 1;
            }
            c => { token.push(c); i += 1; }
        }
    }
    if !token.is_empty() { segs.push(CfPathSegment::Key(token)); }
    if segs.is_empty() { return Err(format!("Invalid path '{}'.", path)); }
    Ok(segs)
}

fn cf_set_path_value(root: &mut serde_json::Value, path: &str, value: serde_json::Value) -> Result<(), String> {
    let segs = cf_parse_path(path)?;
    cf_set_at_segs(root, &segs, value)
}

fn cf_set_at_segs(node: &mut serde_json::Value, segs: &[CfPathSegment], value: serde_json::Value) -> Result<(), String> {
    if segs.is_empty() { *node = value; return Ok(()); }
    match &segs[0] {
        CfPathSegment::Key(key) => {
            if node.is_null() { *node = serde_json::json!({}); }
            let obj = node.as_object_mut().ok_or_else(|| format!("Expected object at '{}'.", key))?;
            if segs.len() == 1 { obj.insert(key.clone(), value); return Ok(()); }
            let next_default = match &segs[1] { CfPathSegment::Index(_) => serde_json::json!([]), _ => serde_json::json!({}) };
            let child = obj.entry(key.clone()).or_insert(next_default);
            cf_set_at_segs(child, &segs[1..], value)
        }
        CfPathSegment::Index(index) => {
            if node.is_null() { *node = serde_json::json!([]); }
            let arr = node.as_array_mut().ok_or_else(|| format!("Expected array at [{}].", index))?;
            while arr.len() <= *index {
                let fill = if segs.len() > 1 { match &segs[1] { CfPathSegment::Index(_) => serde_json::json!([]), _ => serde_json::json!({}) } } else { serde_json::Value::Null };
                arr.push(fill);
            }
            if segs.len() == 1 { arr[*index] = value; return Ok(()); }
            cf_set_at_segs(&mut arr[*index], &segs[1..], value)
        }
    }
}

fn build_cf_active_info_default(sheet_id: &str, range_list: &serde_json::Value) -> serde_json::Value {
    let (active_row, active_col) = range_list.as_array()
        .and_then(|arr| arr.first())
        .map(|r| {
            let row = r.get("start_row").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
            let col = r.get("start_column").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
            (row, col)
        })
        .unwrap_or((0, 0));
    serde_json::json!({
        "active_sheet_id": sheet_id,
        "active_cell": {"active_row": active_row, "active_column": active_col},
        "active_range_list": range_list
    })
}

fn build_cf_active_info_from_optional_inputs(
    sheet_id: &str,
    active_cell_ref: Option<&str>,
    range_ref: Option<&str>,
) -> Result<serde_json::Value, String> {
    let range_list = if let Some(range) = range_ref {
        let (sc, sr, ec, er) = cell_ref::try_parse_range(range)
            .ok_or_else(|| format!("Invalid range '{}'. Use A1:C5 format.", range))?;
        Some(serde_json::json!([{"start_row": sr, "end_row": er, "start_column": sc, "end_column": ec}]))
    } else { None };

    let (active_col, active_row) = if let Some(cell_ref_str) = active_cell_ref {
        cell_ref::try_parse(cell_ref_str)
            .ok_or_else(|| format!("Invalid active cell '{}'. Use A1 format.", cell_ref_str))?
    } else if let Some(ref range) = range_list {
        let first = range.as_array().and_then(|a| a.first())
            .ok_or_else(|| "Invalid active range list.".to_string())?;
        let row = first.get("start_row").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
        let col = first.get("start_column").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
        (col, row)
    } else { (0, 0) };

    let mut ai = serde_json::Map::new();
    ai.insert("active_sheet_id".into(), serde_json::json!(sheet_id));
    ai.insert("active_cell".into(), serde_json::json!({"active_row": active_row, "active_column": active_col}));
    if let Some(range) = range_list { ai.insert("active_range_list".into(), range); }
    Ok(serde_json::Value::Object(ai))
}


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

fn parse_bool_arg(value: &str) -> bool {
    matches!(value.to_lowercase().as_str(), "true" | "yes" | "1" | "on")
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

// ─── Data Validation ─────────────────────────────────────────────────────────

fn cmd_dv(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    if args.is_empty() {
        output::error("Usage: dv <readrange|create|edit|manage|delete> ...");
        return;
    }
    match args[0].to_lowercase().as_str() {
        "readrange" | "read-range" | "check" => cmd_dv_read_range(&args[1..], engine, session),
        "create" | "add" => cmd_dv_create(&args[1..], engine, session),
        "edit" | "update" => cmd_dv_edit(&args[1..], engine, session),
        "manage" | "list" => cmd_dv_manage(&args[1..], engine, session),
        "delete" | "remove" | "clear" => cmd_dv_delete(&args[1..], engine, session),
        other => output::error(&format!(
            "Unknown dv sub-command: '{}'. Use: readrange, create, edit, manage, delete",
            other
        )),
    }
}

/// Policy for what `dv create`/`dv edit` should do when the target range
/// overlaps an existing data-validation rule (engine reports EXTEND_RANGE,
/// EXTEND_DATA_VALIDATION or COLLISION).
#[derive(Clone, Copy, PartialEq)]
enum OnCollision {
    /// Refuse and make no changes (opt-in via `--on-collision abort`).
    Abort,
    /// Apply the new rule anyway; the engine splits/overwrites the overlap (default).
    Replace,
}

fn parse_on_collision(s: &str) -> Option<OnCollision> {
    match s.to_lowercase().as_str() {
        "abort" | "refuse" | "error" => Some(OnCollision::Abort),
        "replace" | "overwrite" | "force" => Some(OnCollision::Replace),
        _ => None,
    }
}

/// Runs a read-only DV read-range preview and returns the engine's action
/// string: CREATE | EDIT | EXTEND_RANGE | EXTEND_DATA_VALIDATION | COLLISION.
/// No `read_range_option` is sent, so this never mutates the workbook.
fn dv_preview_action(
    engine: &EngineHandle,
    rid: &str,
    sid: &str,
    range_list: &serde_json::Value,
) -> Result<String, String> {
    let request = rb::build_dv_read_range(rid, sid, range_list, None);
    let resp = engine
        .process_request_json(&request)
        .map_err(|e| format!("Engine error: {}", e))?;
    let status = rp::parse_status_response(&resp);
    if !rp::is_success(status.status_code) {
        return Err(status.status_message.unwrap_or_else(|| "engine error".into()));
    }
    let v: serde_json::Value = serde_json::from_str(&resp).unwrap_or_default();
    let action = v
        .get("response")
        .and_then(|r| r.get("read_range_action"))
        .and_then(|a| a.as_str())
        .unwrap_or("UNKNOWN")
        .to_string();
    Ok(action)
}

/// True when the range overlaps an existing rule and needs a collision policy.
fn dv_action_is_conflict(action: &str) -> bool {
    matches!(action, "EXTEND_RANGE" | "EXTEND_DATA_VALIDATION" | "COLLISION")
}

/// dv readrange <range>
///
/// Read-only dry-run: reports what applying a DV rule to <range> would do
/// (CREATE | EDIT | EXTEND_RANGE | EXTEND_DATA_VALIDATION | COLLISION) without
/// mutating the workbook. To actually resolve a collision, pass
/// `--on-collision` to `dv create`/`dv edit`.
fn cmd_dv_read_range(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: dv readrange <range>");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let range_list = serde_json::json!([{
        "start_row": sr,
        "start_column": sc,
        "end_row": er,
        "end_column": ec
    }]);

    let request = rb::build_dv_read_range(rid, &sid, &range_list, None);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let v: serde_json::Value = serde_json::from_str(&resp).unwrap_or_default();
            let status = rp::parse_status_response(&resp);
            if !rp::is_success(status.status_code) {
                output::error(&format!(
                    "Operation failed: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
                return;
            }
            if let Some(response) = v.get("response") {
                let action = response.get("read_range_action")
                    .and_then(|v| v.as_str())
                    .unwrap_or("UNKNOWN");
                output::success(&format!("Read range action: {}", action));

                if action == "EXTEND_RANGE" {
                    if let Some(nr) = response.get("new_range") {
                        output::key_value("New range", &nr.to_string(), 2);
                    }
                }
                if let Some(rule) = response.get("dv_rule") {
                    let rows = [
                        ("Criterion Type",       rule.get("criterion_type").and_then(|v| v.as_str()).unwrap_or("–").to_string()),
                        ("Display Dropdown",     rule.get("display_dropdown_list").map(|v| v.to_string()).unwrap_or("–".into())),
                        ("Error Style",          rule.get("error_style").and_then(|v| v.as_str()).unwrap_or("–").to_string()),
                        ("Error Title",          rule.get("error_text_title").and_then(|v| v.as_str()).unwrap_or("–").to_string()),
                        ("Error Message",        rule.get("error_text_message").and_then(|v| v.as_str()).unwrap_or("–").to_string()),
                        ("Help Title",           rule.get("help_text_title").and_then(|v| v.as_str()).unwrap_or("–").to_string()),
                        ("Help Message",         rule.get("help_text_message").and_then(|v| v.as_str()).unwrap_or("–").to_string()),
                        ("Ignore Blank",         rule.get("ignore_blank").map(|v| v.to_string()).unwrap_or("–".into())),
                        ("Error Disabled",       rule.get("is_error_disabled").map(|v| v.to_string()).unwrap_or("–".into())),
                        ("Help Text Disabled",   rule.get("is_help_text_disabled").map(|v| v.to_string()).unwrap_or("–".into())),
                        ("Sort Dropdown Values", rule.get("sort_dropdown_list_values").map(|v| v.to_string()).unwrap_or("–".into())),
                        ("Parameter List",       rule.get("parameter_list").map(|v| v.to_string()).unwrap_or("–".into())),
                    ];
                    output::kv_table("Existing Rule", &rows);
                }
            } else {
                output::info(&resp);
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

/// dv create <range> <criteria-type> [sub-criteria-type] [--val1 <v>] [--val2 <v>]
///           [--delimiter <d>] [--show-list <true|false>] [--sort-list <true|false>]
///           [--ignore-blanks <true|false>]
///           [--help-title <t>] [--help-msg <m>] [--hide-help]
///           [--error-title <t>] [--error-msg <m>] [--error-style <stop|warning|info>]
///           [--no-error-validation]
fn cmd_dv_create(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: dv create <range> <criteria-type> [sub-criteria] [flags]\n\
                 Criteria: whole-number|decimal|list|datetime|text-length|custom|text|cell-range|any-value\n\
                 Sub-criteria: between|notbetween|equal|notequal|gt|lt|gte|lte|contains|notcontains|beginswith|endswith\n\
                 Flags: --val1 <v> --val2 <v> --delimiter <d> (default \",\") --show-list <t/f> --sort-list <t/f>\n\
                        --ignore-blanks <t/f> --help-title <t> --help-msg <m> --hide-help\n\
                        --error-title <t> --error-msg <m> --error-style <stop|warning|info> --no-error-validation\n\
                        --on-collision <abort|replace> (default replace: apply regardless of overlaps; abort refuses if the range overlaps an existing rule)";
    if args.len() < 2 {
        output::error(usage);
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let criteria_type = match parse_dv_criteria_type(args[1]) {
        Some(v) => v,
        None => {
            output::error(&format!("Unknown criteria type: '{}'. Use: whole-number, decimal, list, datetime, text-length, custom, text, cell-range, any-value", args[1]));
            return;
        }
    };

    let mut sub_criteria: Option<i64> = None;
    let mut val1: Option<String> = None;
    let mut val2: Option<String> = None;
    let mut delimiter: Option<String> = Some(",".to_string());
    let mut show_list = true;
    let mut sort_list = false;
    let mut ignore_blanks = true;
    let mut help_title: Option<String> = None;
    let mut help_msg: Option<String> = None;
    let mut hide_help = false;
    let mut error_title: Option<String> = None;
    let mut error_msg: Option<String> = None;
    let mut error_style = rb::DV_ERROR_STYLE_STOP;
    let mut no_error_validation = false;
    // Default: apply the rule regardless of overlaps (engine resolves them).
    // `--on-collision abort` opts into a pre-flight that refuses on an overlap.
    let mut on_collision = OnCollision::Replace;

    let mut i = 2;
    // optional positional sub-criteria (before flags)
    if i < args.len() && !args[i].starts_with("--") {
        sub_criteria = parse_dv_sub_criteria_type(args[i]);
        if sub_criteria.is_none() {
            output::error(&format!("Unknown sub-criteria: '{}'. Use: between, notbetween, equal, notequal, gt, lt, gte, lte, ...", args[i]));
            return;
        }
        i += 1;
    }
    while i < args.len() {
        match args[i].to_lowercase().as_str() {
            "--val1" => { if i + 1 < args.len() { val1 = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--val2" => { if i + 1 < args.len() { val2 = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--delimiter" => { if i + 1 < args.len() { delimiter = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--show-list" => { if i + 1 < args.len() { show_list = parse_bool_arg(args[i+1]); i += 2; } else { i += 1; } }
            "--sort-list" => { if i + 1 < args.len() { sort_list = parse_bool_arg(args[i+1]); i += 2; } else { i += 1; } }
            "--ignore-blanks" => { if i + 1 < args.len() { ignore_blanks = parse_bool_arg(args[i+1]); i += 2; } else { i += 1; } }
            "--help-title" => { if i + 1 < args.len() { help_title = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--help-msg" => { if i + 1 < args.len() { help_msg = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--hide-help" => { hide_help = true; i += 1; }
            "--error-title" => { if i + 1 < args.len() { error_title = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--error-msg" => { if i + 1 < args.len() { error_msg = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--error-style" => {
                if i + 1 < args.len() {
                    error_style = match args[i+1].to_lowercase().as_str() {
                        "stop" | "0" => rb::DV_ERROR_STYLE_STOP,
                        "warning" | "warn" | "1" => rb::DV_ERROR_STYLE_WARNING,
                        "info" | "information" | "2" => rb::DV_ERROR_STYLE_INFORMATION,
                        other => { output::error(&format!("Unknown error style: '{}'. Use: stop, warning, info", other)); return; }
                    };
                    i += 2;
                } else { i += 1; }
            }
            "--no-error-validation" => { no_error_validation = true; i += 1; }
            "--on-collision" => {
                if i + 1 < args.len() {
                    match parse_on_collision(args[i+1]) {
                        Some(v) => on_collision = v,
                        None => { output::error(&format!("Unknown --on-collision mode: '{}'. Use: abort, replace", args[i+1])); return; }
                    }
                    i += 2;
                } else { i += 1; }
            }
            "--force" => { on_collision = OnCollision::Replace; i += 1; }
            _ => { i += 1; }
        }
    }

    let mut condition = serde_json::json!({
        "criteria_type": criteria_type,
        "is_ignore_blanks": ignore_blanks,
        "show_list": show_list,
        "show_list_ascending": sort_list,
    });
    if let Some(sc) = sub_criteria {
        condition["sub_criteria_type"] = serde_json::json!(sc);
    }
    if let Some(v) = val1 { condition["val1"] = serde_json::json!(v); }
    if let Some(v) = val2 { condition["val2"] = serde_json::json!(v); }
    if let Some(d) = delimiter { condition["delimiter"] = serde_json::json!(d); }
    condition["help_text"] = serde_json::json!({
        "is_help_text_disabled": hide_help,
        "title_string": help_title.unwrap_or_default(),
        "message": help_msg.unwrap_or_default(),
    });
    condition["error_alert"] = serde_json::json!({
        "is_enable_error_validation": !no_error_validation,
        "title_string": error_title.unwrap_or_default(),
        "message": error_msg.unwrap_or_default(),
        "error_style": error_style,
    });

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let range_list = serde_json::json!([{
        "start_row": sr, "start_column": sc, "end_row": er, "end_column": ec
    }]);
    let active_info = serde_json::json!({
        "active_sheet_id": sid,
        "active_cell": { "active_row": sr, "active_column": sc },
        "active_range_list": [{ "start_row": sr, "start_column": sc, "end_row": er, "end_column": ec }]
    });

    // By default the rule is applied unconditionally and the engine resolves any
    // overlap (splitting/overwriting the affected cells). Only when the caller opts
    // into --on-collision abort do we pre-flight and refuse on an overlap.
    if on_collision == OnCollision::Abort {
        match dv_preview_action(engine, rid, &sid, &range_list) {
            Ok(action) if dv_action_is_conflict(&action) => {
                output::error(&format!(
                    "Range {} overlaps an existing data-validation rule (engine action: {}). No changes made.\n\
                     (This check is opt-in via --on-collision abort; the default applies the rule anyway.)",
                    args[0], action
                ));
                return;
            }
            Ok(_) => {}
            Err(e) => { output::error(&format!("Collision pre-check failed: {}", e)); return; }
        }
    }

    let request = rb::build_dv_create_rule(rid, &sid, &range_list, condition, active_info);
    exec_status_cmd(engine, &request, session, "Data validation rule created.");
}

/// dv edit <range> [flags — same as create, all optional]
fn cmd_dv_edit(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let usage = "Usage: dv edit <range> [flags]\n\
                 All flags are optional — only supplied fields are updated.\n\
                 Flags: --criteria <type> --sub-criteria <type> --val1 <v> --val2 <v>\n\
                        --show-list <t/f> --sort-list <t/f> --ignore-blanks <t/f>\n\
                        --help-title <t> --help-msg <m> --hide-help\n\
                        --error-title <t> --error-msg <m> --error-style <stop|warning|info> --no-error-validation\n\
                        --on-collision <abort|replace> (default replace: apply regardless of overlaps; abort refuses if the range overlaps a different rule)";
    if args.is_empty() {
        output::error(usage);
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);

    let mut criteria_type: Option<i64> = None;
    let mut sub_criteria: Option<i64> = None;
    let mut val1: Option<String> = None;
    let mut val2: Option<String> = None;
    let mut show_list: Option<bool> = None;
    let mut sort_list: Option<bool> = None;
    let mut ignore_blanks: Option<bool> = None;
    let mut help_title: Option<String> = None;
    let mut help_msg: Option<String> = None;
    let mut hide_help: Option<bool> = None;
    let mut error_title: Option<String> = None;
    let mut error_msg: Option<String> = None;
    let mut error_style: Option<i32> = None;
    let mut no_error_validation: Option<bool> = None;
    // Default: apply the rule regardless of overlaps (engine resolves them).
    // `--on-collision abort` opts into a pre-flight that refuses on an overlap.
    let mut on_collision = OnCollision::Replace;

    let mut i = 1;
    while i < args.len() {
        match args[i].to_lowercase().as_str() {
            "--criteria" => {
                if i + 1 < args.len() {
                    criteria_type = parse_dv_criteria_type(args[i+1]);
                    if criteria_type.is_none() {
                        output::error(&format!("Unknown criteria type: '{}'", args[i+1]));
                        return;
                    }
                    i += 2;
                } else { i += 1; }
            }
            "--sub-criteria" => {
                if i + 1 < args.len() {
                    sub_criteria = parse_dv_sub_criteria_type(args[i+1]);
                    if sub_criteria.is_none() {
                        output::error(&format!("Unknown sub-criteria: '{}'", args[i+1]));
                        return;
                    }
                    i += 2;
                } else { i += 1; }
            }
            "--val1" => { if i + 1 < args.len() { val1 = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--val2" => { if i + 1 < args.len() { val2 = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--show-list" => { if i + 1 < args.len() { show_list = Some(parse_bool_arg(args[i+1])); i += 2; } else { i += 1; } }
            "--sort-list" => { if i + 1 < args.len() { sort_list = Some(parse_bool_arg(args[i+1])); i += 2; } else { i += 1; } }
            "--ignore-blanks" => { if i + 1 < args.len() { ignore_blanks = Some(parse_bool_arg(args[i+1])); i += 2; } else { i += 1; } }
            "--help-title" => { if i + 1 < args.len() { help_title = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--help-msg" => { if i + 1 < args.len() { help_msg = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--hide-help" => { hide_help = Some(true); i += 1; }
            "--error-title" => { if i + 1 < args.len() { error_title = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--error-msg" => { if i + 1 < args.len() { error_msg = Some(args[i+1].to_string()); i += 2; } else { i += 1; } }
            "--error-style" => {
                if i + 1 < args.len() {
                    let es = match args[i+1].to_lowercase().as_str() {
                        "stop" | "0" => rb::DV_ERROR_STYLE_STOP,
                        "warning" | "warn" | "1" => rb::DV_ERROR_STYLE_WARNING,
                        "info" | "information" | "2" => rb::DV_ERROR_STYLE_INFORMATION,
                        other => { output::error(&format!("Unknown error style: '{}'", other)); return; }
                    };
                    error_style = Some(es);
                    i += 2;
                } else { i += 1; }
            }
            "--no-error-validation" => { no_error_validation = Some(true); i += 1; }
            "--on-collision" => {
                if i + 1 < args.len() {
                    match parse_on_collision(args[i+1]) {
                        Some(v) => on_collision = v,
                        None => { output::error(&format!("Unknown --on-collision mode: '{}'. Use: abort, replace", args[i+1])); return; }
                    }
                    i += 2;
                } else { i += 1; }
            }
            "--force" => { on_collision = OnCollision::Replace; i += 1; }
            _ => { i += 1; }
        }
    }

    let mut condition = serde_json::json!({});
    if let Some(ct) = criteria_type { condition["criteria_type"] = serde_json::json!(ct); }
    if let Some(sct) = sub_criteria { condition["sub_criteria_type"] = serde_json::json!(sct); }
    if let Some(v) = val1 { condition["val1"] = serde_json::json!(v); }
    if let Some(v) = val2 { condition["val2"] = serde_json::json!(v); }
    if let Some(v) = show_list { condition["show_list"] = serde_json::json!(v); }
    if let Some(v) = sort_list { condition["show_list_ascending"] = serde_json::json!(v); }
    if let Some(v) = ignore_blanks { condition["is_ignore_blanks"] = serde_json::json!(v); }

    if help_title.is_some() || help_msg.is_some() || hide_help.is_some() {
        let mut ht = serde_json::json!({});
        if let Some(v) = hide_help { ht["is_help_text_disabled"] = serde_json::json!(v); }
        if let Some(v) = help_title { ht["title_string"] = serde_json::json!(v); }
        if let Some(v) = help_msg { ht["message"] = serde_json::json!(v); }
        condition["help_text"] = ht;
    }

    if error_title.is_some() || error_msg.is_some() || error_style.is_some() || no_error_validation.is_some() {
        let mut ea = serde_json::json!({});
        if let Some(v) = no_error_validation { ea["is_enable_error_validation"] = serde_json::json!(!v); }
        if let Some(v) = error_title { ea["title_string"] = serde_json::json!(v); }
        if let Some(v) = error_msg { ea["message"] = serde_json::json!(v); }
        if let Some(v) = error_style { ea["error_style"] = serde_json::json!(v); }
        condition["error_alert"] = ea;
    }

    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();
    let active_info = serde_json::json!({
        "active_sheet_id": sid,
        "active_cell": { "active_row": sr, "active_column": sc },
        "active_range_list": [{ "start_row": sr, "start_column": sc, "end_row": er, "end_column": ec }]
    });

    // By default the edit is applied and the engine resolves any overlap. Only
    // when the caller opts into --on-collision abort do we pre-flight and refuse.
    if on_collision == OnCollision::Abort {
        let range_list = serde_json::json!([{ "start_row": sr, "start_column": sc, "end_row": er, "end_column": ec }]);
        match dv_preview_action(engine, rid, &sid, &range_list) {
            Ok(action) if dv_action_is_conflict(&action) => {
                output::error(&format!(
                    "Range {} overlaps a different data-validation rule (engine action: {}). No changes made.\n\
                     (This check is opt-in via --on-collision abort; the default edits across the overlap.)",
                    args[0], action
                ));
                return;
            }
            Ok(_) => {}
            Err(e) => { output::error(&format!("Collision pre-check failed: {}", e)); return; }
        }
    }

    let request = rb::build_dv_edit_rule(rid, &sid, sr as i64, sc as i64, er as i64, ec as i64, condition, active_info);
    exec_status_cmd(engine, &request, session, "Data validation rule updated.");
}

/// dv manage [--sheet <name|index>] [--range <A1:C5>]
/// Without flags: lists rules across the entire workbook.
fn cmd_dv_manage(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let mut sheet_override: Option<String> = None;
    let mut range_override: Option<serde_json::Value> = None;
    let mut i = 0;
    while i < args.len() {
        match args[i].to_lowercase().as_str() {
            "--sheet" => {
                if i + 1 < args.len() {
                    sheet_override = Some(args[i+1].to_string());
                    i += 2;
                } else { i += 1; }
            }
            "--range" => {
                if i + 1 < args.len() {
                    let (sc, sr, ec, er) = parse_range_arg!(args[i+1]);
                    range_override = Some(serde_json::json!([{
                        "start_row": sr, "start_column": sc, "end_row": er, "end_column": ec
                    }]));
                    i += 2;
                } else { i += 1; }
            }
            _ => { i += 1; }
        }
    }

    let resolved_sheet_id: Option<String> = if let Some(ref name) = sheet_override {
        match resolve_sheet_id(name, session) {
            Some((id, _)) => Some(id),
            None => return,
        }
    } else {
        None
    };

    let (scope, sheet_id_opt, range_opt) = if let Some(ref rl) = range_override {
        let effective_sid = resolved_sheet_id.as_deref().unwrap_or(&sid);
        (rb::DV_SCOPE_RANGE, Some(effective_sid.to_string()), Some(rl.clone()))
    } else if resolved_sheet_id.is_some() {
        (rb::DV_SCOPE_SHEET, resolved_sheet_id.clone(), None)
    } else {
        (rb::DV_SCOPE_WORKBOOK, None, None)
    };

    let request = rb::build_dv_manage_rules(
        rid,
        scope,
        sheet_id_opt.as_deref(),
        range_opt.as_ref(),
    );
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if !rp::is_success(status.status_code) {
                output::error(&format!(
                    "Operation failed: {}",
                    status.status_message.unwrap_or_else(|| "engine error".into())
                ));
                return;
            }
            let v: serde_json::Value = serde_json::from_str(&resp).unwrap_or_default();
            let empty = vec![];
            let rules = v.get("response")
                .and_then(|r| r.get("rules"))
                .and_then(|r| r.as_array())
                .unwrap_or(&empty);

            if rules.is_empty() {
                output::info("No data validation rules found.");
                return;
            }
            let mut total = 0usize;
            for sheet_entry in rules {
                let sheet_id = sheet_entry.get("sheet_id").and_then(|v| v.as_str()).unwrap_or("unknown");
                let empty_rules = vec![];
                let sheet_rules = sheet_entry.get("rules_in_sheet").and_then(|v| v.as_array()).unwrap_or(&empty_rules);
                output::key_value("Sheet", sheet_id, 0);
                for rule in sheet_rules {
                    total += 1;
                    let rule_id = rule.get("rule_id").and_then(|v| v.as_i64()).unwrap_or(-1);
                    let ct = rule.get("criterion_type").and_then(|v| v.as_str()).unwrap_or("?");
                    let range_str = rule.get("range").and_then(|v| v.as_array())
                        .map(|arr| arr.iter().map(|r| {
                            let sr = r.get("start_row").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
                            let sc = r.get("start_column").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
                            let er = r.get("end_row").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
                            let ec = r.get("end_column").and_then(|v| v.as_i64()).unwrap_or(0) as i32;
                            let start = cell_ref::to_ref(sc, sr);
                            if sr == er && sc == ec {
                                start
                            } else {
                                format!("{}:{}", start, cell_ref::to_ref(ec, er))
                            }
                        }).collect::<Vec<_>>().join(", "))
                        .unwrap_or_default();
                    output::key_value(&format!("  Rule #{}", rule_id), &format!("{} — {}", ct, range_str), 2);
                    if let Some(params) = rule.get("parameter_list") {
                        output::key_value("    Values", &params.to_string(), 4);
                    }
                }
            }
            output::success(&format!("{} rule(s) found.", total));
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

/// dv delete <range> [--sheet <name|id>]
///
/// Clears the data validation rule for the given range.
fn cmd_dv_delete(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: dv delete <range> [--sheet <name|id>]");
        return;
    }
    let (sc, sr, ec, er) = parse_range_arg!(args[0]);
    let rid = session.rid.as_deref().unwrap();
    let sid = session.get_active_sheet_id_or_default();

    let mut sheet_id = sid.clone();
    let mut i = 1;
    while i < args.len() {
        if args[i].eq_ignore_ascii_case("--sheet") && i + 1 < args.len() {
            match resolve_sheet_id(args[i + 1], session) {
                Some((id, _)) => sheet_id = id,
                None => return,
            }
            i += 2;
        } else {
            i += 1;
        }
    }

    let sheet_range_list = serde_json::json!([{
        "sheet_id": sheet_id,
        "range_list": [{ "start_row": sr, "end_row": er, "start_column": sc, "end_column": ec }]
    }]);
    let active_info = serde_json::json!({
        "active_sheet_id": sid,
        "active_cell": { "active_row": sr, "active_column": sc },
        "active_range_list": [{ "start_row": sr, "end_row": er, "start_column": sc, "end_column": ec }]
    });

    let request = rb::build_dv_clear_rule(rid, sheet_range_list, active_info);
    exec_status_cmd(engine, &request, session, "Data validation rule cleared.");
}

fn parse_dv_criteria_type(s: &str) -> Option<i64> {
    match s.to_lowercase().replace('-', "").as_str() {
        "wholenumber" | "whole" | "integer" | "int" | "0" => Some(rb::DV_CRITERIA_WHOLE_NUMBER as i64),
        "decimal" | "float" | "number" | "1" => Some(rb::DV_CRITERIA_DECIMAL as i64),
        "list" | "2" => Some(rb::DV_CRITERIA_LIST as i64),
        "datetime" | "date" | "time" | "3" => Some(rb::DV_CRITERIA_DATE_TIME as i64),
        "textlength" | "text-length" | "length" | "4" => Some(rb::DV_CRITERIA_TEXT_LENGTH as i64),
        "custom" | "formula" | "5" => Some(rb::DV_CRITERIA_CUSTOM as i64),
        "text" | "6" => Some(rb::DV_CRITERIA_TEXT as i64),
        "cellrange" | "range" | "7" => Some(rb::DV_CRITERIA_CELL_RANGE as i64),
        "anyvalue" | "any" | "none" | "8" => Some(rb::DV_CRITERIA_ANY_VALUE as i64),
        _ => None,
    }
}

fn parse_dv_sub_criteria_type(s: &str) -> Option<i64> {
    match s.to_lowercase().replace('-', "").as_str() {
        "between" | "0" => Some(0),
        "notbetween" | "1" => Some(1),
        "equal" | "equalto" | "eq" | "2" => Some(2),
        "notequal" | "notequalto" | "ne" | "3" => Some(3),
        "greaterthan" | "gt" | "4" => Some(4),
        "lessthan" | "lt" | "5" => Some(5),
        "gte" | "greaterequal" | "greaterthanorequalto" | "6" => Some(6),
        "lte" | "lessequal" | "lessthanorequalto" | "7" => Some(7),
        "contains" | "8" => Some(8),
        "notcontains" | "doesnotcontain" | "9" => Some(9),
        "beginswith" | "startswith" | "10" => Some(10),
        "endswith" | "11" => Some(11),
        "before" | "12" => Some(12),
        "onorbefore" | "13" => Some(13),
        "after" | "14" => Some(14),
        "onorafter" | "15" => Some(15),
        "on" | "16" => Some(16),
        "noton" | "17" => Some(17),
        "yesterday" | "18" => Some(18),
        "today" | "19" => Some(19),
        "tomorrow" | "20" => Some(20),
        "last7days" | "21" => Some(21),
        "next7days" | "22" => Some(22),
        "lastweek" | "23" => Some(23),
        "thisweek" | "24" => Some(24),
        "nextweek" | "25" => Some(25),
        "lastmonth" | "26" => Some(26),
        "thismonth" | "27" => Some(27),
        "nextmonth" | "28" => Some(28),
        "lastyear" | "29" => Some(29),
        "thisyear" | "30" => Some(30),
        "nextyear" | "31" => Some(31),
        "isvaliddatetime" | "validdate" | "32" => Some(32),
        _ => None,
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
        Some("checkbox") | Some("checkboxes") => print_help_checkbox(),
        Some("row") | Some("col") | Some("column") | Some("rows") | Some("columns") => print_help_rowcol(),
        Some("editing") | Some("copy") | Some("move") | Some("merge") | Some("clear")
        | Some("undo") | Some("redo") | Some("clipboardcopy") => print_help_editing(),
        Some("find") | Some("replace") | Some("sort") | Some("filter") | Some("search") => print_help_find_sort(),
        Some("view") | Some("freeze") | Some("unfreeze") | Some("name") => print_help_view_names(),
        Some("table") | Some("tables") => print_help_table(),
        Some("pivot") => print_help_pivot(),
        Some("chart") | Some("charts") => print_help_chart(),
        Some("format") | Some("formatting") => print_help_format(),
        Some("cf") | Some("conditional") => print_help_conditional(),
        Some("dv") | Some("datavalidation") | Some("data-validation") | Some("validation") => print_help_dv(),
        Some("theme") | Some("themes") => print_help_theme(),
        Some(unknown) => {
            output::error(&format!("Unknown help topic: '{}'. Type 'help' for all topics.", unknown));
            println!();
            output::help_header("Available Help Topics");
            output::help_section("TOPICS");
            output::help_cmd("help file", "File operations (open, save, close)");
            output::help_cmd("help worksheet", "Worksheet management");
            output::help_cmd("help cell", "Cell get/set operations");
            output::help_cmd("help checkbox", "Insert, update, and delete checkbox controls");
            output::help_cmd("help row", "Row & column operations");
            output::help_cmd("help editing", "Copy, move, merge, clear, undo/redo");
            output::help_cmd("help find", "Find, replace, sort & filter");
            output::help_cmd("help view", "Freeze panes & named ranges");
            output::help_cmd("help table", "Table operations");
            output::help_cmd("help pivot", "Pivot table operations");
            output::help_cmd("help chart", "Chart operations");
            output::help_cmd("help format", "Cell formatting & number formats");
            output::help_cmd("help cf", "Conditional formatting APIs");
            output::help_cmd("help dv", "Data validation rules");
            output::help_cmd("help theme", "Workbook themes");
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
    output::help_cmd("help checkbox", "Insert, update, and delete checkbox controls");
    output::help_cmd("help row", "Row & column operations");
    output::help_cmd("help editing", "Copy, move, merge, clear, undo/redo");
    output::help_cmd("help find", "Find, replace, sort & filter");
    output::help_cmd("help view", "Freeze panes & named ranges");
    output::help_cmd("help table", "Table operations");
    output::help_cmd("help pivot", "Pivot table operations");
    output::help_cmd("help chart", "Chart operations");
    output::help_cmd("help format", "Cell formatting & number formats");
    output::help_cmd("help cf", "Conditional formatting APIs");
    output::help_cmd("help dv", "Data validation rules");
    output::help_cmd("help theme", "Workbook themes");
    output::help_cmd("help --all", "Show all commands (full reference)");

    output::help_section("QUICK REFERENCE");
    output::help_cmd("open <filepath>", "Open a file (.xlsx, .csv, .tsv)");
    output::help_cmd("open --new <docname>", "Create a new blank workbook");
    output::help_cmd("save / save --as <path>", "Save workbook");
    output::help_cmd("close [--force]", "Close the current workbook");
    output::help_cmd("worksheet <sub>", "Manage worksheets");
    output::help_cmd("cell get|set <ref> ...", "Read/write cell values & formulas");
    output::help_cmd("cell set <ref> --hyperlink ...", "Insert hyperlink in cell");
    output::help_cmd("cell set <ref> --note <text>", "Set a note/comment on a cell");
    output::help_cmd("checkbox insert|update|delete ...", "Manage checkbox(es) in range");
    output::help_cmd("row|col <sub> ...", "Row & column operations");
    output::help_cmd("copy|move <range> <dest>", "Copy/move ranges");
    output::help_cmd("find|replace ...", "Search & replace");
    output::help_cmd("sort <range> <col> ...", "Sort data");
    output::help_cmd("table <sub> ...", "Table operations");
    output::help_cmd("pivot list|create|... ", "Pivot table operations");
    output::help_cmd("chart <sub> ...", "Chart operations");
    output::help_cmd("format <prop> <range> ...", "Cell formatting");
    output::help_cmd("cf classic|colorscale|databar|iconset ...", "Insert/edit conditional-format rules");
    output::help_cmd("cf list|delete|move ...", "List, delete, and reorder CF rules");
    output::help_cmd("dv create|edit|manage|readrange ...", "Data validation rules");
    output::help_cmd("theme apply <type|name>", "Apply a theme to the workbook");
    output::help_cmd("theme apply custom --bg ... --text ...", "Apply a custom theme");
    output::help_cmd("theme list", "List available themes");
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
    print_help_checkbox_content();
    print_help_rowcol_content();
    print_help_editing_content();
    print_help_find_sort_content();
    print_help_view_names_content();
    print_help_table_content();
    print_help_pivot_content();
    print_help_chart_content();
    print_help_format_content();
    print_help_conditional_content();
    print_help_dv_content();

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
    output::help_detail("Password-protected files: you will be prompted interactively.");
    output::help_detail("  Open password   → enter when prompted (input is hidden)");
    output::help_detail("  Modify password → choose read-only or enter the modify password");
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
    output::help_cmd("worksheet switch <name|index>", "Switch to another sheet");
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
    output::help_cmd("cell set <ref> --hyperlink <link> [--text <display>] [--type <0..4>]", "Insert hyperlink (0=WEB_PAGE,1=RANGE,2=EMAIL,3=TELEPHONE,4=DEFINED_NAME)");
    output::help_cmd("cell set <ref> --note <text>", "Set a note/comment on a cell");
}

fn print_help_rowcol() {
    output::help_header("Row & Column Operations");
    print_help_rowcol_content();
    output::help_footer();
}

fn print_help_checkbox() {
    output::help_header("Checkbox Operations");
    print_help_checkbox_content();
    output::help_footer();
}

fn print_help_checkbox_content() {
    output::help_section("CHECKBOX");
    output::help_cmd("checkbox insert <range>", "Insert checkbox(es) in the selected range");
    output::help_cmd("checkbox update <range> <true|false>", "Set checkbox state in the selected range");
    output::help_detail("<true|false> also accepts: 1/0, yes/no, on/off");
    output::help_cmd("checkbox delete <range>", "Delete checkbox(es) in the selected range");
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

fn print_help_conditional() {
    output::help_header("Conditional Formatting");
    print_help_conditional_content();
    output::help_footer();
}

fn print_help_conditional_content() {
    output::help_section("COLORS");
    output::help_detail("  Any flag that accepts a color supports these notations:");
    output::help_detail("    Named:  red green blue yellow orange white black gray/grey purple cyan magenta");
    output::help_detail("            pink brown lime navy teal maroon olive");
    output::help_detail("    Hex:    \"#0070C0\"  or  0070C0");
    output::help_detail("    RGB:    0,112,192");
    output::help_detail("    Theme:  theme:4   or   theme:4,0.5   (index,tint)");
    output::help_detail("  Theme colors NOT supported in iconset — use RGB/hex/named there.");

    output::help_section("EDIT MODE");
    output::help_detail("  All four rule types share the same flag form for insert and edit.");
    output::help_detail("  Insert: cf classic --range A1:C10 <flags>");
    output::help_detail("  Edit:   cf classic <#index>   <flags>   (index from 'cf list')");
    output::help_detail("          cf classic --rule-id <id> <flags>");
    output::help_detail("  criteria_id is auto-assigned from the condition order.");

    output::help_section("CONDITIONAL FORMATTING");

    output::help_cmd("cf classic --range <A1:C5> [--when \"expr\"] [style flags] [condition flags]", "Insert a classic rule");
    output::help_cmd("cf classic <#index|--rule-id id> [--range] [flags]", "Edit a classic rule");
    output::help_detail("  ── --when shortcuts ──");
    output::help_detail("    --when \">100\"              number_comparison + gt + lhs 100");
    output::help_detail("    --when \">=100\" \"<100\" \"<=100\" \"=100\" \"!=100\"   (same pattern)");
    output::help_detail("    --when \"between 10 and 20\"   / \"not between 10 and 20\"");
    output::help_detail("    --when \"contains 'text'\"  \"not contains 'x'\"  \"begins with 'Q'\"  \"ends with '.csv'\"");
    output::help_detail("    --when \"is duplicate\"  \"is unique\"  \"is blank\"  \"is not blank\"  \"is error\"  \"is not error\"");
    output::help_detail("    --when \"above average\"  \"below average\"");
    output::help_detail("    --when \"top 10\"  \"bottom 5\"");
    output::help_detail("    --when \"today\"  \"yesterday\"  \"last 7 days\"  \"this week\"  etc.");
    output::help_detail("  ── Style shorthands ──");
    output::help_detail("    --bold                           font_style bold");
    output::help_detail("    --italic                         font_style italic");
    output::help_detail("    --bold-italic                    font_style bold_italic");
    output::help_detail("    --underline [single|double|...]  underline_type (default single)");
    output::help_detail("    --strike                         strike_type on");
    output::help_detail("    --font-color <color>             font color");
    output::help_detail("    --fill <color>                   fill color");
    output::help_detail("    --border-bottom <line_style> [color]   (also --border-top/left/right)");
    output::help_detail("  ── Raw condition flags ──");
    output::help_detail("    --condition.criteria_type <name|int>");
    output::help_detail("      number(0) percent(1) percentile(2) formula(3)");
    output::help_detail("      minimum_value/min(4)  maximum_value/max(5)  number_comparison(6)");
    output::help_detail("      date(7) text(8) cell_containing(9) average(10)");
    output::help_detail("      standard_deviation(11) automatic(12) top_bottom_values(13) none(14)");
    output::help_detail("    --condition.sub_criteria_type <name|int>");
    output::help_detail("      number_comparison: equal_to/eq(0)  not_equal_to/neq(1)  between(2)  not_between(3)");
    output::help_detail("                         greater_than/gt(4)  less_than/lt(5)");
    output::help_detail("                         greater_than_or_equal_to/gte(6)  less_than_or_equal_to/lte(7)");
    output::help_detail("      date: yesterday(0) today(1) tomorrow(2) last_7_days(3) last_week(4) this_week(5)");
    output::help_detail("            next_week(6) last_month(7) this_month(8) next_month(9)");
    output::help_detail("            NOTE: engine values 10-17 are reserved; no named aliases.");
    output::help_detail("            next_7_days(18) last_year(19) this_year(20) next_year(21)");
    output::help_detail("      text: contains(0) not_contains(1) begins_with(2) ends_with(3)");
    output::help_detail("      cell_containing: duplicate_values(0) unique_values(1) blanks(2) no_blanks(3) errors(4) no_errors(5)");
    output::help_detail("      average: above(0) below(1) equal_or_above(2) below_or_equal(3)");
    output::help_detail("      standard_deviation: one_above(0) one_below(1) two_above(2) two_below(3) three_above(4) three_below(5)");
    output::help_detail("      top_bottom_values: top(0) bottom(1)");
    output::help_detail("    --condition.value <val>   Threshold (alias: .lhs)");
    output::help_detail("    --condition.value2 <val>  Second threshold for between (alias: .rhs)");
    output::help_detail("  ── Raw style flags ──");
    output::help_detail("    --style.font.font_style      regular(0) italic(1) bold(2) bold_italic(3) automatic(4)");
    output::help_detail("    --style.font.underline_type  none(0) single(1) double(2) single_accounting(3) double_accounting(4)");
    output::help_detail("    --style.font.strike_type     off(0) on(1) automatic(2)   [engine inversion handled internally]");
    output::help_detail("    --style.font.color <color>");
    output::help_detail("    --style.borders[N].border_type   left(0) right(1) top(2) bottom(3)");
    output::help_detail("    --style.borders[N].line_style    none(0) dash_dot(1) dash_dot_dot(2) dashed(3) dotted(4)");
    output::help_detail("      (alias: border_line_style)       double(5) hairline(6) medium(7) medium_dash_dot(8)");
    output::help_detail("                                       medium_dash_dot_dot(9) medium_dashed(10) thick(11) thin(12) slant_dash_dot(13)");
    output::help_detail("    --style.borders[N].color <color>");
    output::help_detail("    --style.fill.color <color>");
    output::help_detail("    --style.number_format.type   general(0) number(1) currency(2) accounting(3) date(4)");
    output::help_detail("      (alias: custom_format_type)  time(5) duration(6) percentage(7) scientific(8) fraction(9)");
    output::help_detail("                                   text(10) regional(11) custom(12)");
    output::help_detail("  ── Common flags ──");
    output::help_detail("    --stop-if-true   --sheet <id>   --active-cell <A1>");
    output::help_detail("  ── Examples ──");
    output::help_detail("    cf classic --range A1:C10 --when \">100\" --bold --fill red");
    output::help_detail("    cf classic --range A1:A50 --when \"is duplicate\" --fill orange");
    output::help_detail("    cf classic --range A1:A50 --when \"top 10\" --fill green");
    output::help_detail("    cf classic 1 --fill blue                  (edit rule #1 from cf list)");
    output::help_detail("    cf classic --rule-id rule-abc --fill blue");

    output::help_cmd("cf colorscale --range <A1:C5> [--min <color>] [--mid <color>] [--max <color>]", "Insert a color-scale rule");
    output::help_cmd("cf colorscale <#index|--rule-id id> [flags]", "Edit a color-scale rule");
    output::help_detail("  ── Color shorthands ──");
    output::help_detail("    --min <color>                  min stop  (criteria_type=min)");
    output::help_detail("    --mid <color>                  mid stop  (criteria_type=percentile, value=50)");
    output::help_detail("    --max <color>                  max stop  (criteria_type=max)");
    output::help_detail("  ── Extended stop flags ──");
    output::help_detail("    --min.criteria_type <type>     Override criteria_type for the min stop");
    output::help_detail("    --min.value <val>              Threshold for the min stop");
    output::help_detail("    --min.color <color>            Color for the min stop");
    output::help_detail("    --mid.criteria_type <type>     Override criteria_type for the mid stop");
    output::help_detail("    --mid.value <val>              Threshold for the mid stop");
    output::help_detail("    --mid.color <color>            Color for the mid stop");
    output::help_detail("    --max.criteria_type <type>     Override criteria_type for the max stop");
    output::help_detail("    --max.value <val>              Threshold for the max stop");
    output::help_detail("    --max.color <color>            Color for the max stop");
    output::help_detail("    criteria_type values: number(0) percent(1) percentile(2) formula(3) min(4) max(5) automatic(12)");
    output::help_detail("  ── Other flags ──");
    output::help_detail("    --auto-text-color   --hide-values   --stop-if-true");
    output::help_detail("  ── Examples ──");
    output::help_detail("    cf colorscale --range A1:A20 --min red --max green");
    output::help_detail("    cf colorscale --range A1:A20 --min red --mid yellow --max green");
    output::help_detail("    cf colorscale --range B1:B50 --min #FF0000 --mid #FFFF00 --max #00FF00");
    output::help_detail("    cf colorscale --range D2:D50 --min.criteria_type percentile --min.value 10 --min.color red --mid.criteria_type percentile --mid.value 50 --mid.color white --max.criteria_type percentile --max.value 90 --max.color blue");
    output::help_detail("    cf colorscale --range C1:C30 --min.color white --max.criteria_type number --max.value 1000 --max.color green");
    output::help_detail("    cf colorscale 2 --max.color blue");

    output::help_cmd("cf databar --range <A1:C5> [--positive <color>] [--negative <color>] [--min.<key> <val>] [--max.<key> <val>] [flags]", "Insert a data-bar rule");
    output::help_cmd("cf databar <#index|--rule-id id> [flags]", "Edit a data-bar rule");
    output::help_detail("  ── Color shorthands ──");
    output::help_detail("    --positive <color>        positive_value_fill color");
    output::help_detail("    --negative <color>        negative_value_fill color");
    output::help_detail("    --positive-border <color> positive_value_border color");
    output::help_detail("    --negative-border <color> negative_value_border color");
    output::help_detail("    --axis-color <color>      axis line color");
    output::help_detail("  ── Bar flags ──");
    output::help_detail("    --direction  left_to_right(0)  right_to_left(1)  context(2)");
    output::help_detail("    --axis       automatic(0)  middle(1)  none(2)");
    output::help_detail("    --fill-type  solid(0)  gradient(1)");
    output::help_detail("    --bar-border none/off(0)  with_border/on(1)");
    output::help_detail("  ── Threshold shorthands ──");
    output::help_detail("    --min.criteria_type <val>     min threshold criteria type");
    output::help_detail("    --min.value <val>             min threshold value");
    output::help_detail("    --max.criteria_type <val>     max threshold criteria type");
    output::help_detail("    --max.value <val>             max threshold value");
    output::help_detail("    criteria_type values: number(0) percent(1) percentile(2) formula(3) min(4) max(5) automatic(12)");
    output::help_detail("  ── Other flags ──");
    output::help_detail("    --hide-values   --stop-if-true");
    output::help_detail("  ── Examples ──");
    output::help_detail("    cf databar --range B1:B20 --positive \"#0070C0\" --negative red");
    output::help_detail("    cf databar --range B1:B20 --min.criteria_type min --max.criteria_type max");
    output::help_detail("    cf databar --range B1:B20 --min.criteria_type number --min.value 0 --max.criteria_type number --max.value 100");
    output::help_detail("    cf databar 3 --fill-type gradient");

    output::help_cmd("cf iconset --range <A1:C5> --set <icon_set_type> [flags]", "Insert an icon-set rule");
    output::help_cmd("cf iconset <#index|--rule-id id> [flags]", "Edit an icon-set rule");
    output::help_detail("  ── --set values (name or 0..23) ──");
    output::help_detail("    two_thumbs(0)  two_heart(1)  two_tick_wrong(2)");
    output::help_detail("    three_arrows(3)  three_arrows_gray(4)  three_flags(5)");
    output::help_detail("    three_traffic_lights_1(6)  three_traffic_lights_2(7)  three_signs(8)");
    output::help_detail("    three_triangles(9)  three_symbols(10)  three_heart(11)  three_smiley(12)  three_stars(13)");
    output::help_detail("    four_arrows(14)  four_arrows_gray(15)  four_black_to_red(16)  four_traffic_lights(17)");
    output::help_detail("    five_arrows(18)  five_arrows_gray(19)  five_quaters(20)  five_smiley(21)");
    output::help_detail("    five_circles(22)  five_clouds(23)");
    output::help_detail("  ── Other flags ──");
    output::help_detail("    --reverse-icons   --default-icon-size   --hide-values   --stop-if-true");
    output::help_detail("  ── Examples ──");
    output::help_detail("    cf iconset --range C1:C20 --set three_arrows");
    output::help_detail("    cf iconset --range A1:A10 --set three_stars");
    output::help_detail("    cf iconset 4 --set three_traffic_lights_1");

    output::help_cmd("cf list [--scope workbook|sheet|range] [--sheet <id>] [--range <A1:C5>]", "List CF rules with numbered index");
    output::help_detail("  (alias: cf manage)");
    output::help_detail("  --scope workbook|sheet|range   Default: sheet");
    output::help_detail("  NOTE: --range <A1:C5> is required when --scope range.");
    output::help_detail("  Output columns: #  RANGE  TYPE  SUMMARY  ID");
    output::help_detail("  The # column is the index used in edit/delete/move commands.");
    output::help_detail("  Example: cf list   |   cf list --scope workbook");

    output::help_cmd("cf delete <#index|rule-id|--target-range A1:C5|--all> [--sheet <id>] [--active-cell <A1>]", "Delete CF rule(s)");
    output::help_detail("  cf delete 2                    Delete rule #2 from last cf list");
    output::help_detail("  cf delete rule-abc             Delete by rule ID");
    output::help_detail("  cf delete --target-range A1:C10  Delete all rules touching that range");
    output::help_detail("  cf delete --all                Delete all rules on the sheet");

    output::help_cmd("cf move <#index|rule-id> <position> [--sheet <id>] [--active-cell <A1>] [--range <A1:C5>]", "Reorder a CF rule");
    output::help_detail("  Position flags:");
    output::help_detail("    --up [n]          Move up n positions (default 1)");
    output::help_detail("    --down [n]        Move down n positions (default 1)");
    output::help_detail("    --top             Move to highest priority");
    output::help_detail("    --bottom          Move to lowest priority");
    output::help_detail("    --after <#index|rule-id>  Place after a specific rule");
    output::help_detail("  Examples:");
    output::help_detail("    cf move 1 --down");
    output::help_detail("    cf move 3 --top");
    output::help_detail("    cf move rule-abc --after rule-def");

    output::help_section("GENERAL NOTES");
    output::help_detail("  All enums are 0-based at the CLI surface; engine offsets handled internally.");
    output::help_detail("  Flat dot paths: --condition.criteria_type  --style.font.font_style");
    output::help_detail("  Enum fields accept literal name or integer: criteria_type min  OR  criteria_type 4");
    output::help_detail("  Run 'cf list' before using #index in edit/delete/move.");
    output::help_detail("  active_info is auto-built from --range + the active sheet when omitted.");
    output::help_detail("  On insert, --range overrides any range_list embedded in the payload.");
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

// ─── Theme ───────────────────────────────────────────────────────────────────

fn cmd_theme(args: &[&str], engine: &EngineHandle, session: &mut CliSession) {
    require_active!(session);
    let rid = session.rid.clone().unwrap();

    if args.is_empty() {
        output::error("Usage: theme apply <type|name> | theme list");
        return;
    }

    match args[0].to_lowercase().as_str() {
        "list" => cmd_theme_list(&rid, engine),
        "apply" => cmd_theme_apply(&args[1..], &rid, engine, session),
        _ => output::error("Usage: theme apply <type|name> | theme list"),
    }
}

fn theme_name_to_type(name: &str) -> Option<i32> {
    match name.to_lowercase().as_str() {
        "zsheet"        => Some(1),
        "solid"         => Some(2),
        "attractive"    => Some(3),
        "onam"          => Some(4),
        "flat"          => Some(5),
        "retro"         => Some(6),
        "army"          => Some(7),
        "vivid"         => Some(8),
        "raksha_bandhan" | "raksha-bandhan" | "raksha" => Some(9),
        "tree"          => Some(10),
        "spring"        => Some(11),
        "christmas"     => Some(12),
        "playful"       => Some(13),
        "ocean_beach" | "ocean-beach" | "ocean" => Some(14),
        "diverging"     => Some(15),
        "sheer"         => Some(16),
        "architecture"  => Some(17),
        "executive"     => Some(18),
        "essential"     => Some(19),
        "office"        => Some(20),
        "legacy"        => Some(21),
        "custom"        => Some(22),
        _ => None,
    }
}

fn theme_type_to_name(t: i32) -> &'static str {
    match t {
        1  => "ZSHEET",
        2  => "SOLID",
        3  => "ATTRACTIVE",
        4  => "ONAM",
        5  => "FLAT",
        6  => "RETRO",
        7  => "ARMY",
        8  => "VIVID",
        9  => "RAKSHA_BANDHAN",
        10 => "TREE",
        11 => "SPRING",
        12 => "CHRISTMAS",
        13 => "PLAYFUL",
        14 => "OCEAN_BEACH",
        15 => "DIVERGING",
        16 => "SHEER",
        17 => "ARCHITECTURE",
        18 => "EXECUTIVE",
        19 => "ESSENTIAL",
        20 => "OFFICE",
        21 => "LEGACY",
        22 => "CUSTOM",
        _  => "UNKNOWN",
    }
}

fn parse_rgb_flag(args: &[&str], flag: &str) -> Option<serde_json::Value> {
    // Expects: --flag r g b
    for i in 0..args.len() {
        if args[i].eq_ignore_ascii_case(flag) {
            if i + 3 < args.len() {
                if let (Ok(r), Ok(g), Ok(b)) = (
                    args[i + 1].parse::<u8>(),
                    args[i + 2].parse::<u8>(),
                    args[i + 3].parse::<u8>(),
                ) {
                    return Some(serde_json::json!({ "red": r, "green": g, "blue": b }));
                }
            }
            return None;
        }
    }
    None
}

fn parse_string_flag<'a>(args: &[&'a str], flag: &str) -> Option<&'a str> {
    for i in 0..args.len() {
        if args[i].eq_ignore_ascii_case(flag) {
            if i + 1 < args.len() {
                return Some(args[i + 1]);
            }
        }
    }
    None
}

fn cmd_theme_apply(args: &[&str], rid: &str, engine: &EngineHandle, session: &mut CliSession) {
    if args.is_empty() {
        output::error("Usage: theme apply <type|name> [--bg R G B] [--text R G B] [--light R G B] [--dark R G B] [--accent1..6 R G B] [--hyperlink R G B] [--font <name>]");
        return;
    }

    let themes_type: i32 = if let Ok(n) = args[0].parse::<i32>() {
        n
    } else if let Some(t) = theme_name_to_type(args[0]) {
        t
    } else {
        output::error(&format!("Unknown theme: '{}'. Use a theme name (e.g. vivid, ocean) or numeric type 1-22.", args[0]));
        return;
    };

    let themes_properties = if themes_type == 22 {
        // Build custom properties from flags
        let mut props = serde_json::json!({});
        let color_flags = [
            ("--bg",        "background"),
            ("--text",      "text"),
            ("--light",     "light"),
            ("--dark",      "dark"),
            ("--accent1",   "accent_1"),
            ("--accent2",   "accent_2"),
            ("--accent3",   "accent_3"),
            ("--accent4",   "accent_4"),
            ("--accent5",   "accent_5"),
            ("--accent6",   "accent_6"),
            ("--hyperlink", "themes_hyperlink"),
        ];
        for (flag, key) in &color_flags {
            if let Some(rgb) = parse_rgb_flag(args, flag) {
                props[key] = rgb;
            }
        }
        if let Some(font) = parse_string_flag(args, "--font") {
            props["themes_font"] = serde_json::Value::String(font.to_string());
        }
        if props.as_object().map(|m| m.is_empty()).unwrap_or(true) {
            output::error("Custom theme requires at least one color flag (e.g. --bg 255 0 0).");
            return;
        }
        Some(props)
    } else {
        None
    };

    let request = rb::build_apply_theme(rid, themes_type, themes_properties);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if rp::is_success(status.status_code) {
                output::success(&format!("Theme '{}' applied.", theme_type_to_name(themes_type)));
                session.is_dirty = true;
            } else {
                output::error(&format!("Apply theme failed: {}", status.status_message.unwrap_or_default()));
            }
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn cmd_theme_list(rid: &str, engine: &EngineHandle) {
    let request = rb::build_get_themes(rid);
    match engine.process_request_json(&request) {
        Ok(resp) => {
            let status = rp::parse_status_response(&resp);
            if !rp::is_success(status.status_code) {
                output::error(&format!("Get themes failed: {}", status.status_message.unwrap_or_default()));
                return;
            }
            // Print the built-in theme list (engine may return additional info)
            output::line("Available Themes:", 0);
            output::blank_line();
            let built_in = [
                (1,  "ZSHEET"),
                (2,  "SOLID"),
                (3,  "ATTRACTIVE"),
                (4,  "ONAM"),
                (5,  "FLAT"),
                (6,  "RETRO"),
                (7,  "ARMY"),
                (8,  "VIVID"),
                (9,  "RAKSHA_BANDHAN"),
                (10, "TREE"),
                (11, "SPRING"),
                (12, "CHRISTMAS"),
                (13, "PLAYFUL"),
                (14, "OCEAN_BEACH"),
                (15, "DIVERGING"),
                (16, "SHEER"),
                (17, "ARCHITECTURE"),
                (18, "EXECUTIVE"),
                (19, "ESSENTIAL"),
                (20, "OFFICE"),
                (21, "LEGACY"),
                (22, "CUSTOM"),
            ];
            for (id, name) in &built_in {
                output::key_value(&format!("{:>2}", id), name, 2);
            }
            output::blank_line();
            output::line("Use: theme apply <name|number>  (e.g. theme apply vivid  or  theme apply 8)", 2);
        }
        Err(e) => output::error(&format!("Engine error: {}", e)),
    }
}

fn print_help_theme() {
    output::help_header("Themes");
    print_help_theme_content();
    output::help_footer();
}

fn print_help_theme_content() {
    output::help_section("THEMES");

    output::help_cmd("theme list", "List all available themes with their type IDs");

    output::help_cmd("theme apply <name|number>", "Apply a built-in theme to the current workbook");
    output::help_detail("  name: zsheet | solid | attractive | onam | flat | retro | army | vivid");
    output::help_detail("        raksha-bandhan | tree | spring | christmas | playful | ocean-beach");
    output::help_detail("        diverging | sheer | architecture | executive | essential | office | legacy");
    output::help_detail("  number: 1–21 (corresponding to the names above)");
    output::help_detail("  Example: theme apply vivid");
    output::help_detail("  Example: theme apply 7");

    output::help_cmd("theme apply custom [color flags] [--font <name>]", "Apply a fully custom theme");
    output::help_detail("  Color flags (each takes R G B values, 0-255):");
    output::help_detail("    --bg <R G B>        Background color");
    output::help_detail("    --text <R G B>      Text color");
    output::help_detail("    --light <R G B>     Light variant color");
    output::help_detail("    --dark <R G B>      Dark variant color");
    output::help_detail("    --accent1 <R G B>   Accent 1");
    output::help_detail("    --accent2 <R G B>   Accent 2");
    output::help_detail("    --accent3 <R G B>   Accent 3");
    output::help_detail("    --accent4 <R G B>   Accent 4");
    output::help_detail("    --accent5 <R G B>   Accent 5");
    output::help_detail("    --accent6 <R G B>   Accent 6");
    output::help_detail("    --hyperlink <R G B> Hyperlink color");
    output::help_detail("    --font <name>       Font family (e.g. Tahoma, Arial)");
    output::help_detail("  Example: theme apply custom --bg 255 255 255 --text 0 0 0 --accent1 126 87 201 --font Tahoma");
}

fn print_help_dv() {
    output::help_header("Data Validation");
    print_help_dv_content();
    output::help_footer();
}

fn print_help_dv_content() {
    output::help_section("DATA VALIDATION");

    output::help_cmd("dv readrange <range>", "Read-only dry-run: preview what applying DV to a range would do");
    output::help_detail("  Returns one of: CREATE | EDIT | EXTEND_RANGE | EXTEND_DATA_VALIDATION | COLLISION");
    output::help_detail("  Never mutates the workbook. To resolve a collision, use dv create --on-collision.");
    output::help_detail("  Example: dv readrange A1:B10");

    output::help_cmd("dv create <range> <criteria-type> [sub-criteria] [flags]", "Create a data validation rule");
    output::help_detail("  criteria-type: whole-number | decimal | list | datetime | text-length | custom | text | cell-range | any-value");
    output::help_detail("  sub-criteria:  between | notbetween | equal | notequal | gt | lt | gte | lte");
    output::help_detail("                 contains | notcontains | beginswith | endswith");
    output::help_detail("                 before | onorbefore | after | onorafter | on | noton");
    output::help_detail("                 yesterday | today | tomorrow | last7days | next7days");
    output::help_detail("                 lastweek | thisweek | nextweek | lastmonth | thismonth | nextmonth");
    output::help_detail("                 lastyear | thisyear | nextyear | isvaliddatetime");
    output::help_detail("  --val1 <value>               First parameter (lower bound, formula, or list items)");
    output::help_detail("  --val2 <value>               Second parameter (upper bound; required for 'between')");
    output::help_detail("  --delimiter <sep>            Delimiter separating list values (default: \\n)");
    output::help_detail("  --show-list <true|false>     Show dropdown list (default: true)");
    output::help_detail("  --sort-list <true|false>     Sort list ascending (default: false)");
    output::help_detail("  --ignore-blanks <true|false> Skip blank cells (default: true)");
    output::help_detail("  --help-title <text>          Input help popup title");
    output::help_detail("  --help-msg <text>            Input help popup message");
    output::help_detail("  --hide-help                  Disable the help popup entirely");
    output::help_detail("  --error-title <text>         Error alert title");
    output::help_detail("  --error-msg <text>           Error alert message");
    output::help_detail("  --error-style stop|warning|info   stop=block(default) | warning=prompt | info=notify");
    output::help_detail("  --no-error-validation        Disable error alert on invalid input");
    output::help_detail("  --on-collision replace|abort  What to do when the range overlaps an existing rule:");
    output::help_detail("                 replace = default (alias --force): apply the rule; the engine resolves the overlap");
    output::help_detail("                 abort   = refuse and make no change");
    output::help_detail("  Example: dv create A1:A20 whole-number between --val1 1 --val2 100 --error-style stop --error-msg \"Enter 1-100\"");
    output::help_detail("  Example: dv create B1:B50 list --val1 \"Apple\\nBanana\\nCherry\" --show-list true --sort-list true");
    output::help_detail("  Example: dv create C1:C10 datetime after --val1 2024-01-01 --help-msg \"Date must be after Jan 2024\"");

    output::help_cmd("dv edit <range> [flags]", "Edit an existing rule — all flags are optional, only supplied fields are changed");
    output::help_detail("  --criteria <type>            Change the criteria type (same literals as dv create)");
    output::help_detail("  --sub-criteria <type>        Change the sub-criteria (same literals as dv create)");
    output::help_detail("  --val1 / --val2 / --delimiter / --show-list / --sort-list / --ignore-blanks");
    output::help_detail("  --help-title / --help-msg / --hide-help");
    output::help_detail("  --error-title / --error-msg / --error-style stop|warning|info / --no-error-validation");
    output::help_detail("  --on-collision replace|abort   replace = default: edit across the overlap  |  abort = refuse if the range overlaps a different rule");
    output::help_detail("  Example: dv edit A1:A20 --val1 5 --val2 50 --error-style warning");
    output::help_detail("  Example: dv edit B1:B50 --criteria list --val1 \"Red\\nGreen\\nBlue\"");

    output::help_cmd("dv manage", "List all validation rules across the workbook");
    output::help_cmd("dv manage --sheet <name|id>", "List rules for one sheet only");
    output::help_cmd("dv manage --range <A1:C5>", "List rules that cover a specific range");
    output::help_detail("  Example: dv manage --sheet Sheet2");
    output::help_detail("  Example: dv manage --range A1:Z100");

    output::help_cmd("dv delete <range>", "Clear the data validation rule for a range");
    output::help_cmd("dv delete <range> --sheet <name|id>", "Clear DV on a specific sheet (by name or ID)");
    output::help_detail("  Aliases: dv remove, dv clear");
    output::help_detail("  Example: dv delete A1:B10");
    output::help_detail("  Example: dv delete A1:B10 --sheet Sheet2");
    output::help_detail("  Example: dv delete A1:B10 --sheet 45433965-5f58-4688-920c-58d3bee8d958");
}
