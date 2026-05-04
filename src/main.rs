mod commands;
mod engine;
mod output;
mod session;
mod util;

use std::env;
use std::io::{self, BufRead};

use rustyline::error::ReadlineError;
use rustyline::DefaultEditor;

use engine::ffi::EngineHandle;
use session::CliSession;

/// Maximum allowed length for any single argument.
const MAX_INPUT_ARG_LEN: usize = 32_768;

fn main() {
    let args: Vec<String> = env::args().collect();

    // Initialise native engine
    let engine = match EngineHandle::load(None) {
        Ok(e) => e,
        Err(err) => {
            output::error(&format!("Failed to load native engine: {}", err));
            std::process::exit(1);
        }
    };

    if !engine::initializer::initialize(&engine) {
        output::error("Engine initialization failed.");
        std::process::exit(1);
    }

    let mut session = CliSession::default();

    // If arguments are given on the command line, run in single-command mode.
    // Syntax: zs-cli <command> [args...]  OR  zs-cli --script <file>
    if args.len() > 1 {
        if args[1] == "--script" || args[1] == "-s" {
            if args.len() < 3 {
                output::error("Usage: zs-cli --script <file>");
                std::process::exit(1);
            }
            run_script(&args[2], &engine, &mut session);
        } else {
            let cmd_tokens: Vec<&str> = args[1..].iter().map(|s| s.as_str()).collect();
            commands::dispatch(&cmd_tokens, &engine, &mut session);
        }
        return;
    }

    // Interactive REPL
    run_repl(&engine, &mut session);
}

/// Run the interactive REPL (read-eval-print loop).
fn run_repl(engine: &EngineHandle, session: &mut CliSession) {
    output::print_banner();
    output::info("Type 'help' for available commands, 'exit' to quit.\n");

    let mut rl = match DefaultEditor::new() {
        Ok(r) => r,
        Err(_) => {
            output::warning("Line editing unavailable. Falling back to basic input.");
            run_repl_basic(engine, session);
            return;
        }
    };

    // Try to load history
    let history_path = dirs::data_local_dir()
        .unwrap_or_else(|| std::path::PathBuf::from("."))
        .join("ZsCLI")
        .join("history.txt");
    let _ = std::fs::create_dir_all(history_path.parent().unwrap_or(std::path::Path::new(".")));
    let _ = rl.load_history(&history_path);

    loop {
        let prompt = build_prompt(session);
        match rl.readline(&prompt) {
            Ok(line) => {
                let trimmed = line.trim();
                if trimmed.is_empty() || trimmed.starts_with('#') {
                    continue;
                }
                let _ = rl.add_history_entry(trimmed);

                let lower = trimmed.to_lowercase();
                if lower == "exit" || lower == "quit" {
                    if session.is_active() && session.is_dirty {
                        let resp = output::confirm("You have unsaved changes. Save before exit?");
                        match resp.as_str() {
                            "c" => continue,
                            "y" => { commands::dispatch(&["save"], engine, session); },
                            _ => {}
                        }
                    }
                    output::info("Goodbye.");
                    break;
                }

                let tokens = tokenize_input(trimmed);
                let refs: Vec<&str> = tokens.iter().map(|s| s.as_str()).collect();
                commands::dispatch(&refs, engine, session);
            }
            Err(ReadlineError::Interrupted | ReadlineError::Eof) => {
                output::info("Goodbye.");
                break;
            }
            Err(err) => {
                output::error(&format!("Input error: {}", err));
                break;
            }
        }
    }

    let _ = rl.save_history(&history_path);
}

/// Fallback basic REPL for environments without terminal capabilities.
fn run_repl_basic(engine: &EngineHandle, session: &mut CliSession) {
    let stdin = io::stdin();
    loop {
        let prompt = build_prompt(session);
        eprint!("{}", prompt);
        let mut line = String::new();
        match stdin.lock().read_line(&mut line) {
            Ok(0) => break, // EOF
            Ok(_) => {
                let trimmed = line.trim();
                if trimmed.is_empty() || trimmed.starts_with('#') {
                    continue;
                }
                let lower = trimmed.to_lowercase();
                if lower == "exit" || lower == "quit" {
                    if session.is_active() && session.is_dirty {
                        let resp = output::confirm("You have unsaved changes. Save before exit?");
                        match resp.as_str() {
                            "c" => continue,
                            "y" => { commands::dispatch(&["save"], engine, session); },
                            _ => {}
                        }
                    }
                    output::info("Goodbye.");
                    break;
                }
                let tokens = tokenize_input(trimmed);
                let refs: Vec<&str> = tokens.iter().map(|s| s.as_str()).collect();
                commands::dispatch(&refs, engine, session);
            }
            Err(e) => {
                output::error(&format!("Input error: {}", e));
                break;
            }
        }
    }
}

/// Build the REPL prompt string.
///
/// ```text
/// zs-cli> (no workbook open)
/// zs-cli [Budget | Sheet1]> (workbook named "Budget", active sheet "Sheet1")
/// ```
fn build_prompt(session: &CliSession) -> String {
    if !session.is_active() {
        return "zs-cli> ".to_string();
    }
    let wb = session
        .workbook_name
        .as_deref()
        .unwrap_or("workbook");
    let sheet = session
        .active_sheet_name
        .as_deref()
        .unwrap_or("Sheet1");
    format!("zs-cli [{} | {}]> ", wb, sheet)
}

/// Quote-aware tokenizer.
///
/// `cell set A1 "hello world"` → `["cell", "set", "A1", "hello world"]`
///
/// Handles both single and double quotes. Escaped quotes inside a quoted
/// section are not currently supported.
fn tokenize_input(input: &str) -> Vec<String> {
    let mut tokens: Vec<String> = Vec::new();
    let mut current = String::new();
    let mut in_quote = false;
    let mut quote_char: char = '"';

    for ch in input.chars() {
        if in_quote {
            if ch == quote_char {
                in_quote = false;
            } else {
                current.push(ch);
            }
        } else if ch == '"' || ch == '\'' {
            in_quote = true;
            quote_char = ch;
        } else if ch.is_whitespace() {
            if !current.is_empty() {
                tokens.push(std::mem::take(&mut current));
            }
        } else {
            current.push(ch);
        }
    }
    if !current.is_empty() {
        tokens.push(current);
    }
    tokens
}

/// Execute commands from a script file, one line at a time.
fn run_script(path: &str, engine: &EngineHandle, session: &mut CliSession) {
    let content = match std::fs::read_to_string(path) {
        Ok(c) => c,
        Err(e) => {
            output::error(&format!("Cannot read script '{}': {}", path, e));
            return;
        }
    };

    for (lineno, line) in content.lines().enumerate() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            continue;
        }
        output::info(&format!("[{}:{}] {}", path, lineno + 1, trimmed));
        let tokens = tokenize_input(trimmed);
        let refs: Vec<&str> = tokens.iter().map(|s| s.as_str()).collect();
        let should_continue = commands::dispatch(&refs, engine, session);
        if !should_continue {
            break;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_tokenize_simple() {
        let tokens = tokenize_input("cell set A1 hello");
        assert_eq!(tokens, vec!["cell", "set", "A1", "hello"]);
    }

    #[test]
    fn test_tokenize_quoted() {
        let tokens = tokenize_input(r#"cell set A1 "hello world""#);
        assert_eq!(tokens, vec!["cell", "set", "A1", "hello world"]);
    }

    #[test]
    fn test_tokenize_single_quoted() {
        let tokens = tokenize_input("cell set A1 'hello world'");
        assert_eq!(tokens, vec!["cell", "set", "A1", "hello world"]);
    }

    #[test]
    fn test_tokenize_empty() {
        let tokens = tokenize_input("   ");
        assert!(tokens.is_empty());
    }

    #[test]
    fn test_build_prompt_no_session() {
        let session = CliSession::default();
        assert_eq!(build_prompt(&session), "zs-cli> ");
    }

    #[test]
    fn test_build_prompt_active() {
        let mut session = CliSession::default();
        session.rid = Some("abc".to_string());
        session.workbook_name = Some("Budget".to_string());
        session.active_sheet_name = Some("Sheet1".to_string());
        assert_eq!(build_prompt(&session), "zs-cli [Budget | Sheet1]> ");
    }
}
