/// Centralizes all console output for the CLI.
/// Port of C# `CliPrinter` using `crossterm` for cross-platform coloured output.
use std::io::{self, Write};

use crossterm::style::{Color, ResetColor, SetForegroundColor};
use crossterm::{execute, queue};

// ─── Public API ──────────────────────────────────────────────────────────────

/// Prints a success message prefixed with a green check mark.
pub fn success(msg: &str) {
    let mut out = io::stdout();
    let _ = execute!(out, SetForegroundColor(Color::Green));
    print!("\u{2714} ");
    let _ = execute!(out, ResetColor);
    println!("{msg}");
}

/// Prints an error message prefixed with a red cross mark.
pub fn error(msg: &str) {
    let mut out = io::stdout();
    let _ = execute!(out, SetForegroundColor(Color::Red));
    print!("\u{2718} ");
    let _ = execute!(out, ResetColor);
    println!("{msg}");
}

/// Prints an informational message prefixed with a cyan info symbol.
pub fn info(msg: &str) {
    let mut out = io::stdout();
    let _ = execute!(out, SetForegroundColor(Color::Cyan));
    print!("\u{2139} ");
    let _ = execute!(out, ResetColor);
    println!("{msg}");
}

/// Prints a warning message prefixed with a yellow warning symbol.
pub fn warning(msg: &str) {
    let mut out = io::stdout();
    let _ = execute!(out, SetForegroundColor(Color::Yellow));
    print!("\u{26A0} ");
    let _ = execute!(out, ResetColor);
    println!("{msg}");
}

/// Prints a plain line of text with optional indentation.
pub fn line(msg: &str, indent: usize) {
    if indent > 0 {
        print!("{}", " ".repeat(indent));
    }
    println!("{msg}");
}

/// Prints a key-value pair formatted with consistent alignment.
pub fn key_value(key: &str, value: &str, indent: usize) {
    if indent > 0 {
        print!("{}", " ".repeat(indent));
    }
    println!("{:<14}: {}", key, value);
}

/// Prints a blank line for visual separation.
pub fn blank_line() {
    println!();
}

/// Prompts the user with a yes/no/cancel question and returns `"y"`, `"n"`, or `"c"`.
pub fn confirm(msg: &str) -> String {
    let mut out = io::stdout();
    let _ = execute!(out, SetForegroundColor(Color::Yellow));
    print!("\u{26A0} ");
    let _ = execute!(out, ResetColor);
    print!("{msg} ");
    let _ = execute!(out, SetForegroundColor(Color::DarkGrey));
    print!("(y/n/c) ");
    let _ = execute!(out, ResetColor);
    let _ = out.flush();

    loop {
        let mut buf = String::new();
        if io::stdin().read_line(&mut buf).is_err() {
            return "c".to_string();
        }
        let trimmed = buf.trim().to_lowercase();
        if trimmed.is_empty() {
            continue;
        }
        match trimmed.as_str() {
            "y" | "yes" => return "y".to_string(),
            "n" | "no" => return "n".to_string(),
            "c" | "cancel" => return "c".to_string(),
            _ => {
                print!("  Please enter y (yes), n (no), or c (cancel): ");
                let _ = out.flush();
            }
        }
    }
}

// ─── Box-drawing constants ───────────────────────────────────────────────────

const BOX_H: char = '\u{2550}'; // ═
const BOX_V: char = '\u{2551}'; // ║
const BOX_TL: char = '\u{2554}'; // ╔
const BOX_TR: char = '\u{2557}'; // ╗
const BOX_BL: char = '\u{255A}'; // ╚
const BOX_BR: char = '\u{255D}'; // ╝
const THIN_H: char = '\u{2500}'; // ─
const THIN_LT: char = '\u{255F}'; // ╟
const THIN_RT: char = '\u{2562}'; // ╢
const BOX_W: usize = 80; // inner width between ║…║

/// Prints the startup banner with the ASCII logo.
pub fn print_banner() {
    let mut out = io::stdout();
    println!();

    // ── Top border
    print_box_border(&mut out, BOX_TL, BOX_TR);

    // ── Empty line above logo
    print_box_empty(&mut out);

    // ── ASCII logo
    let logo: &[&str] = &[
        r#" ______     _              _____ _               _      _____ _      _____  "#,
        r#"|___  /    | |            / ____| |             | |    / ____| |    |_   _| "#,
        r#"   / / ___ | |__   ___   | (___ | |__   ___  ___| |_  | |    | |      | |   "#,
        r#"  / / / _ \| '_ \ / _ \   \___ \| '_ \ / _ \/ _ \ __| | |    | |      | |   "#,
        r#" / /_| (_) | | | | (_) |  ____) | | | |  __/  __/ |_  | |____| |____ _| |_  "#,
        r#"/_____\___/|_| |_|\___/  |_____/|_| |_|\___|\___|\__|  \_____|______|_____| "#,
    ];
    for l in logo {
        print_box_line(&mut out, l, Color::Green, true);
    }

    // ── Empty line below logo
    print_box_empty(&mut out);

    // ── Thin divider
    print_thin_divider(&mut out);

    // ── Tagline + version
    print_box_empty(&mut out);
    print_box_line(
        &mut out,
        "Interactive Spreadsheet Engine for Desktop",
        Color::DarkCyan,
        true,
    );
    print_box_line(
        &mut out,
        "v1.0.20  \u{2502}  Type 'help' for commands",
        Color::DarkGrey,
        true,
    );
    print_box_empty(&mut out);

    // ── Bottom border
    print_box_border(&mut out, BOX_BL, BOX_BR);
    println!();
}

// ─── Box-drawing helpers ─────────────────────────────────────────────────────

fn print_box_border(out: &mut io::Stdout, left: char, right: char) {
    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    let bar: String = std::iter::repeat(BOX_H).take(BOX_W).collect();
    print!("{left}{bar}{right}");
    let _ = queue!(out, ResetColor);
    println!();
}

fn print_thin_divider(out: &mut io::Stdout) {
    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    let bar: String = std::iter::repeat(THIN_H).take(BOX_W).collect();
    print!("{THIN_LT}{bar}{THIN_RT}");
    let _ = queue!(out, ResetColor);
    println!();
}

fn print_box_empty(out: &mut io::Stdout) {
    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    let padding = " ".repeat(BOX_W);
    print!("{BOX_V}{padding}{BOX_V}");
    let _ = queue!(out, ResetColor);
    println!();
}

fn print_box_line(out: &mut io::Stdout, text: &str, color: Color, center: bool) {
    let visible_len = text.len();
    let pad = BOX_W.saturating_sub(visible_len);
    let left_pad = if center { pad / 2 } else { 0 };
    let right_pad = pad - left_pad;

    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    print!("{BOX_V}");
    let _ = queue!(out, ResetColor);

    let _ = queue!(out, SetForegroundColor(color));
    print!("{}{text}{}", " ".repeat(left_pad), " ".repeat(right_pad));
    let _ = queue!(out, ResetColor);

    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    print!("{BOX_V}");
    let _ = queue!(out, ResetColor);
    println!();
}

fn print_box_content(out: &mut io::Stdout, fragments: &[(Color, &str)]) {
    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    print!("{BOX_V}");
    let _ = queue!(out, ResetColor);

    let mut used: usize = 0;
    for &(color, text) in fragments {
        let _ = queue!(out, SetForegroundColor(color));
        print!("{text}");
        let _ = queue!(out, ResetColor);
        used += text.len();
    }
    let remaining = BOX_W.saturating_sub(used);
    if remaining > 0 {
        print!("{}", " ".repeat(remaining));
    }

    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    print!("{BOX_V}");
    let _ = queue!(out, ResetColor);
    println!();
}

fn print_section_header(out: &mut io::Stdout, title: &str) {
    let pad = BOX_W.saturating_sub(title.len() + 2); // 2 for leading spaces

    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    print!("{BOX_V}");
    let _ = queue!(out, ResetColor);

    let _ = queue!(out, SetForegroundColor(Color::White));
    print!("  {title}");
    let _ = queue!(out, ResetColor);
    print!("{}", " ".repeat(pad));

    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    print!("{BOX_V}");
    let _ = queue!(out, ResetColor);
    println!();
}

fn print_cmd_row(out: &mut io::Stdout, command: &str, description: &str) {
    const CMD_WIDTH: usize = 36;
    const INDENT: usize = 4;

    let cmd_part = if command.len() > CMD_WIDTH {
        &command[..CMD_WIDTH]
    } else {
        command
    };
    let cmd_pad = CMD_WIDTH.saturating_sub(cmd_part.len());

    let desc_space = BOX_W.saturating_sub(INDENT + CMD_WIDTH);
    let desc_part = if description.len() > desc_space {
        &description[..desc_space]
    } else {
        description
    };
    let desc_pad = desc_space.saturating_sub(desc_part.len());

    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    print!("{BOX_V}");
    let _ = queue!(out, ResetColor);

    print!("{}", " ".repeat(INDENT));

    let _ = queue!(out, SetForegroundColor(Color::Yellow));
    print!("{cmd_part}{}", " ".repeat(cmd_pad));
    let _ = queue!(out, ResetColor);

    let _ = queue!(out, SetForegroundColor(Color::DarkGrey));
    print!("{desc_part}{}", " ".repeat(desc_pad));
    let _ = queue!(out, ResetColor);

    let _ = queue!(out, SetForegroundColor(Color::DarkCyan));
    print!("{BOX_V}");
    let _ = queue!(out, ResetColor);
    println!();
}
