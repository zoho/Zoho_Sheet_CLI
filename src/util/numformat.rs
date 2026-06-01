/// Parameterized number format pattern generator.
/// Ports the logic from the Zoho Sheet UI's FormatCellPatternGenerator.

pub struct NumFormatParams {
    pub decimals: u8,
    pub separator: bool,
    pub leading_zeros: u8,
    pub negative: NegativeStyle,
    pub prefix: String,
    pub suffix: String,
    pub currency: String,
    pub digits: u8,
    pub date: String,
    pub time: String,
}

/// Named negative number styles (replaces magic integer codes).
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum NegativeStyle {
    Minus,      // 0: default, e.g. -1,234.00
    Red,        // 1: red font
    RedMinus,   // 2: red font with minus sign
    Parens,     // 3: parentheses, e.g. (1,234.00)
    RedParens,  // 4: red font with parentheses
}

impl NegativeStyle {
    /// Parse from named string or deprecated integer code.
    /// Returns (style, is_deprecated) where is_deprecated=true when an integer code is used.
    pub fn parse(s: &str) -> Result<(Self, bool), String> {
        match s {
            "minus" => Ok((Self::Minus, false)),
            "red" => Ok((Self::Red, false)),
            "red-minus" => Ok((Self::RedMinus, false)),
            "parens" => Ok((Self::Parens, false)),
            "red-parens" => Ok((Self::RedParens, false)),
            "0" => Ok((Self::Minus, true)),
            "1" => Ok((Self::Red, true)),
            "2" => Ok((Self::RedMinus, true)),
            "3" => Ok((Self::Parens, true)),
            "4" => Ok((Self::RedParens, true)),
            _ => Err(format!(
                "Invalid --negative style '{}'. Use: minus, red, red-minus, parens, red-parens",
                s
            )),
        }
    }

    pub fn as_u8(self) -> u8 {
        match self {
            Self::Minus => 0,
            Self::Red => 1,
            Self::RedMinus => 2,
            Self::Parens => 3,
            Self::RedParens => 4,
        }
    }
}

impl Default for NumFormatParams {
    fn default() -> Self {
        Self {
            decimals: 2,
            separator: true,
            leading_zeros: 1,
            negative: NegativeStyle::Minus,
            prefix: String::new(),
            suffix: String::new(),
            currency: String::new(),
            digits: 1,
            date: String::new(),
            time: String::new(),
        }
    }
}

/// Parse `--flag value` style arguments into NumFormatParams.
/// Returns None if no flags are detected (so caller can fall through to legacy behavior).
/// Prints deprecation warnings for --zeros and integer --negative codes.
pub fn parse_flags(args: &[&str]) -> Option<NumFormatParams> {
    if args.is_empty() || !args.iter().any(|a| a.starts_with("--")) {
        return None;
    }
    let mut params = NumFormatParams::default();
    let mut i = 0;
    while i < args.len() {
        match args[i] {
            "--decimals" | "-d" => {
                i += 1;
                if i < args.len() {
                    params.decimals = args[i].parse().unwrap_or(2);
                }
            }
            "--noseparator" => {
                params.separator = false;
            }
            "--leading-zeros" => {
                i += 1;
                if i < args.len() {
                    params.leading_zeros = args[i].parse().unwrap_or(1);
                }
            }
            // Hidden backward-compat alias for --leading-zeros
            "--zeros" | "-z" => {
                i += 1;
                if i < args.len() {
                    params.leading_zeros = args[i].parse().unwrap_or(1);
                }
            }
            "--negative" | "-n" => {
                i += 1;
                if i < args.len() {
                    match NegativeStyle::parse(args[i]) {
                        Ok((style, _deprecated)) => {
                            params.negative = style;
                        }
                        Err(_) => {
                            // Leave default
                        }
                    }
                }
            }
            "--prefix" => {
                i += 1;
                if i < args.len() {
                    params.prefix = args[i].to_string();
                }
            }
            "--suffix" => {
                i += 1;
                if i < args.len() {
                    params.suffix = args[i].to_string();
                }
            }
            "--currency" | "-c" => {
                i += 1;
                if i < args.len() {
                    params.currency = normalize_locale_key(args[i]);
                }
            }
            "--digits" => {
                i += 1;
                if i < args.len() {
                    params.digits = args[i].parse().unwrap_or(1);
                }
            }
            "--date" => {
                i += 1;
                if i < args.len() {
                    params.date = args[i].to_string();
                }
            }
            "--time" => {
                i += 1;
                if i < args.len() {
                    params.time = args[i].to_string();
                }
            }
            _ => {}
        }
        i += 1;
    }
    Some(params)
}

/// Normalize BCP 47 locale key (en-US) to engine format en(US).
/// Accepts both en-US and en(US) as input. Public for use by commands.
pub fn normalize_locale(input: &str) -> String {
    normalize_locale_key(input)
}

/// Normalize BCP 47 locale key (en-US) to engine format en(US).
/// Accepts both en-US and en(US) as input.
fn normalize_locale_key(input: &str) -> String {
    // If already in en(XX) format, pass through
    if input.contains('(') && input.contains(')') {
        return input.to_string();
    }
    // Convert en-US → en(US)
    if let Some(pos) = input.find('-') {
        let lang = &input[..pos];
        let region = &input[pos + 1..];
        return format!("{}({})", lang, region);
    }
    input.to_string()
}

/// Convert locale key to hyphen form for use in format patterns.
/// en(US) → en-US, already-hyphenated passes through.
fn to_hyphen_locale(input: &str) -> String {
    if let (Some(open), Some(close)) = (input.find('('), input.find(')')) {
        let lang = &input[..open];
        let region = &input[open + 1..close];
        return format!("{}-{}", lang, region);
    }
    input.to_string()
}

/// Build the base decimal format pattern (equivalent to C# GetDecimalFormat).
fn get_decimal_format(separator: bool, leading_zeros: u8, decimals: u8, prefix: &str, suffix: &str) -> String {
    let leading_zeros = leading_zeros as usize;
    let decimals = decimals as usize;
    let hash_count = if leading_zeros == 0 { 4 } else { (4usize).saturating_sub(leading_zeros).max(1) };

    let mut pattern = "#".repeat(hash_count);
    if leading_zeros > 0 {
        pattern.push_str(&"0".repeat(leading_zeros));
    }
    if separator {
        let mut chars: Vec<char> = pattern.chars().collect();
        if chars.len() > 1 {
            chars.insert(1, ',');
        }
        pattern = chars.into_iter().collect();
    }
    if decimals > 0 {
        pattern.push('.');
        pattern.push_str(&"0".repeat(decimals));
    }
    if !prefix.is_empty() {
        pattern = format!("\"{}\"{}", prefix, pattern);
    }
    if !suffix.is_empty() {
        pattern = format!("{}\"{}\"", pattern, suffix);
    }
    pattern
}

/// Generate a complete format pattern for the given type and parameters.
pub fn generate_pattern(type_str: &str, params: &NumFormatParams) -> Result<String, String> {
    // Validate --date/--time flag scope
    if !params.time.is_empty() && type_str == "date" {
        return Err(String::from(
            "--time flag is not valid for type 'date'. Use type 'custom' for combined date/time patterns."
        ));
    }
    if !params.date.is_empty() && type_str == "time" {
        return Err(String::from(
            "--date flag is not valid for type 'time'. Use type 'custom' for combined date/time patterns."
        ));
    }
    // Validate --prefix/--suffix flag scope
    let has_prefix_suffix = !params.prefix.is_empty() || !params.suffix.is_empty();
    if has_prefix_suffix {
        match type_str {
            "number" | "custom" => {} // supported
            _ => {
                let flag = if !params.prefix.is_empty() { "--prefix" } else { "--suffix" };
                return Err(format!(
                    "{} is not supported for type '{}'. Supported types: number, custom.",
                    flag, type_str
                ));
            }
        }
    }
    match type_str {
        "number" => Ok(generate_number_pattern(params)),
        "currency" => Ok(generate_currency_pattern(params)),
        "accounting" => Ok(generate_accounting_pattern(params)),
        "percentage" => Ok(generate_percentage_pattern(params)),
        "scientific" => Ok(generate_scientific_pattern(params)),
        "fraction" => Ok(generate_fraction_pattern(params)),
        "date" | "time" | "custom" => Ok(generate_datetime_pattern(params)),
        _ => Err(format!("Parameterized generation not supported for type '{}'", type_str)),
    }
}

fn generate_number_pattern(params: &NumFormatParams) -> String {
    let base = get_decimal_format(params.separator, params.leading_zeros, params.decimals, &params.prefix, &params.suffix);
    apply_negative_style_number(&base, params.negative)
}

fn apply_negative_style_number(pattern: &str, negative: NegativeStyle) -> String {
    match negative {
        NegativeStyle::Red => format!("{};[RED]{}", pattern, pattern),
        NegativeStyle::RedMinus => format!("{};[RED]-{}", pattern, pattern),
        NegativeStyle::Parens => format!("{}_);({})", pattern, pattern),
        NegativeStyle::RedParens => format!("{}_);[RED]({})", pattern, pattern),
        NegativeStyle::Minus => pattern.to_string(),
    }
}

fn generate_currency_pattern(params: &NumFormatParams) -> String {
    // Matches C# GetEngineCurrencyPattern: Pattern = "[" + currency + "]" + DecimalFormat
    // currency value from C# is "$en-US", so pattern becomes "[$en-US]#,##0.00"
    let base = get_decimal_format(params.separator, params.leading_zeros, params.decimals, "", "");
    let pattern = if !params.currency.is_empty() {
        let locale_for_pattern = to_hyphen_locale(&params.currency);
        format!("[${}]{}", locale_for_pattern, base)
    } else {
        base.clone()
    };
    apply_negative_style_currency(&pattern, params.negative)
}

fn generate_accounting_pattern(params: &NumFormatParams) -> String {
    let base = get_decimal_format(params.separator, params.leading_zeros, params.decimals, "", "");
    if !params.currency.is_empty() && !params.currency.eq_ignore_ascii_case("none") {
        // With currency: use accounting alignment wrappers around locale symbol
        // _([$en-US]* #,##0.00_);_([$en-US]* (#,##0.00);_([$en-US]* "-"??_);_(@_)
        let locale_for_pattern = to_hyphen_locale(&params.currency);
        let sym = format!("[${}]", locale_for_pattern);
        return format!(
            "_({}* {}_);_({}* ({});_({}* \"-\"??_);_(@_)",
            sym, base, sym, base, sym
        );
    }
    // No currency: use accounting alignment wrappers without symbol
    format!("_(* {}_);_(* ({});_(* \"-\"??_);_(@_)", base, base)
}

fn generate_percentage_pattern(params: &NumFormatParams) -> String {
    let base = get_decimal_format(false, 1, params.decimals, "", "");
    format!("{}%", base)
}

fn generate_scientific_pattern(params: &NumFormatParams) -> String {
    let decimals = params.decimals as usize;
    if decimals > 0 {
        format!("0.{}E+00", "0".repeat(decimals))
    } else {
        String::from("0E+00")
    }
}

fn generate_fraction_pattern(params: &NumFormatParams) -> String {
    let q_count = (params.digits as usize) + 1;
    let qs = "?".repeat(q_count);
    format!("# {}/{}", qs, qs)
}

fn generate_datetime_pattern(params: &NumFormatParams) -> String {
    let mut pattern = String::new();
    let date_str = if params.date == "none" { "" } else { params.date.as_str() };
    let time_str = if params.time == "none" { "" } else { params.time.as_str() };
    if !date_str.is_empty() {
        pattern.push_str(date_str);
        if !time_str.is_empty() {
            pattern.push(' ');
        }
    }
    if !time_str.is_empty() {
        pattern.push_str(time_str);
    }
    pattern
}

/// Apply negative style for currency patterns (matches C# switch on Negative).
fn apply_negative_style_currency(pattern: &str, negative: NegativeStyle) -> String {
    match negative {
        NegativeStyle::Minus => pattern.to_string(), // default: no explicit negative section
        NegativeStyle::Red => format!("{};[RED]{}", pattern, pattern),
        NegativeStyle::RedMinus => format!("{};[RED]-{}", pattern, pattern),
        NegativeStyle::Parens => format!("{}_);({})", pattern, pattern),
        NegativeStyle::RedParens => format!("{}_);[RED]({})", pattern, pattern),
    }
}

/// Resolve a locale key to a complete currency format pattern.
/// Used by the non-flag path in resolve_numformat_text when the user passes
/// a locale key directly (e.g. `format numformat A1 currency en-US`).
pub fn resolve_currency_locale(locale: &str) -> String {
    // Pattern uses [$en-US] form ($ prefix inside brackets)
    let hyphen = to_hyphen_locale(&normalize_locale_key(locale));
    let base = "#,##0.00"; // default 2 decimals with separator
    format!("[${}]{}", hyphen, base)
}

/// Resolve a locale key to a complete accounting format pattern.
pub fn resolve_accounting_locale(locale: &str) -> String {
    // With currency locale: accounting alignment wrappers around locale symbol
    let hyphen = to_hyphen_locale(&normalize_locale_key(locale));
    let base = "#,##0.00";
    let sym = format!("[${}]", hyphen);
    format!("_({}* {}_);_({}* ({});_({}* \"-\"??_);_(@_)", sym, base, sym, base, sym)
}
