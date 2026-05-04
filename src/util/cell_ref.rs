/// Cell reference parser — converts between "A1" notation and (col, row) indices.
/// Pure arithmetic, no external dependencies.

/// Maximum column index (0-based): 16383 = XFD
const MAX_COL: i32 = 16383;
/// Maximum row index (0-based): 1048575
const MAX_ROW: i32 = 1_048_575;

/// Parses a cell reference string (e.g., "A1", "BC42") into zero-based (col, row).
pub fn try_parse(cell_ref: &str) -> Option<(i32, i32)> {
    if cell_ref.is_empty() {
        return None;
    }

    let upper = cell_ref.to_ascii_uppercase();
    let bytes = upper.as_bytes();

    // Find boundary between letters and digits
    let mut split = 0;
    while split < bytes.len() && bytes[split].is_ascii_uppercase() {
        split += 1;
    }

    // Must have at least one letter and at least one digit
    if split == 0 || split == bytes.len() {
        return None;
    }

    // Verify remaining characters are all digits
    for &b in &bytes[split..] {
        if !b.is_ascii_digit() {
            return None;
        }
    }

    // Parse column: A=0, B=1, ..., Z=25, AA=26, AB=27, ...
    let mut col_value: i32 = 0;
    for &b in &bytes[..split] {
        col_value = col_value * 26 + (b - b'A') as i32 + 1;
    }
    let col = col_value - 1;

    // Parse row: "1" -> 0, "2" -> 1, etc.
    let row_str = &upper[split..];
    let row_one_based: i32 = row_str.parse().ok()?;
    if row_one_based < 1 {
        return None;
    }
    let row = row_one_based - 1;

    // Bounds check
    if col > MAX_COL || row > MAX_ROW {
        return None;
    }

    Some((col, row))
}

/// Parses a range string (e.g., "A1:C5") or a single cell reference into
/// zero-based (start_col, start_row, end_col, end_row).
pub fn try_parse_range(range_ref: &str) -> Option<(i32, i32, i32, i32)> {
    if range_ref.is_empty() {
        return None;
    }

    if let Some(colon_idx) = range_ref.find(':') {
        let start_part = &range_ref[..colon_idx];
        let end_part = &range_ref[colon_idx + 1..];

        let (start_col, start_row) = try_parse(start_part)?;
        let (end_col, end_row) = try_parse(end_part)?;

        Some((start_col, start_row, end_col, end_row))
    } else {
        // Single cell — treat as a 1x1 range
        let (col, row) = try_parse(range_ref)?;
        Some((col, row, col, row))
    }
}

/// Converts zero-based (col, row) back to a cell reference string (e.g., "A1").
pub fn to_ref(col: i32, row: i32) -> String {
    let mut col_ref = String::new();
    let mut c = col + 1; // one-based
    while c > 0 {
        c -= 1;
        col_ref.insert(0, (b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    format!("{}{}", col_ref, row + 1)
}

/// Parses a column letter string (e.g., "A", "BC") into a zero-based column index.
pub fn try_parse_col_letter(input: &str) -> Option<i32> {
    if input.is_empty() {
        return None;
    }

    let upper = input.to_ascii_uppercase();
    let mut value: i32 = 0;
    for b in upper.bytes() {
        if !b.is_ascii_uppercase() {
            return None;
        }
        value = value * 26 + (b - b'A') as i32 + 1;
    }

    let col = value - 1;
    if col >= 0 && col <= MAX_COL {
        Some(col)
    } else {
        None
    }
}

/// Converts a zero-based column index to a column letter string (e.g., 0 -> "A").
pub fn col_to_letter(col_index: i32) -> String {
    let mut result = String::new();
    let mut c = col_index + 1;
    while c > 0 {
        c -= 1;
        result.insert(0, (b'A' + (c % 26) as u8) as char);
        c /= 26;
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_basic() {
        assert_eq!(try_parse("A1"), Some((0, 0)));
        assert_eq!(try_parse("B2"), Some((1, 1)));
        assert_eq!(try_parse("Z1"), Some((25, 0)));
        assert_eq!(try_parse("AA1"), Some((26, 0)));
        assert_eq!(try_parse("BC42"), Some((54, 41)));
    }

    #[test]
    fn test_parse_case_insensitive() {
        assert_eq!(try_parse("a1"), Some((0, 0)));
        assert_eq!(try_parse("bc42"), Some((54, 41)));
    }

    #[test]
    fn test_parse_invalid() {
        assert_eq!(try_parse(""), None);
        assert_eq!(try_parse("1"), None);
        assert_eq!(try_parse("A"), None);
        assert_eq!(try_parse("A0"), None);
        assert_eq!(try_parse("A-1"), None);
        assert_eq!(try_parse("1A"), None);
    }

    #[test]
    fn test_to_ref() {
        assert_eq!(to_ref(0, 0), "A1");
        assert_eq!(to_ref(1, 1), "B2");
        assert_eq!(to_ref(25, 0), "Z1");
        assert_eq!(to_ref(26, 0), "AA1");
    }

    #[test]
    fn test_roundtrip() {
        for col in [0, 1, 25, 26, 27, 701, 702] {
            for row in [0, 1, 99, 1000] {
                let r = to_ref(col, row);
                assert_eq!(try_parse(&r), Some((col, row)), "Failed roundtrip for ({}, {})", col, row);
            }
        }
    }

    #[test]
    fn test_parse_range() {
        assert_eq!(try_parse_range("A1:C5"), Some((0, 0, 2, 4)));
        assert_eq!(try_parse_range("A1"), Some((0, 0, 0, 0)));
    }

    #[test]
    fn test_col_letter() {
        assert_eq!(try_parse_col_letter("A"), Some(0));
        assert_eq!(try_parse_col_letter("Z"), Some(25));
        assert_eq!(try_parse_col_letter("AA"), Some(26));
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
    }
}
