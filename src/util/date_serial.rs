/// Converts a date string (YYYY-MM-DD) to a serial number.
/// Serial: days since 1899-12-30 (so 1900-01-01 = 1).
/// Returns None if the string is not a valid date.
pub fn parse_date_to_serial(s: &str) -> Option<f64> {
    let parts: Vec<&str> = s.split('-').collect();
    if parts.len() != 3 {
        return None;
    }
    let year: i32 = parts[0].parse().ok()?;
    let month: u32 = parts[1].parse().ok()?;
    let day: u32 = parts[2].parse().ok()?;

    if month < 1 || month > 12 || day < 1 || day > 31 || year < 1900 {
        return None;
    }

    Some(date_to_serial(year, month, day))
}

/// Converts (year, month, day) to serial number.
/// Uses 1900-01-01 = 1 (with the Lotus 1-2-3 bug: 1900-02-29 exists).
fn date_to_serial(year: i32, month: u32, day: u32) -> f64 {
    // Count days from a fixed reference using a Julian Day Number approach
    fn jdn(y: i32, m: u32, d: u32) -> i64 {
        let a = (14 - m as i64) / 12;
        let y2 = y as i64 + 4800 - a;
        let m2 = m as i64 + 12 * a - 3;
        d as i64 + (153 * m2 + 2) / 5 + 365 * y2 + y2 / 4 - y2 / 100 + y2 / 400 - 32045
    }

    // Epoch: 1899-12-31 = serial 1... actually 1900-01-01 = 1
    // JDN of 1899-12-31 (the day before serial 1)
    let epoch_jdn = jdn(1899, 12, 31);
    let target_jdn = jdn(year, month, day);
    let serial = (target_jdn - epoch_jdn) as f64;

    // Lotus bug: dates after 1900-02-28 are off by 1 (it thinks 1900-02-29 exists)
    if serial > 59.0 {
        serial + 1.0
    } else {
        serial
    }
}

/// Tries to parse as a date (YYYY-MM-DD) first, then falls back to f64.
pub fn parse_date_or_number(s: &str) -> Option<f64> {
    if let Some(serial) = parse_date_to_serial(s) {
        Some(serial)
    } else {
        s.parse::<f64>().ok()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_known_dates() {
        // 1900-01-01 = serial 1
        assert_eq!(parse_date_to_serial("1900-01-01"), Some(1.0));
        // 2024-01-01 should be 45292
        assert_eq!(parse_date_to_serial("2024-01-01"), Some(45292.0));
    }

    #[test]
    fn test_fallback_to_number() {
        assert_eq!(parse_date_or_number("45292"), Some(45292.0));
        assert_eq!(parse_date_or_number("2024-01-01"), Some(45292.0));
    }
}
