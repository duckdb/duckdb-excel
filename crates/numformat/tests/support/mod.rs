//! The text encoding of the corpus files in `tests/data`, see `tests/data/README.md`.

/// Decodes an escaped field into the bytes it stands for.
pub fn unescape(text: &str) -> Vec<u8> {
    let bytes = text.as_bytes();
    let mut result = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        index += 1;
        if byte != b'\\' || index >= bytes.len() {
            result.push(byte);
            continue;
        }
        let escaped = bytes[index];
        index += 1;
        match escaped {
            b't' => result.push(b'\t'),
            b'n' => result.push(b'\n'),
            b'r' => result.push(b'\r'),
            b'x' => {
                let hex = text.get(index..index + 2);
                if let Some(value) = hex.and_then(|hex| u8::from_str_radix(hex, 16).ok()) {
                    result.push(value);
                    index += 2;
                }
            }
            other => result.push(other),
        }
    }
    result
}

/// Encodes bytes as printable ASCII without tabs.
pub fn escape(bytes: &[u8]) -> String {
    let mut result = String::with_capacity(bytes.len());
    for &byte in bytes {
        match byte {
            b'\\' => result.push_str("\\\\"),
            b'\t' => result.push_str("\\t"),
            b'\n' => result.push_str("\\n"),
            b'\r' => result.push_str("\\r"),
            0x20..=0x7e => result.push(char::from(byte)),
            _ => result.push_str(&format!("\\x{byte:02X}")),
        }
    }
    result
}

/// Parses a value field: a decimal number or `0x` and the 16 hex digits of the IEEE 754 bits.
pub fn parse_value(text: &str) -> Option<f64> {
    match text.strip_prefix("0x") {
        Some(bits) if bits.len() == 16 => u64::from_str_radix(bits, 16).ok().map(f64::from_bits),
        _ => text.parse().ok(),
    }
}
