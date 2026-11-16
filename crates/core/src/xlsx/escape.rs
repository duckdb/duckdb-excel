//! Escaping of text for XML, and its reverse.

/// Escapes the five characters XML reserves and drops NUL, which XML cannot hold.
pub fn escape_xml(text: &str) -> String {
    let mut escaped = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '\0' => {}
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

/// Replaces the five predefined entities; any other `&` stays as it is.
pub fn unescape_xml_entities(text: &str) -> String {
    let mut unescaped = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(start) = rest.find('&') {
        let (before, entity_start) = rest.split_at(start);
        unescaped.push_str(before);
        let replacement = entity_start.find(';').and_then(|end| {
            let (entity, after) = entity_start.split_at(end + 1);
            let character = match entity {
                "&lt;" => '<',
                "&gt;" => '>',
                "&amp;" => '&',
                "&quot;" => '"',
                "&apos;" => '\'',
                _ => return None,
            };
            Some((character, after))
        });
        match replacement {
            Some((character, after)) => {
                unescaped.push(character);
                rest = after;
            }
            None => {
                unescaped.push('&');
                rest = entity_start.get(1..).unwrap_or_default();
            }
        }
    }
    unescaped.push_str(rest);
    unescaped
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unescaping_undoes_escaping() {
        for text in ["", "plain", "a&b<c>d\"e'f", "&amp;", "&&;;", "&unknown; &"] {
            assert_eq!(unescape_xml_entities(&escape_xml(text)), text);
        }
    }

    #[test]
    fn unescaping_keeps_what_is_not_an_entity() {
        assert_eq!(unescape_xml_entities("a &amp; b"), "a & b");
        assert_eq!(unescape_xml_entities("&lt;&gt;&quot;&apos;"), "<>\"'");
        assert_eq!(unescape_xml_entities("a & b; c"), "a & b; c");
        assert_eq!(unescape_xml_entities("&nbsp;&amp;"), "&nbsp;&");
        assert_eq!(unescape_xml_entities("fish & chips"), "fish & chips");
        assert_eq!(unescape_xml_entities("&amp;amp;"), "&amp;");
    }
}
