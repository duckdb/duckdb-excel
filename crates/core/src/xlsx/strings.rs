//! The shared strings of a workbook, which cells refer to by index.

/// Strings kept back to back in one buffer, so that a large table costs no allocation per string.
///
/// The number of strings a workbook announces can be anything, so only so much room is reserved for it.
#[derive(Debug, Default)]
pub struct StringTable {
    text: String,
    ends: Vec<usize>,
}

impl StringTable {
    const MAX_RESERVED: usize = 1 << 20;

    pub fn new() -> StringTable {
        StringTable::default()
    }

    pub fn reserve(&mut self, count: usize) {
        self.ends.reserve(count.min(Self::MAX_RESERVED));
    }

    pub fn push(&mut self, string: &str) {
        self.text.push_str(string);
        self.ends.push(self.text.len());
    }

    pub fn get(&self, index: usize) -> Option<&str> {
        let end = *self.ends.get(index)?;
        let start = match index.checked_sub(1) {
            Some(previous) => *self.ends.get(previous)?,
            None => 0,
        };
        self.text.get(start..end)
    }

    pub fn len(&self) -> usize {
        self.ends.len()
    }

    pub fn is_empty(&self) -> bool {
        self.ends.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strings_come_back_by_index() {
        let mut table = StringTable::new();
        table.reserve(usize::MAX);
        assert!(table.is_empty());
        for string in ["first", "", "żółw", "first"] {
            table.push(string);
        }
        assert_eq!(table.len(), 4);
        assert_eq!(table.get(0), Some("first"));
        assert_eq!(table.get(1), Some(""));
        assert_eq!(table.get(2), Some("żółw"));
        assert_eq!(table.get(3), Some("first"));
        assert_eq!(table.get(4), None);
    }
}
