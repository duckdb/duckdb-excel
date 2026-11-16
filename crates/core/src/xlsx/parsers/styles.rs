//! `xl/styles.xml`: which cell styles show their number as a date or a time.

use std::collections::HashMap;

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::{
        cell::parse_integer,
        xml::{Attributes, Flow, XmlHandler},
    },
};

/// What the number format of a cell style makes of a number.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberKind {
    Double,
    Date,
    Time,
    Timestamp,
}

/// The number formats below this id are built into Excel; the others are defined by the workbook.
const FIRST_CUSTOM_FORMAT: i64 = 164;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Start,
    StyleSheet,
    NumFmts,
    NumFmt,
    CellXfs,
    Xf,
}

/// Collects the number kind of every cell style, in the order cells refer to them with their `s` attribute.
/// A style whose custom number format the workbook does not define is left out.
#[derive(Debug)]
pub struct StylesParser {
    state: State,
    number_formats: HashMap<i64, NumberKind>,
    pub cell_styles: Vec<NumberKind>,
}

impl StylesParser {
    pub fn new() -> StylesParser {
        StylesParser {
            state: State::Start,
            number_formats: HashMap::new(),
            cell_styles: Vec::new(),
        }
    }
}

impl Default for StylesParser {
    fn default() -> Self {
        StylesParser::new()
    }
}

fn custom_format_kind(format_code: &str) -> NumberKind {
    let has_date_part = ["DD", "dd", "YY", "yy"]
        .iter()
        .any(|part| format_code.contains(part));
    let has_time_part = format_code.contains(['h', 'H']);
    match (has_date_part, has_time_part) {
        (true, true) => NumberKind::Timestamp,
        (true, false) => NumberKind::Date,
        (false, true) => NumberKind::Time,
        (false, false) => NumberKind::Double,
    }
}

fn builtin_format_kind(id: i64) -> NumberKind {
    match id {
        14..=17 => NumberKind::Date,
        18..=21 => NumberKind::Time,
        22 => NumberKind::Timestamp,
        _ => NumberKind::Double,
    }
}

impl XmlHandler for StylesParser {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        match self.state {
            State::Start if name == "styleSheet" => self.state = State::StyleSheet,
            State::StyleSheet if name == "numFmts" => self.state = State::NumFmts,
            State::StyleSheet if name == "cellXfs" => self.state = State::CellXfs,
            State::NumFmts => {
                self.state = State::NumFmt;
                let [id, format_code] = attributes.get(["numFmtId", "formatCode"])?;
                let Some(id) = id else {
                    return Err(invalid_input("Invalid numFmt entry in styles.xml"));
                };
                let id = parse_integer(&id);
                if let Some(format_code) = format_code.filter(|_| id >= FIRST_CUSTOM_FORMAT) {
                    self.number_formats
                        .entry(id)
                        .or_insert_with(|| custom_format_kind(&format_code));
                }
            }
            State::CellXfs => {
                self.state = State::Xf;
                let [Some(id)] = attributes.get(["numFmtId"])? else {
                    return Err(invalid_input("Invalid xf entry in styles.xml"));
                };
                let id = parse_integer(&id);
                if id < FIRST_CUSTOM_FORMAT {
                    self.cell_styles.push(builtin_format_kind(id));
                } else if let Some(kind) = self.number_formats.get(&id) {
                    self.cell_styles.push(*kind);
                }
            }
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn end_element(&mut self, name: &str) -> Result<Flow> {
        match self.state {
            State::NumFmt if name == "numFmt" => self.state = State::NumFmts,
            State::Xf if name == "xf" => self.state = State::CellXfs,
            State::NumFmts if name == "numFmts" => self.state = State::StyleSheet,
            State::CellXfs if name == "cellXfs" => self.state = State::StyleSheet,
            State::StyleSheet if name == "styleSheet" => return Ok(Flow::Stop),
            _ => {}
        }
        Ok(Flow::Continue)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::xml::XmlReader;
    use duckdb_neo::error::DuckDBError;

    fn parse(xml: &str) -> Result<Vec<NumberKind>> {
        let mut parser = StylesParser::new();
        XmlReader::new(xml.as_bytes(), None).parse_all(&mut parser)?;
        Ok(parser.cell_styles)
    }

    fn cell_xfs(ids: &[i64]) -> String {
        let xfs: String = ids
            .iter()
            .map(|id| format!(r#"<xf numFmtId="{id}" fontId="0"/>"#))
            .collect();
        format!("<cellXfs>{xfs}</cellXfs>")
    }

    #[test]
    fn builtin_formats() {
        use NumberKind::{Date, Double, Time, Timestamp};
        let ids = [
            0, 1, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 49, 163, -5,
        ];
        let styles = parse(&format!("<styleSheet>{}</styleSheet>", cell_xfs(&ids))).unwrap();
        assert_eq!(
            styles,
            [
                Double, Double, Double, Date, Date, Date, Date, Time, Time, Time, Time, Timestamp,
                Double, Double, Double, Double
            ]
        );
    }

    #[test]
    fn custom_formats_are_dates_or_times_by_their_letters() {
        use NumberKind::{Date, Double, Time, Timestamp};
        let formats = [
            (164, "yyyy-mm-dd", Date),
            (165, "hh:mm:ss", Time),
            (166, "dd/mm/yy h:mm", Timestamp),
            (167, "0.00", Double),
            (168, "H", Time),
            (169, "DD", Date),
            (170, "mm:ss", Double),
            (171, "\"th\" 0", Time),
            (172, "y", Double),
        ];
        let num_fmts: String = formats
            .iter()
            .map(|(id, code, _)| {
                format!(
                    r#"<numFmt numFmtId="{id}" formatCode="{}"/>"#,
                    code.replace('"', "&quot;")
                )
            })
            .collect();
        let ids: Vec<i64> = formats.iter().map(|(id, _, _)| *id).collect();
        let xml = format!(
            "<styleSheet><numFmts>{num_fmts}</numFmts>{}</styleSheet>",
            cell_xfs(&ids)
        );
        let kinds: Vec<NumberKind> = formats.iter().map(|(_, _, kind)| *kind).collect();
        assert_eq!(parse(&xml).unwrap(), kinds);
    }

    #[test]
    fn styles_with_an_undefined_custom_format_are_left_out() {
        let xml = format!(
            r#"<styleSheet><numFmts><numFmt numFmtId="170" formatCode="yyyy"/><numFmt numFmtId="100" formatCode="yyyy"/>
               <numFmt numFmtId="171"/></numFmts>{}</styleSheet>"#,
            cell_xfs(&[14, 999, 170, 171, 100, 0])
        );
        use NumberKind::{Date, Double};
        assert_eq!(parse(&xml).unwrap(), [Date, Date, Double, Double]);
    }

    #[test]
    fn the_first_definition_of_a_format_counts() {
        let xml = format!(
            r#"<styleSheet><numFmts><numFmt numFmtId="164" formatCode="hh"/><numFmt numFmtId="164" formatCode="yyyy"/></numFmts>{}</styleSheet>"#,
            cell_xfs(&[164])
        );
        assert_eq!(parse(&xml).unwrap(), [NumberKind::Time]);
    }

    #[test]
    fn only_the_cell_formats_of_the_style_sheet_count() {
        let xml = r#"<styleSheet><cellStyleXfs><xf numFmtId="14"/></cellStyleXfs>
            <cellXfs><xf numFmtId="0"><alignment horizontal="left"/><protection locked="1"/></xf><xf numFmtId="14"/></cellXfs>
            <dxfs><dxf><numFmt numFmtId="170" formatCode="yyyy"/></dxf></dxfs></styleSheet>"#;
        assert_eq!(parse(xml).unwrap(), [NumberKind::Double, NumberKind::Date]);
        assert!(
            parse(r#"<other><cellXfs><xf numFmtId="14"/></cellXfs></other>"#)
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn entries_without_a_format_id_are_errors() {
        let error = parse("<styleSheet><numFmts><numFmt formatCode=\"0\"/></numFmts></styleSheet>")
            .unwrap_err();
        assert_eq!(error.code, DuckDBError::InputInvalid);
        assert_eq!(error.message, "Invalid numFmt entry in styles.xml");
        let error = parse("<styleSheet><cellXfs><xf/></cellXfs></styleSheet>").unwrap_err();
        assert_eq!(error.message, "Invalid xf entry in styles.xml");
    }
}
