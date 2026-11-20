//! `xl/styles.xml` as appending a sheet needs it: the number formats and cell styles a workbook already has.

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::{
        cell::parse_integer,
        xml::{Attributes, Flow, XmlHandler},
    },
};

/// A `numFmt` element: a number format the workbook defines.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumberFormat {
    pub id: u64,
    pub format_code: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Start,
    StyleSheet,
    NumFmts,
    NumFmt,
    CellXfs,
    Xf,
}

/// Collects the number formats, and the number format id of every cell style in the order cells refer to the
/// styles. A cell style without a number format has format 0, and a negative id wraps around.
#[derive(Debug)]
pub struct StylesAppendParser {
    state: State,
    pub number_formats: Vec<NumberFormat>,
    pub cell_styles: Vec<u64>,
}

impl StylesAppendParser {
    pub fn new() -> StylesAppendParser {
        StylesAppendParser {
            state: State::Start,
            number_formats: Vec::new(),
            cell_styles: Vec::new(),
        }
    }
}

impl Default for StylesAppendParser {
    fn default() -> Self {
        StylesAppendParser::new()
    }
}

impl XmlHandler for StylesAppendParser {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        match self.state {
            State::Start if name == "styleSheet" => self.state = State::StyleSheet,
            State::StyleSheet if name == "numFmts" => self.state = State::NumFmts,
            State::StyleSheet if name == "cellXfs" => self.state = State::CellXfs,
            State::NumFmts if name == "numFmt" => {
                self.state = State::NumFmt;
                let [id, format_code] = attributes.get(["numFmtId", "formatCode"])?;
                let Some(id) = id else {
                    return Err(invalid_input("Invalid numFmt entry in styles.xml"));
                };
                self.number_formats.push(NumberFormat {
                    id: parse_integer(&id) as u64,
                    format_code: format_code
                        .map(|code| code.into_owned())
                        .unwrap_or_default(),
                });
            }
            State::CellXfs if name == "xf" => {
                self.state = State::Xf;
                let [id] = attributes.get(["numFmtId"])?;
                self.cell_styles
                    .push(id.map_or(0, |id| parse_integer(&id) as u64));
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

    fn parse(xml: &str) -> Result<StylesAppendParser> {
        let mut parser = StylesAppendParser::new();
        XmlReader::new(xml.as_bytes(), None).parse_all(&mut parser)?;
        Ok(parser)
    }

    fn format(id: u64, code: &str) -> NumberFormat {
        NumberFormat {
            id,
            format_code: code.to_owned(),
        }
    }

    #[test]
    fn formats_and_cell_styles_come_in_document_order() {
        let parsed = parse(
            r#"<styleSheet><numFmts count="3">
                 <numFmt numFmtId="170" formatCode="&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;"/>
                 <numFmt formatCode="dd/mm/yyyy\ hh:mm:ss" numFmtId="165"/>
                 <numFmt numFmtId="166"/>
               </numFmts>
               <cellStyleXfs><xf numFmtId="99"/></cellStyleXfs>
               <cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="170"><alignment wrapText="1"/></xf><xf fontId="1"/><xf numFmtId="14"/></cellXfs>
               <dxfs><dxf><numFmt numFmtId="171" formatCode="0"/></dxf></dxfs></styleSheet>"#,
        )
        .unwrap();
        assert_eq!(
            parsed.number_formats,
            [
                format(170, "\"TRUE\";\"TRUE\";\"FALSE\""),
                format(165, "dd/mm/yyyy\\ hh:mm:ss"),
                format(166, ""),
            ]
        );
        assert_eq!(parsed.cell_styles, [0, 170, 0, 14]);
    }

    #[test]
    fn ids_are_read_leniently() {
        let parsed = parse(
            r#"<styleSheet><numFmts><numFmt numFmtId="-1" formatCode="a"/><numFmt numFmtId=" 12x" formatCode="b"/></numFmts>
               <cellXfs><xf numFmtId="abc"/><xf numFmtId="-2"/></cellXfs></styleSheet>"#,
        )
        .unwrap();
        assert_eq!(
            parsed.number_formats,
            [format(u64::MAX, "a"), format(12, "b")]
        );
        assert_eq!(parsed.cell_styles, [0, u64::MAX - 1]);
    }

    #[test]
    fn other_elements_are_not_formats_or_styles() {
        for xml in [
            r#"<styleSheet><numFmts><other numFmtId="164" formatCode="a"/></numFmts>
               <cellXfs><other numFmtId="164"/></cellXfs></styleSheet>"#,
            r#"<other><cellXfs><xf numFmtId="14"/></cellXfs></other>"#,
        ] {
            let parsed = parse(xml).unwrap();
            assert!(parsed.number_formats.is_empty() && parsed.cell_styles.is_empty());
        }
    }

    #[test]
    fn a_format_needs_an_id() {
        let error =
            parse(r#"<styleSheet><numFmts><numFmt formatCode="0"/></numFmts></styleSheet>"#)
                .unwrap_err();
        assert_eq!(error.code, DuckDBError::InputInvalid);
        assert_eq!(error.message, "Invalid numFmt entry in styles.xml");
    }
}
