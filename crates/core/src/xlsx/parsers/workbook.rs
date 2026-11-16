//! `xl/workbook.xml`: the sheets of a workbook.

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::xml::{Attributes, Flow, XmlHandler},
};

/// A `sheet` element: the name of a sheet and the id of the relationship that leads to its part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetEntry {
    pub name: String,
    pub relation_id: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Start,
    Workbook,
    Sheets,
    Sheet,
}

/// Collects the sheets of a workbook in the order the workbook shows them.
#[derive(Debug)]
pub struct WorkbookParser {
    state: State,
    pub sheets: Vec<SheetEntry>,
}

impl WorkbookParser {
    pub fn new() -> WorkbookParser {
        WorkbookParser {
            state: State::Start,
            sheets: Vec::new(),
        }
    }
}

impl Default for WorkbookParser {
    fn default() -> Self {
        WorkbookParser::new()
    }
}

impl XmlHandler for WorkbookParser {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        match self.state {
            State::Start if name == "workbook" => self.state = State::Workbook,
            State::Workbook if name == "sheets" => self.state = State::Sheets,
            State::Sheets if name == "sheet" => {
                self.state = State::Sheet;
                let [Some(name), Some(relation_id)] = attributes.get(["name", "r:id"])? else {
                    return Err(invalid_input("Invalid sheet entry in workbook.xml"));
                };
                self.sheets.push(SheetEntry {
                    name: name.into_owned(),
                    relation_id: relation_id.into_owned(),
                });
            }
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn end_element(&mut self, name: &str) -> Result<Flow> {
        match self.state {
            State::Sheet if name == "sheet" => self.state = State::Sheets,
            State::Sheets if name == "sheets" => self.state = State::Workbook,
            State::Workbook if name == "workbook" => return Ok(Flow::Stop),
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

    fn parse(xml: &str) -> Result<Vec<(String, String)>> {
        let mut parser = WorkbookParser::new();
        XmlReader::new(xml.as_bytes(), None).parse_all(&mut parser)?;
        Ok(parser
            .sheets
            .into_iter()
            .map(|sheet| (sheet.name, sheet.relation_id))
            .collect())
    }

    fn pairs(sheets: &[(&str, &str)]) -> Vec<(String, String)> {
        sheets
            .iter()
            .map(|(name, id)| ((*name).to_owned(), (*id).to_owned()))
            .collect()
    }

    #[test]
    fn sheets_come_in_workbook_order() {
        let sheets = parse(
            r#"<workbook xmlns:r="rel"><bookViews><workbookView/></bookViews><sheets>
                 <sheet name="Second" sheetId="2" r:id="rId7"/>
                 <sheet r:id="rId3" name="A &amp; B" sheetId="1" state="hidden"/>
               </sheets><definedNames/></workbook>"#,
        );
        assert_eq!(
            sheets.unwrap(),
            pairs(&[("Second", "rId7"), ("A & B", "rId3")])
        );
    }

    #[test]
    fn prefixed_elements_are_found() {
        let sheets = parse(
            r#"<x:workbook><x:sheets><x:sheet name="S" r:id="rId1"/></x:sheets></x:workbook>"#,
        );
        assert_eq!(sheets.unwrap(), pairs(&[("S", "rId1")]));
    }

    #[test]
    fn the_relationship_id_is_found_by_its_written_name() {
        let error =
            parse(r#"<workbook><sheets><sheet name="S" rel:id="rId1"/></sheets></workbook>"#)
                .unwrap_err();
        assert_eq!(error.code, DuckDBError::InputInvalid);
        assert_eq!(error.message, "Invalid sheet entry in workbook.xml");
        assert!(parse(r#"<workbook><sheets><sheet r:id="rId1"/></sheets></workbook>"#).is_err());
    }

    #[test]
    fn sheets_outside_a_sheets_element_are_ignored() {
        let sheets = parse(
            r#"<workbook><sheet name="S" r:id="rId1"/><other><sheets><sheet name="N" r:id="rId2"/></sheets></other></workbook>"#,
        );
        assert_eq!(sheets.unwrap(), pairs(&[("N", "rId2")]));
        let sheets =
            parse(r#"<notworkbook><sheets><sheet name="S" r:id="rId1"/></sheets></notworkbook>"#);
        assert!(sheets.unwrap().is_empty());
    }

    #[test]
    fn reading_stops_at_the_end_of_the_workbook() {
        let sheets =
            parse(r#"<workbook><sheets><sheet name="S" r:id="rId1"/></sheets></workbook><broken"#);
        assert_eq!(sheets.unwrap(), pairs(&[("S", "rId1")]));
    }
}
