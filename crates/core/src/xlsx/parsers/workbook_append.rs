//! `xl/workbook.xml` as appending a sheet needs it: the sheets with their ids.

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::xml::{Attributes, Flow, XmlHandler},
};

/// A `sheet` element; `sheet_id` is empty when the element has no `sheetId`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AppendSheetEntry {
    pub name: String,
    pub relation_id: String,
    pub sheet_id: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Start,
    Workbook,
    Sheets,
    Sheet,
}

/// Collects the sheets of a workbook in the order the workbook shows them; a sheet needs a name and a
/// relationship id that are not empty.
#[derive(Debug)]
pub struct WorkbookAppendParser {
    state: State,
    pub sheets: Vec<AppendSheetEntry>,
}

impl WorkbookAppendParser {
    pub fn new() -> WorkbookAppendParser {
        WorkbookAppendParser {
            state: State::Start,
            sheets: Vec::new(),
        }
    }
}

impl Default for WorkbookAppendParser {
    fn default() -> Self {
        WorkbookAppendParser::new()
    }
}

impl XmlHandler for WorkbookAppendParser {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        match self.state {
            State::Start if name == "workbook" => self.state = State::Workbook,
            State::Workbook if name == "sheets" => self.state = State::Sheets,
            State::Sheets if name == "sheet" => {
                self.state = State::Sheet;
                let [name, relation_id, sheet_id] = attributes
                    .get(["name", "r:id", "sheetId"])?
                    .map(|value| value.map(|value| value.into_owned()).unwrap_or_default());
                if name.is_empty() || relation_id.is_empty() {
                    return Err(invalid_input("Invalid sheet entry in workbook.xml"));
                }
                self.sheets.push(AppendSheetEntry {
                    name,
                    relation_id,
                    sheet_id,
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

    fn parse(xml: &str) -> Result<Vec<(String, String, String)>> {
        let mut parser = WorkbookAppendParser::new();
        XmlReader::new(xml.as_bytes(), None).parse_all(&mut parser)?;
        Ok(parser
            .sheets
            .into_iter()
            .map(|sheet| (sheet.name, sheet.relation_id, sheet.sheet_id))
            .collect())
    }

    fn sheet(name: &str, relation_id: &str, sheet_id: &str) -> (String, String, String) {
        (name.to_owned(), relation_id.to_owned(), sheet_id.to_owned())
    }

    #[test]
    fn sheets_come_with_their_ids() {
        let sheets = parse(
            r#"<workbook xmlns:r="rel"><bookViews/><sheets>
                 <sheet name="Second" sheetId="7" r:id="rId4"/>
                 <x:sheet r:id="rId1" name="A &amp; B" state="hidden"/>
               </sheets><definedNames/></workbook><ignored"#,
        );
        assert_eq!(
            sheets.unwrap(),
            [sheet("Second", "rId4", "7"), sheet("A & B", "rId1", "")]
        );
    }

    #[test]
    fn a_sheet_needs_a_name_and_a_relationship() {
        for xml in [
            r#"<workbook><sheets><sheet name="" r:id="rId1"/></sheets></workbook>"#,
            r#"<workbook><sheets><sheet name="S" r:id=""/></sheets></workbook>"#,
            r#"<workbook><sheets><sheet name="S" sheetId="1"/></sheets></workbook>"#,
            r#"<workbook><sheets><sheet r:id="rId1"/></sheets></workbook>"#,
        ] {
            let error = parse(xml).unwrap_err();
            assert_eq!(error.code, DuckDBError::InputInvalid, "{xml}");
            assert_eq!(error.message, "Invalid sheet entry in workbook.xml");
        }
    }

    #[test]
    fn sheets_outside_the_sheets_of_a_workbook_are_ignored() {
        let sheets = parse(r#"<workbook><sheet name="S" r:id="rId1"/></workbook>"#);
        assert!(sheets.unwrap().is_empty());
        let sheets = parse(r#"<other><sheets><sheet name="S" r:id="rId1"/></sheets></other>"#);
        assert!(sheets.unwrap().is_empty());
    }
}
