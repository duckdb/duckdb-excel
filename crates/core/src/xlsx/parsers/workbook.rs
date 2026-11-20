//! `xl/workbook.xml`: the sheets of a workbook and the day it counts serial numbers from.

use std::collections::HashMap;

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::{
        serial::DateSystem,
        xml::{Attributes, Flow, XmlHandler},
    },
};

/// The namespaces a relationship id is in, in a transitional and in a strict workbook.
const RELATIONSHIP_NAMESPACES: [&str; 2] = [
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
];

/// A `sheet` element: the name of a sheet, the id of the relationship that leads to its part, and its
/// `sheetId`, which is empty when the element has none.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetEntry {
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

/// A namespace prefix an element declares, and what the prefix was bound to outside the element.
#[derive(Debug)]
struct Declaration {
    depth: usize,
    prefix: String,
    outer: Option<bool>,
}

/// Collects the sheets of a workbook in the order the workbook shows them. A sheet needs a name and a
/// relationship id, which is the `id` attribute in the relationships namespace or else the one written `r:id`.
#[derive(Debug)]
pub struct WorkbookParser {
    state: State,
    is_for_append: bool,
    depth: usize,
    declarations: Vec<Declaration>,
    relationship_prefixes: HashMap<String, bool>,
    pub sheets: Vec<SheetEntry>,
    pub date_system: DateSystem,
}

impl WorkbookParser {
    pub fn new() -> WorkbookParser {
        WorkbookParser {
            state: State::Start,
            is_for_append: false,
            depth: 0,
            declarations: Vec::new(),
            relationship_prefixes: HashMap::new(),
            sheets: Vec::new(),
            date_system: DateSystem::Since1900,
        }
    }

    /// A parser for a workbook that gets a sheet added: the relationship id of a sheet must be written
    /// `r:id`, as the one of the new sheet is, and neither it nor the name may be empty.
    pub fn for_append() -> WorkbookParser {
        WorkbookParser {
            is_for_append: true,
            ..WorkbookParser::new()
        }
    }

    fn declare(&mut self, prefix: &str, namespace: &str) {
        let is_relationships = RELATIONSHIP_NAMESPACES.contains(&namespace);
        let outer = self
            .relationship_prefixes
            .insert(prefix.to_owned(), is_relationships);
        self.declarations.push(Declaration {
            depth: self.depth,
            prefix: prefix.to_owned(),
            outer,
        });
    }

    fn leave_element(&mut self) {
        while let Some(declaration) = self
            .declarations
            .pop_if(|declaration| declaration.depth >= self.depth)
        {
            match declaration.outer {
                Some(outer) => self.relationship_prefixes.insert(declaration.prefix, outer),
                None => self.relationship_prefixes.remove(&declaration.prefix),
            };
        }
        self.depth = self.depth.saturating_sub(1);
    }

    fn is_relationship_id(&self, attribute: &str) -> bool {
        attribute
            .strip_suffix(":id")
            .and_then(|prefix| self.relationship_prefixes.get(prefix))
            .is_some_and(|is_relationships| *is_relationships)
    }
}

impl Default for WorkbookParser {
    fn default() -> Self {
        WorkbookParser::new()
    }
}

impl XmlHandler for WorkbookParser {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        self.depth += 1;
        let attributes = attributes.all()?;
        for (attribute, value) in &attributes {
            if let Some(prefix) = attribute.strip_prefix("xmlns:") {
                self.declare(prefix, value);
            }
        }
        match self.state {
            State::Start if name == "workbook" => self.state = State::Workbook,
            State::Workbook if name == "sheets" => self.state = State::Sheets,
            State::Workbook if name == "workbookPr" => {
                let uses_1904 = attributes.iter().any(|(attribute, value)| {
                    *attribute == "date1904" && matches!(value.as_ref(), "1" | "true" | "on")
                });
                if uses_1904 {
                    self.date_system = DateSystem::Since1904;
                }
            }
            State::Sheets if name == "sheet" => {
                self.state = State::Sheet;
                let value_of = |wanted: &str| {
                    let found = attributes
                        .iter()
                        .find(|(attribute, _)| *attribute == wanted);
                    found.map(|(_, value)| value.as_ref())
                };
                let namespaced = attributes
                    .iter()
                    .find(|(attribute, _)| self.is_relationship_id(attribute))
                    .map(|(_, value)| value.as_ref())
                    .filter(|_| !self.is_for_append);
                let relation_id = namespaced.or_else(|| value_of("r:id"));
                let invalid = || invalid_input("Invalid sheet entry in workbook.xml");
                let (Some(name), Some(relation_id)) = (value_of("name"), relation_id) else {
                    return Err(invalid());
                };
                if self.is_for_append && (name.is_empty() || relation_id.is_empty()) {
                    return Err(invalid());
                }
                self.sheets.push(SheetEntry {
                    name: name.to_owned(),
                    relation_id: relation_id.to_owned(),
                    sheet_id: value_of("sheetId").unwrap_or_default().to_owned(),
                });
            }
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn end_element(&mut self, name: &str) -> Result<Flow> {
        self.leave_element();
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

    fn parse_with(mut parser: WorkbookParser, xml: &str) -> Result<Vec<(String, String)>> {
        XmlReader::new(xml.as_bytes(), None).parse_all(&mut parser)?;
        Ok(parser
            .sheets
            .into_iter()
            .map(|sheet| (sheet.name, sheet.relation_id))
            .collect())
    }

    fn parse(xml: &str) -> Result<Vec<(String, String)>> {
        parse_with(WorkbookParser::new(), xml)
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
    fn sheets_come_with_their_ids() {
        let mut parser = WorkbookParser::for_append();
        let xml = r#"<workbook><sheets><sheet name="A" sheetId="7" r:id="rId4"/><sheet r:id="rId1" name="B"/>
            <sheet name="C" sheetId="" r:id="rId2"/></sheets></workbook>"#;
        XmlReader::new(xml.as_bytes(), None)
            .parse_all(&mut parser)
            .unwrap();
        let ids: Vec<&str> = parser
            .sheets
            .iter()
            .map(|sheet| sheet.sheet_id.as_str())
            .collect();
        assert_eq!(ids, ["7", "", ""]);
    }

    #[test]
    fn empty_names_and_relationship_ids_are_refused_on_request() {
        for xml in [
            r#"<workbook><sheets><sheet name="" r:id="rId1"/></sheets></workbook>"#,
            r#"<workbook><sheets><sheet name="S" r:id=""/></sheets></workbook>"#,
        ] {
            assert_eq!(parse(xml).unwrap().len(), 1, "{xml}");
            let error = parse_with(WorkbookParser::for_append(), xml).unwrap_err();
            assert_eq!(error.code, DuckDBError::InputInvalid, "{xml}");
            assert_eq!(error.message, "Invalid sheet entry in workbook.xml");
        }
        for xml in [
            r#"<workbook><sheets><sheet name="S" sheetId="1"/></sheets></workbook>"#,
            r#"<workbook><sheets><sheet r:id="rId1"/></sheets></workbook>"#,
        ] {
            assert!(parse(xml).is_err(), "{xml}");
            assert!(parse_with(WorkbookParser::for_append(), xml).is_err());
        }
    }

    #[test]
    fn prefixed_elements_are_found() {
        let sheets = parse(
            r#"<x:workbook><x:sheets><x:sheet name="S" r:id="rId1"/></x:sheets></x:workbook>"#,
        );
        assert_eq!(sheets.unwrap(), pairs(&[("S", "rId1")]));
    }

    const TRANSITIONAL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    #[test]
    fn the_relationship_id_is_the_id_in_the_relationships_namespace() {
        for namespace in [TRANSITIONAL, STRICT] {
            let xml = format!(
                r#"<workbook xmlns:rel="{namespace}" xmlns:o="other"><sheets>
                     <sheet name="A" o:id="no" rel:id="rId1" id="no"/>
                     <sheet name="B" r:id="rId2"/>
                     <sheet name="C" r:id="no" rel:id="rId3"/>
                     <sheet xmlns:q="{namespace}" name="D" q:id="rId4"/>
                     <x:sheet xmlns:x="{namespace}" name="E" x:id="rId5" rel:id="no"/>
                   </sheets></workbook>"#
            );
            assert_eq!(
                parse(&xml).unwrap(),
                pairs(&[
                    ("A", "rId1"),
                    ("B", "rId2"),
                    ("C", "rId3"),
                    ("D", "rId4"),
                    ("E", "rId5")
                ])
            );
        }
    }

    #[test]
    fn a_prefix_counts_where_its_declaration_reaches() {
        let xml = format!(
            r#"<workbook xmlns:a="{TRANSITIONAL}"><other xmlns:b="{TRANSITIONAL}" xmlns:a="other"><inner/></other>
                 <sheets xmlns:c="other">
                   <sheet name="A" a:id="rId1" b:id="no"/>
                   <sheet xmlns:a="other" xmlns:c="{STRICT}" name="B" a:id="no" c:id="rId2"/>
                   <sheet name="C" c:id="no" a:id="rId3"/>
                 </sheets></workbook>"#
        );
        assert_eq!(
            parse(&xml).unwrap(),
            pairs(&[("A", "rId1"), ("B", "rId2"), ("C", "rId3")])
        );
    }

    #[test]
    fn an_id_under_a_prefix_bound_elsewhere_is_no_relationship_id() {
        for xml in [
            r#"<workbook><sheets><sheet name="S" rel:id="rId1"/></sheets></workbook>"#,
            r#"<workbook xmlns:rel="other"><sheets><sheet name="S" rel:id="rId1"/></sheets></workbook>"#,
            r#"<workbook><other xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
               <sheets><sheet name="S" rel:id="rId1"/></sheets></workbook>"#,
            r#"<workbook xmlns="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>
               <sheet name="S" id="rId1"/></sheets></workbook>"#,
            r#"<workbook><sheets><sheet r:id="rId1"/></sheets></workbook>"#,
        ] {
            let error = parse(xml).unwrap_err();
            assert_eq!(error.code, DuckDBError::InputInvalid, "{xml}");
            assert_eq!(error.message, "Invalid sheet entry in workbook.xml");
        }
    }

    #[test]
    fn a_workbook_that_gets_a_sheet_added_must_write_its_ids_as_the_new_sheet_does() {
        let xml = format!(
            r#"<workbook xmlns:rel="{TRANSITIONAL}"><sheets><sheet name="S" rel:id="rId1"/></sheets></workbook>"#
        );
        assert_eq!(parse(&xml).unwrap(), pairs(&[("S", "rId1")]));
        let error = parse_with(WorkbookParser::for_append(), &xml).unwrap_err();
        assert_eq!(error.message, "Invalid sheet entry in workbook.xml");
        let xml = xml.replace("rel:id=", "r:id=\"rId9\" rel:id=");
        assert_eq!(parse(&xml).unwrap(), pairs(&[("S", "rId1")]));
        assert_eq!(
            parse_with(WorkbookParser::for_append(), &xml).unwrap(),
            pairs(&[("S", "rId9")])
        );
    }

    fn date_system(workbook: &str) -> DateSystem {
        let mut parser = WorkbookParser::new();
        XmlReader::new(workbook.as_bytes(), None)
            .parse_all(&mut parser)
            .unwrap();
        parser.date_system
    }

    #[test]
    fn a_workbook_says_when_it_counts_from_1904() {
        for value in ["1", "true", "on"] {
            let xml = format!(
                r#"<workbook><fileVersion/><workbookPr codeName="x" date1904="{value}"/><sheets/></workbook>"#
            );
            assert_eq!(date_system(&xml), DateSystem::Since1904, "{value}");
        }
        for xml in [
            r#"<workbook><sheets/></workbook>"#,
            r#"<workbook><workbookPr/><sheets/></workbook>"#,
            r#"<workbook><workbookPr date1904="0"/><sheets/></workbook>"#,
            r#"<workbook><workbookPr date1904="false"/><sheets/></workbook>"#,
            r#"<workbook><workbookPr date1904="TRUE"/><sheets/></workbook>"#,
            r#"<workbook><workbookPr x:date1904="1"/><sheets/></workbook>"#,
            r#"<workbook><sheets><workbookPr date1904="1"/></sheets></workbook>"#,
            r#"<other><workbookPr date1904="1"/></other>"#,
            r#"<workbook><sheets/></workbook><workbookPr date1904="1"/>"#,
        ] {
            assert_eq!(date_system(xml), DateSystem::Since1900, "{xml}");
        }
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
