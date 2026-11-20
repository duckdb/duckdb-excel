//! Relationship parts such as `xl/_rels/workbook.xml.rels`, which map relationship ids to other parts.

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::xml::{Attributes, Flow, XmlHandler},
};

/// One `Relationship` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relation {
    pub id: String,
    pub kind: String,
    pub target: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Start,
    Relationships,
    Relationship,
}

/// Collects the relationships of a part in document order.
#[derive(Debug)]
pub struct RelationshipsParser {
    state: State,
    pub relations: Vec<Relation>,
}

impl RelationshipsParser {
    pub fn new() -> RelationshipsParser {
        RelationshipsParser {
            state: State::Start,
            relations: Vec::new(),
        }
    }
}

impl Default for RelationshipsParser {
    fn default() -> Self {
        RelationshipsParser::new()
    }
}

impl XmlHandler for RelationshipsParser {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        match self.state {
            State::Start if name == "Relationships" => self.state = State::Relationships,
            State::Relationships if name == "Relationship" => {
                self.state = State::Relationship;
                let [Some(id), Some(kind), Some(target)] =
                    attributes.get(["Id", "Type", "Target"])?
                else {
                    return Err(invalid_input("Invalid relationship entry in _rels/.rels"));
                };
                self.relations.push(Relation {
                    id: id.into_owned(),
                    kind: kind.into_owned(),
                    target: target.into_owned(),
                });
            }
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn end_element(&mut self, name: &str) -> Result<Flow> {
        match self.state {
            State::Relationship if name == "Relationship" => self.state = State::Relationships,
            State::Relationships if name == "Relationships" => return Ok(Flow::Stop),
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

    fn parse(xml: &str) -> Result<Vec<Relation>> {
        let mut parser = RelationshipsParser::new();
        XmlReader::new(xml.as_bytes(), None).parse_all(&mut parser)?;
        Ok(parser.relations)
    }

    #[test]
    fn relationships_come_in_document_order() {
        let relations = parse(
            r#"<Relationships xmlns="p">
                 <Relationship Id="rId2" Type="t/worksheet" Target="worksheets/sheet2.xml"/>
                 <Relationship Target="styles.xml" Type="t/styles" Id="rId1" TargetMode="Internal"/>
               </Relationships>"#,
        )
        .unwrap();
        let fields: Vec<_> = relations
            .iter()
            .map(|relation| {
                (
                    relation.id.as_str(),
                    relation.kind.as_str(),
                    relation.target.as_str(),
                )
            })
            .collect();
        assert_eq!(
            fields,
            [
                ("rId2", "t/worksheet", "worksheets/sheet2.xml"),
                ("rId1", "t/styles", "styles.xml")
            ]
        );
    }

    #[test]
    fn a_relationship_needs_id_type_and_target() {
        for xml in [
            r#"<Relationships><Relationship Id="a" Type="b"/></Relationships>"#,
            r#"<Relationships><Relationship Id="a" Target="c"/></Relationships>"#,
            r#"<Relationships><Relationship Type="b" Target="c"/></Relationships>"#,
        ] {
            let error = parse(xml).unwrap_err();
            assert_eq!(error.code, DuckDBError::InputInvalid);
            assert_eq!(error.message, "Invalid relationship entry in _rels/.rels");
        }
    }

    #[test]
    fn other_documents_have_no_relationships() {
        assert!(
            parse(r#"<Other><Relationship/></Other>"#)
                .unwrap()
                .is_empty()
        );
        assert!(parse("").unwrap().is_empty());
    }
}
