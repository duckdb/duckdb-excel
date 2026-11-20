//! `[Content_Types].xml`: nothing in it is needed to find a sheet, but a workbook must have a well-formed one.

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::xml::{Attributes, Flow, XmlHandler},
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Start,
    Types,
    Override,
}

/// Checks that every part the content types list has a name and a type.
#[derive(Debug)]
pub struct ContentTypesParser {
    state: State,
}

impl ContentTypesParser {
    pub fn new() -> ContentTypesParser {
        ContentTypesParser {
            state: State::Start,
        }
    }
}

impl Default for ContentTypesParser {
    fn default() -> Self {
        ContentTypesParser::new()
    }
}

impl XmlHandler for ContentTypesParser {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        match self.state {
            State::Start if name == "Types" => self.state = State::Types,
            State::Types if name == "Override" => {
                self.state = State::Override;
                let [content_type, part_name] = attributes.get(["ContentType", "PartName"])?;
                if content_type.is_none() || part_name.is_none() {
                    return Err(invalid_input(
                        "Invalid content type entry in [Content_Types].xml",
                    ));
                }
            }
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn end_element(&mut self, name: &str) -> Result<Flow> {
        match self.state {
            State::Override if name == "Override" => self.state = State::Types,
            State::Types if name == "Types" => return Ok(Flow::Stop),
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

    fn parse(xml: &str) -> Result<()> {
        XmlReader::new(xml.as_bytes(), None).parse_all(&mut ContentTypesParser::new())
    }

    #[test]
    fn overrides_need_a_name_and_a_type() {
        assert!(parse(r#"<Types><Default Extension="xml"/><Override PartName="/a" ContentType="b"/></Types>"#).is_ok());
        for xml in [
            r#"<Types><Override PartName="/a"/></Types>"#,
            r#"<Types><Override ContentType="b"/></Types>"#,
        ] {
            let error = parse(xml).unwrap_err();
            assert_eq!(error.code, DuckDBError::InputInvalid);
            assert_eq!(
                error.message,
                "Invalid content type entry in [Content_Types].xml"
            );
        }
    }

    #[test]
    fn only_overrides_of_the_types_element_are_checked() {
        assert!(parse(r#"<Other><Override/></Other>"#).is_ok());
        assert!(parse(r#"<Types><Default><Override/></Default></Types>"#).is_err());
        assert!(parse(r#"<Types></Types><Override/>"#).is_ok());
    }
}
