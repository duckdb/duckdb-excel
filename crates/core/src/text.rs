//! `text(number, format)` and its alias `excel_text`: a number as Excel's `TEXT` function shows it.

use duckdb_neo::{
    DuckDBType, Result,
    connection::{Context, FFILink},
    data_chunk::VectorCollection,
    scalar::{ScalarCallbacks, ScalarFunctionBuilder},
    signature::{FunctionParameter, FunctionParameterKind, SignatureBuilder},
    vector::{StorageKind, Unknown, Vector},
};
use excel_numformat::format_number;

/// Formats a DOUBLE with an Excel number format; NULL when either argument is NULL.
#[derive(Debug)]
pub struct TextFunction;

impl ScalarCallbacks for TextFunction {
    type BindData = ();
    type InitData = ();

    fn exec(
        &self,
        _context: &Context,
        _bind_data: Option<&()>,
        _init_data: Option<&mut ()>,
        vectors: &mut VectorCollection,
        output: Vector<'_, Unknown>,
    ) -> Result<()> {
        let rows = vectors.row_count()?;
        let mut output = output.cast::<String>()?;
        output.set_size(rows)?;
        let [mut numbers, mut formats] = vectors.get_vector_array_mut([0, 1])?;
        for vector in [&mut numbers, &mut formats] {
            if vector.storage_kind() == StorageKind::Other {
                vector.flatten()?;
            }
        }
        let numbers = numbers.cast::<f64>()?;
        let formats = formats.cast::<String>()?;
        let (numbers, formats) = (numbers.reader()?, formats.reader()?);
        let mut writer = output.writer()?;
        for row in 0..rows {
            match (numbers.get(row)?, formats.get(row)?) {
                (Some(number), Some(format)) => writer.set(row, &format_number(format, *number))?,
                _ => writer.set_null(row)?,
            }
        }
        Ok(())
    }
}

/// The scalar function under `name`: `text` or `excel_text`.
pub fn text_function<C: FFILink>(
    link: &C,
    name: &str,
) -> Result<ScalarFunctionBuilder<TextFunction>> {
    let signature = SignatureBuilder::new()
        .param(FunctionParameter::new::<f64>("col0").kind(FunctionParameterKind::PositionalOnly))
        .param(FunctionParameter::new::<String>("col1").kind(FunctionParameterKind::PositionalOnly))
        .return_type(String::logical_type(link)?);
    Ok(ScalarFunctionBuilder::new(name, signature, TextFunction))
}
