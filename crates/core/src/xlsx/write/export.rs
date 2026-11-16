//! Writes the rows of a `COPY ... TO` into a workbook.

use std::{
    sync::{
        Mutex, MutexGuard,
        atomic::{AtomicU64, Ordering},
    },
    thread,
    time::{Duration, SystemTime, UNIX_EPOCH},
};

use duckdb_neo::{
    DuckDBType, Result, column_data_collection::ColumnDataCollection, connection::Context,
    copy_function::CopyToBindInfo, file::FileSystem, logical_type::LogicalTypeID, value::Value,
};
use excel_zip::ZipWriter;

use crate::{
    error::{binder, from_io, internal, not_implemented, reraise},
    file::FileWriter,
    xlsx::{
        archive::Archive,
        cell::MAX_CELL_ROWS,
        write::{
            append::AppendPlan,
            convert::{ConversionPlan, TextConverter},
            package::{FIRST_SHEET_FILE, WORKSHEETS_DIRECTORY, write_workbook_parts},
            sheet::{CellKind, CellStyles, SheetWriter},
        },
    },
};

const DEFAULT_SHEET_NAME: &str = "Sheet1";

/// What `MODE` makes of a workbook that already exists; without one, every mode writes a new workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WriteMode {
    /// The workbook is written anew.
    Create,
    /// The sheet is added to the sheets of the workbook.
    Append,
    /// The sheet takes the place of the sheet of the same name.
    Replace,
}

/// What a `COPY ... TO` writes, decided when the statement is bound.
///
/// The path is the one the statement has, which is where an existing workbook is looked for. The row limit
/// counts the header row, and a negative limit wraps around to one no sheet reaches.
///
/// Adding to a workbook takes reading it and moving the new one into its place, which only local files can be
/// relied on to do: a remote path is refused when the statement is bound, before DuckDB does anything about
/// the output file.
#[derive(Debug)]
pub struct WritePlan {
    file_path: String,
    sheet_name: String,
    header: bool,
    row_limit: u64,
    mode: WriteMode,
    column_names: Vec<String>,
    cell_kinds: Vec<CellKind>,
    conversions: ConversionPlan,
}

impl WritePlan {
    /// Reads the options `header`, `sheet`, `sheet_row_limit` and `mode`; other options are ignored.
    pub fn new(context: &Context, bind_info: &CopyToBindInfo) -> Result<WritePlan> {
        let mut plan = WritePlan {
            file_path: bind_info.file_path()?,
            sheet_name: DEFAULT_SHEET_NAME.to_owned(),
            header: false,
            row_limit: MAX_CELL_ROWS,
            mode: WriteMode::Create,
            column_names: Vec::new(),
            cell_kinds: Vec::new(),
            conversions: ConversionPlan::default(),
        };
        // DuckDB hands the options over sorted by name; of several that are wrong, the first of this list is
        // the one reported.
        let [header, sheet, row_limit, mode] =
            given_options(bind_info, ["header", "sheet", "sheet_row_limit", "mode"])?;
        if let Some(value) = header {
            plan.header = header_option(context, &value)?;
        }
        if let Some(value) = sheet {
            plan.sheet_name = sheet_option(&value)?;
        }
        if let Some(value) = row_limit {
            plan.row_limit = row_limit_option(context, &value)?;
        }
        if let Some(value) = mode {
            plan.mode = mode_option(&value)?;
        }
        plan.check_mode(context)?;
        for index in 0..bind_info.column_count()? {
            let (name, logical_type) = bind_info.get_column(index)?;
            plan.column_names.push(name.to_owned());
            let kind = plan.conversions.add_column(logical_type.id()?);
            plan.cell_kinds.push(kind);
        }
        Ok(plan)
    }

    fn check_mode(&self, context: &Context) -> Result<()> {
        let mode = match self.mode {
            WriteMode::Create => return Ok(()),
            WriteMode::Append => "append",
            WriteMode::Replace => "replace",
        };
        if !FileSystem::new(context)?.is_remote(&self.file_path)? {
            return Ok(());
        }
        Err(not_implemented(format!(
            "XLSX MODE '{mode}' is only supported for local files; '{}' is on a remote filesystem. \
             Write the workbook to a local path and upload it instead.",
            self.file_path
        )))
    }
}

fn given_options<const N: usize>(
    bind_info: &CopyToBindInfo,
    names: [&str; N],
) -> Result<[Option<Value>; N]> {
    let mut values = [const { None }; N];
    for index in 0..bind_info.option_count()? {
        let (name, value) = bind_info.get_option(index)?;
        let wanted = names
            .iter()
            .position(|wanted| name.eq_ignore_ascii_case(wanted));
        if let Some(slot) = wanted.and_then(|position| values.get_mut(position)) {
            *slot = Some(value);
        }
    }
    Ok(values)
}

fn cast_option<T: DuckDBType + duckdb_neo::types::FromValue>(
    context: &Context,
    value: &Value,
) -> Result<Option<T>> {
    if value.logical_type()?.id()? == LogicalTypeID::Tuple {
        return Ok(None);
    }
    match value.cast(context, &T::logical_type(context)?) {
        Ok(cast) => cast.get(),
        Err(_) => Ok(None),
    }
}

fn header_option(context: &Context, value: &Value) -> Result<bool> {
    cast_option(context, value)?
        .ok_or_else(|| binder("Header option must be a single boolean value"))
}

fn sheet_option(value: &Value) -> Result<String> {
    let not_a_string = || binder("Sheet name option must be a single string value");
    if value.logical_type()?.id()? != LogicalTypeID::Varchar {
        return Err(not_a_string());
    }
    value.get()?.ok_or_else(not_a_string)
}

fn row_limit_option(context: &Context, value: &Value) -> Result<u64> {
    let not_an_integer = || binder("Sheet row limit option must be a single integer value");
    // An option without a value arrives as a boolean, which is not a limit.
    if value.logical_type()?.id()? == LogicalTypeID::Boolean {
        return Err(not_an_integer());
    }
    let limit: i32 = cast_option(context, value)?.ok_or_else(not_an_integer)?;
    Ok(i64::from(limit) as u64)
}

fn mode_option(value: &Value) -> Result<WriteMode> {
    let not_a_string = || binder("MODE option must be a string ('create', 'append', or 'replace')");
    if value.logical_type()?.id()? != LogicalTypeID::Varchar {
        return Err(not_a_string());
    }
    let mode: String = value.get()?.ok_or_else(not_a_string)?;
    let mode = mode.to_ascii_lowercase();
    match mode.as_str() {
        "create" => Ok(WriteMode::Create),
        "append" => Ok(WriteMode::Append),
        "replace" => Ok(WriteMode::Replace),
        _ => Err(binder(format!(
            "Invalid MODE '{mode}' (expected 'create', 'append', or 'replace')"
        ))),
    }
}

/// How often a file operation that fails is tried, and how long to wait in between: a virus scanner may hold
/// a file that was just written for a moment.
const FILE_ATTEMPTS: u32 = 40;
const FILE_RETRY_PAUSE: Duration = Duration::from_millis(50);

/// What putting a file in the place of another takes of a file system.
trait FileOperations {
    fn exists(&self, path: &str) -> Result<bool>;

    fn remove(&self, path: &str) -> Result<bool>;

    fn move_item(&self, from: &str, to: &str) -> Result<()>;

    fn pause(&self) {
        thread::sleep(FILE_RETRY_PAUSE);
    }
}

impl FileOperations for FileSystem<'_> {
    fn exists(&self, path: &str) -> Result<bool> {
        FileSystem::exists(self, path)
    }

    fn remove(&self, path: &str) -> Result<bool> {
        FileSystem::remove(self, path)
    }

    fn move_item(&self, from: &str, to: &str) -> Result<()> {
        FileSystem::move_item(self, from, to)
    }
}

fn retry(files: &impl FileOperations, mut operation: impl FnMut() -> Result<()>) -> Result<()> {
    let mut attempt = 1;
    loop {
        match operation() {
            Err(_) if attempt < FILE_ATTEMPTS => {
                attempt += 1;
                files.pause();
            }
            result => return result,
        }
    }
}

/// The file an append writes to when the workbook is to change in place: the archive cannot be written
/// while its entries are still being copied out of it, so the new one is built next to it and moved over it
/// at the end.
///
/// The move replaces the workbook in one step where the file system can do that. Where it cannot, the
/// workbook is removed first; if the move then still fails, the file is the only copy of the workbook and is
/// kept. In every other case the file is removed when the append does not get to its end. DuckDB's file
/// system is only there during a callback, so a statement that ends without one is cleaned up after through
/// the file system of the operating system.
#[derive(Debug)]
struct SiblingFile {
    path: String,
    target: String,
    is_kept: bool,
}

impl SiblingFile {
    fn new(target: &str) -> SiblingFile {
        static COUNTER: AtomicU64 = AtomicU64::new(0);
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map_or(0, |elapsed| elapsed.as_nanos());
        let sequence = COUNTER.fetch_add(1, Ordering::Relaxed);
        SiblingFile {
            path: format!("{target}.xlsxapp.tmp.{nanos}.{sequence}"),
            target: target.to_owned(),
            is_kept: false,
        }
    }

    fn replace_target(&mut self, files: &impl FileOperations) -> Result<()> {
        if files.move_item(&self.path, &self.target).is_ok() {
            return Ok(());
        }
        let mut target_removed = false;
        let made_room = retry(files, || {
            // Without the new workbook there is nothing to make room for.
            if files.exists(&self.path)? && files.exists(&self.target)? {
                target_removed |= files.remove(&self.target)?;
            }
            Ok(())
        });
        let moved =
            made_room.and_then(|()| retry(files, || files.move_item(&self.path, &self.target)));
        let Err(error) = moved else {
            return Ok(());
        };
        let error = reraise(error);
        if target_removed || !files.exists(&self.target).unwrap_or(false) {
            self.is_kept = true;
            return Err(duckdb_neo::Error {
                code: error.code,
                message: format!(
                    "{}. '{}' is gone; the new workbook was kept as '{}'",
                    error.message, self.target, self.path
                ),
            });
        }
        self.remove(files);
        Err(error)
    }

    fn remove(&self, files: &impl FileOperations) {
        let _ = files.remove(&self.path);
    }
}

impl Drop for SiblingFile {
    fn drop(&mut self) {
        if !self.is_kept {
            let _ = std::fs::remove_file(&self.path);
        }
    }
}

/// What it takes to turn the archive with the finished sheet into a workbook: the parts of a new workbook, or
/// the patched parts of the existing one, which may be built in a file next to it.
#[derive(Debug)]
enum Completion {
    NewWorkbook,
    Appended {
        plan: AppendPlan,
        sibling: Option<SiblingFile>,
    },
}

/// A workbook that is being written, in the file DuckDB chose for it, which need not be the path of the
/// statement. The sheet comes before the completion so that the file is closed before a sibling file is
/// removed.
#[derive(Debug)]
struct Export {
    sheet: SheetWriter<ZipWriter<FileWriter>>,
    completion: Completion,
}

impl Export {
    fn begin(context: &Context, plan: &WritePlan, output_path: &str) -> Result<Export> {
        let file_system = FileSystem::new(context)?;
        let existing_path = file_system.expand(&plan.file_path).map_err(reraise)?;
        let appends = plan.mode != WriteMode::Create
            && file_system.exists(&existing_path).map_err(reraise)?;

        let mut export = if appends {
            Export::begin_append(&file_system, plan, &existing_path, output_path)?
        } else {
            let zip = ZipWriter::new(FileWriter::create(&file_system, output_path)?);
            Export::begin_sheet(zip, plan, FIRST_SHEET_FILE, CellStyles::NEW_WORKBOOK).map(
                |sheet| Export {
                    sheet,
                    completion: Completion::NewWorkbook,
                },
            )?
        };
        if plan.header
            && let Err(error) = export.write_header(plan)
        {
            export.discard(&file_system);
            return Err(error);
        }
        Ok(export)
    }

    fn begin_append(
        file_system: &FileSystem<'_>,
        plan: &WritePlan,
        existing_path: &str,
        output_path: &str,
    ) -> Result<Export> {
        let mut existing = Archive::open(file_system, existing_path)?;
        let replace = plan.mode == WriteMode::Replace;
        let append = AppendPlan::new(&mut existing, existing_path, &plan.sheet_name, replace)?;

        let sibling = (existing_path == output_path).then(|| SiblingFile::new(existing_path));
        let write_path = sibling.as_ref().map_or(output_path, |file| &file.path);
        if file_system.exists(write_path).map_err(reraise)? {
            file_system.remove(write_path).map_err(reraise)?;
        }
        let mut zip = ZipWriter::new(FileWriter::create(file_system, write_path)?);
        let sheet = append
            .copy_kept_entries(&mut existing, &mut zip)
            .and_then(|()| Export::begin_sheet(zip, plan, &append.sheet_file, append.styles));
        match sheet {
            Ok(sheet) => Ok(Export {
                sheet,
                completion: Completion::Appended {
                    plan: append,
                    sibling,
                },
            }),
            Err(error) => {
                if let Some(sibling) = &sibling {
                    sibling.remove(file_system);
                }
                Err(error)
            }
        }
    }

    fn begin_sheet(
        mut zip: ZipWriter<FileWriter>,
        plan: &WritePlan,
        sheet_file: &str,
        styles: CellStyles,
    ) -> Result<SheetWriter<ZipWriter<FileWriter>>> {
        zip.begin_file(&format!("{WORKSHEETS_DIRECTORY}{sheet_file}"))
            .map_err(|error| from_io(&error))?;
        Ok(SheetWriter::begin(
            zip,
            &plan.cell_kinds,
            styles,
            plan.row_limit,
        ))
    }

    fn write_header(&mut self, plan: &WritePlan) -> Result<()> {
        self.sheet.begin_row();
        for (column, name) in plan.column_names.iter().enumerate() {
            self.sheet.text_cell(column, name)?;
        }
        self.sheet.end_row()
    }

    fn write_rows(
        &mut self,
        context: &Context,
        plan: &WritePlan,
        input: &ColumnDataCollection<'_>,
    ) -> Result<()> {
        if plan.conversions.column_count() == 0 {
            for _ in 0..input.len()? {
                self.sheet.begin_row();
                self.sheet.end_row()?;
            }
            return Ok(());
        }
        let mut converter = TextConverter::new(context, &plan.conversions)?;
        let mut scan = input.to_scan(context)?;
        while let Some(chunk) = scan.next_chunk()? {
            converter.convert(&chunk)?;
            let columns = converter.texts()?;
            for row in 0..chunk.row_count()? {
                self.sheet.begin_row();
                for (column, texts) in columns.iter().enumerate() {
                    if let Some(text) = texts.get(row)? {
                        self.sheet.cell(column, text.as_bytes())?;
                    }
                }
                self.sheet.end_row()?;
            }
        }
        Ok(())
    }

    fn finish(self, context: &Context, plan: &WritePlan) -> Result<()> {
        let Export {
            sheet,
            mut completion,
        } = self;
        let file_system = FileSystem::new(context)?;
        let written = sheet.finish().and_then(|mut zip| {
            zip.end_file().map_err(|error| from_io(&error))?;
            match &completion {
                Completion::NewWorkbook => write_workbook_parts(&mut zip, &plan.sheet_name)
                    .map_err(|error| from_io(&error))?,
                Completion::Appended { plan, .. } => plan.write_patched_parts(&mut zip)?,
            }
            zip.finish().map_err(|error| from_io(&error))?.close()
        });
        let Completion::Appended {
            sibling: Some(sibling),
            ..
        } = &mut completion
        else {
            return written;
        };
        match written {
            Ok(()) => sibling.replace_target(&file_system),
            Err(error) => {
                sibling.remove(&file_system);
                Err(error)
            }
        }
    }

    fn discard(self, file_system: &FileSystem<'_>) {
        let Export { sheet, completion } = self;
        drop(sheet);
        if let Completion::Appended {
            sibling: Some(sibling),
            ..
        } = &completion
        {
            sibling.remove(file_system);
        }
    }
}

/// The workbook one output file of a `COPY ... TO` becomes.
///
/// DuckDB hands over the batches of a file one at a time and in the order of their rows, but not from one
/// thread, hence the lock. A workbook that failed is given up, and whatever comes after fails as well.
#[derive(Debug)]
pub struct WorkbookOutput {
    export: Mutex<Option<Export>>,
}

impl WorkbookOutput {
    /// Starts the workbook in `output_path`, the file DuckDB chose for it, and writes the header row.
    pub fn begin(context: &Context, plan: &WritePlan, output_path: &str) -> Result<WorkbookOutput> {
        let export = Export::begin(context, plan, output_path)?;
        Ok(WorkbookOutput {
            export: Mutex::new(Some(export)),
        })
    }

    /// Writes the rows of `input` after those the sheet has.
    pub fn write_rows(
        &self,
        context: &Context,
        plan: &WritePlan,
        input: &ColumnDataCollection<'_>,
    ) -> Result<()> {
        let mut guard = self.lock()?;
        let export = guard.as_mut().ok_or_else(gone)?;
        let written = export.write_rows(context, plan, input);
        if written.is_err()
            && let Some(export) = guard.take()
        {
            export.discard(&FileSystem::new(context)?);
        }
        written
    }

    /// Ends the sheet, writes the other parts of the workbook and closes the file.
    pub fn finish(&self, context: &Context, plan: &WritePlan) -> Result<()> {
        let export = self.lock()?.take().ok_or_else(gone)?;
        export.finish(context, plan)
    }

    fn lock(&self) -> Result<MutexGuard<'_, Option<Export>>> {
        self.export
            .lock()
            .map_err(|_| internal("COPY TO xlsx: writing the workbook was interrupted"))
    }
}

fn gone() -> duckdb_neo::Error {
    internal("COPY TO xlsx: the workbook is not being written any more")
}

#[cfg(test)]
mod tests {
    use std::{
        cell::{Cell, RefCell},
        collections::BTreeSet,
    };

    use super::*;
    use crate::error::io;

    const TARGET: &str = "book.xlsx";

    /// Files as names only; the first `failing_moves` moves and `failing_removes` removals fail.
    #[derive(Default)]
    struct FakeFiles {
        names: RefCell<BTreeSet<String>>,
        failing_moves: Cell<u32>,
        failing_removes: Cell<u32>,
        calls: RefCell<Vec<&'static str>>,
    }

    impl FakeFiles {
        fn with(names: &[&str]) -> FakeFiles {
            let files = FakeFiles::default();
            files
                .names
                .borrow_mut()
                .extend(names.iter().map(|name| (*name).to_owned()));
            files
        }

        fn names(&self) -> Vec<String> {
            self.names.borrow().iter().cloned().collect()
        }

        fn fails(&self, left: &Cell<u32>) -> bool {
            let fails = left.get() > 0;
            left.set(left.get().saturating_sub(1));
            fails
        }
    }

    impl FileOperations for FakeFiles {
        fn exists(&self, path: &str) -> Result<bool> {
            Ok(self.names.borrow().contains(path))
        }

        fn remove(&self, path: &str) -> Result<bool> {
            self.calls.borrow_mut().push("remove");
            if self.fails(&self.failing_removes) {
                return Err(io("IO Error: the file is in use"));
            }
            Ok(self.names.borrow_mut().remove(path))
        }

        fn move_item(&self, from: &str, to: &str) -> Result<()> {
            self.calls.borrow_mut().push("move");
            if self.fails(&self.failing_moves) || !self.names.borrow_mut().remove(from) {
                return Err(io("IO Error: Could not rename file"));
            }
            self.names.borrow_mut().insert(to.to_owned());
            Ok(())
        }

        fn pause(&self) {}
    }

    fn sibling() -> SiblingFile {
        SiblingFile {
            path: "book.xlsx.new".to_owned(),
            target: TARGET.to_owned(),
            is_kept: false,
        }
    }

    #[test]
    fn the_new_workbook_is_moved_over_the_old_one_in_one_step() {
        let files = FakeFiles::with(&[TARGET, "book.xlsx.new"]);
        let mut sibling = sibling();
        sibling.replace_target(&files).unwrap();
        assert_eq!(files.names(), [TARGET]);
        assert_eq!(*files.calls.borrow(), ["move"]);
        assert!(!sibling.is_kept);
    }

    #[test]
    fn a_file_system_that_does_not_replace_gets_room_made_first() {
        let files = FakeFiles::with(&[TARGET, "book.xlsx.new"]);
        files.failing_moves.set(1);
        let mut sibling = sibling();
        sibling.replace_target(&files).unwrap();
        assert_eq!(files.names(), [TARGET]);
        assert_eq!(*files.calls.borrow(), ["move", "remove", "move"]);
    }

    #[test]
    fn operations_are_tried_again_for_a_while() {
        let files = FakeFiles::with(&[TARGET, "book.xlsx.new"]);
        files.failing_moves.set(FILE_ATTEMPTS);
        files.failing_removes.set(FILE_ATTEMPTS - 1);
        let mut sibling = sibling();
        sibling.replace_target(&files).unwrap();
        assert_eq!(files.names(), [TARGET]);
        let calls = files.calls.borrow();
        let moves = calls.iter().filter(|call| **call == "move").count();
        assert_eq!(moves as u32, FILE_ATTEMPTS + 1);
        assert_eq!(calls.len() as u32, 2 * FILE_ATTEMPTS + 1);
    }

    #[test]
    fn a_workbook_that_cannot_be_removed_stays_and_the_new_one_goes() {
        let files = FakeFiles::with(&[TARGET, "book.xlsx.new"]);
        files.failing_moves.set(1);
        files.failing_removes.set(FILE_ATTEMPTS);
        let mut sibling = sibling();
        let error = sibling.replace_target(&files).unwrap_err();
        assert_eq!(error.message, "the file is in use");
        assert_eq!(files.names(), [TARGET]);
        assert!(!sibling.is_kept);
    }

    #[test]
    fn the_new_workbook_is_kept_once_the_old_one_is_removed() {
        let files = FakeFiles::with(&[TARGET, "book.xlsx.new"]);
        files.failing_moves.set(u32::MAX);
        let mut sibling = sibling();
        let error = sibling.replace_target(&files).unwrap_err();
        assert_eq!(
            error.message,
            "Could not rename file. 'book.xlsx' is gone; the new workbook was kept as 'book.xlsx.new'"
        );
        assert_eq!(error.code, duckdb_neo::error::DuckDBError::IoGeneral);
        assert_eq!(files.names(), ["book.xlsx.new"]);
        assert!(sibling.is_kept);
        let removals = files
            .calls
            .borrow()
            .iter()
            .filter(|call| **call == "remove")
            .count();
        assert_eq!(removals, 1);
    }

    #[test]
    fn the_new_workbook_is_kept_when_the_old_one_has_vanished() {
        let files = FakeFiles::with(&["book.xlsx.new"]);
        files.failing_moves.set(u32::MAX);
        let mut sibling = sibling();
        let error = sibling.replace_target(&files).unwrap_err();
        assert!(
            error.message.ends_with("was kept as 'book.xlsx.new'"),
            "{}",
            error.message
        );
        assert_eq!(files.names(), ["book.xlsx.new"]);
        assert!(sibling.is_kept);
        assert!(!files.calls.borrow().contains(&"remove"));
    }

    #[test]
    fn the_old_workbook_is_not_removed_for_a_new_one_that_is_missing() {
        let files = FakeFiles::with(&[TARGET]);
        let mut sibling = sibling();
        let error = sibling.replace_target(&files).unwrap_err();
        assert_eq!(error.message, "Could not rename file");
        assert_eq!(files.names(), [TARGET]);
        assert!(!sibling.is_kept);
    }

    #[test]
    fn a_kept_file_outlives_its_owner() {
        let path = std::env::temp_dir().join(format!(
            "excel_core_{}_kept_sibling.tmp",
            std::process::id()
        ));
        let path_text = path.to_str().unwrap().to_owned();
        for is_kept in [false, true] {
            std::fs::write(&path, b"workbook").unwrap();
            drop(SiblingFile {
                path: path_text.clone(),
                target: TARGET.to_owned(),
                is_kept,
            });
            assert_eq!(path.exists(), is_kept);
        }
        std::fs::remove_file(&path).unwrap();
    }
}
