#include "xlsx/read_xlsx.hpp"

#include "duckdb/common/helper.hpp"
#include "duckdb/common/types/time.hpp"
#include "duckdb/function/replacement_scan.hpp"
#include "duckdb/function/table_function.hpp"
#include "duckdb/main/database.hpp"
#include "duckdb/main/extension_util.hpp"
#include "duckdb/main/query_result.hpp"
#include "duckdb/parser/expression/constant_expression.hpp"
#include "duckdb/parser/expression/function_expression.hpp"
#include "duckdb/parser/tableref/table_function_ref.hpp"
#include "xlsx/parsers/content_types_parser.hpp"
#include "xlsx/parsers/relationship_parser.hpp"
#include "xlsx/parsers/shared_strings_parser.hpp"
#include "xlsx/parsers/stylesheet_parser.hpp"
#include "xlsx/parsers/workbook_parser.hpp"
#include "xlsx/parsers/worksheet_parser.hpp"
#include "xlsx/string_table.hpp"
#include "xlsx/xlsx_parts.hpp"
#include "xlsx/xml_parser.hpp"
#include "xlsx/xml_util.hpp"
#include "xlsx/zip_file.hpp"

#include <utf8proc.hpp>

namespace duckdb {

//-------------------------------------------------------------------
// Util
//-------------------------------------------------------------------
// Helper function for UTF-8 aware space trimming
static string TrimWhitespace(const string &col_name) {
	utf8proc_int32_t codepoint;
	auto str = reinterpret_cast<const utf8proc_uint8_t *>(col_name.c_str());
	idx_t size = col_name.size();
	// Find the first character that is not left trimmed
	idx_t begin = 0;
	while (begin < size) {
		auto bytes = utf8proc_iterate(str + begin, NumericCast<utf8proc_ssize_t>(size - begin), &codepoint);
		D_ASSERT(bytes > 0);
		if (utf8proc_category(codepoint) != UTF8PROC_CATEGORY_ZS) {
			break;
		}
		begin += NumericCast<idx_t>(bytes);
	}

	// Find the last character that is not right trimmed
	idx_t end = begin;
	for (auto next = begin; next < col_name.size();) {
		auto bytes = utf8proc_iterate(str + next, NumericCast<utf8proc_ssize_t>(size - next), &codepoint);
		D_ASSERT(bytes > 0);
		next += NumericCast<idx_t>(bytes);
		if (utf8proc_category(codepoint) != UTF8PROC_CATEGORY_ZS) {
			end = next;
		}
	}

	// return the trimmed string
	return col_name.substr(begin, end - begin);
}

static string NormalizeColumnName(const string &col_name) {
	// normalize UTF8 characters to NFKD
	auto nfkd = utf8proc_NFKD(reinterpret_cast<const utf8proc_uint8_t *>(col_name.c_str()),
	                          NumericCast<utf8proc_ssize_t>(col_name.size()));
	const string col_name_nfkd = string(const_char_ptr_cast(nfkd), strlen(const_char_ptr_cast(nfkd)));
	free(nfkd);

	// only keep ASCII characters 0-9 a-z A-Z and replace spaces with regular whitespace
	string col_name_ascii = "";
	for (idx_t i = 0; i < col_name_nfkd.size(); i++) {
		if (col_name_nfkd[i] == '_' || (col_name_nfkd[i] >= '0' && col_name_nfkd[i] <= '9') ||
		    (col_name_nfkd[i] >= 'A' && col_name_nfkd[i] <= 'Z') ||
		    (col_name_nfkd[i] >= 'a' && col_name_nfkd[i] <= 'z')) {
			col_name_ascii += col_name_nfkd[i];
		} else if (StringUtil::CharacterIsSpace(col_name_nfkd[i])) {
			col_name_ascii += " ";
		}
	}

	// trim whitespace and replace remaining whitespace by _
	string col_name_trimmed = TrimWhitespace(col_name_ascii);
	string col_name_cleaned = "";
	bool in_whitespace = false;
	for (idx_t i = 0; i < col_name_trimmed.size(); i++) {
		if (col_name_trimmed[i] == ' ') {
			if (!in_whitespace) {
				col_name_cleaned += "_";
				in_whitespace = true;
			}
		} else {
			col_name_cleaned += col_name_trimmed[i];
			in_whitespace = false;
		}
	}

	// don't leave string empty; if not empty, make lowercase
	if (col_name_cleaned.empty()) {
		col_name_cleaned = "_";
	} else {
		col_name_cleaned = StringUtil::Lower(col_name_cleaned);
	}

	// prepend _ if name starts with a digit or is a reserved keyword
	auto keyword = KeywordHelper::KeywordCategoryType(col_name_cleaned);
	if (keyword == KeywordCategory::KEYWORD_TYPE_FUNC || keyword == KeywordCategory::KEYWORD_RESERVED ||
	    (col_name_cleaned[0] >= '0' && col_name_cleaned[0] <= '9')) {
		col_name_cleaned = "_" + col_name_cleaned;
	}
	return col_name_cleaned;
}

static void CleanColumnNames(vector<string> &names, bool normalize) {
	for (auto &name : names) {
		// normalize names or at least trim whitespace
		if (normalize) {
			name = NormalizeColumnName(name);
		} else {
			name = TrimWhitespace(name);
		}
	}
}

//-------------------------------------------------------------------
// Meta
//-------------------------------------------------------------------
static void ParseXLSXFileMeta(const unique_ptr<XLSXReadData> &result, ZipFileReader &reader) {

	unordered_map<string, string> candidate_sheets;
	string primary_sheet;

	// Extract the content types to get the primary sheet
	if (!reader.TryOpenEntry("[Content_Types].xml")) {
		throw BinderException("No [Content_Types].xml found in xlsx file");
	}
	const auto ctypes = ContentParser::ParseContentTypes(reader);
	reader.CloseEntry();

	if (!reader.TryOpenEntry("xl/workbook.xml")) {
		throw BinderException("No xl/workbook.xml found in xlsx file");
	}
	const auto sheets = WorkBookParser::GetSheets(reader);
	reader.CloseEntry();

	if (!reader.TryOpenEntry("xl/_rels/workbook.xml.rels")) {
		throw BinderException("No xl/_rels/workbook.xml.rels found in xlsx file");
	}
	const auto wbrels = RelParser::ParseRelations(reader);
	reader.CloseEntry();

	// TODO: Detect if we have a shared string table

	// Resolve the sheet names to the paths
	// Start by mapping rid to sheet path
	unordered_map<string, string> rid_to_sheet_map;
	for (auto &rel : wbrels) {
		if (StringUtil::EndsWith(rel.type, "/worksheet")) {
			rid_to_sheet_map[rel.id] = rel.target;
		}
	}

	// Now map name to rid and rid to sheet path
	for (auto &sheet : sheets) {
		const auto found = rid_to_sheet_map.find(sheet.second);
		if (found != rid_to_sheet_map.end()) {

			// Normalize everything to absolute paths
			if (StringUtil::StartsWith(found->second, "/xl/")) {
				candidate_sheets[sheet.first] = found->second.substr(1);
			} else {
				candidate_sheets[sheet.first] = "xl/" + found->second;
			}

			// Set the first sheet we find as the primary sheet
			if (primary_sheet.empty()) {
				primary_sheet = sheet.first;
			}
		}
	}

	if (candidate_sheets.empty()) {
		throw BinderException("No sheets found in xlsx file (is the file corrupt?)");
	}

	// Default to the primary sheet if no option is given
	auto &options = result->options;
	if (options.sheet.empty()) {
		options.sheet = primary_sheet;
	}

	const auto found = candidate_sheets.find(options.sheet);
	if (found == candidate_sheets.end()) {
		// Throw a helpful error message
		vector<string> all_sheets;
		for (auto &sheet : candidate_sheets) {
			all_sheets.push_back(sheet.first);
		}
		auto suggestions = StringUtil::CandidatesErrorMessage(all_sheets, options.sheet, "Did you mean");
		throw BinderException("Sheet \"%s\" not found in xlsx file \"%s\"%s", options.sheet, result->file_path,
		                      suggestions);
	}
	result->sheet_path = found->second;
}

static void ResolveColumnNames(vector<XLSXCell> &header_cells, ZipFileReader &archive) {

	vector<idx_t> shared_string_ids;
	vector<idx_t> shared_string_pos;

	for (idx_t i = 0; i < header_cells.size(); i++) {
		auto &cell = header_cells[i];
		if (cell.type == XLSXCellType::SHARED_STRING) {
			shared_string_ids.push_back(std::strtol(cell.data.c_str(), nullptr, 10));
			shared_string_pos.push_back(i);
		}
	}

	// There is nothing to resolve
	if (shared_string_ids.empty()) {
		return;
	}

	// Resolve the shared strings
	if (!archive.TryOpenEntry("xl/sharedStrings.xml")) {
		throw BinderException("No shared strings found in xlsx file");
	}
	SharedStringSearcher searcher(shared_string_ids);
	searcher.ParseAll(archive);
	archive.CloseEntry();

	auto &shared_strings = searcher.GetResult();

	// Replace the shared strings with the resolved strings
	for (idx_t i = 0; i < shared_string_pos.size(); i++) {
		header_cells[shared_string_pos[i]].data = shared_strings.at(shared_string_ids[i]);
	}
}

void ReadXLSX::ParseOptions(XLSXReadOptions &options, const named_parameter_map_t &input) {

	// Check which sheet to use, default to the primary sheet
	const auto sheet_opt = input.find("sheet");
	if (sheet_opt != input.end()) {
		// We need to escape all user-supplied strings when searching for them in the XML
		options.sheet = EscapeXMLString(StringValue::Get(sheet_opt->second));
	}

	// Get the header mode
	const auto header_mode_opt = input.find("header");
	if (header_mode_opt != input.end()) {
		options.header_mode =
		    BooleanValue::Get(header_mode_opt->second) ? XLSXHeaderMode::FORCE : XLSXHeaderMode::NEVER;
	}

	const auto all_varchar_opt = input.find("all_varchar");
	if (all_varchar_opt != input.end()) {
		options.all_varchar = BooleanValue::Get(all_varchar_opt->second);
	}

	const auto ignore_errors_opt = input.find("ignore_errors");
	if (ignore_errors_opt != input.end()) {
		options.ignore_errors = BooleanValue::Get(ignore_errors_opt->second);
	}

	const auto normalize_names_opt = input.find("normalize_names");
	if (normalize_names_opt != input.end()) {
		options.normalize_names = BooleanValue::Get(normalize_names_opt->second);
	}

	const auto range_opt = input.find("range");
	if (range_opt != input.end()) {
		const auto range_str = StringValue::Get(range_opt->second);
		XLSXCellRange range;
		if (!range.TryParse(range_str.c_str())) {
			throw BinderException("Invalid range '%s' specified", range_str);
		}
		if (!range.IsValid()) {
			throw BinderException("Invalid range '%s' specified", range_str);
		}
		// We do allow more rows than the maximum, but not more columns, because DuckDB kind of breaks down otherwise.
		if (range.Width() > XLSX_MAX_CELL_COLS) {
			throw BinderException("Invalid range '%s' specified", range_str);
		}

		// Make sure the range is inclusive of the last cell
		range.end.col++;
		range.end.row++;

		options.range = range;
		options.has_explicit_range = true;

		// Default to not stop at empty if a range is specified
		options.stop_at_empty = false;
	}

	const auto stop_at_empty_op = input.find("stop_at_empty");
	if (stop_at_empty_op != input.end()) {
		options.stop_at_empty = BooleanValue::Get(stop_at_empty_op->second);
	}

	const auto empty_as_varchar_opt = input.find("empty_as_varchar");
	if (empty_as_varchar_opt != input.end()) {
		options.default_cell_type =
		    BooleanValue::Get(empty_as_varchar_opt->second) ? XLSXCellType::INLINE_STRING : XLSXCellType::NUMBER;
	}
}

static void ParseStyleSheet(const unique_ptr<XLSXReadData> &result, ZipFileReader &archive) {
	// Parse the styles (so we can handle dates)
	if (archive.TryOpenEntry("xl/styles.xml")) {
		XLSXStyleParser style_parser;
		style_parser.ParseAll(archive);
		result->style_sheet = XLSXStyleSheet(std::move(style_parser.cell_styles));
		archive.CloseEntry();
	}
}

static void SniffRange(const unique_ptr<XLSXReadData> &result, ZipFileReader &archive) {
	if (!archive.TryOpenEntry(result->sheet_path)) {
		throw BinderException("Sheet '%s' not found in xlsx file", result->sheet_path);
	}
	RangeSniffer range_sniffer;
	range_sniffer.ParseAll(archive);
	archive.CloseEntry();
	result->options.range = range_sniffer.GetRange();
}

static void SniffHeader(const unique_ptr<XLSXReadData> &result, ZipFileReader &archive) {
	auto &options = result->options;

	if (!archive.TryOpenEntry(result->sheet_path)) {
		throw BinderException("Sheet '%s' not found in xlsx file", result->sheet_path);
	}
	HeaderSniffer sniffer(result->options.range, result->options.header_mode, result->options.has_explicit_range,
	                      result->options.default_cell_type);
	sniffer.ParseAll(archive);
	archive.CloseEntry();

	// This is the range of actual data in the sheet (header not included)
	options.range = sniffer.GetRange();

	auto &header_cells = sniffer.GetHeaderCells();
	auto &column_cells = sniffer.GetColumnCells();

	if (column_cells.empty()) {
		if (header_cells.empty()) {
			if (!options.has_explicit_range) {
				throw BinderException("No rows found in xlsx file");
			}
			// Otherwise, add a header row with the column names in the range
			for (idx_t i = options.range.beg.col; i < options.range.end.col; i++) {
				XLSXCellPos pos(options.range.beg.row, i);
				header_cells.emplace_back(result->options.default_cell_type, pos, pos.ToString(), 0);
			}
		}
		// Else, we have a header row but no data rows
		// Users seem to expect this to work, so we allow it by creating an empty dummy row of varchars
		for (auto &cell : header_cells) {
			column_cells.emplace_back(result->options.default_cell_type, cell.cell, "", 0);
		}
	}

	// Resolve any shared strings in the header
	ResolveColumnNames(header_cells, archive);

	// Set the return names
	for (auto &cell : header_cells) {
		result->column_names.push_back(cell.data);
	}

	// Convert excel types to duckdb types
	for (auto &cell : column_cells) {
		auto duckdb_type = cell.GetDuckDBType(result->options.all_varchar, result->style_sheet);
		result->return_types.push_back(duckdb_type);
		result->source_types.push_back(cell.type);
	}
}

void ReadXLSX::ResolveSheet(const unique_ptr<XLSXReadData> &result, ZipFileReader &archive) {
	// Parse the meta
	ParseXLSXFileMeta(result, archive);
	// Parse the style sheet
	ParseStyleSheet(result, archive);
	if (!result->options.has_explicit_range) {
		// Sniff content range if required
		SniffRange(result, archive);
	}
	// Sniff header
	SniffHeader(result, archive);
}

//-------------------------------------------------------------------
// Bind
//-------------------------------------------------------------------
static unique_ptr<FunctionData> Bind(ClientContext &context, TableFunctionBindInput &input,
                                     vector<LogicalType> &return_types, vector<string> &names) {
	auto result = make_uniq<XLSXReadData>();
	// Get the file name
	const auto file_path = StringValue::Get(input.inputs[0]);

	// Glob here so that we auto load any required extension filesystems
	auto &fs = FileSystem::GetFileSystem(context);
	auto files = fs.GlobFiles(file_path, context, FileGlobOptions::ALLOW_EMPTY);
	if (files.empty()) {
		// Use our own error message to be less confusing
		throw IOException("Cannot open file \"%s\": No such file or directory", file_path);
	}

	// We dont support multi-file-reading, so just take the first file for now
	result->file_path = files.front();

	// Open the archive
	ZipFileReader archive(context, result->file_path);

	// Parse the options
	ReadXLSX::ParseOptions(result->options, input.named_parameters);

	// Resolve the sheet
	ReadXLSX::ResolveSheet(result, archive);

	return_types = result->return_types;
	names = result->column_names;

	// Clean and normalize names if requested
	CleanColumnNames(names, result->options.normalize_names);

	// Deduplicate column names
	QueryResult::DeduplicateColumns(names);

	return std::move(result);
}

//-------------------------------------------------------------------
// Global State
//-------------------------------------------------------------------
class XLSXGlobalState final : public GlobalTableFunctionState {
public:
	explicit XLSXGlobalState(ClientContext &context, const string &file_name, const XLSXCellRange &range,
	                         bool stop_at_empty)
	    : archive(context, file_name), strings(BufferAllocator::Get(context)),
	      parser(context, range, strings, stop_at_empty),
	      buffer(make_unsafe_uniq_array_uninitialized<char>(BUFFER_SIZE)) {

		cast_vec.Initialize(context, {LogicalType::DOUBLE});
	}

	ZipFileReader archive;
	StringTable strings;
	SheetParser parser;
	unsafe_unique_array<char> buffer;

	XMLParseResult status = XMLParseResult::OK;

	string cast_err;
	DataChunk cast_vec;

	atomic<idx_t> stream_pos = {0};
	idx_t stream_len = 0;

	// 8kb buffer
	static constexpr auto BUFFER_SIZE = 8096;
};

static unique_ptr<GlobalTableFunctionState> InitGlobal(ClientContext &context, TableFunctionInitInput &input) {
	auto &data = input.bind_data->Cast<XLSXReadData>();
	auto &options = data.options;
	auto state = make_uniq<XLSXGlobalState>(context, data.file_path, data.options.range, options.stop_at_empty);

	// Check if there is a string table. If there is, extract it
	if (state->archive.TryOpenEntry("xl/sharedStrings.xml")) {
		SharedStringParser::ParseStringTable(state->archive, state->strings);
		state->archive.CloseEntry();
	}

	// Open the main sheet for reading
	if (!state->archive.TryOpenEntry(data.sheet_path)) {
		// This should never happen, we've already checked this in the bind function
		throw InvalidInputException("Sheet '%s' not found in xlsx file", data.sheet_path);
	}

	// Set the progress counters
	state->stream_len = state->archive.GetEntryLen();
	state->stream_pos = 0;

	return std::move(state);
}

//-------------------------------------------------------------------
// Execute
//-------------------------------------------------------------------

int64_t ExcelToEpochUS(const double serial) {
	// Convert to microseconds since epoch
	static constexpr auto SECONDS_PER_DAY = 86400UL;
	static constexpr auto MICROSECONDS_PER_SECOND = 1000000UL;
	static constexpr auto DAYS_BETWEEN_1900_AND_1970 = 25569UL;

	// Excel serial is days since 1900-01-01
	const auto serial_days = serial;
	auto serial_secs = serial_days * SECONDS_PER_DAY;

	if (std::fabs(serial_secs - std::round(serial_secs)) < 1e-3) {
		serial_secs = std::round(serial_secs);
	}

	const auto epoch_secs = serial_secs - (DAYS_BETWEEN_1900_AND_1970 * SECONDS_PER_DAY);
	const auto epoch_micros = epoch_secs * MICROSECONDS_PER_SECOND;

	// Clamp to the range. Theres not much we can do if the value is out of range
	if (epoch_micros <= static_cast<double>(NumericLimits<int64_t>::Minimum())) {
		return NumericLimits<int64_t>::Minimum();
	}
	if (epoch_micros >= static_cast<double>(NumericLimits<int64_t>::Maximum())) {
		return NumericLimits<int64_t>::Maximum();
	}

	return static_cast<int64_t>(epoch_micros);
}

static void TryCastFromString(XLSXGlobalState &state, bool ignore_errors, const idx_t col_idx, ClientContext &context,
                              Vector &target_col) {

	auto &chunk = state.parser.GetChunk();
	auto &source_col = chunk.data[col_idx];
	const auto row_count = chunk.size();

	const auto ok = VectorOperations::TryCast(context, source_col, target_col, row_count, &state.cast_err);
	if (!ok && !ignore_errors) {
		// Figure out which cell failed
		const auto &source_validity = FlatVector::Validity(source_col);
		const auto &target_validity = FlatVector::Validity(target_col);
		for (idx_t row_idx = 0; row_idx < row_count; row_idx++) {
			if (source_validity.RowIsValid(row_idx) != target_validity.RowIsValid(row_idx)) {
				// If the string is empty, allow it to be cast to NULL
				if (!FlatVector::GetData<string_t>(source_col)[row_idx].Empty()) {
					const auto cell_name = state.parser.GetCellName(row_idx, col_idx);
					throw InvalidInputException("read_xlsx: Failed to parse cell '%s': %s", cell_name, state.cast_err);
				}
			}
		}
	}
}

static void TryCastTime(XLSXGlobalState &state, bool ignore_errors, const idx_t col_idx, ClientContext &context,
                        Vector &target_col) {
	// First cast it to a double
	TryCastFromString(state, ignore_errors, col_idx, context, state.cast_vec.data[0]);

	// Then convert the double to a time
	const auto row_count = state.parser.GetChunk().size();
	UnaryExecutor::Execute<double, dtime_t>(state.cast_vec.data[0], target_col, row_count, [&](const double &input) {
		const auto epoch_us = ExcelToEpochUS(input);
		const auto stamp = Timestamp::FromEpochMicroSeconds(epoch_us);
		return Timestamp::GetTime(stamp);
	});
}

static void TryCastDate(XLSXGlobalState &state, bool ignore_errors, const idx_t col_idx, ClientContext &context,
                        Vector &target_col) {
	// First cast it to a double
	TryCastFromString(state, ignore_errors, col_idx, context, state.cast_vec.data[0]);

	// Then convert the double to a date
	const auto row_count = state.parser.GetChunk().size();
	UnaryExecutor::Execute<double, date_t>(state.cast_vec.data[0], target_col, row_count, [&](const double &input) {
		const auto epoch_us = ExcelToEpochUS(input);
		const auto stamp = Timestamp::FromEpochMicroSeconds(epoch_us);
		return Timestamp::GetDate(stamp);
	});
}

static void TryCastTimestamp(XLSXGlobalState &state, bool ignore_errors, const idx_t col_idx, ClientContext &context,
                             Vector &target_col) {
	// First cast it to a double
	TryCastFromString(state, ignore_errors, col_idx, context, state.cast_vec.data[0]);

	// Then convert the double to a timestamp
	const auto row_count = state.parser.GetChunk().size();
	UnaryExecutor::Execute<double, timestamp_t>(state.cast_vec.data[0], target_col, row_count,
	                                            [&](const double &input) {
		                                            const auto epoch_us = ExcelToEpochUS(input);
		                                            return Timestamp::FromEpochMicroSeconds(epoch_us);
	                                            });
}

static void Execute(ClientContext &context, TableFunctionInput &data, DataChunk &output) {
	auto &bind_data = data.bind_data->Cast<XLSXReadData>();
	auto &options = bind_data.options;
	auto &gstate = data.global_state->Cast<XLSXGlobalState>();
	const auto buffer = gstate.buffer.get();

	auto &stream = gstate.archive;
	auto &parser = gstate.parser;
	auto &status = gstate.status;

	// Ready the chunk
	auto &chunk = parser.GetChunk();
	chunk.Reset();

	while (chunk.size() != STANDARD_VECTOR_SIZE) {
		if (status == XMLParseResult::SUSPENDED) {
			if (parser.FoundSkippedRow()) {
				if (options.stop_at_empty) {
					status = XMLParseResult::ABORTED;
					break;
				}
				parser.SkipRows();
				continue;
			}

			// Resume normally
			status = parser.Resume();
			continue;
		}
		if (stream.IsDone()) {
			break;
		}
		if (status == XMLParseResult::ABORTED) {
			break;
		}

		// Otherwise, read more data
		const auto read_size = stream.Read(buffer, XLSXGlobalState::BUFFER_SIZE);

		// Update the progess
		gstate.stream_pos += read_size;

		status = parser.Parse(buffer, read_size, stream.IsDone());
	}

	// Pad with empty rows if wanted (and needed)
	if (options.has_explicit_range && !options.stop_at_empty) {
		parser.FillRows();
	}

	// Cast all the strings to the correct types, unless they are already strings in which case we reference them
	const auto row_count = chunk.size();

	for (idx_t col_idx = 0; col_idx < output.ColumnCount(); col_idx++) {
		auto &source_col = chunk.data[col_idx];
		auto &target_col = output.data[col_idx];
		auto &xlsx_type = bind_data.source_types[col_idx];

		const auto source_type = source_col.GetType().id();
		const auto target_type = target_col.GetType().id();

		if (source_type == target_type) {
			// If the types are the same, reference the column
			target_col.Reference(source_col);
			continue;
		}

		// Clear the cast vector
		gstate.cast_vec.Reset();

		if (xlsx_type == XLSXCellType::NUMBER && target_type == LogicalTypeId::TIME) {
			TryCastTime(gstate, options.ignore_errors, col_idx, context, target_col);
		} else if (xlsx_type == XLSXCellType::NUMBER && target_type == LogicalTypeId::DATE) {
			TryCastDate(gstate, options.ignore_errors, col_idx, context, target_col);
		} else if (xlsx_type == XLSXCellType::NUMBER && target_type == LogicalTypeId::TIMESTAMP) {
			TryCastTimestamp(gstate, options.ignore_errors, col_idx, context, target_col);
		} else {
			// Cast the from string to the target type
			TryCastFromString(gstate, options.ignore_errors, col_idx, context, target_col);
		}
	}
	output.SetCapacity(row_count);
	output.SetCardinality(row_count);

	output.Verify();
}

//-------------------------------------------------------------------
// Progress
//-------------------------------------------------------------------
static double Progress(ClientContext &context, const FunctionData *bind_data_p,
                       const GlobalTableFunctionState *global_state) {
	if (!global_state) {
		return 0;
	}

	const auto &state = global_state->Cast<XLSXGlobalState>();
	const auto pos = static_cast<double>(state.stream_pos.load());
	const auto len = static_cast<double>(state.stream_len);

	return (pos == 0 || len == 0) ? 0 : (pos / len) * 100.0;
}

static unique_ptr<TableRef> XLSXReplacementScan(ClientContext &context, ReplacementScanInput &input,
                                                optional_ptr<ReplacementScanData> data) {
	const auto table_name = ReplacementScan::GetFullPath(input);
	const auto lower_name = StringUtil::Lower(table_name);

	if (!StringUtil::EndsWith(lower_name, ".xlsx")) {
		return nullptr;
	}

	auto result = make_uniq<TableFunctionRef>();
	vector<unique_ptr<ParsedExpression>> children;
	children.push_back(make_uniq<ConstantExpression>(Value(table_name)));
	result->function = make_uniq_base<ParsedExpression, FunctionExpression>("read_xlsx", std::move(children));

	return std::move(result);
}

//-------------------------------------------------------------------
// Register
//-------------------------------------------------------------------
TableFunction ReadXLSX::GetFunction() {

	TableFunction read_xlsx("read_xlsx", {LogicalType::VARCHAR}, Execute, Bind);
	read_xlsx.init_global = InitGlobal;
	read_xlsx.table_scan_progress = Progress;

	// Parameters
	read_xlsx.named_parameters["header"] = LogicalType::BOOLEAN;
	read_xlsx.named_parameters["all_varchar"] = LogicalType::BOOLEAN;
	read_xlsx.named_parameters["ignore_errors"] = LogicalType::BOOLEAN;
	read_xlsx.named_parameters["range"] = LogicalType::VARCHAR;
	read_xlsx.named_parameters["sheet"] = LogicalType::VARCHAR;
	read_xlsx.named_parameters["stop_at_empty"] = LogicalType::BOOLEAN;
	read_xlsx.named_parameters["empty_as_varchar"] = LogicalType::BOOLEAN;
	read_xlsx.named_parameters["normalize_names"] = LogicalType::BOOLEAN;

	return read_xlsx;
}

void ReadXLSX::Register(DatabaseInstance &db) {
	ExtensionUtil::RegisterFunction(db, GetFunction());
	db.config.replacement_scans.emplace_back(XLSXReplacementScan);
}

} // namespace duckdb