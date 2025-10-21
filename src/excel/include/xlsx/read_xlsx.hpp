#pragma once
#include "duckdb/common/named_parameter_map.hpp"
#include "duckdb/function/table_function.hpp"
#include "xlsx/xlsx_parts.hpp"

namespace duckdb {

class DatabaseInstance;

struct WriteXLSX {
	static void Register(DatabaseInstance &db);
};

enum class XLSXHeaderMode : uint8_t { NEVER, MAYBE, FORCE };

class XLSXReadOptions {
public:
	string sheet;
	XLSXHeaderMode header_mode = XLSXHeaderMode::MAYBE;
	bool all_varchar = false;
	bool ignore_errors = false;
	bool stop_at_empty = true;
	bool has_explicit_range = false;
	bool normalize_names = false;
	XLSXCellType default_cell_type = XLSXCellType::NUMBER;
	XLSXCellRange range;
};

class XLSXReadData final : public TableFunctionData {
public:
	string file_path;
	string sheet_path;
	vector<string> all_files;

	vector<LogicalType> return_types;
	vector<XLSXCellType> source_types;
	vector<string> column_names;

	XLSXReadOptions options;
	XLSXStyleSheet style_sheet;
};

class ZipFileReader;

struct ReadXLSX {
	// options and file path need to be resolved already
	static void ParseOptions(XLSXReadOptions &options, const named_parameter_map_t &input);
	static void ResolveSheet(const unique_ptr<XLSXReadData> &result, ZipFileReader &archive);

	static void Register(DatabaseInstance &db);
	static TableFunction GetFunction();
};

} // namespace duckdb