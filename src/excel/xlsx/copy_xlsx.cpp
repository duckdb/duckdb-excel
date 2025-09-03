#include "duckdb/common/exception/conversion_exception.hpp"
#include "duckdb/function/copy_function.hpp"
#include "duckdb/main/database.hpp"
#include "duckdb/main/extension/extension_loader.hpp"
#include "duckdb/planner/expression/bound_cast_expression.hpp"
#include "duckdb/planner/expression/bound_function_expression.hpp"
#include "duckdb/planner/expression/bound_reference_expression.hpp"
#include "xlsx/read_xlsx.hpp"
#include "xlsx/xlsx_writer.hpp"

namespace duckdb {

//------------------------------------------------------------------------------
// Conversion Expressions
//------------------------------------------------------------------------------
static constexpr auto DAYS_BETWEEN_1900_AND_1970 = 25569;
static constexpr auto SECONDS_PER_DAY = 86400;

static void TimestampToExcelNumberFunction(DataChunk &args, ExpressionState &state, Vector &result) {
	const auto count = args.size();
	UnaryExecutor::Execute<timestamp_t, double>(args.data[0], result, count, [&](timestamp_t timestamp) {
		const auto epoch_us = Timestamp::GetEpochMicroSeconds(timestamp);
		// Turn ms to seconds
		const auto epoch_s = static_cast<double>(epoch_us) / 1000000.0;
		// Add the seconds to the days
		const auto res = epoch_s / SECONDS_PER_DAY + DAYS_BETWEEN_1900_AND_1970;
		return res;
	});
}

static void TimeToExcelNumberFunction(DataChunk &args, ExpressionState &state, Vector &result) {
	const auto count = args.size();
	UnaryExecutor::Execute<dtime_t, double>(args.data[0], result, count, [&](dtime_t time) {
		// 1.0 is a full day;
		return static_cast<double>(time.micros) / Interval::MICROS_PER_DAY;
	});
}

static void DateToExcelNumberFunction(DataChunk &args, ExpressionState &state, Vector &result) {
	const auto count = args.size();
	UnaryExecutor::Execute<date_t, double>(args.data[0], result, count, [&](date_t date) {
		// Date only has days, so there is no need to get more precision
		const auto epoch_s = Date::Epoch(date);
		// Add the seconds to the days
		const auto res = epoch_s / SECONDS_PER_DAY + DAYS_BETWEEN_1900_AND_1970;
		return res;
	});
}

static unique_ptr<Expression> TimestampConversionExpr(unique_ptr<Expression> ref) {
	vector<unique_ptr<Expression>> children;
	children.push_back(std::move(ref));

	ScalarFunction sfunc("timestamp_to_excel_number", {LogicalType::TIMESTAMP}, LogicalType::DOUBLE,
	                     TimestampToExcelNumberFunction);
	auto func = make_uniq<BoundFunctionExpression>(LogicalType::DOUBLE, sfunc, std::move(children), nullptr);
	return std::move(func);
}

static unique_ptr<Expression> TimeConversionExpr(unique_ptr<Expression> ref) {
	vector<unique_ptr<Expression>> children;
	children.push_back(std::move(ref));

	ScalarFunction sfunc("time_to_excel_number", {LogicalType::TIME}, LogicalType::DOUBLE, TimeToExcelNumberFunction);
	auto func = make_uniq<BoundFunctionExpression>(LogicalType::DOUBLE, sfunc, std::move(children), nullptr);
	return std::move(func);
}

static unique_ptr<Expression> DateConversionExpr(unique_ptr<Expression> ref) {
	vector<unique_ptr<Expression>> children;
	children.push_back(std::move(ref));

	ScalarFunction sfunc("date_to_excel_number", {LogicalType::DATE}, LogicalType::DOUBLE, DateToExcelNumberFunction);
	auto func = make_uniq<BoundFunctionExpression>(LogicalType::DOUBLE, sfunc, std::move(children), nullptr);
	return std::move(func);
}

//------------------------------------------------------------------------------
// Bind
//------------------------------------------------------------------------------
struct WriteXLSXData final : TableFunctionData {
	vector<LogicalType> column_types;
	vector<string> column_names;

	string file_path;
	string sheet_name;
	idx_t sheet_row_limit;
	bool header;
};

static void ParseCopyToOptions(const unique_ptr<WriteXLSXData> &data,
                               const case_insensitive_map_t<vector<Value>> &options) {
	// Find the header options
	const auto header_opt = options.find("header");
	if (header_opt != options.end()) {
		if (header_opt->second.size() != 1) {
			throw BinderException("Header option must be a single boolean value");
		}
		string error_msg;
		Value bool_val;
		if (!header_opt->second.back().DefaultTryCastAs(LogicalType::BOOLEAN, bool_val, &error_msg)) {
			throw BinderException("Header option must be a single boolean value");
		}
		if (bool_val.IsNull()) {
			throw BinderException("Header option must be a single boolean value");
		}
		data->header = BooleanValue::Get(bool_val);
	} else {
		data->header = false;
	}

	// Find the sheet name option
	const auto sheet_name_opt = options.find("sheet");
	if (sheet_name_opt != options.end()) {
		if (sheet_name_opt->second.size() != 1) {
			throw BinderException("Sheet name option must be a single string value");
		}
		if (sheet_name_opt->second.back().type() != LogicalType::VARCHAR) {
			throw BinderException("Sheet name option must be a single string value");
		}
		if (sheet_name_opt->second.back().IsNull()) {
			throw BinderException("Sheet name option must be a single string value");
		}
		data->sheet_name = StringValue::Get(sheet_name_opt->second.back());
	} else {
		data->sheet_name = "Sheet1";
	}

	// Find the row limit option
	const auto sheet_row_limit_opt = options.find("sheet_row_limit");
	if (sheet_row_limit_opt != options.end()) {
		if (sheet_row_limit_opt->second.size() != 1) {
			throw BinderException("Sheet row limit option must be a single integer value");
		}
		string error_msg;
		Value int_val;
		if (!sheet_row_limit_opt->second.back().DefaultTryCastAs(LogicalType::INTEGER, int_val, &error_msg)) {
			throw BinderException("Sheet row limit option must be a single integer value");
		}
		if (int_val.IsNull()) {
			throw BinderException("Sheet row limit option must be a single integer value");
		}
		data->sheet_row_limit = IntegerValue::Get(int_val);
	} else {
		data->sheet_row_limit = XLSX_MAX_CELL_ROWS;
	}
}

static unique_ptr<FunctionData> Bind(ClientContext &context, CopyFunctionBindInput &input, const vector<string> &names,
                                     const vector<LogicalType> &sql_types) {
	auto data = make_uniq<WriteXLSXData>();

	// Parse the options
	ParseCopyToOptions(data, input.info.options);

	data->column_types = sql_types;
	data->column_names = names;
	data->file_path = input.info.file_path;

	return std::move(data);
}

//------------------------------------------------------------------------------
// Init Global
//------------------------------------------------------------------------------
struct GlobalWriteXLSXData final : public GlobalFunctionData {

	XLXSWriter writer;
	DataChunk cast_chunk;
	ExpressionExecutor executor;
	vector<unique_ptr<Expression>> conversion_expressions;

	GlobalWriteXLSXData(ClientContext &context, const string &file_path, const WriteXLSXData &data)
	    : writer(context, file_path, data.sheet_row_limit), executor(context) {

		// Initialize the expression executor
		for (idx_t col_idx = 0; col_idx < data.column_types.size(); col_idx++) {
			auto &col_type = data.column_types[col_idx];
			auto expr = make_uniq_base<Expression, BoundReferenceExpression>(col_type, col_idx);

			switch (col_type.id()) {
			case LogicalTypeId::TIMESTAMP_TZ:
			case LogicalTypeId::TIMESTAMP_SEC:
			case LogicalTypeId::TIMESTAMP_MS:
			case LogicalTypeId::TIMESTAMP_NS:
				expr = BoundCastExpression::AddCastToType(context, std::move(expr), LogicalType::TIMESTAMP);
			case LogicalTypeId::TIMESTAMP: // Fall through
				expr = TimestampConversionExpr(std::move(expr));
				break;
			case LogicalTypeId::TIME_TZ:
				expr = BoundCastExpression::AddCastToType(context, std::move(expr), LogicalType::TIME);
			case LogicalTypeId::TIME: // Fall through
				expr = TimeConversionExpr(std::move(expr));
				break;
			case LogicalTypeId::DATE:
				expr = DateConversionExpr(std::move(expr));
				break;
			case LogicalTypeId::BOOLEAN:
				// Convert booleans to numbers first, then number to varchar
				expr = BoundCastExpression::AddCastToType(context, std::move(expr), LogicalType::INTEGER);
				break;
			default:
				break;
			}

			// Finally, cast the column to a VARCHAR
			expr = BoundCastExpression::AddCastToType(context, std::move(expr), LogicalType::VARCHAR);

			conversion_expressions.push_back(std::move(expr));
			executor.AddExpression(*conversion_expressions.back());
		}

		// Initialize the cast chunk;
		const vector<LogicalType> types(data.column_types.size(), LogicalType::VARCHAR);
		cast_chunk.Initialize(BufferAllocator::Get(context), types);
	}
};

static unique_ptr<GlobalFunctionData> InitGlobal(ClientContext &context, FunctionData &bind_data,
                                                 const string &file_path) {
	auto &data = bind_data.Cast<WriteXLSXData>();
	auto gstate = make_uniq<GlobalWriteXLSXData>(context, file_path, data);

	auto &writer = gstate->writer;

	// Begin writing the worksheet
	writer.BeginSheet(data.sheet_name, data.column_names, data.column_types);

	// Write the header
	if (data.header) {
		writer.BeginRow();
		for (const auto &col_name : data.column_names) {
			writer.WriteInlineStringCell(string_t(col_name));
		}
		writer.EndRow();
	}
	return std::move(gstate);
}

//------------------------------------------------------------------------------
// Init Local
//------------------------------------------------------------------------------
struct LocalWriteXLSXData final : public LocalFunctionData {};

static unique_ptr<LocalFunctionData> InitLocal(ExecutionContext &context, FunctionData &bind_data) {
	return make_uniq<LocalWriteXLSXData>();
}

//------------------------------------------------------------------------------
// Sink
//------------------------------------------------------------------------------
static void Sink(ExecutionContext &context, FunctionData &bind_data, GlobalFunctionData &gstate,
                 LocalFunctionData &lstate, DataChunk &input) {
	auto &data = bind_data.Cast<WriteXLSXData>();
	auto &state = gstate.Cast<GlobalWriteXLSXData>();
	auto &writer = state.writer;

	const auto row_count = input.size();
	const auto col_count = input.data.size();

	// First, cast the input columns to the target columns
	state.executor.Execute(input, state.cast_chunk);

	// Then, setup unified formats for the cast columns
	vector<UnifiedVectorFormat> formats;
	formats.resize(col_count);

	for (idx_t col_idx = 0; col_idx < col_count; col_idx++) {
		state.cast_chunk.data[col_idx].ToUnifiedFormat(row_count, formats[col_idx]);
	}

	// Now write the rows as xml
	for (idx_t in_idx = 0; in_idx < row_count; in_idx++) {
		writer.BeginRow();

		for (idx_t col_idx = 0; col_idx < col_count; col_idx++) {
			const auto &format = formats[col_idx];
			const auto row_idx = format.sel->get_index(in_idx);
			if (!format.validity.RowIsValid(row_idx)) {
				writer.WriteEmptyCell();
				continue;
			}

			const auto &val = UnifiedVectorFormat::GetData<string_t>(format)[row_idx];
			const auto &type = data.column_types[col_idx];

			switch (type.id()) {
			case LogicalTypeId::BOOLEAN:
				writer.WriteBooleanCell(val);
				break;
			case LogicalTypeId::DATE:
				writer.WriteDateCell(val);
				break;
			case LogicalTypeId::TIME_TZ:
			case LogicalTypeId::TIME:
				writer.WriteTimeCell(val);
				break;
			case LogicalTypeId::TIMESTAMP_TZ:
			case LogicalTypeId::TIMESTAMP_MS:
			case LogicalTypeId::TIMESTAMP_NS:
			case LogicalTypeId::TIMESTAMP:
				writer.WriteTimestampCell(val);
				break;
			case LogicalTypeId::TIMESTAMP_SEC:
				// Style this differently (no milliseconds)
				writer.WriteTimestampCellNoMilliseconds(val);
				break;
			case LogicalTypeId::TINYINT:
			case LogicalTypeId::SMALLINT:
			case LogicalTypeId::INTEGER:
			case LogicalTypeId::BIGINT:
			case LogicalTypeId::HUGEINT:
			case LogicalTypeId::FLOAT:
			case LogicalTypeId::DOUBLE:
			case LogicalTypeId::DECIMAL:
			case LogicalTypeId::UTINYINT:
			case LogicalTypeId::USMALLINT:
			case LogicalTypeId::UINTEGER:
			case LogicalTypeId::UBIGINT:
			case LogicalTypeId::UHUGEINT:
				writer.WriteNumberCell(val);
				break;
			default:
				writer.WriteInlineStringCell(val);
				break;
			}
		}
		writer.EndRow();
	}
}

//------------------------------------------------------------------------------
// Combine
//------------------------------------------------------------------------------
static void Combine(ExecutionContext &context, FunctionData &bind_data, GlobalFunctionData &gstate,
                    LocalFunctionData &lstate) {
}

//------------------------------------------------------------------------------
// Finalize
//------------------------------------------------------------------------------
static void Finalize(ClientContext &context, FunctionData &bind_data, GlobalFunctionData &gstate) {
	auto &state = gstate.Cast<GlobalWriteXLSXData>();

	// Finish writing the worksheet
	state.writer.EndSheet();
	state.writer.Finish();
}

//------------------------------------------------------------------------------
// Execution Mode
//------------------------------------------------------------------------------
CopyFunctionExecutionMode ExecutionMode(bool preserve_insertion_order, bool supports_batch_index) {
	// TODO: Support parallel ordered batch writes
	return CopyFunctionExecutionMode::REGULAR_COPY_TO_FILE;
}

//------------------------------------------------------------------------------
// Copy From
//------------------------------------------------------------------------------
static void SetBooleanValue(named_parameter_map_t &params, const string &key, const vector<Value> &val) {
	static constexpr auto error_msg = "'%s' option must be standalone or a BOOLEAN value";
	if (val.size() > 1) {
		throw BinderException(error_msg, key);
	}
	if (val.size() == 1) {
		auto &value = val.back();
		if (value.IsNull()) {
			throw BinderException(error_msg, key);
		}
		if (value.type() == LogicalType::FLOAT || value.type() == LogicalType::DOUBLE ||
		    value.type().id() == LogicalTypeId::DECIMAL) {
			throw BinderException("'%s' options expects a boolean value (e.g. TRUE or 1)", key);
		}
		params[key] = value.DefaultCastAs(LogicalType::BOOLEAN);
	} else {
		params[key] = Value::BOOLEAN(true);
	}
}

static void SetVarcharValue(named_parameter_map_t &params, const string &key, const vector<Value> &val) {
	static constexpr auto error_msg = "'%s' option must be a single VARCHAR value";
	if (val.size() != 1) {
		throw BinderException(error_msg, key);
	}
	if (val.back().type() != LogicalType::VARCHAR) {
		throw BinderException(error_msg, key);
	}
	if (val.back().IsNull()) {
		throw BinderException(error_msg, key);
	}
	params[key] = val.back();
}

static void ParseCopyFromOptions(XLSXReadData &data, const case_insensitive_map_t<vector<Value>> &options) {

	// Just make it really easy for us, extract everything into a named parameter map
	named_parameter_map_t named_parameters;

	for (auto &kv : options) {
		auto key = StringUtil::Lower(kv.first);
		auto &val = kv.second;
		if (key == "sheet") {
			SetVarcharValue(named_parameters, key, val);
		} else if (key == "header") {
			SetBooleanValue(named_parameters, key, val);
		} else if (key == "all_varchar") {
			SetBooleanValue(named_parameters, key, val);
		} else if (key == "ignore_errors") {
			SetBooleanValue(named_parameters, key, val);
		} else if (key == "range") {
			SetVarcharValue(named_parameters, key, val);
		} else if (key == "stop_at_empty") {
			SetBooleanValue(named_parameters, key, val);
		} else if (key == "empty_as_varchar") {
			SetBooleanValue(named_parameters, key, val);
		}
	}

	// Now just pass this to the table function data
	ReadXLSX::ParseOptions(data.options, named_parameters);
}

static unique_ptr<FunctionData> CopyFromBind(ClientContext &context, CopyFromFunctionBindInput &input, vector<string> &expected_names,
                                             vector<LogicalType> &expected_types) {

	auto result = make_uniq<XLSXReadData>();
	result->file_path = input.info.file_path; 

	// TODO: Parse options
	ParseCopyFromOptions(*result, input.info.options);

	ZipFileReader archive(context, input.info.file_path);
	ReadXLSX::ResolveSheet(result, archive);

	// Column count mismatch!
	if (expected_types.size() != result->return_types.size()) {
		string extended_error = "Table schema: ";
		for (idx_t col_idx = 0; col_idx < expected_types.size(); col_idx++) {
			if (col_idx > 0) {
				extended_error += ", ";
			}
			extended_error += expected_names[col_idx] + " " + expected_types[col_idx].ToString();
		}
		extended_error += "\nXLSX schema: ";
		for (idx_t col_idx = 0; col_idx < result->return_types.size(); col_idx++) {
			if (col_idx > 0) {
				extended_error += ", ";
			}
			extended_error += result->column_names[col_idx] + " " + result->return_types[col_idx].ToString();
		}
		extended_error += "\n\nPossible solutions:";
		extended_error += "\n* Manually specify which columns to insert using \"INSERT INTO tbl SELECT ... "
		                  "FROM read_xlsx(...)\"";
		extended_error += "\n* Provide an explicit range option with the same width as the table schema using e.g.";
		extended_error += "\"COPY tbl FROM ... (FORMAT 'xlsx', range 'A1:Z10')\"";

		throw ConversionException(
		    "Failed to read file(s) \"%s\" - column count mismatch: expected %d columns but found %d\n%s",
		    result->file_path, expected_types.size(), result->return_types.size(), extended_error);
	}

	// Override the column names and types with the expected ones
	result->return_types = expected_types;
	result->column_names = expected_names;

	return std::move(result);
}

//------------------------------------------------------------------------------
// Register
//------------------------------------------------------------------------------
void WriteXLSX::Register(ExtensionLoader &loader) {
	CopyFunction info("xlsx");

	info.copy_to_bind = Bind;
	info.copy_to_initialize_global = InitGlobal;
	info.copy_to_initialize_local = InitLocal;
	info.copy_to_sink = Sink;
	info.copy_to_combine = Combine;
	info.copy_to_finalize = Finalize;
	info.execution_mode = ExecutionMode;

	info.copy_from_bind = CopyFromBind;
	info.copy_from_function = ReadXLSX::GetFunction();

	info.extension = "xlsx";
	loader.RegisterFunction(info);
}

} // namespace duckdb