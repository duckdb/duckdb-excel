#include "excel_extension.hpp"

#include "duckdb.hpp"
#include "duckdb/common/exception.hpp"
#include "duckdb/common/string_util.hpp"
#include "duckdb/function/scalar_function.hpp"
#include "duckdb/main/extension/extension_loader.hpp"
#include "nf_calendar.h"
#include "nf_localedata.h"
#include "nf_zformat.h"
#include "xlsx/read_xlsx.hpp"

#include <duckdb/common/types/time.hpp>

namespace duckdb {

static std::string GetNumberFormatString(std::string &format, double num_value) {
	duckdb_excel::LocaleData locale_data;
	duckdb_excel::ImpSvNumberInputScan input_scan(&locale_data);
	uint16_t nCheckPos;
	std::string out_str;

	duckdb_excel::SvNumberformat num_format(format, &locale_data, &input_scan, nCheckPos);

	if (!num_format.GetOutputString(num_value, out_str)) {
		return out_str;
	}

	return "";
}

static string_t NumberFormatScalarFunction(Vector &result, double num_value, string_t format) {
	try {
		string in_str = format.GetString();
		string out_str = GetNumberFormatString(in_str, num_value);

		if (out_str.length() > 0) {
			auto result_string = StringVector::EmptyString(result, out_str.size());
			auto result_data = result_string.GetDataWriteable();
			memcpy(result_data, out_str.c_str(), out_str.size());
			result_string.Finalize();
			return result_string;
		} else {
			auto result_string = StringVector::EmptyString(result, 0);
			result_string.Finalize();
			return result_string;
		}
	} catch (...) {
		throw InternalException("Unexpected result for number format");
	}

	return string_t();
}

static void NumberFormatFunction(DataChunk &args, ExpressionState &state, Vector &result) {
	auto &number_vector = args.data[0];
	auto &format_vector = args.data[1];
	BinaryExecutor::Execute<double, string_t, string_t>(
	    number_vector, format_vector, result, args.size(),
	    [&](double value, string_t format) { return NumberFormatScalarFunction(result, value, format); });
}

//--------------------------------------------------------------------------------------------------
// Time conversion functions
//--------------------------------------------------------------------------------------------------

//--------------------------------------------------------------------------------------------------
// Load
//--------------------------------------------------------------------------------------------------

static void LoadInternal(ExtensionLoader &loader) {

	ScalarFunction text_func("text", {LogicalType::DOUBLE, LogicalType::VARCHAR}, LogicalType::VARCHAR,
							 NumberFormatFunction);
	loader.RegisterFunction(text_func);

	ScalarFunction excel_text_func("excel_text", {LogicalType::DOUBLE, LogicalType::VARCHAR}, LogicalType::VARCHAR,
								   NumberFormatFunction);

	loader.RegisterFunction(excel_text_func);

	// Register the XLSX functions
	ReadXLSX::Register(loader);
	WriteXLSX::Register(loader);
}

void ExcelExtension::Load(ExtensionLoader &loader) {
	LoadInternal(loader);
}

std::string ExcelExtension::Name() {
	return "excel";
}

} // namespace duckdb

extern "C" {

DUCKDB_CPP_EXTENSION_ENTRY(excel, loader) {
	duckdb::LoadInternal(loader);
}

}
