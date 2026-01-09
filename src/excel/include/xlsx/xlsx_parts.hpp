#pragma once

#include "duckdb/common/string.hpp"
#include "duckdb/common/typedefs.hpp"
#include "duckdb/common/types.hpp"
#include "duckdb/common/exception/binder_exception.hpp"

namespace duckdb {

//-------------------------------------------------------------------
// Constants
//-------------------------------------------------------------------

constexpr auto XLSX_MAX_CELL_SIZE = 32767UL;
constexpr auto XLSX_MAX_CELL_ROWS = 1048576UL;
constexpr auto XLSX_MAX_CELL_COLS = 16384UL;

//-------------------------------------------------------------------------
// Cell position
//-------------------------------------------------------------------------

// A cell position in a worksheet
struct XLSXCellPos {
	idx_t row; // 1-indexed
	idx_t col; // 1-indexed

	XLSXCellPos(idx_t row_p, idx_t col_p) : row(row_p), col(col_p) {
	}
	XLSXCellPos() : row(1), col(1) {
	}

	// Try to parse a cell from a string, e.g. "A1"
	// returns the position after the cell, or nullptr if parsing failed
	const char *TryParse(const char *str);
	string ToString() const;
	string GetColumnName() const;
};

inline const char *XLSXCellPos::TryParse(const char *str) {
	// Parse the column

	auto did_not_parse_col = false;
	if (*str >= 'A' && *str <= 'Z') {
		idx_t p_col = *str - 'A' + 1;
		str++;

		while (*str >= 'A' && *str <= 'Z') {
			p_col = p_col * 26 + (*str - 'A' + 1);
			str++;
		}
		col = p_col;
	} else {
		did_not_parse_col = true;
	}

	// Parse the row
	auto did_not_parse_row = false;
	if (*str >= '1' && *str <= '9') {
		idx_t p_row = *str - '0';
		str++;

		while (*str >= '0' && *str <= '9') {
			p_row = p_row * 10 + (*str - '0');
			str++;
		}
		row = p_row;
	} else {
		did_not_parse_row = true;
	}

	if (did_not_parse_col && did_not_parse_row) {
		// Well, we had to parse something...
		return nullptr;
	}

	return str;
}

inline string XLSXCellPos::ToString() const {
	D_ASSERT(col != 0 && row != 0);
	string result = GetColumnName();
	result += std::to_string(row);
	return result;
}

inline string XLSXCellPos::GetColumnName() const {
	D_ASSERT(col != 0);
	string result;
	idx_t col = this->col;
	while (col > 0) {
		idx_t remainder = (col - 1) % 26;
		result = static_cast<char>('A' + remainder) + result;
		col = (col - 1) / 26;
	}
	return result;
}

//-------------------------------------------------------------------------
// Cell Range
//-------------------------------------------------------------------------

// A range of cells in a worksheet
struct XLSXCellRange {

	// 1-indexed, inclusive
	XLSXCellPos beg;
	// 1-indexed, exlusive
	XLSXCellPos end;

	XLSXCellRange(idx_t beg_row, idx_t beg_col, idx_t end_row, idx_t end_col)
	    : beg(beg_row, beg_col), end(end_row, end_col) {
	}

	XLSXCellRange() : beg(1, 1), end(NumericLimits<idx_t>::Maximum(), XLSX_MAX_CELL_COLS) {
	}

	// Try to parse a range from a string, e.g. "A1:B2"
	// returns the position after the range, or nullptr if parsing failed
	const char *TryParse(const char *str);

	bool ContainsRow(const idx_t row) const {
		return row >= beg.row && row < end.row;
	}
	bool ContainsCol(const idx_t col) const {
		return col >= beg.col && col < end.col;
	}
	bool ContainsPos(const XLSXCellPos &pos) const {
		return ContainsCol(pos.col) && ContainsRow(pos.row);
	}

	idx_t Width() const {
		return end.col - beg.col;
	}
	idx_t Height() const {
		return end.row - beg.row;
	}
	bool IsValid() const {
		return beg.row <= end.row && beg.col <= end.col;
	}
};

inline const char *XLSXCellRange::TryParse(const char *str) {
	// Parse the beginning cell

	XLSXCellPos beg_result = {1, 1};
	str = beg_result.TryParse(str);
	if (!str) {
		return nullptr;
	}

	// Parse the colon
	if (*str != ':') {
		return nullptr;
	}
	str++;

	// Parse the ending cell
	XLSXCellPos end_result = {XLSX_MAX_CELL_ROWS, XLSX_MAX_CELL_COLS};
	str = end_result.TryParse(str);
	if (!str) {
		return nullptr;
	}

	// Success
	beg = beg_result;
	end = end_result;

	return str;
}

enum class XLSXCellType : uint8_t {
	UNKNOWN,
	NUMBER,
	BOOLEAN,
	SHARED_STRING,
	INLINE_STRING,
	DATE,
	ERROR,
	FORMULA_STRING
};

inline XLSXCellType ParseCellType(const char *ctype) {
	// If no type is specified, assume it is a number
	if (!ctype) {
		return XLSXCellType::NUMBER;
	}
	switch (*ctype++) {
	case 'n':
		return XLSXCellType::NUMBER;
	case 's': {
		if (*ctype == 0) {
			return XLSXCellType::SHARED_STRING;
		}
		if (strcmp(ctype, "tr") == 0) {
			return XLSXCellType::FORMULA_STRING;
		}
		return XLSXCellType::UNKNOWN;
	}
	case 'd':
		return XLSXCellType::DATE;
	case 'b':
		return XLSXCellType::BOOLEAN;
	case 'e':
		return XLSXCellType::ERROR;
	case 'i': {
		if (strcmp(ctype, "nlineStr") == 0) {
			return XLSXCellType::INLINE_STRING;
		}
		return XLSXCellType::UNKNOWN;
	}
	default:
		return XLSXCellType::UNKNOWN;
	}
}

//-------------------------------------------------------------------------
// Style Sheet
//-------------------------------------------------------------------------
class XLSXStyleSheet {
public:
	XLSXStyleSheet() = default;
	explicit XLSXStyleSheet(vector<LogicalType> &&formats_p) : formats(std::move(formats_p)) {
	}
	optional_ptr<const LogicalType> GetFormat(const idx_t idx) const {
		if (idx < formats.size()) {
			return &formats[idx];
		}
		return nullptr;
	}

private:
	vector<LogicalType> formats;
};

//-------------------------------------------------------------------------
// Cell
//-------------------------------------------------------------------------

struct XLSXCell {
	XLSXCellType type;
	XLSXCellPos cell;
	string data;
	idx_t style;

	XLSXCell(XLSXCellType type_p, XLSXCellPos cell_p, string data_p, idx_t style_p)
	    : type(type_p), cell(cell_p), data(std::move(data_p)), style(style_p) {
	}

	LogicalType GetDuckDBType(bool all_varchar, const XLSXStyleSheet &style_sheet) {
		if (all_varchar) {
			return LogicalType::VARCHAR;
		}

		// Check if we get a TRUE or FALSE value, if so, treat it as a boolean
		if (StringUtil::CIEquals(data, "true")) {
			return LogicalType::BOOLEAN;
		}
		if (StringUtil::CIEquals(data, "false")) {
			return LogicalType::BOOLEAN;
		}

		switch (type) {
		case XLSXCellType::NUMBER: {
			// The logical type of a number is dependent on the style of the cell
			// Some styles are dates, some are doubles, some are integers
			// (some are even postcodes or phone numbers, but we don't care about those for now)
			auto optional_type = style_sheet.GetFormat(style);
			if (optional_type) {
				return *optional_type;
			}
			// Default to double
			return LogicalType::DOUBLE;
		}

		case XLSXCellType::BOOLEAN:
			return LogicalType::BOOLEAN;
		case XLSXCellType::SHARED_STRING:
		case XLSXCellType::INLINE_STRING:
		case XLSXCellType::FORMULA_STRING:
		case XLSXCellType::ERROR:
			return LogicalType::VARCHAR;
		case XLSXCellType::DATE:
			return LogicalType::DATE;
		default:
			throw BinderException("Unknown cell type in xlsx file");
		}
	}
};

} // namespace duckdb
