#pragma once

#include "xlsx/xml_parser.hpp"

namespace duckdb {

//-------------------------------------------------------------------
// Base Worksheet Parser
//-------------------------------------------------------------------
// Traverses the worksheet, extracts the data from cells and calls
// the appropriate callbacks.
//-------------------------------------------------------------------
class SheetParserBase : public XMLParser {
public:
	void OnText(const char *text, idx_t len) override;
	void OnStartElement(const char *name, const char **atts) override;
	void OnEndElement(const char *name) override;

protected:
	virtual void OnBeginRow(idx_t row_idx) {};
	virtual void OnEndRow(idx_t row_idx) {};
	virtual void OnCell(const XLSXCellPos &pos, XLSXCellType type, vector<char> &data, idx_t style) {
	}

private:
	enum class State : uint8_t { START, SHEETDATA, ROW, EMPTY_ROW, CELL, V, IS, T };
	State state = State::START;

	XLSXCellPos cell_pos = {0, 0};
	XLSXCellType cell_type = XLSXCellType::NUMBER;
	vector<char> cell_data = {};
	idx_t cell_style = 0;
};

inline void SheetParserBase::OnText(const char *text, idx_t len) {
	if (cell_data.size() + len > XLSX_MAX_CELL_SIZE * 2) {
		// Something is obviously wrong, error out!
		throw InvalidInputException("XLSX: Cell data too large (is the file corrupted?)");
	}
	cell_data.insert(cell_data.end(), text, text + len);
}

inline void SheetParserBase::OnStartElement(const char *name, const char **atts) {
	if (state == State::START && MatchTag("sheetData", name)) {
		state = State::SHEETDATA;
	} else if (state == State::SHEETDATA && MatchTag("row", name)) {
		state = State::ROW;

		// Reset the column position
		cell_pos.col = 0;

		const char *rref_ptr = nullptr;
		for (idx_t i = 0; atts[i]; i += 2) {
			if (strcmp(atts[i], "r") == 0) {
				rref_ptr = atts[i + 1];
			}
		}
		// Default: Increment the row
		if (!rref_ptr) {
			cell_pos.row++;
		} else {
			// Else, jump to the row reference
			// TODO: Verify this
			cell_pos.row = strtol(rref_ptr, nullptr, 10);
		}

		OnBeginRow(cell_pos.row);
	} else if (state == State::ROW && MatchTag("c", name)) {
		state = State::CELL;

		// Reset the cell data
		cell_data.clear();

		// We're entering a cell. Parse the attributes
		const char *type_ptr = nullptr;
		const char *cref_ptr = nullptr;
		const char *style_ptr = nullptr;
		for (idx_t i = 0; atts[i]; i += 2) {
			if (strcmp(atts[i], "t") == 0) {
				type_ptr = atts[i + 1];
			} else if (strcmp(atts[i], "r") == 0) {
				cref_ptr = atts[i + 1];
			} else if (strcmp(atts[i], "s") == 0) {
				style_ptr = atts[i + 1];
			}
		}

		// Default: 0
		cell_style = style_ptr ? strtol(style_ptr, nullptr, 10) : 0;
		// Default: NUMBER
		cell_type = type_ptr ? ParseCellType(type_ptr) : XLSXCellType::NUMBER;
		// Default: next cell
		if (!cref_ptr) {
			cell_pos.col++;
		} else {
			XLSXCellPos cref;
			if (!cref.TryParse(cref_ptr)) {
				throw InvalidInputException("Invalid cell reference in sheet: %s", cref_ptr);
			}
			if (cref.row != cell_pos.row) {
				throw InvalidInputException("Cell reference does not match row reference in sheet");
			}
			cell_pos.col = cref.col;
		}
	} else if (state == State::CELL && MatchTag("v", name)) {
		state = State::V;
		EnableTextHandler(true);
	} else if (state == State::CELL && MatchTag("is", name)) {
		state = State::IS;
	} else if (state == State::IS && MatchTag("t", name)) {
		state = State::T;
		EnableTextHandler(true);
	}
}

inline void SheetParserBase::OnEndElement(const char *name) {
	if (state == State::SHEETDATA && MatchTag("sheetData", name)) {
		Stop(false);
	} else if (state == State::ROW && MatchTag("row", name)) {
		OnEndRow(cell_pos.row);
		state = State::SHEETDATA;
	} else if (state == State::CELL && MatchTag("c", name)) {
		OnCell(cell_pos, cell_type, cell_data, cell_style);
		state = State::ROW;
	} else if (state == State::V && MatchTag("v", name)) {
		state = State::CELL;
		EnableTextHandler(false);
	} else if (state == State::IS && MatchTag("is", name)) {
		state = State::CELL;
	} else if (state == State::T && MatchTag("t", name)) {
		state = State::IS;
		EnableTextHandler(false);
	}
}

//-------------------------------------------------------------------
// Range Sniffer
//-------------------------------------------------------------------
// The range sniffer is used to determine the range of the sheet to scan
// It will scan the sheet until it finds a row with data, and then use that
// to infer the column range of the sheet
//-------------------------------------------------------------------
class RangeSniffer final : public SheetParserBase {
public:
	XLSXCellRange GetRange() const;

private:
	void OnEndRow(idx_t row_idx) override;
	void OnCell(const XLSXCellPos &pos, XLSXCellType type, vector<char> &data, idx_t style) override;

private:
	// The range of the first consecutive non-empty row
	idx_t beg_col = 0;
	idx_t end_col = 0;
	idx_t beg_row = 0;

	// The total range of the sheet
	idx_t min_col = 0;
	idx_t max_col = 0;

	enum class State : uint8_t { EMPTY, FOUND, ENDED };
	State state = State::EMPTY;
};

inline XLSXCellRange RangeSniffer::GetRange() const {
	if (beg_row == 0) {
		// We didnt find any rows... return the whole sheet
		return XLSXCellRange();
	}
	// Otherwise, return the sniffed range
	return XLSXCellRange(beg_row, beg_col, NumericLimits<idx_t>::Maximum(), end_col + 1);
}

inline void RangeSniffer::OnCell(const XLSXCellPos &pos, XLSXCellType type, vector<char> &data, idx_t style) {
	min_col = std::min(min_col, pos.col);
	max_col = std::max(max_col, pos.col);

	switch (state) {
	case State::EMPTY:
		if (!data.empty()) {
			state = State::FOUND;
			beg_col = pos.col;
			end_col = pos.col;
		}
		break;
	case State::FOUND:
		if (data.empty()) {
			state = State::ENDED;
		} else {
			end_col = pos.col;
		}
		break;
	case State::ENDED:
		// We're done
		break;
	}
}

inline void RangeSniffer::OnEndRow(const idx_t row_idx) {
	if (state == State::FOUND || state == State::ENDED) {
		// We found a row with data, between beg_col and end_col
		// We can now use this as the range for the sheet
		beg_row = row_idx;
		Stop(false);
	} else {
		// Reset, continue on to the next row
		state = State::EMPTY;
		beg_col = 0;
		end_col = 0;
	}
}

//-------------------------------------------------------------------
// Header Sniffer
//-------------------------------------------------------------------
// The header sniffer is used to determine the header and the types
// of the columns in the sheet (within the range)
//-------------------------------------------------------------------
class HeaderSniffer final : public SheetParserBase {
public:
	HeaderSniffer(const XLSXCellRange &range_p, const XLSXHeaderMode header_mode_p, const bool absolute_range_p,
	              XLSXCellType default_cell_type_p)
	    : range(range_p), header_mode(header_mode_p), absolute_range(absolute_range_p),
	      default_cell_type(default_cell_type_p) {
	}

	const XLSXCellRange &GetRange() const {
		return range;
	}
	vector<XLSXCell> &GetColumnCells() {
		return column_cells;
	}
	vector<XLSXCell> &GetHeaderCells() {
		return header_cells;
	}

private:
	void OnBeginRow(idx_t row_idx) override;
	void OnEndRow(idx_t row_idx) override;
	void OnCell(const XLSXCellPos &pos, XLSXCellType type, vector<char> &data, idx_t style) override;

private:
	vector<XLSXCell> header_cells;
	vector<XLSXCell> column_cells;

	XLSXCellRange range;
	XLSXHeaderMode header_mode;

	idx_t last_col = 0;

	bool first_row = true;
	bool absolute_range;
	XLSXCellType default_cell_type;
};

inline void HeaderSniffer::OnBeginRow(const idx_t row_idx) {
	if (!range.ContainsRow(row_idx)) {
		return;
	}
	column_cells.clear();
	last_col = range.beg.col - 1;
}

inline void HeaderSniffer::OnCell(const XLSXCellPos &pos, XLSXCellType type, vector<char> &data, idx_t style) {
	if (!range.ContainsCol(pos.col)) {
		return;
	}

	// Now, add the cell to the data cells, but make sure to pad with empty varchars if needed.
	if (last_col + 1 < pos.col) {
		// Pad with empty cells
		for (idx_t i = last_col + 1; i < pos.col; i++) {
			column_cells.emplace_back(default_cell_type, XLSXCellPos(pos.row, i), "", 0);
		}
	}

	// Add the cell
	column_cells.emplace_back(type, pos, string(data.data(), data.size()), style);
	last_col = pos.col;
}

inline void HeaderSniffer::OnEndRow(const idx_t row_idx) {
	if (!range.ContainsRow(row_idx)) {
		column_cells.clear();
		last_col = range.beg.col - 1;
		return;
	}

	// If there are columns missing at the end, pad with empty string cells
	if (last_col + 1 < range.end.col) {
		for (idx_t i = last_col + 1; i < range.end.col; i++) {
			column_cells.emplace_back(default_cell_type, XLSXCellPos(row_idx, i), "", 0);
		}
	}

	// Now we have all the cells in the row, we can inspect them
	if (!first_row) {
		// This is the data row. We can stop here
		Stop(false);
		return;
	}

	// Now its time to determine the header row
	auto has_header = false;
	switch (header_mode) {
	case XLSXHeaderMode::NEVER:
		// We're not looking for a header, so we're done
		has_header = false;
		break;
	case XLSXHeaderMode::FORCE:
		has_header = true;
		break;
	case XLSXHeaderMode::MAYBE: {
		// We're looking for a header, but we're not sure if we found it yet
		// We need to inspect the cells to determine if this is a header row or not
		// We need all the rows to be non empty and strings
		auto all_strings = true;
		for (const auto &cell : column_cells) {
			auto &type = cell.type;
			const auto is_str = type == XLSXCellType::SHARED_STRING || type == XLSXCellType::INLINE_STRING;
			const auto is_empty = cell.data.empty();
			if (!is_str || is_empty) {
				all_strings = false;
				break;
			}
		}
		has_header = all_strings;
	}
	}

	if (!has_header) {
		// Generate a dummy header
		header_cells = column_cells;
		for (auto &cell : header_cells) {
			cell.type = XLSXCellType::INLINE_STRING;
			cell.style = 0;
			if (!absolute_range) {
				cell.data = cell.cell.ToString();
			} else {
				cell.data = cell.cell.GetColumnName();
			}
		}
		Stop(false);
		return;
	}

	// Save the header cells
	header_cells = column_cells;
	column_cells.clear();
	last_col = range.beg.col - 1;

	// Try to parse another row to see if we can find the data row
	first_row = false;

	// Move the range down one row
	range.beg.row = row_idx + 1;
}

//-------------------------------------------------------------------
// Sheet Parser
//-------------------------------------------------------------------
// The sheet parser is used to parse the actual data from the sheet
//-------------------------------------------------------------------
class SheetParser final : public SheetParserBase {
public:
	explicit SheetParser(ClientContext &context, const XLSXCellRange &range_p, const StringTable &table,
	                     bool stop_at_empty_p)
	    : string_table(table), range(range_p), stop_at_empty(stop_at_empty_p) {

		// Initialize the chunk
		const vector<LogicalType> types(range.Width(), LogicalType::VARCHAR);
		auto &buffer_alloc = BufferAllocator::Get(context);
		chunk.Initialize(buffer_alloc, types);

		// Set the beginning column
		// Allocate the sheet row number mapping
		sheet_row_number = make_unsafe_uniq_array<idx_t>(STANDARD_VECTOR_SIZE);

		last_row = range.beg.row - 1;
		curr_row = range.beg.row;
		last_col = range.beg.col - 1;
	}

	DataChunk &GetChunk() {
		return chunk;
	}
	string GetCellName(idx_t chunk_row, idx_t chunk_col) const;

	// Returns true if the chunk is full
	bool FoundSkippedRow() const;
	void SkipRows();
	// Fill empty rows to the end of the range
	void FillRows();

protected:
	void OnBeginRow(idx_t row_idx) override;
	void OnEndRow(idx_t row_idx) override;
	void OnCell(const XLSXCellPos &pos, XLSXCellType type, vector<char> &data, idx_t style) override;

private:
	// Shared String Table
	const StringTable &string_table;
	// Range to read
	XLSXCellRange range;
	// Mapping from chunk row to sheet row
	unsafe_unique_array<idx_t> sheet_row_number;
	// Current chunk
	DataChunk chunk;
	// Current row in the chunk
	idx_t out_index = 0;

	// The last column we wrote to
	idx_t last_col;
	// The last row we wrote to
	idx_t last_row;
	idx_t curr_row;

	bool stop_at_empty = false;
	bool is_row_empty = false;
};

inline string SheetParser::GetCellName(idx_t chunk_row, idx_t chunk_col) const {
	// Get the cell name and row given a chunk row and column
	const auto sheet_row = sheet_row_number[chunk_row];
	const auto sheet_col = chunk_col + range.beg.col;

	const XLSXCellPos pos = {static_cast<idx_t>(sheet_row), sheet_col};
	return pos.ToString();
}

inline bool SheetParser::FoundSkippedRow() const {
	return last_row + 1 < curr_row;
}

inline void SheetParser::SkipRows() {
	// Pad empty rows
	while (last_row + 1 < curr_row) {
		last_row++;

		for (auto &col : chunk.data) {
			FlatVector::SetNull(col, out_index, true);
		}
		sheet_row_number[out_index] = last_row;
		out_index++;
		chunk.SetCardinality(out_index);

		if (out_index == STANDARD_VECTOR_SIZE) {
			// We have filled up the chunk, yield!
			out_index = 0;
			return;
		}
	}
}

inline void SheetParser::FillRows() {
	if(chunk.size() == STANDARD_VECTOR_SIZE) {
		// Chunk is full, no more to fill
		return;
	}

	const auto total_remaining = range.end.row - 1 - last_row;
	const auto local_remaining = STANDARD_VECTOR_SIZE - out_index;

	const auto remaining = MinValue(total_remaining, local_remaining);
	for (idx_t i = 0; i < remaining; i++) {
		for (auto &col : chunk.data) {
			FlatVector::SetNull(col, out_index, true);
		}
		sheet_row_number[out_index] = last_row;
		out_index++;
		last_row++;
	}
	chunk.SetCardinality(chunk.size() + remaining);
	out_index = 0;
}

inline void SheetParser::OnBeginRow(idx_t row_idx) {
	if (!range.ContainsRow(row_idx)) {
		// not in range, skip
		return;
	}

	last_col = range.beg.col - 1;
	is_row_empty = true;

	curr_row = row_idx;

	// Check if we need to pad empty rows
	if (last_row + 1 < curr_row) {
		Stop(true);
	}
}

inline void SheetParser::OnCell(const XLSXCellPos &pos, XLSXCellType type, vector<char> &data, idx_t style) {
	if (!range.ContainsPos(pos)) {
		// not in range, skip
		return;
	}

	// If we jumped over some columns, pad with nulls
	if (last_col + 1 < pos.col) {
		for (idx_t i = last_col + 1; i < pos.col; i++) {
			auto &vec = chunk.data[i - range.beg.col];
			FlatVector::SetNull(vec, out_index, true);
		}
	}

	// Get the column data
	auto &vec = chunk.data[pos.col - range.beg.col];

	// Push the cell data to our chunk
	const auto ptr = FlatVector::GetData<string_t>(vec);

	if (type == XLSXCellType::SHARED_STRING) {
		// Push a null to the buffer so that the string is null-terminated
		data.push_back('\0');
		// Now we can use strtol to get the shared string index
		const auto ssi = std::strtol(data.data(), nullptr, 10);
		// Look up the string in the string table
		ptr[out_index] = string_table.Get(ssi);
	} else if (data.empty() && type != XLSXCellType::INLINE_STRING) {
		// If the cell is empty (and not a string), we wont be able to convert it
		// so just null it immediately
		FlatVector::SetNull(vec, out_index, true);
	} else {
		// Otherwise just pass along the call data, we will cast it later.
		ptr[out_index] = StringVector::AddString(vec, data.data(), data.size());
	}

	if (!data.empty()) {
		is_row_empty = false;
	}

	last_col = pos.col;
}

inline void SheetParser::OnEndRow(idx_t row_idx) {
	if (!range.ContainsRow(row_idx)) {
		// not in range, skip
		return;
	}

	last_row = row_idx;

	if (stop_at_empty && is_row_empty) {
		Stop(false);
		return;
	}

	// If we didnt write out all the columns, pad with nulls
	if (last_col + 1 < range.end.col) {
		for (idx_t i = last_col + 1; i < range.end.col; i++) {
			auto &vec = chunk.data[i - range.beg.col];
			FlatVector::SetNull(vec, out_index, true);
		}
	}

	// Map the chunk row to the sheet row
	sheet_row_number[out_index] = UnsafeNumericCast<int32_t>(row_idx);

	out_index++;
	chunk.SetCardinality(out_index);
	if (out_index == STANDARD_VECTOR_SIZE) {
		// We have filled up the chunk, yield!
		out_index = 0;
		Stop(true);
	}
}

} // namespace duckdb