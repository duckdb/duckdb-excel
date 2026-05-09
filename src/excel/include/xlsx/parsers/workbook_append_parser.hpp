#pragma once

#include "xlsx/xml_parser.hpp"

namespace duckdb {

struct XLSXSheetEntry {
	string name;
	string rid;
	string sheet_id;
};

class WorkBookAppendParser final : public XMLParser {
public:
	static vector<XLSXSheetEntry> GetSheets(ZipFileReader &stream) {
		WorkBookAppendParser parser;
		parser.ParseAll(stream);
		return std::move(parser.sheets);
	}

private:
	void OnStartElement(const char *name, const char **atts) override;
	void OnEndElement(const char *name) override;

	enum class State : uint8_t { START, WORKBOOK, SHEETS, SHEET };
	State state = State::START;
	vector<XLSXSheetEntry> sheets;
};

inline void WorkBookAppendParser::OnStartElement(const char *name, const char **atts) {
	switch (state) {
	case State::START:
		if (MatchTag("workbook", name)) {
			state = State::WORKBOOK;
		}
		break;
	case State::WORKBOOK:
		if (MatchTag("sheets", name)) {
			state = State::SHEETS;
		}
		break;
	case State::SHEETS:
		if (MatchTag("sheet", name)) {
			state = State::SHEET;
			XLSXSheetEntry entry;
			for (idx_t i = 0; atts[i]; i += 2) {
				if (strcmp(atts[i], "name") == 0) {
					entry.name = atts[i + 1];
				} else if (strcmp(atts[i], "r:id") == 0) {
					entry.rid = atts[i + 1];
				} else if (strcmp(atts[i], "sheetId") == 0) {
					entry.sheet_id = atts[i + 1];
				}
			}
			if (entry.name.empty() || entry.rid.empty()) {
				throw InvalidInputException("Invalid sheet entry in workbook.xml");
			}
			sheets.push_back(std::move(entry));
		}
		break;
	default:
		break;
	}
}

inline void WorkBookAppendParser::OnEndElement(const char *name) {
	switch (state) {
	case State::SHEET:
		if (MatchTag("sheet", name)) {
			state = State::SHEETS;
		}
		break;
	case State::SHEETS:
		if (MatchTag("sheets", name)) {
			state = State::WORKBOOK;
		}
		break;
	case State::WORKBOOK:
		if (MatchTag("workbook", name)) {
			Stop(false);
		}
		break;
	default:
		break;
	}
}

} // namespace duckdb
