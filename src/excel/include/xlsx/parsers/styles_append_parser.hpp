#pragma once

#include "xlsx/xml_parser.hpp"

namespace duckdb {

struct XLSXNumFmtEntry {
	idx_t num_fmt_id;
	string format_code;
};

struct XLSXCellXfEntry {
	idx_t num_fmt_id;
};

class StylesAppendParser final : public XMLParser {
public:
	vector<XLSXNumFmtEntry> num_fmts;
	vector<XLSXCellXfEntry> cell_xfs;

protected:
	void OnStartElement(const char *name, const char **atts) override;
	void OnEndElement(const char *name) override;

private:
	enum class State : uint8_t { START, STYLESHEET, NUMFMTS, NUMFMT, CELLXFS, XF };
	State state = State::START;
};

inline void StylesAppendParser::OnStartElement(const char *name, const char **atts) {
	switch (state) {
	case State::START:
		if (MatchTag("styleSheet", name)) {
			state = State::STYLESHEET;
		}
		break;
	case State::STYLESHEET:
		if (MatchTag("numFmts", name)) {
			state = State::NUMFMTS;
		} else if (MatchTag("cellXfs", name)) {
			state = State::CELLXFS;
		}
		break;
	case State::NUMFMTS: {
		if (!MatchTag("numFmt", name)) {
			break;
		}
		state = State::NUMFMT;
		const char *id_ptr = nullptr;
		const char *format_ptr = nullptr;
		for (idx_t i = 0; atts[i]; i += 2) {
			if (strcmp(atts[i], "numFmtId") == 0) {
				id_ptr = atts[i + 1];
			} else if (strcmp(atts[i], "formatCode") == 0) {
				format_ptr = atts[i + 1];
			}
		}
		if (!id_ptr) {
			throw InvalidInputException("Invalid numFmt entry in styles.xml");
		}
		XLSXNumFmtEntry entry;
		entry.num_fmt_id = static_cast<idx_t>(strtol(id_ptr, nullptr, 10));
		entry.format_code = format_ptr ? string(format_ptr) : string();
		num_fmts.push_back(std::move(entry));
	} break;
	case State::CELLXFS: {
		if (!MatchTag("xf", name)) {
			break;
		}
		state = State::XF;
		const char *id_ptr = nullptr;
		for (idx_t i = 0; atts[i]; i += 2) {
			if (strcmp(atts[i], "numFmtId") == 0) {
				id_ptr = atts[i + 1];
			}
		}
		XLSXCellXfEntry entry;
		entry.num_fmt_id = id_ptr ? static_cast<idx_t>(strtol(id_ptr, nullptr, 10)) : 0;
		cell_xfs.push_back(entry);
	} break;
	default:
		break;
	}
}

inline void StylesAppendParser::OnEndElement(const char *name) {
	switch (state) {
	case State::NUMFMT:
		if (MatchTag("numFmt", name)) {
			state = State::NUMFMTS;
		}
		break;
	case State::XF:
		if (MatchTag("xf", name)) {
			state = State::CELLXFS;
		}
		break;
	case State::NUMFMTS:
		if (MatchTag("numFmts", name)) {
			state = State::STYLESHEET;
		}
		break;
	case State::CELLXFS:
		if (MatchTag("cellXfs", name)) {
			state = State::STYLESHEET;
		}
		break;
	case State::STYLESHEET:
		if (MatchTag("styleSheet", name)) {
			Stop(false);
		}
		break;
	default:
		break;
	}
}

} // namespace duckdb
