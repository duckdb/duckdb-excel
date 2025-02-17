#include "xlsx/xml_parser.hpp"

namespace duckdb {

class XLSXStyleParser final : public XMLParser {
public:
	unordered_map<idx_t, LogicalType> number_formats;
	vector<LogicalType> cell_styles;

protected:
	void OnStartElement(const char *name, const char **atts) override;
	void OnEndElement(const char *name) override;

private:
	template <class... ARGS>
	static bool StringContainsAny(const char *str, ARGS &&...args) {
		for (auto &&substr : {args...}) {
			if (strstr(str, substr) != nullptr) {
				return true;
			}
		}
		return false;
	}
	enum class State : uint8_t { START, STYLESHEET, NUMFMTS, NUMFMT, CELLXFS, XF };
	State state = State::START;
};

inline void XLSXStyleParser::OnStartElement(const char *name, const char **atts) {
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
		const auto id = strtol(id_ptr, nullptr, 10);
		if (id <= 163 || format_ptr == nullptr) {
			break;
		}

		const auto has_date_part = StringContainsAny(format_ptr, "DD", "dd", "YY", "yy");
		const auto has_time_part = StringContainsAny(format_ptr, "HH", "hh", "h", "H");

		if (has_date_part && has_time_part) {
			number_formats.emplace(id, LogicalType::TIMESTAMP);
		} else if (has_date_part) {
			number_formats.emplace(id, LogicalType::DATE);
		} else if (has_time_part) {
			number_formats.emplace(id, LogicalType::TIME);
		} else {
			// If we dont know how to handle the format, default to the numeric value.
			number_formats.emplace(id, LogicalType::DOUBLE); // TODO: Or double?
		}
	} break;
	case State::CELLXFS: {
		state = State::XF;
		const char *id_ptr = nullptr;

		for (idx_t i = 0; atts[i]; i += 2) {
			if (strcmp(atts[i], "numFmtId") == 0) {
				id_ptr = atts[i + 1];
			}
		}
		if (!id_ptr) {
			throw InvalidInputException("Invalid xf entry in styles.xml");
		}
		const auto id = strtol(id_ptr, nullptr, 10);
		if (id < 164) {
			// Special cases
			if (id >= 14 && id <= 17) {
				cell_styles.push_back(LogicalType::DATE);
			} else if (id >= 18 && id <= 21) {
				cell_styles.push_back(LogicalType::TIME);
			} else if (id == 22) {
				cell_styles.push_back(LogicalType::TIMESTAMP);
			} else {
				// Else, just push a double
				cell_styles.push_back(LogicalType::DOUBLE);
			}
		} else {
			// Look up the ID in the format map
			const auto it = number_formats.find(id);
			if (it != number_formats.end()) {
				cell_styles.push_back(it->second);
			}
		}
	} break;
	default:
		break;
	}
}

inline void XLSXStyleParser::OnEndElement(const char *name) {
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