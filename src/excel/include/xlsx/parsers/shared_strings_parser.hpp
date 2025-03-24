#pragma once

#include "xlsx/xml_parser.hpp"
#include "xlsx/string_table.hpp"

namespace duckdb {

//-------------------------------------------------------------------
// Base Shared Strings Parser
//-------------------------------------------------------------------
// Base class for parsing the "sharedStrings.xml" entry
// in the XLSX file. This class is used by both the SharedStringSearcher
// and SharedStringParser classes.
//-------------------------------------------------------------------
class SharedStringParserBase : public XMLParser {
protected:
	virtual void OnUniqueCount(idx_t count) {
	}
	virtual void OnString(const vector<char> &str) = 0;

private:
	void OnStartElement(const char *name, const char **atts) override {
		switch (state) {
		case State::START:
			if (MatchTag("sst", name)) {
				state = State::SST;
				// Optionally look for the uniqueCount attributes
				// TODO: Do we also look for count?
				for (idx_t i = 0; atts[i]; i += 2) {
					if (strcmp(atts[i], "uniqueCount") == 0) {
						// TODO: Check that this succeeds!
						const auto unique_count = atoi(atts[i + 1]);
						OnUniqueCount(unique_count);
					}
				}
			}
			break;
		case State::SST:
			if (MatchTag("si", name)) {
				state = State::SI;
			}
			break;
		case State::SI:
			if (MatchTag("t", name)) {
				state = State::T;
				// Enable text handling
				EnableTextHandler(true);
			}
			break;
		case State::T:
			// TODO: Throw? We should never get here
			break;
		default:
			// TODO: Count garbage iterations where we dont find anything and throw if we get too far
			// to guard against corrupted files.
			// we're probably going to have to set that in the parser base class itself
			break;
		}
	}
	void OnEndElement(const char *name) override {
		switch (state) {
		case State::T:
			if (MatchTag("t", name)) {
				// Disable text handling
				EnableTextHandler(false);
				state = State::SI;
			}
			break;
		case State::SI:
			if (MatchTag("si", name)) {
				state = State::SST;
				// Pass the string we've collected from the <t> tags to the handler
				OnString(data);
				// Reset the buffer
				data.clear();
			}
			break;
		case State::SST:
			if (MatchTag("sst", name)) {
				Stop(false);
			}
			break;
		default:
			break;
		}
	}
	void OnText(const char *text, const idx_t len) override {
		data.insert(data.end(), text, text + len);
	}

private:
	enum class State : uint8_t { START, SST, SI, T };
	State state = State::START;
	vector<char> data;
};

//-------------------------------------------------------------------
// Shared Strings Searcher
//-------------------------------------------------------------------
// Parses the string table for a specific set of strings
// and returns the strings for those specific indices
//-------------------------------------------------------------------
class SharedStringSearcher final : public SharedStringParserBase {
public:
	explicit SharedStringSearcher(const vector<idx_t> &ids_p) : ids(ids_p) {

		// Sort and remove duplicates
		std::sort(ids.begin(), ids.end());
		ids.erase(std::unique(ids.begin(), ids.end()), ids.end());
	}

	const unordered_map<idx_t, string> &GetResult() const {
		return result;
	}

protected:
	void OnString(const vector<char> &str) override {

		if (current_idx >= ids.size()) {
			// We're done, no more strings to find
			Stop(false);
			return;
		}
		const auto &id = ids[current_idx];
		if (id == current_str) {
			result[id] = string(str.data(), str.size());
			current_idx++;
		}
		current_str++;
	}

private:
	idx_t current_idx = 0;
	idx_t current_str = 0;

	vector<idx_t> ids;
	unordered_map<idx_t, string> result;
};

//-------------------------------------------------------------------
// Shared Strings Parser
//-------------------------------------------------------------------
// Parses the string table and populate it completely
// with all the strings found in the file
//-------------------------------------------------------------------
class SharedStringParser final : public SharedStringParserBase {
public:
	static void ParseStringTable(ZipFileReader &stream, StringTable &table) {
		SharedStringParser parser(table);
		parser.ParseAll(stream);
	}

private:
	explicit SharedStringParser(StringTable &table_p) : table(table_p) {
	}

protected:
	void OnString(const vector<char> &str) override {
		table.Add(string_t(str.data(), str.size()));
	}
	void OnUniqueCount(const idx_t count) override {
		table.Reserve(count);
	}

private:
	StringTable &table;
};

} // namespace duckdb