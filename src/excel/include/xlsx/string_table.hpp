#pragma once

#include "duckdb/common/allocator.hpp"
#include "duckdb/common/string_map_set.hpp"
#include "duckdb/common/types/string_type.hpp"

namespace duckdb {

//-------------------------------------------------------------------
// String Table
//-------------------------------------------------------------------
// StringTable is a simple class that stores a unique set of strings
// in arena backed memory and allows fast access by index.

class StringTable {
public:
	explicit StringTable(Allocator &alloc) : arena(alloc) {
	}
	idx_t Add(const string_t &str);
	const string_t &Get(idx_t val) const;
	void Reserve(idx_t count);

private:
	ArenaAllocator arena;
	string_map_t<idx_t> table;
	vector<string_t> index; // todo: make this a reference to the key in the map?
};

inline idx_t StringTable::Add(const string_t &str) {

	// Check if the string is already in the map
	const auto found = table.find(str);
	if (found != table.end()) {
		// We still need to update the index vector
		index.push_back(found->first);
		return found->second;
	}

	// Create a new entry
	const auto val = index.size();

	// If the string is inlined, just store it in the map and vector
	if (str.IsInlined()) {
		table[str] = val;
		index.push_back(str);
		return val;
	}

	// Otherwise copy the string to the arena
	const auto len = str.GetSize();
	const auto src = str.GetData();
	const auto dst = arena.Allocate(len);
	memcpy(dst, src, len);

	const auto key = string_t(const_char_ptr_cast(dst), len);
	table[key] = val;
	index.push_back(key);

	return val;
}

inline const string_t &StringTable::Get(const idx_t val) const {
	return index[val];
}

inline void StringTable::Reserve(const idx_t count) {
	table.reserve(count);
	index.reserve(count);
}

} // namespace duckdb