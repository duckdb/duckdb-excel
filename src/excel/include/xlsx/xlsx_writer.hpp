#pragma once

#include "xlsx/zip_file.hpp"
#include "xlsx/xml_util.hpp"

namespace duckdb {

class XLXSWriter {
public:
	void BeginSheet(const string &sheet_name, const vector<string> &sql_column_names,
	                const vector<LogicalType> &sql_column_types);
	void EndSheet();

	explicit XLXSWriter(ClientContext &context, const string &file_name, idx_t sheet_row_limit_p)
	    : stream(context, file_name), sheet_row_limit(sheet_row_limit_p) {
	}

	void WriteNumberCell(const string_t &value);
	void WriteInlineStringCell(const string_t &value);
	void WriteBooleanCell(const string_t &value);
	void WriteDateCell(const string_t &value);
	void WriteTimeCell(const string_t &value);
	void WriteTimestampCell(const string_t &value);
	void WriteTimestampCellNoMilliseconds(const string_t &value);
	void WriteEmptyCell();
	void BeginRow();
	void EndRow();

	void Finish();

private:
	idx_t WriteEscapedXML(const char *str);
	idx_t WriteEscapedXML(const string &str);
	idx_t WriteEscapedXML(const char *buffer, idx_t write_size);

	void WriteStyles();
	void WriteWorkbook();
	void WriteRels();
	void WriteContentTypes();
	void WriteDocProps();
	void WriteSharedStrings();
	void WritePackageRels();

	class XLSXSheet {
	public:
		string sheet_name;
		string sheet_file;
		vector<string> sheet_column_names; // A1... Z1, AA1... ZZ1, etc.
		vector<string> sheet_column_types; // e.g. "str", "n", etc.
		vector<string> sql_column_names;
		vector<LogicalType> sql_column_types;
	};

	ZipFileWriter stream;
	idx_t sheet_row_limit = XLSX_MAX_CELL_ROWS;

	// Current sheet data;
	string row_str = "1";
	idx_t row_idx = 0;
	idx_t col_idx = 0;
	bool has_active_sheet = false;

	XLSXSheet active_sheet;
	vector<XLSXSheet> written_sheets;

	vector<char> escaped_buffer;
};

static constexpr auto ENCODING_FRAGMENT = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
)";

inline void XLXSWriter::BeginSheet(const string &sheet_name, const vector<string> &sql_column_names,
                                   const vector<LogicalType> &sql_column_types) {

	if (written_sheets.empty()) {
		// We need to create the directory for sheets
		stream.AddDirectory("xl/");
		stream.AddDirectory("xl/worksheets/");
	}

	D_ASSERT(!has_active_sheet);
	has_active_sheet = true;
	active_sheet.sheet_name = EscapeXMLString(sheet_name);
	active_sheet.sheet_file = "sheet" + std::to_string(written_sheets.size() + 1) + ".xml";
	active_sheet.sql_column_names = sql_column_names;
	active_sheet.sql_column_types = sql_column_types;

	D_ASSERT(sql_column_names.size() == sql_column_types.size());
	const auto column_count = sql_column_names.size();

	// Generate sheet column names
	active_sheet.sheet_column_names.resize(column_count);
	for (idx_t col_idx = 0; col_idx < column_count; col_idx++) {
		// Convert number to excel column name, e.g. 0 -> A, 1 -> B, 25 -> Z, 26 -> AA, etc.
		string col_name;
		idx_t col_num = col_idx + 1;
		while (col_num > 0) {
			col_name = static_cast<char>('A' + (col_num - 1) % 26) + col_name;
			col_num = (col_num - 1) / 26;
		}
		active_sheet.sheet_column_names[col_idx] = col_name;
	}

	// Generate sheet column types
	active_sheet.sheet_column_types.resize(column_count);
	for (idx_t col_idx = 0; col_idx < column_count; col_idx++) {
		// Convert the logical type to the excel cell type identifier
		const auto &type = sql_column_types[col_idx];
		if (type.IsNumeric()) {
			active_sheet.sheet_column_types[col_idx] = "n";
		} else {
			active_sheet.sheet_column_types[col_idx] = "inlineStr";
		}
	}

	static constexpr auto WORKSHEET_XML_START = R"(
	<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	           xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main"
	           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
	           xmlns:mv="urn:schemas-microsoft-com:mac:vml"
	           xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
	           xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
	           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
	           xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
	<sheetData>
	)";

	stream.BeginFile("xl/worksheets/" + active_sheet.sheet_file);
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(WORKSHEET_XML_START);
}

inline void XLXSWriter::EndSheet() {
	D_ASSERT(has_active_sheet);
	has_active_sheet = false;

	static constexpr auto WORKSHEET_XML_END = R"(</sheetData></worksheet>)";
	stream.Write(WORKSHEET_XML_END);
	stream.EndFile();

	// Save the sheet
	written_sheets.push_back(std::move(active_sheet));

	row_str = "1";
	row_idx = 0;
	col_idx = 0;
}

inline void XLXSWriter::WriteNumberCell(const string_t &value) {
	stream.Write("<c r=\"" + active_sheet.sheet_column_names[col_idx] + row_str + "\"><v>");
	stream.Write(value.GetData(), value.GetSize());
	stream.Write("</v></c>");

	col_idx++;
}

inline void XLXSWriter::WriteBooleanCell(const string_t &value) {
	stream.Write("<c r=\"" + active_sheet.sheet_column_names[col_idx] + row_str + "\" t=\"b\" s=\"5\"><v>");
	stream.Write(value.GetData(), value.GetSize());
	stream.Write("</v></c>");

	col_idx++;
}

inline void XLXSWriter::WriteInlineStringCell(const string_t &value) {
	stream.Write("<c r=\"" + active_sheet.sheet_column_names[col_idx] + row_str + "\" t=\"inlineStr\"><is><t>");
	// We need to escape this string in case it contains XML special characters
	WriteEscapedXML(value.GetData(), value.GetSize());
	stream.Write("</t></is></c>");

	col_idx++;
}

inline void XLXSWriter::WriteDateCell(const string_t &value) {
	stream.Write("<c r=\"" + active_sheet.sheet_column_names[col_idx] + row_str + "\" s=\"1\"><v>");
	stream.Write(value.GetData(), value.GetSize());
	stream.Write("</v></c>");

	col_idx++;
}

inline void XLXSWriter::WriteTimeCell(const string_t &value) {
	stream.Write("<c r=\"" + active_sheet.sheet_column_names[col_idx] + row_str + "\" s=\"3\"><v>");
	stream.Write(value.GetData(), value.GetSize());
	stream.Write("</v></c>");

	col_idx++;
}

inline void XLXSWriter::WriteTimestampCell(const string_t &value) {
	stream.Write("<c r=\"" + active_sheet.sheet_column_names[col_idx] + row_str + "\" s=\"4\"><v>");
	stream.Write(value.GetData(), value.GetSize());
	stream.Write("</v></c>");

	col_idx++;
}

inline void XLXSWriter::WriteTimestampCellNoMilliseconds(const string_t &value) {
	stream.Write("<c r=\"" + active_sheet.sheet_column_names[col_idx] + row_str + "\" s=\"2\"><v>");
	stream.Write(value.GetData(), value.GetSize());
	stream.Write("</v></c>");

	col_idx++;
}

inline void XLXSWriter::WriteEmptyCell() {
	col_idx++;
}

inline void XLXSWriter::BeginRow() {
	stream.Write("<row r=\"" + row_str + "\">");
}

inline void XLXSWriter::EndRow() {
	stream.Write("</row>");
	col_idx = 0;

	row_idx++;
	row_str = std::to_string(row_idx + 1);

	if (row_idx > sheet_row_limit) {
		if (sheet_row_limit >= XLSX_MAX_CELL_ROWS) {
			const auto msg = "XLSX: Sheet row limit of '%d' rows exceeded!\n"
			                 " * XLSX files and compatible applications generally have a limit of '%d' rows\n"
			                 " * You can export larger sheets at your own risk by setting the 'sheet_row_limit' "
			                 "parameter to a higher value";
			throw InvalidInputException(msg, sheet_row_limit, XLSX_MAX_CELL_ROWS);
		} else {
			throw InvalidInputException("XLSX: Sheet row limit of '%d' rows exceeded!", sheet_row_limit);
		}
	}
}

inline void XLXSWriter::Finish() {

	WriteContentTypes();
	WritePackageRels();
	WriteWorkbook();
	WriteRels();
	WriteStyles();
	WriteSharedStrings();
	WriteDocProps();

	// Done!
	stream.Finalize();
}

inline idx_t XLXSWriter::WriteEscapedXML(const char *str) {
	return WriteEscapedXML(str, strlen(str));
}

inline idx_t XLXSWriter::WriteEscapedXML(const string &str) {
	return WriteEscapedXML(str.c_str(), str.size());
}

inline idx_t XLXSWriter::WriteEscapedXML(const char *buffer, idx_t write_size) {
	escaped_buffer.clear();
	EscapeXMLString(buffer, write_size, escaped_buffer);
	return stream.Write(escaped_buffer.data(), escaped_buffer.size());
}

inline void XLXSWriter::WriteStyles() {
	//--------------------------------------------------------------------------------------------------
	// The number formats we write to the styles.xml file
	//--------------------------------------------------------------------------------------------------
	// 0 | 164: GENERAL					(default)
	// 1 | 165: DD/MM/YY				(date)
	// 2 | 166: DD/MM/YYYY HH:MM:SS		(timestamp)
	// 3 | 167: HH:MM:SS				(time)
	// 4 | 168: DD/MM/YYYY HH:MM:SS.000	(timestamp with milliseconds)*
	// 5 | 169: TRUE/FALSE				(bool)
	//--------------------------------------------------------------------------------------------------
	// * Note: Excel can only display up to millisecond precision (even if we can store in microseconds)

	static constexpr auto STYLES_XML = R"(
	<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		<numFmts count="6">
		    <numFmt formatCode="General" numFmtId="164"/>
		    <numFmt formatCode="dd/mm/yy" numFmtId="165"/>
		    <numFmt formatCode="dd/mm/yyyy\ hh:mm:ss" numFmtId="166"/>
		    <numFmt formatCode="hh:mm:ss" numFmtId="167"/>
			<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss.000" numFmtId="168"/>
			<numFmt formatCode="&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;" numFmtId="169"/>
		</numFmts>
		<fonts count="1">
			<font>
				<name val="Arial"/>
				<family val="2"/>
				<sz val="12"/>
			</font>
		</fonts>
		<fills count="1">
			<fill>
				<patternFill patternType="none"/>
			</fill>
		</fills>
		<borders count="1">
			<border diagonalDown="false" diagonalUp="false">
				<left/>
				<right/>
				<top/>
				<bottom/>
				<diagonal/>
			</border>
		</borders>
		<cellStyleXfs count="1">
			<xf numFmtId="164"></xf>
		</cellStyleXfs>
		<cellXfs count="6">
			<xf numFmtId="164" xfId="0"/>
			<xf numFmtId="165" xfId="0"/>
			<xf numFmtId="166" xfId="0"/>
			<xf numFmtId="167" xfId="0"/>
			<xf numFmtId="168" xfId="0"/>
			<xf numFmtId="169" xfId="0"/>
		</cellXfs>
		<cellStyles count="1">
			<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>
		</cellStyles>
	</styleSheet>
	)";

	stream.BeginFile("xl/styles.xml");
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(STYLES_XML);
	stream.EndFile();
}

inline void XLXSWriter::WriteContentTypes() {
	static constexpr auto CONTENT_TYPES_XML_START =
	    R"(<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">)"
	    R"(<Default Extension="xml" ContentType="application/xml"/>)"
	    R"(<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>)"
	    R"(<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>)"
	    R"(<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>)"
	    R"(<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>)"
		R"(<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>)"
	    R"(<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>)";
	static constexpr auto CONTENT_TYPES_XML_END = R"(</Types>)";

	stream.BeginFile("[Content_Types].xml");
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(CONTENT_TYPES_XML_START);
	for (const auto &sheet : written_sheets) {
		stream.Write(StringUtil::Format(
		    "<Override PartName=\"/xl/worksheets/%s\" "
		    "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
		    sheet.sheet_file));
	}
	stream.Write(CONTENT_TYPES_XML_END);
	stream.EndFile();
}

inline void XLXSWriter::WriteDocProps() {

	stream.AddDirectory("docProps/");

	static constexpr auto APP_PROPS_XML =
	    R"(<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">)"
	    R"(<Application>DuckDB</Application>)"
	    R"(<TotalTime>0</TotalTime>)"
	    R"(</Properties>)";

	stream.BeginFile("docProps/app.xml");
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(APP_PROPS_XML);
	stream.EndFile();

	static constexpr auto CORE_PROPS_XML =
		R"(<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">)"
		R"(<dcterms:created xsi:type="dcterms:W3CDTF">2024-11-15T13:37:00.00Z</dcterms:created>)"
		R"(<dc:creator>DuckDB</dc:creator>)"
		R"(<cp:lastModifiedBy>DuckDB</cp:lastModifiedBy>)"
		R"(<dcterms:modified xsi:type="dcterms:W3CDTF">2024-11-15T13:37:00.00Z</dcterms:modified>)"
		R"(<cp:revision>1</cp:revision>)"
		R"(</cp:coreProperties>)";

	stream.BeginFile("docProps/core.xml");
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(CORE_PROPS_XML);
	stream.EndFile();

}


inline void XLXSWriter::WriteRels() {
	static constexpr auto WORKBOOK_REL_XML_START =
	    R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>)";
	static constexpr auto WORKBOOK_REL_XML_END = R"(</Relationships>)";
	static constexpr auto REL_SHEET_XML =
	    R"(<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/%s"/>)";

	stream.AddDirectory("xl/_rels/");

	stream.BeginFile("xl/_rels/workbook.xml.rels");
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(WORKBOOK_REL_XML_START);
	idx_t sheet_offset = 3;
	for (const auto &sheet : written_sheets) {
		stream.Write(StringUtil::Format(REL_SHEET_XML, sheet_offset++, sheet.sheet_file));
	}
	stream.Write(WORKBOOK_REL_XML_END);
	stream.EndFile();
}

inline void XLXSWriter::WriteWorkbook() {
	static constexpr auto WORKBOOK_XML_START =
	    R"(<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><workbookPr/><sheets>)";
	static constexpr auto WORKBOOK_XML_END = R"(</sheets><definedNames/><calcPr/></workbook>)";

	stream.BeginFile("xl/workbook.xml");
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(WORKBOOK_XML_START);
	idx_t sheet_offset = 3;
	idx_t sheet_id = 1;
	for (const auto &sheet : written_sheets) {
		static constexpr auto SHEET_XML = R"(<sheet name="%s" state="visible" sheetId="%d" r:id="rId%d"/>)";
		stream.Write(StringUtil::Format(SHEET_XML, sheet.sheet_name, sheet_id++, sheet_offset++));
	}
	stream.Write(WORKBOOK_XML_END);
	stream.EndFile();
}

inline void XLXSWriter::WriteSharedStrings() {
	// We dont use shared strings for now, but still create a dummy file
	static constexpr auto SHARED_STRINGS_XML =
	    R"(<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>)";

	stream.BeginFile("xl/sharedStrings.xml");
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(SHARED_STRINGS_XML);
	stream.EndFile();
}

inline void XLXSWriter::WritePackageRels() {

	static constexpr auto ROOT_RELS =
	    R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
	    R"(<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>)"
	    R"(<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>)"
	    R"(<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>)"
	    R"(</Relationships>)";

	stream.AddDirectory("_rels/");
	stream.BeginFile("_rels/.rels");
	stream.Write(ENCODING_FRAGMENT);
	stream.Write(ROOT_RELS);
	stream.EndFile();
}

} // namespace duckdb