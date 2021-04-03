#ifndef FD_READ_XLSX_HEADER_ONLY_HPP
#define FD_READ_XLSX_HEADER_ONLY_HPP

#include "fd-read-xlsx.hpp"

namespace fd_read_xlsx {

typedef std::string str_t;

class Exception : public std::exception
{
public:
	Exception(str_t const& msg)
	  : msg_("fd-read-xslx library: " + msg + '.')
	{}
	const char* what() const throw() { return msg_.c_str(); }

private:
	str_t const msg_;
};

// Representation of a cell: std::variant of string, int64_t and double. A int64_t is choosen for
// int: the size is like the size of a double.
typedef std::variant<str_t, int64_t, double> cell_t;

// Type of the table returned by the read function.
typedef std::vector<std::vector<cell_t>> table_t;

// This function returns the value of the attribute “attr” of the tag “tag” in the “str” string from
// “pos”. This function returns [value, pos, end, error].  “end” is true if the tag is not found.
// “error” is true if the tag is found but if the closed quote is not found. “pos” is the new
// position; “pos” can be equal to str.size() but this value can be reused in the next call to this
// function.
std::tuple<str_t, str_t::size_type, bool, bool>
get_attribute(str_t const& str,
              str_t::size_type pos,
              str_t const& nmspace,
              char const* const tag,
              char const* const attr)
{
	auto const beg_tag{ str.find('<' + ((nmspace == "") ? nmspace : nmspace + ':') + tag, pos) };
	if (beg_tag == str_t::npos)
		return std::make_tuple("", 0, true, false);
	auto const end_tag{ str.find("/>", beg_tag) };
	auto const pos1{ str.find(' ' + str_t(attr) + '=', beg_tag) };
	if ((pos1 > end_tag) || (pos1 == str_t::npos))
		return std::make_tuple("", 0, true, false);
	auto const quote_or_dbl_quote = str[pos1 + strlen(attr) + 2];
	if (quote_or_dbl_quote == '"') {
		auto const pos2{ str.find('"', pos1 + strlen(attr) + 3) };
		if (pos2 == str_t::npos)
			return std::make_tuple("", 0, false, true);
		auto const value{ str_t(cbegin(str) + pos1 + strlen(attr) + 3, cbegin(str) + pos2) };
		return std::make_tuple(value, end_tag + 2, false, false);
	} else if (quote_or_dbl_quote == '\'') {
		auto const pos2{ str.find('\'', pos1 + strlen(attr) + 3) };
		if (pos2 == str_t::npos)
			return std::make_tuple("", 0, false, true);
		auto const value{ str_t(cbegin(str) + pos1 + strlen(attr) + 3, cbegin(str) + pos2) };
		return std::make_tuple(value, end_tag + 2, false, false);
	} else
		throw Exception{ "the attribute “" + str_t(attr) +
			               "” is not followed by either “=\"” or “='” (file corrupted?)" };
}
// Get the file contents from the archive into a string.
str_t
get_contents(zip_t* archive_ptr, char const* const file_name)
{
	auto const file_ptr{ zip_fopen(archive_ptr, file_name, 0) };
	if (!file_ptr)
		throw Exception{ "unable to open the “" + str_t(file_name) +
			               "” workbook (or the file is not a xlsx workbook)" };
	str_t rvo;
	while (true) {
		char buffer[1024];
		auto const n{ zip_fread(file_ptr, buffer, sizeof(buffer)) };
		if (n == 0)
			break;
		rvo += str_t(buffer, buffer + n);
	}
	return rvo;
}
str_t
get_contents(zip_t* archive_ptr, str_t const& file_name)
{
	return get_contents(archive_ptr, file_name.c_str());
}
void
replace_all(str_t& str, str_t that, char c)
{
	auto pos{ str.find(that) };
	while (pos != str_t::npos) {
		str.replace(pos, that.size(), str_t(1, c));
		pos = str.find(that, pos + 1);
	}
}
// Get the shared strings in the xml file from a Microsoft xlsx workbook.  We only concatenate the
// text between <t ...> and </t> tags within <si> and </si> tags to populate the vector.
std::vector<str_t>
get_shared_strings(zip_t* archive_ptr, str_t const& file_name, str_t const& nmspace)
{

	std::vector<str_t> rvo;

	// We presume that the file is not so big ; so we can get it in memory.
	auto const contents{ get_contents(archive_ptr, file_name) };

	auto const beg_si_tag{ '<' + ((nmspace == "") ? nmspace : (nmspace + ':')) + "si>" };
	auto const end_si_tag{ "</" + ((nmspace == "") ? nmspace : (nmspace + ':')) + "si>" };
	auto const beg_t_tag{ '<' + ((nmspace == "") ? nmspace : (nmspace + ':')) + 't' };
	auto const end_t_tag{ "</" + ((nmspace == "") ? nmspace : (nmspace + ':')) + "t>" };
	str_t::size_type pos{ 0 };
	while (true) {
		auto const pos_si_0{ contents.find(beg_si_tag, pos) };
		if (pos_si_0 == str_t::npos)
			break;
		auto const pos_si_1{ contents.find(end_si_tag, pos_si_0 + beg_si_tag.size()) };
		if (pos_si_1 == str_t::npos)
			throw Exception{ "unable to found the “" + end_si_tag + "” string after the “" + beg_si_tag +
				               "” tag (" + file_name + " corrupted?)" };
		str_t str;
		str_t::size_type pos_t{ pos_si_0 + beg_si_tag.size() };
		while (true) {
			auto const pos_t_0{ contents.find(beg_t_tag, pos_t) };
			if (pos_t_0 > pos_si_1)
				break;
			auto const pos_t_1{ contents.find('>', pos_t_0 + beg_t_tag.size()) };
			if (pos_t_1 == str_t::npos)
				throw Exception{ "unable to found the '>' char after the “" + beg_t_tag + "” tag (" +
					               file_name + " corrupted?)" };
			auto const pos_t_2{ contents.find(end_t_tag, pos_t_1 + 1) };
			if (pos_t_2 == str_t::npos)
				throw Exception{ "unable to found the “" + end_t_tag + "” tag after the “" + beg_t_tag +
					               "” tag (" + file_name + " corrupted?)" };
			str += str_t(cbegin(contents) + pos_t_1 + 1, cbegin(contents) + pos_t_2);
			pos_t = pos_t_2 + end_t_tag.size();
		}
		replace_all(str, "&lt;", '<');
		replace_all(str, "&gt;", '>');
		replace_all(str, "&quot;", '"');
		replace_all(str, "&apos;", '\'');
		rvo.emplace_back(str);
		pos = pos_si_1 + end_si_tag.size();
	}

	return rvo;
}

// This function returns the tuple of the xml namespace, the map of (sheet ids, sheet names) and
// active sheet name.
std::tuple<str_t, std::map<str_t, str_t>, str_t>
get_ns_ids_and_active(zip_t* archive_ptr, str_t const& wb_base, str_t const& wb_name)
{
	// We presume that the file is not so big ; so we can get it in memory.
	auto const contents{ get_contents(archive_ptr, wb_base + '/' + wb_name) };

	// Guess the namespace: if we find a tag with <NAMESPACE:workbook ... xmlns:NAMESPACE=... then
	// NAMESPACE is the namespace.
	auto const nmspace{ [&]() -> str_t {
		auto const pos = contents.find("workbook ");
		if (pos == str_t::npos)
			throw Exception{ "unable to found the “workbook” tag (" + wb_base + '/' + wb_name +
				               " corrupted?)" };
		if (pos == 0)
			throw Exception{ "the string “workbook” is found at position 0 (" + wb_base + '/' + wb_name +
				               " corrupted?)" };
		if (contents[pos - 1] == '<')
			return "";
		if (contents[pos - 1] != ':')
			throw Exception{ "the string “workbook” is not preceded by either “<” or “:” (" + wb_base +
				               '/' + wb_name + " corrupted?)" };
		auto const pos0 = contents.rfind('<', pos - 2);
		if (pos0 == str_t::npos)
			throw Exception{ "unable to found the “<” for the “workbook” tag (" + wb_base + '/' +
				               wb_name + " corrupted?)" };
		return str_t{ cbegin(contents) + pos0 + 1, cbegin(contents) + pos - 1 };
	}() };

	auto const active_tab{ [&]() -> int {
		auto const [str, pos, end, err] =
		  get_attribute(contents, 0, nmspace, "workbookView ", "activeTab");
		if (err)
			throw Exception{
				"unable to found the closing “\"” or “'”  char after the “activeTab” attribute (" +
				wb_base + '/' + wb_name + " corrupted?)"
			};
		// If the workbook has only one worksheet, the active worksheet is not indicated.
		return end ? 0 : std::stoi(str);
	}() };

	std::map<str_t, str_t> ids;
	str_t::size_type pos{};
	str_t active_name;
	// <sheet name="sheet_name" sheetId="1" r:id="rId1"/>
	while (true) {
		auto const [name, pos_name, end_name, err_name]{ get_attribute(
			contents, pos, nmspace, "sheet ", "name") };
		auto const [id, pos_id, end_id, err_id]{ get_attribute(
			contents, pos, nmspace, "sheet ", "sheetId") };
		auto const [rid, pos_rid, end_rid, err_rid]{ get_attribute(
			contents, pos, nmspace, "sheet ", "r:id") };
		if (end_name || err_name || end_id || err_id || end_rid || err_rid) {
			if (ids.empty())
				throw Exception{ "unable to found the  sheet names (" + wb_base + '/' + wb_name +
					               " corrupted?)" };
			return { nmspace, ids, active_name };
		}
		if (std::stoi(id) == (active_tab + 1))
			active_name = name;
		ids[name] = rid;
		pos = pos_rid;
	}
}
// For debug.
std::vector<str_t>
get_shared_strings(char const* const xlsx_file_name)
{
	int zip_error;
	auto const archive_ptr{ zip_open(xlsx_file_name, ZIP_RDONLY, &zip_error) };
	if (!archive_ptr)
		throw Exception{ "unable to open the “" + str_t(xlsx_file_name) +
			               "” workbook (or the file is not a xlsx workbook)" };

	return get_shared_strings(archive_ptr, "xl/sharedStrings.xml", "");
}

std::pair<str_t, str_t>
get_wb_base_and_name(zip_t* archive_ptr)
{
	// We presume that the file is not so big ; so we can get it in memory.
	auto const contents{ get_contents(archive_ptr, "_rels/.rels") };
	// We are looking for <Relationship
	// Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	// Target="xl/workbook.xml"/>
	str_t::size_type pos{};
	while (true) {
		auto const [type, pos_type, end_type, err_type]{ get_attribute(
			contents, pos, "", "Relationship ", "Type") };
		auto const [target, pos_target, end_target, err_target]{ get_attribute(
			contents, pos, "", "Relationship ", "Target") };
		if (end_type || err_type || end_target || err_type)
			throw Exception{ "unable to found the workbook name (file corrupted?)" };
		if (type.find("relationships/officeDocument") != str_t::npos) {
			auto const pos{ target.find('/') };
			if (pos == str_t::npos)
				throw Exception{ "unable to found the workbook base (file corrupted?)" };
			return { str_t{ cbegin(target), cbegin(target) + pos },
				       str_t{ cbegin(target) + pos + 1, cend(target) } };
		}
		pos = pos_target;
	}
}
std::tuple<str_t, std::map<str_t, str_t>, str_t>
get_ws_and_shared(zip_t* archive_ptr, str_t const& wb_base, str_t const& wb_name)
{
	// We presume that the file is not so big ; so we can get it in memory.
	auto const contents{ get_contents(archive_ptr, wb_base + "/_rels/" + wb_name + ".rels") };
	// We are looking for
	// <Relationship
	//   Id="rId1"
	//   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
	//   Target="worksheets/sheet1.xml"/>
	// or
	// <Relationship
	//   Id="rId2"
	//   Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
	//   Target="sharedStrings.xml"/>

	str_t base;
	std::map<str_t, str_t> names;
	str_t shared;
	str_t::size_type pos{};
	while (true) {
		auto const [id, pos_id, end_id, err_id]{ get_attribute(
			contents, pos, "", "Relationship ", "Id") };
		auto const [type, pos_type, end_type, err_type]{ get_attribute(
			contents, pos, "", "Relationship ", "Type") };
		auto const [target, pos_target, end_target, err_target]{ get_attribute(
			contents, pos, "", "Relationship ", "Target") };
		if (end_id || err_id || end_type || err_type || end_target || err_type) {
			if (base.empty())
				throw Exception{ "unable to found the worksheet names (file corrupted?)" };
			return { base, names, shared };
		}
		if (type.find("relationships/worksheet") != str_t::npos) {
			auto const pos{ target.find('/') };
			if (pos == str_t::npos)
				throw Exception{ "unable to found the worksheet base (file corrupted?)" };
			if (base.empty())
				base = str_t{ cbegin(target), cbegin(target) + pos };
			else if (base != str_t{ cbegin(target), cbegin(target) + pos })
				throw Exception{ "different worksheet bases (file corrupted?)" };
			names[id] = str_t{ cbegin(target) + pos + 1, cend(target) };
		} else if (type.find("relationships/sharedStrings") != str_t::npos) {
			// Some time, it is only the relative path, some time, it is the absolute...
			if (target.find('/') != str_t::npos) {
				if (target[0] == '/')
					shared = str_t{ cbegin(target) + 1, cend(target) };
				else
					shared = target;
			} else
				shared = wb_base + '/' + target;
		}
		pos = pos_target;
	}
}

// Class for RAII.
struct Zip
{
	Zip(char const* const file_name)
	{
		int zip_error;
		archive_ptr_ = zip_open(file_name, ZIP_RDONLY, &zip_error);
		if (!archive_ptr_)
			throw Exception{ "unable to open the “" + str_t(file_name) +
				               "” workbook (or the file is not a xlsx workbook)" };
	}
	~Zip()
	{
		if (archive_ptr_)
			// Do not check the return code as is bad to throw an exception in a destructor...
			zip_close(archive_ptr_);
	}
	zip_t* archive_ptr_;
};

// Read a sheet and returns a table (vectors of vectors) of variants.
std::pair<std::vector<std::vector<cell_t>>, str_t>
get_table_sheetname(char const* const xlsx_file_name, char const* const sheet_name)
{

	auto const zip{ Zip{ xlsx_file_name } };

	// The archive tree is
	//          _rels
	//          xl
	//          [Content_Types].xml
	// We do not read [Content_Types].xml, assuming that the file is a xlsx file.
	// We read the “_rels/.rels” file to get the workbook base and name (the base is
	// usually “xl” and the name “workbook.xml”).

	auto const [wb_base, wb_name]{ get_wb_base_and_name(zip.archive_ptr_) };

	// The xl directory is
	//          _rels
	//          worksheets
	//          workbook.xml
	//          sharedStrings.xml
	// We read the “_rels/workbook.xml.rels” file to get the worksheets base, the Ids and names of the
	// worksheets (the base is usually “worksheets", the Ids “rId1”, “rId2”, ... and the names
	// “sheet1.xml”, “sheet2.xml”, ...) and the shared file name. ws_names is a map with rid as key
	// and effective file name as value.
	auto const [ws_base, ws_names, shared]{ get_ws_and_shared(zip.archive_ptr_, wb_base, wb_name) };

	// We read the “workbook.xml” to get the namespace, the worksheets effective names and the active
	// sheet. ids is a map with sheet name as key and rid as value.
	auto const [nmspace, ids, active]{ get_ns_ids_and_active(zip.archive_ptr_, wb_base, wb_name) };

	auto const [sheet_file_name, sheetname]{ [&]() {
		// The user asks for the active sheet.
		if (sheet_name[0] == '\0') {
			if (active == "")
				return std::pair{ wb_base + '/' + ws_base + '/' + cbegin(ws_names)->second,
					                cbegin(ws_names)->second };
			else {
				auto const it_ids{ ids.find(active) };
				if (it_ids == cend(ids))
					throw Exception{ "unable to get the active sheet (file corrupted?)" };
				auto const it_names{ ws_names.find(it_ids->second) };
				if (it_names == cend(ws_names))
					throw Exception{ "unable to get the requested sheet (file corrupted?)" };
				return std::pair{ wb_base + '/' + ws_base + '/' + it_names->second, active };
			}
		}
		// The user asks for a requested sheet.
		auto const it_ids{ ids.find(sheet_name) };
		if (it_ids == cend(ids))
			throw Exception{ "the requested sheet “" + str_t{ sheet_name } + "” is not in the workbook" };
		auto const it_names{ ws_names.find(it_ids->second) };
		if (it_names == cend(ws_names))
			throw Exception{ "unable to get the requested sheet (file corrupted?)" };
		return std::pair{ wb_base + '/' + ws_base + '/' + it_names->second, str_t{ sheet_name } };
	}() };

	auto const shared_strings{ shared.empty()
		                           ? std::vector<str_t>{}
		                           : get_shared_strings(zip.archive_ptr_, shared, nmspace) };

	std::vector<std::vector<cell_t>> rvo;

	auto const file_ptr{ zip_fopen(zip.archive_ptr_, sheet_file_name.c_str(), 0) };
	if (!file_ptr)
		throw Exception{ "unable to open the “" + sheet_file_name + "” file" };

	char buffer[1024];
	size_t n{};
	size_t i{};

	auto const next_char{ [&]() -> int {
		if (i == n) {
			n = zip_fread(file_ptr, buffer, sizeof(buffer));
			if (n == 0)
				return -1;
			i = 0;
		}
		return buffer[i++];
	} };

	std::vector<cell_t> row;

	auto const push_value{ [&](str_t& ref, str_t& type, str_t& value) {
		cell_t v;
		// String inline.
		if (type == "inlineStr") {
			replace_all(value, "&lt;", '<');
			replace_all(value, "&gt;", '>');
			replace_all(value, "&quot;", '"');
			replace_all(value, "&apos;", '\'');
			v = value;
		}
		// Shared string.
		else if (type == "s") {
			auto const i = std::stoi(value);
			if ((0 <= i) && (size_t(i) < shared_strings.size()))
				v = shared_strings[i];
			else
				throw Exception{ "invalid index for the a shared string (workbook corrupted?)" };
		} else {
			if (value.find('.') == str_t::npos) {
				// Integer: use “stoll” because the size of a “long long int” is at
				// least 64 bytes.
				auto const tmp = std::stoll(value);
				if ((std::numeric_limits<int64_t>::min() <= tmp) &&
				    (tmp <= std::numeric_limits<int64_t>::max()))
					v = int64_t(tmp);
				else
					v = double(tmp);
			} else
				v = std::stod(value);
		}
		size_t i{}, j{};
		for (auto const& c : ref) {
			if (('A' <= c) && (c <= 'Z'))
				j = 26 * j + size_t(1 + c - 'A');
			else if (('0' <= c) && (c <= '9'))
				i = 10 * i + size_t(c - '0');
		}
		if ((i == 0) || (j == 0))
			throw Exception{ "invalid cell ref (workbook corrupted?)" };
		--i;
		--j;
		// rvo.size() == 3 and i == 2 : error
		// rvo.size() == 2 and i == 2 : do nothing
		// rvo.size() == 1 and i == 2 : push the current row and clear it
		// rvo.size() == 0 and i == 2 : push the current row, clear it and add 1
		// empty row
		if (rvo.size() > i)
			throw Exception{ "rows not sorted (workbook corrupted?)" };
		else if (rvo.size() < i) {
			rvo.emplace_back(row);
			row.clear();
			auto const count = i - rvo.size();
			for (size_t ii{}; ii < count; ++ii)
				rvo.emplace_back(row);
		}
		// row.size() == 2 and j == 1 : error
		// row.size() == 1 and j == 1 : push the element on the current row
		// row.size() == 0 and j == 1 : add an empty cell and push the element
		if (row.size() > j)
			throw Exception{ "columns not sorted (workbook corrupted?)" };
		else {
			auto const count = j - row.size();
			for (size_t jj{}; jj < count; ++jj)
				row.emplace_back(cell_t{});
			row.emplace_back(v);
		}
	} };
	// Integer value : <c r="A1"> <v>12</v> </c>
	// Double value : <c r="A1"> <v>1.2</v> </c>
	// Shared string : <c r="A1" t="s"> <v>0</v> </c>
	// Inline string : <c r="A1" t="inlineStr"> <is> <t>a string</t> </is> </c>
	// <is> is for strings inline.
	enum class State
	{
		start,
		lt,
		c,
		space,
		r,
		re,
		red,
		res,
		t,
		te,
		ted,
		tes,
		u,
		uu,
		ue,
		ued,
		ues,
		next,
		nlt,
		nv,
		nvgt,
		nt,
		ntgt,
	};

	str_t ref, type, value;
	State state{ State::start };

	for (int c{ next_char() }; c != -1; c = next_char()) {
		switch (state) {
			// Waiting for “<”.
			case State::start:
				if (c == '<')
					state = State::lt;
				break;
				// Waiting for “xml_namespace:c”.
			case State::lt:
				if (nmspace != "") {
					for (auto const& cc : nmspace) {
						if (c != cc) {
							state = State::start;
							break;
						}
						c = next_char();
					}
					if (state == State::lt) {
						if (c != ':') {
							state = State::start;
							break;
						}
						c = next_char();
					} else
						break;
				}
				if (c == 'c')
					state = State::c;
				else
					state = State::start;
				break;
			// Waiting for a attibute within the “c” tag. This state is also the return state after the
			// end of a attribute.
			case State::c:
				if (c == ' ')
					state = State::space;
				else if (c == '>')
					state = State::next;
				else
					ref.clear(), type.clear(),
					  state = State::start; // <c r="xx" t="xx"/> : no value
				break;
			// Waiting for attributes “r”, “t” or unknow or for char “>”.
			case State::space:
				if (c == 'r')
					state = State::r;
				else if (c == 't')
					state = State::t;
				else if (c == '>')
					state = State::next;
				else if (c == '/')
					ref.clear(), type.clear(),
					  state = State::start; // <c r="xx" t="xx" /> : no value
				else if (c != ' ')
					state = State::u;
				break;
			// Waiting for “=” after the “r” attribute.
			case State::r:
				if (c == '=')
					state = State::re;
				else
					state = State::start;
				break;
			// Waiting for a single ou a double quote.
			case State::re:
				if (c == '"')
					state = State::red;
				else if (c == '\'')
					state = State::res;
				else
					state = State::start;
				break;
			// Waiting for the value of the “r” attribute after a double quote.
			case State::red:
				if (c == '"')
					state = State::c;
				else
					ref += c;
				break;
			// Waiting for the value of the “r” attribute after a single quote.
			case State::res:
				if (c == '\'')
					state = State::c;
				else
					ref += c;
				break;
			// Waiting for “=” after the “t” attribute.
			case State::t:
				if (c == '=')
					state = State::te;
				else
					state = State::start;
				break;
			// Waiting for a single ou a double quote.
			case State::te:
				if (c == '"')
					state = State::ted;
				else if (c == '\'')
					state = State::tes;
				else
					state = State::start;
				break;
			// Waiting for the value of the “t” attribute after a double quote.
			case State::ted:
				if (c == '"')
					state = State::c;
				else
					type += c;
				break;
			// Waiting for the value of the “t” attribute after a single quote.
			case State::tes:
				if (c == '\'')
					state = State::c;
				else
					type += c;
				break;
			// Waiting for “=” after a unknow attribute.
			case State::u:
				if (c == '=')
					state = State::ue;
				break;
			// Waiting for a single ou a double quote.
			case State::ue:
				if (c == '"')
					state = State::ued;
				else if (c == '\'')
					state = State::ues;
				else
					state = State::start;
				break;
			// Waiting for the value of an unknow attribute after a double quote (to discard it).
			case State::ued:
				if (c == '"')
					state = State::c;
				break;
			// Waiting for the value of an unknow attribute after a single quote (to discard it).
			case State::ues:
				if (c == '\'')
					state = State::c;
				break;
			// Waiting for “<” within the “c” tag.
			case State::next:
				if (c == '<')
					state = State::nlt;
				break;
				// Waiting for “xml_namespace:v” or “xml_namespace:t”.
			case State::nlt:
				if (nmspace != "") {
					for (auto const& cc : nmspace) {
						if (c != cc) {
							state = State::next; // stay in next state to skip <is> tag
							break;
						}
						c = next_char();
					}
					if (state == State::nlt) {
						if (c != ':') {
							state = State::next; // stay in next state to skip <is> tag
							break;
						}
						c = next_char();
					} else
						break;
				}
				if (c == 'v')
					state = State::nv;
				else if (c == 't')
					state = State::nt;
				else
					state = State::next; // stay in next state to skip <is> tag
				break;
				// Waiting for “>” after the “v” tag.
			case State::nv:
				if (c == '>')
					state = State::nvgt;
				else
					state = State::next; // stay in next state to skip <is> tag
				break;
			// Waiting for the value of the “v” tag.
			case State::nvgt:
				if (c == '<') {
					// YES.
					push_value(ref, type, value);
					ref.clear(), type.clear(), value.clear();
					state = State::start;
				} else
					value += c;
				break;
				// Waiting for “>” after the “t” tag.
			case State::nt:
				if (c == '>')
					state = State::ntgt;
				else
					state = State::next; // stay in next state to skip <is> tag
				break;
			// Waiting for the value of the “t” tag.
			case State::ntgt:
				if (c == '<') {
					// YES.
					push_value(ref, type, value);
					ref.clear(), type.clear(), value.clear();
					state = State::start;
				} else
					value += c;
				break;
			// Oups...
			default:
				throw Exception{ "internal error (should never occur...)" };
		}
	}
	// Do not forget to append the last row !
	if (!row.empty())
		rvo.emplace_back(row);
	return { rvo, sheetname };
}
std::pair<std::vector<std::vector<cell_t>>, str_t>
get_table_sheetname(char const* const xlsx_file_name)
{
	return get_table_sheetname(xlsx_file_name, "");
}
std::vector<std::vector<cell_t>>
read(char const* const xlsx_file_name, char const* const sheet_name)
{
	return get_table_sheetname(xlsx_file_name, sheet_name).first;
}
std::vector<std::vector<cell_t>>
read(char const* const xlsx_file_name)
{
	return read(xlsx_file_name, "");
}

// Read a workbook and returns the worksheet name list.
std::vector<str_t>
get_worksheet_names(char const* const xlsx_file_name)
{

	auto const zip{ Zip{ xlsx_file_name } };
	auto const [wb_base, wb_name]{ get_wb_base_and_name(zip.archive_ptr_) };
	auto const [ws_base, ws_names, shared]{ get_ws_and_shared(zip.archive_ptr_, wb_base, wb_name) };
	auto const [nmspace, ids, active]{ get_ns_ids_and_active(zip.archive_ptr_, wb_base, wb_name) };
	std::vector<str_t> rvo;
	rvo.reserve(ids.size());
	for (auto const& p : ids)
		rvo.push_back(p.first);
	return rvo;
}
std::vector<std::vector<cell_t>>
read(str_t const& xlsx_file_name, char const* const sheet_name)
{
	return read(xlsx_file_name.c_str(), sheet_name);
}
std::vector<std::vector<cell_t>>
read(str_t const& xlsx_file_name)
{
	return read(xlsx_file_name, "");
}
std::map<std::string, size_t>
names(std::vector<cell_t> const& v)
{
	std::map<std::string, size_t> rvo;
	for (auto const& c : v)
		if (std::holds_alternative<std::string>(c))
			rvo[std::get<std::string>(c)] = &c - &v[0];
	return rvo;
}
bool
compare(cell_t const& v, char const* ptr)
{
	return std::holds_alternative<str_t>(v) && (std::strcmp(ptr, std::get<str_t>(v).c_str()) == 0);
}
bool
empty(cell_t const& v)
{
	return std::holds_alternative<str_t>(v) && std::get<str_t>(v).empty();
}
bool
holds_string(cell_t const& v)
{
	return std::holds_alternative<str_t>(v);
}
str_t
get_string(cell_t const& v)
{
	return std::get<str_t>(v);
}
bool
holds_int(cell_t const& v)
{
	return std::holds_alternative<int64_t>(v);
}
int64_t
get_int(cell_t const& v)
{
	return std::get<int64_t>(v);
}
bool
holds_double(cell_t const& v)
{
	return std::holds_alternative<double>(v);
}
double
get_double(cell_t const& v)
{
	return std::get<double>(v);
}
bool
holds_num(cell_t const& v)
{
	return holds_int(v) || holds_double(v);
}
double
get_num(cell_t const& v)
{
	return holds_int(v) ? get_int(v) : get_double(v);
}
str_t
to_string(cell_t const& v)
{
	std::ostringstream out;
	std::visit([&](auto&& arg) { out << arg; }, v);
	return out.str();
}
} // namespace fd_read_xlsx
#endif // FD_READ_XLSX_HEADER_ONLY_HPP
