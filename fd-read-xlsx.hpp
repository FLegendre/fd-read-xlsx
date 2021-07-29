#ifndef FD_READ_XLSX_HPP
#define FD_READ_XLSX_HPP

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <iostream>
#include <map>
#include <sstream>
#include <stdexcept>
#include <string>
#include <tuple>
#include <variant>
#include <vector>
#include <zip.h>

#define FD_READ_XLSX_SHOW(arg) std::cout << #arg << '{' << (arg) << '}' << std::endl;

namespace fd_read_xlsx {

typedef std::string str_t;

// Representation of a cell: std::variant of string, int64_t and double. A int64_t is choosen for
// int: the size is like the size of a double.
typedef std::variant<str_t, int64_t, double> cell_t;

// Type of the table returned by the read function.
typedef std::vector<std::vector<cell_t>> table_t;

// A row of this table.
typedef std::vector<cell_t> row_t;

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
              char const* const attr);
// Get the file contents from the archive into a string.
str_t
get_contents(zip_t* archive_ptr, char const* const file_name);
str_t
get_contents(zip_t* archive_ptr, str_t const& file_name);
void
replace_all(str_t& str, str_t that, char c);
// Get the shared strings in the xml file from a Microsoft xlsx workbook.  We only concatenate the
// text between <t ...> and </t> tags within <si> and </si> tags to populate the vector.
std::vector<str_t>
get_shared_strings(zip_t* archive_ptr, str_t const& file_name, str_t const& nmspace);

// This function returns the tuple of the xml namespace, the map of (sheet ids, sheet names) and
// active sheet name.
std::tuple<str_t, std::map<str_t, str_t>, str_t>
get_ns_ids_and_active(zip_t* archive_ptr, str_t const& wb_base, str_t const& wb_name);
// For debug.
std::vector<str_t>
get_shared_strings(char const* const xlsx_file_name);
std::pair<str_t, str_t>
get_wb_base_and_name(zip_t* archive_ptr);
std::tuple<str_t, std::map<str_t, str_t>, str_t>
get_ws_and_shared(zip_t* archive_ptr, str_t const& wb_base, str_t const& wb_name);
// Read a sheet and returns a table (vectors of vectors) of variants.
std::pair<table_t, str_t>
get_table_sheetname(char const* const xlsx_file_name, char const* const sheet_name);
std::pair<table_t, str_t>
get_table_sheetname(char const* const xlsx_file_name);
table_t
read(char const* const xlsx_file_name, char const* const sheet_name);
table_t
read(char const* const xlsx_file_name);
// Read a workbook and returns the worksheet name list.
std::vector<str_t>
get_worksheet_names(char const* const xlsx_file_name);
table_t
read(str_t const& xlsx_file_name, char const* const sheet_name);
table_t
read(str_t const& xlsx_file_name);
std::map<std::string, size_t>
names(row_t const& v);
template<typename T>
bool
compare(cell_t const& cell, T const& t)
{
	return std::holds_alternative<T>(cell) && (std::get<T>(cell) == t);
}
bool
compare(cell_t const& cell, char const* ptr);
template<typename T>
bool inline compare(row_t const& v, size_t j, T const& t)
{
	return (j < v.size()) && compare(v[j], t);
}
bool
empty(cell_t const& cell);
bool inline empty(row_t const& v, size_t j)
{
	return (j >= v.size()) || empty(v[j]);
}
bool
holds_string(cell_t const& cell);
bool inline holds_string(row_t const& v, size_t j)
{
	return (j < v.size()) && holds_string(v[j]);
}
str_t
get_string(cell_t const& cell);
bool
holds_int(cell_t const& cell);
bool inline holds_int(row_t const& v, size_t j)
{
	return (j < v.size()) && holds_int(v[j]);
}
int64_t
get_int(cell_t const& cell);
bool
holds_double(cell_t const& cell);
bool inline holds_double(row_t const& v, size_t j)
{
	return (j < v.size()) && holds_double(v[j]);
}
double
get_double(cell_t const& cell);
bool inline holds_num(cell_t const& cell)
{
	return holds_int(cell) || holds_double(cell);
}
bool inline holds_num(row_t const& v, size_t j)
{
	return (j < v.size()) && holds_num(v[j]);
}
double
get_num(cell_t const& cell);
str_t
to_string(cell_t const& cell);
} // namespace fd_read_xlsx
#endif // FD_READ_XLSX_HPP
