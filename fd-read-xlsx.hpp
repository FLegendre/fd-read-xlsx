#ifndef FD_READ_XLSX_HPP
#define FD_READ_XLSX_HPP

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <iostream>
#include <sstream>
#include <stdexcept>
#include <string>
#include <tuple>
#include <variant>
#include <vector>
#include <zip.h>

#define FD_READ_XLSX_SHOW(arg)                                                 \
  std::cout << #arg << '{' << arg << '}' << std::endl;

namespace fd_read_xlsx {

typedef std::string str_t;

class Exception : public std::exception
{
public:
  Exception(str_t const& msg)
    : msg_(msg)
  {}
  const char* what() const throw() { return msg_.c_str(); }

private:
  str_t const msg_;
};

// Representation of a cell: std::variant of string, int64_t and double. A
// int64_t is choosen for int: the size is like the size of a double.
typedef std::variant<str_t, int64_t, double> cell_t;

// This function returns the value of the attribute “attr” of the tag “tag” in
// the “str” string from “pos”. This function returns [value, pos, end, error].
// “end” is true if the tag is not found. “error” is true if the tag is found
// but if the closed quote is not found. “pos” is the new position; “pos” can be
// equal to str.size() but this value can be reused in the next call to this
// function.
std::tuple<str_t, str_t::size_type, bool, bool>
get_attribute(str_t const& str,
              str_t::size_type pos,
              str_t const& nmspace,
              char const* const tag,
              char const* const attr)
{
  auto const beg_tag{ str.find(
    '<' + ((nmspace == "") ? nmspace : nmspace + ':') + tag, pos) };
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
    auto const value{ str_t(cbegin(str) + pos1 + strlen(attr) + 3,
                            cbegin(str) + pos2) };
    return std::make_tuple(value, end_tag + 2, false, false);
  } else if (quote_or_dbl_quote == '\'') {
    auto const pos2{ str.find('\'', pos1 + strlen(attr) + 3) };
    if (pos2 == str_t::npos)
      return std::make_tuple("", 0, false, true);
    auto const value{ str_t(cbegin(str) + pos1 + strlen(attr) + 3,
                            cbegin(str) + pos2) };
    return std::make_tuple(value, end_tag + 2, false, false);
  } else
    throw Exception{ "fd-read-xslx library: the attribute “" + str_t(attr) +
                     "” is not followed by either "
                     "“=\"” or “='” (file corrupted?)." };
}
// Get the file contents from the archive into a string.
str_t
get_contents(zip_t* archive_ptr, char const* const file_name)
{
  auto const file_ptr{ zip_fopen(archive_ptr, file_name, 0) };
  if (!file_ptr)
    throw Exception{ "fd-read-xslx library: unable to open “" +
                     str_t(file_name) + "” file." };
  str_t rvo;
  while (true) {
    char buffer[1024];
    auto const n = zip_fread(file_ptr, buffer, sizeof(buffer));
    if (n == 0)
      break;
    rvo += str_t(buffer, buffer + n);
  }
  return rvo;
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
// Get the shared strings in the xml file from a Microsoft xlsx workbook.
// We only get the text between <t ...> and </t> tags to populate the vector.
std::vector<str_t>
get_shared_strings(zip_t* archive_ptr, str_t const& nmspace)
{

  std::vector<str_t> rvo;

  try {
    // We presume that the file is not so big ; so we can get it in memory.
    auto const contents{ get_contents(archive_ptr, "xl/sharedStrings.xml") };
    auto const beg_tag{ '<' + ((nmspace == "") ? nmspace : (nmspace + ':')) +
                        't' };
    auto const end_tag{ "</" + ((nmspace == "") ? nmspace : (nmspace + ':')) +
                        "t>" };
    str_t::size_type pos{ 0 };
    while (true) {
      auto const pos0{ contents.find(beg_tag, pos) };
      if (pos0 == str_t::npos)
        break;
      auto const pos1{ contents.find('>', pos0 + beg_tag.size()) };
      if (pos1 == str_t::npos)
        throw Exception{ "fd-read-xslx library: unable to found the '>' char "
                         "after the “" +
                         beg_tag +
                         "” tag "
                         "(xl/sharedStrings.xml corrupted?)." };
      auto const pos2{ contents.find(end_tag, pos1 + 1) };
      if (pos2 == str_t::npos)
        throw Exception{ "fd-read-xslx library: unable to found the “" +
                         end_tag +
                         "” tag "
                         "after the “" +
                         beg_tag +
                         "” "
                         "tag (xl/sharedStrings.xml corrupted?)." };
      auto str = str_t(cbegin(contents) + pos1 + 1, cbegin(contents) + pos2);
      replace_all(str, "&lt;", '<');

      rvo.emplace_back(str);
      pos = pos2 + end_tag.size();
    }

  } catch (Exception& e) {
    // The sharedStrings.xml file does not always exist.
  }
  return rvo;
}

// This function returns the tuple of the xml namespace, the
// vector of (sheet names, sheet id) and active sheet index.
std::tuple<str_t, std::vector<std::pair<str_t, str_t>>, size_t>
get_namespace_sheet_names_active(zip_t* archive_ptr)
{
  // We presume that the file is not so big ; so we can get it in memory.
  auto const contents{ get_contents(archive_ptr, "xl/workbook.xml") };

  // Guess the namespace: if we find a tag with <NAMESPACE:workbook ...
  // xmlns:NAMESPACE=... then NAMESPACE is the namespace.
  auto const nmspace = [&]() -> str_t {
    auto const pos = contents.find("workbook ");
    if (pos == str_t::npos)
      throw Exception{ "fd-read-xslx library: unable to found the “workbook” "
                       "tag (xl/workbook.xml corrupted?)." };
    if (pos == 0)
      throw Exception{ "fd-read-xslx library: the string “workbook” is found "
                       "at position 0 (xl/workbook.xml corrupted?)." };
    if (contents[pos - 1] == '<')
      return "";
    if (contents[pos - 1] != ':')
      throw Exception{
        "fd-read-xslx library: the string “workbook” is not preceded by either "
        "“<” or “:” (xl/workbook.xml corrupted?)."
      };
    auto const pos0 = contents.rfind('<', pos - 2);
    if (pos0 == str_t::npos)
      throw Exception{ "fd-read-xslx library: unable to found the “<” for the "
                       "“workbook” tag (xl/workbook.xml corrupted?)." };
    return str_t{ begin(contents) + pos0 + 1, begin(contents) + pos - 1 };
  }();

  auto const active_index = [&]() -> size_t {
    auto const [str, pos, end, err] =
      get_attribute(contents, 0, nmspace, "workbookView", "activeTab");
    if (err)
      throw Exception{ "fd-read-xslx library: unable to found the closing “\"” "
                       "or “'”  char after the "
                       "“activeTab” attribute (xl/workbook.xml corrupted?)." };
    // If the workbook has only one worksheet, the active worksheet is not
    // indicated.
    return end ? 0 : std::stoi(str);
  }();

  std::vector<std::pair<str_t, str_t>> sheet_names;
  str_t::size_type pos{ 0 };
  while (true) {
    auto const [name, pos1, end1, err1]{ get_attribute(
      contents, pos, nmspace, "sheet ", "name") };
    if (err1)
      throw Exception{ "fd-read-xslx library: unable to found the closing “\"” "
                       "or “'” char after the "
                       "“name” attribute (xl/workbook.xml corrupted?)." };
    if (end1)
      break;
    auto const [id, pos2, end2, err2]{ get_attribute(
      contents, pos, nmspace, "sheet ", "sheetId") };
    if (err2 || end2)
      throw Exception{ "fd-read-xslx library: unable to found the closing “\"” "
                       "or “'” char after the "
                       "“sheetId” attribute (xl/workbook.xml corrupted?)." };
    sheet_names.emplace_back(name, id);
    pos = pos2;
  }
  // Be sure that the active index is valid.
  return std::make_tuple(nmspace,
                         sheet_names,
                         (active_index < sheet_names.size()) ? active_index
                                                             : 0);
}

// Read a sheet and returns a table (vectors of vectors) of variants.
std::vector<std::vector<cell_t>>
read(char const* const xlsx_file_name, char const* const sheet_name = "")
{

  int zip_error;
  auto const archive_ptr{ zip_open(xlsx_file_name, ZIP_RDONLY, &zip_error) };
  if (!archive_ptr)
    throw Exception{ "fd-read-xslx library: unable to open the “" +
                     str_t(xlsx_file_name) + "” workbook." };

  auto const [nmspace, sheet_names, active]{ get_namespace_sheet_names_active(
    archive_ptr) };

  // If the sheet name is empty, then the active sheet is used.
  auto const id = [&] {
    if (sheet_name[0] == '\0') {
      if (active >= sheet_names.size())
        throw Exception{ "fd-read-xslx library: invalid index for the active "
                         "worksheet (workbook corrupted?)." };
      return sheet_names[active].second;
    }
    auto const i =
      std::find_if(cbegin(sheet_names), cend(sheet_names), [&](auto const& p) {
        return p.first == sheet_name;
      });
    if (i == cend(sheet_names))
      throw Exception{ "fd-read-xslx library: unable to found the “" +
                       str_t(sheet_name) + "” worksheet." };
    return sheet_names[i - cbegin(sheet_names)].second;
  }();

  auto const shared_strings{ get_shared_strings(archive_ptr, nmspace) };

  std::vector<std::vector<cell_t>> rvo;

  auto const file_name = "xl/worksheets/sheet" + id + ".xml";

  auto const file_ptr{ zip_fopen(archive_ptr, file_name.c_str(), 0) };
  if (!file_ptr)
    throw Exception{ "fd-read-xslx library: unable to open the “" + file_name +
                     "” file." };

  char buffer[1024];
  size_t n{};
  size_t i{};

  auto const next_char = [&]() -> int {
    if (i == n) {
      n = zip_fread(file_ptr, buffer, sizeof(buffer));
      if (n == 0)
        return -1;
      i = 0;
    }
    return buffer[i++];
  };

  std::vector<cell_t> row;

  auto const push_value = [&](str_t& ref, str_t& type, str_t& value) {
    cell_t v;
    // String inline.
    if (type == "inlineStr") {
      replace_all(value, "&lt;", '<');
      v = value;
    }
    // Shared string.
    else if (type == "s") {
      auto const i = std::stoi(value);
      if ((0 <= i) && (size_t(i) < shared_strings.size()))
        v = shared_strings[i];
      else
        throw Exception{ "fd-read-xslx library: invalid index for the a "
                         "shared string (workbook corrupted?)." };
    } else {
      // Integer.
      if (value.find('.') == str_t::npos) {
        // Use “stoll” because the size of a “long long int” is at least 64
        // bytes.
        auto const tmp = std::stoll(value);
        if ((std::numeric_limits<int64_t>::min() <= tmp) &&
            (tmp <= std::numeric_limits<int64_t>::max()))
          v = int64_t(tmp);
        else
          v = double(tmp);
      }
      // Double.
      else
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
      throw Exception{
        "fd-read-xslx library: invalid cell ref (workbook corrupted?)."
      };
    --i;
    --j;
    // rvo.size() == 3 and i == 2 : error
    // rvo.size() == 2 and i == 2 : do nothing
    // rvo.size() == 1 and i == 2 : push the current row and clear it
    // rvo.size() == 0 and i == 2 : push the current row, clear it and add 1
    // empty row
    if (rvo.size() > i)
      throw Exception{
        "fd-read-xslx library: rows not sorted (workbook corrupted?)."
      };
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
      throw Exception{
        "fd-read-xslx library: columns not sorted (workbook corrupted?)."
      };
    else {
      auto const count = j - row.size();
      for (size_t jj{}; jj < count; ++jj)
        row.emplace_back(cell_t{});
      row.emplace_back(v);
    }
  };
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
  State state = State::start;

  for (int c = next_char(); c != -1; c = next_char()) {
    switch (state) {
      // Waiting for "<c ".
      case State::start:
        if (c == '<')
          state = State::lt;
        break;
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
      // This state is also the return state after the end of a attribute.
      case State::c:
        if (c == ' ')
          state = State::space;
        else if (c == '>')
          state = State::next;
        else
          ref.clear(), type.clear(),
            state = State::start; // <c r="xx" t="xx"/> : no value
        break;
      // Waiting for 'r', 't', unknow or '>'.
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
      // Waiting for ="xxx" or ='xxx'.
      case State::r:
        if (c == '=')
          state = State::re;
        else
          state = State::start;
        break;
      case State::re:
        if (c == '"')
          state = State::red;
        else if (c == '\'')
          state = State::res;
        else
          state = State::start;
        break;
      case State::red:
        if (c == '"')
          state = State::c;
        else
          ref += c;
        break;
      case State::res:
        if (c == '\'')
          state = State::c;
        else
          ref += c;
        break;
      // Waiting for ="xxx" or ='xxx'.
      case State::t:
        if (c == '=')
          state = State::te;
        else
          state = State::start;
        break;
      case State::te:
        if (c == '"')
          state = State::ted;
        else if (c == '\'')
          state = State::tes;
        else
          state = State::start;
        break;
      case State::ted:
        if (c == '"')
          state = State::c;
        else
          type += c;
        break;
      case State::tes:
        if (c == '\'')
          state = State::c;
        else
          type += c;
        break;
      // Waiting for ="xxx" or ='xxx'.
      case State::u:
        if (c == '=')
          state = State::ue;
        break;
      case State::ue:
        if (c == '"')
          state = State::ued;
        else if (c == '\'')
          state = State::ues;
        else
          state = State::start;
        break;
      case State::ued:
        if (c == '"')
          state = State::c;
        break;
      case State::ues:
        if (c == '\'')
          state = State::c;
        break;
      // Waiting for <v> or for <t>
      case State::next:
        if (c == '<')
          state = State::nlt;
        break;
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
      case State::nv:
        if (c == '>')
          state = State::nvgt;
        else
          state = State::next; // stay in next state to skip <is> tag
        break;
      // Yes, a cell is gotten !
      case State::nvgt:
        if (c == '<') {
          push_value(ref, type, value);
          ref.clear(), type.clear(), value.clear();
          state = State::start;
        } else
          value += c;
        break;
      case State::nt:
        if (c == '>')
          state = State::ntgt;
        else
          state = State::next; // stay in next state to skip <is> tag
        break;
      // Yes, a cell is gotten !
      case State::ntgt:
        if (c == '<') {
          push_value(ref, type, value);
          ref.clear(), type.clear(), value.clear();
          state = State::start;
        } else
          value += c;
        break;
      default:
        throw Exception{
          "fd-read-xslx library: internal error (should never occur...)."
        };
    }
  }
  // Do not forget to append the last row !
  if (!row.empty())
    rvo.emplace_back(row);
  return rvo;
}
template<typename T>
bool
compare(cell_t const& v, T const& t)
{
  return std::holds_alternative<T>(v) && (std::get<T>(v) == t);
}
bool
holds_string(cell_t const& v)
{
  return std::holds_alternative<str_t>(v);
}
bool
holds_int(cell_t const& v)
{
  return std::holds_alternative<int64_t>(v);
}
bool
holds_double(cell_t const& v)
{
  return std::holds_alternative<double>(v);
}
str_t
to_string(cell_t const& v)
{
  std::ostringstream out;
  std::visit([&](auto&& arg) { out << arg; }, v);
  return out.str();
}
} // namespace fd_read_xlsx
#endif // FD_READ_XLSX_HPP
