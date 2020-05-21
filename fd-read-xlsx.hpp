#ifndef FD_READ_XLSX_HPP
#define FD_READ_XLSX_HPP

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <iostream>
#include <stdexcept>
#include <string>
#include <tuple>
#include <variant>
#include <vector>
#include <zip.h>

#ifndef FD_READ_XLSX_SHOW
#define FD_READ_XLSX_SHOW(arg)                                                 \
  std::cout << #arg << '{' << arg << '}' << std::endl;
#endif

namespace fd_read_xlsx {

class Exception : public std::exception
{
public:
  Exception(std::string const& msg)
    : msg_(msg)
  {}
  const char* what() const throw() { return msg_.c_str(); }

private:
  std::string const msg_;
};

typedef std::variant<std::string, int64_t, double> fd_read_xlsx_variant_type;

// This function search for “target” in the “str” string. The “target” is the
// name of a attribute including the “=” sign and the opened double quote. The
// returned value is the attribute. This function returns [value, pos, end,
// error]. “end” is true if the target is not found. “error” is true if the
// target is found but if the closed double quote is not found. “pos” is the new
// position; “pos” can be equal to str.size() but this value can be reused in
// the next call to this function. We assume that there is no space around the
// “=” sign.
std::tuple<std::string, std::string::size_type, bool, bool>
get_attribute(std::string const& str,
              std::string::size_type pos,
              char const* const tag,
              char const* const target)
{
  auto const beg_tag{ str.find(tag, pos) };
  if (beg_tag == std::string::npos)
    return std::make_tuple("", 0, true, false);
  auto const end_tag{ str.find("/>", beg_tag) };
  auto const pos1{ str.find(target, beg_tag) };
  if ((pos1 > end_tag) || (pos1 == std::string::npos))
    return std::make_tuple("", 0, true, false);
  auto const pos2{ str.find('"', pos1 + strlen(target)) };
  if (pos2 == std::string::npos)
    return std::make_tuple("", 0, false, true);
  auto const value{ std::string(cbegin(str) + pos1 + strlen(target),
                                cbegin(str) + pos2) };
  return std::make_tuple(value, end_tag + 2, false, false);
}
// Get the file contents from the archive into a string.
std::string
get_contents(zip_t* archive_ptr, char const* const file_name)
{
  auto const file_ptr{ zip_fopen(archive_ptr, file_name, 0) };
  if (!file_ptr)
    throw Exception{ "fd-read-xslx library: unable to open " +
                     std::string(file_name) + " file." };
  std::string rvo;
  while (true) {
    char buffer[1024];
    auto const n = zip_fread(file_ptr, buffer, sizeof(buffer));
    if (n == 0)
      break;
    rvo += std::string(buffer, buffer + n);
  }
  return rvo;
}
void
replace_all(std::string& str, std::string that, char c)
{
  auto pos{ str.find(that) };
  while (pos != std::string::npos) {
    str.replace(pos, that.size(), std::string(1, c));
    pos = str.find(that, pos + 1);
  }
}
// Get the shared strings in the xml file from a Microsoft xlsx workbook.
// We only get the text between <t ...> and </t> tags to populate the vector.
std::vector<std::string>
get_shared_strings(zip_t* archive_ptr)
{

  std::vector<std::string> rvo;

  try {
    // We presume that the file is not so big ; so we can get it in memory.
    auto const contents{ get_contents(archive_ptr, "xl/sharedStrings.xml") };
    std::string::size_type pos{ 0 };
    while (true) {
      char const* const beg_tag{ "<t" };
      char const* const end_tag{ "</t>" };
      auto const pos0{ contents.find(beg_tag, pos) };
      if (pos0 == std::string::npos)
        break;
      auto const pos1{ contents.find('>', pos0 + 2) };
      if (pos1 == std::string::npos)
        throw Exception{ "fd-read-xslx library: unable to found the '>' char "
                         "after the “<t” tag "
                         "(xl/sharedStrings.xml corrupted?)." };
      auto const pos2{ contents.find(end_tag, pos1 + 1) };
      if (pos2 == std::string::npos)
        throw Exception{ "fd-read-xslx library: unable to found the “</t>” tag "
                         "after the “<t>” "
                         "tag (xl/sharedStrings.xml corrupted?)." };
      auto str =
        std::string(cbegin(contents) + pos1 + 1, cbegin(contents) + pos2);
      replace_all(str, "&lt;", '<');

      rvo.emplace_back(str);
      pos = pos2 + strlen(end_tag);
    }

  } catch (Exception& e) {
    // The sharedStrings.xml file does not always exist.
  }
  return rvo;
}

// Vector of (sheet names, sheet id) and active sheet index.
std::pair<std::vector<std::pair<std::string, std::string>>, size_t>
get_sheet_names(zip_t* archive_ptr)
{
  // We presume that the file is not so big ; so we can get it in memory.
  auto const contents{ get_contents(archive_ptr, "xl/workbook.xml") };

  auto const active_index = [&]() -> size_t {
    auto const [str, pos, end, err] =
      get_attribute(contents, 0, "<workbookView ", " activeTab=\"");
    if (err)
      throw Exception{
        "fd-read-xslx library: unable to found the closing '\"' char after the "
        "“activeTab” attribute (xl/workbook.xml corrupted?)."
      };
    // If the workbook has only one worksheet, the active worksheet is not
    // indicated.
    return end ? 0 : std::stoi(str);
  }();

  std::vector<std::pair<std::string, std::string>> sheet_names;
  std::string::size_type pos{ 0 };
  while (true) {
    auto const [name, pos1, end1, err1]{ get_attribute(
      contents, pos, "<sheet ", " name=\"") };
    if (err1)
      throw Exception{
        "fd-read-xslx library: unable to found the closing '\"' char after the "
        "“name” attribute (xl/workbook.xml corrupted?)."
      };
    if (end1)
      break;
    auto const [id, pos2, end2, err2]{ get_attribute(
      contents, pos, "<sheet ", " sheetId=\"") };
    if (err2 || end2)
      throw Exception{
        "fd-read-xslx library: unable to found the closing '\"' char after the "
        "“sheetId” attribute (xl/workbook.xml corrupted?)."
      };
    sheet_names.emplace_back(name, id);
    pos = pos2;
  }
  // Be sure that the active index is valid.
  return std::make_pair(sheet_names,
                        (active_index < sheet_names.size()) ? active_index : 0);
}

// Read a sheet and returns a table (vectors of vectors) of variants.
std::vector<std::vector<fd_read_xlsx_variant_type>>
read(char const* const xlsx_file_name, char const* const sheet_name = "")
{

  int zip_error;
  auto const archive_ptr{ zip_open(xlsx_file_name, ZIP_RDONLY, &zip_error) };
  if (!archive_ptr)
    throw Exception{ "fd-read-xslx library: unable to open the " +
                     std::string(xlsx_file_name) + " workbook." };

  auto const [sheet_names, active]{ get_sheet_names(archive_ptr) };

  // If the sheet name is empty, then the active sheet is used.
  auto const id = [&] {
    if (sheet_name[0] == '\0')
      return sheet_names[active].second;
    auto const i =
      std::find_if(cbegin(sheet_names), cend(sheet_names), [&](auto const& p) {
        return p.first == sheet_name;
      });
    if (i == cend(sheet_names))
      throw Exception{ "fd-read-xslx library: unable to found the " +
                       std::string(sheet_name) + " worksheet." };
    return sheet_names[i - cbegin(sheet_names)].second;
  }();

  auto const shared_strings{ get_shared_strings(archive_ptr) };

  std::vector<std::vector<fd_read_xlsx_variant_type>> rvo;

  auto const file_name = "xl/worksheets/sheet" + id + ".xml";

  auto const file_ptr{ zip_fopen(archive_ptr, file_name.c_str(), 0) };
  if (!file_ptr)
    throw Exception{ "fd-read-xslx library: unable to open the " + file_name +
                     " file." };

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

  std::vector<fd_read_xlsx_variant_type> row;

  auto const push_value =
    [&](std::string& ref, std::string& type, std::string& value) {
      fd_read_xlsx_variant_type v;
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
        if (value.find('.') == std::string::npos)
          v = std::stol(value);
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
      if (i > 0)
        --i;
      if (j > 0)
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
          row.emplace_back(fd_read_xlsx_variant_type{});
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
    t,
    te,
    ted,
    u,
    uu,
    ue,
    ued,
    next,
    nlt,
    nv,
    nvgt,
    nt,
    ntgt,
  };

  std::string ref, type, value;
  State state = State::start;

  for (int c = next_char(); c != -1; c = next_char()) {
    switch (state) {
      // Waiting for "<c ".
      case State::start:
        if (c == '<')
          state = State::lt;
        break;
      case State::lt:
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
        else if (c != ' ')
          state = State::u;
        break;
      // Waiting for ="xxx".
      case State::r:
        if (c == '=')
          state = State::re;
        else
          state = State::start;
        break;
      case State::re:
        if (c == '"')
          state = State::red;
        else
          state = State::start;
        break;
      case State::red:
        if (c == '"')
          state = State::c;
        else
          ref += c;
        break;
      // Waiting for ="xxx".
      case State::t:
        if (c == '=')
          state = State::te;
        else
          state = State::start;
        break;
      case State::te:
        if (c == '"')
          state = State::ted;
        else
          state = State::start;
        break;
      case State::ted:
        if (c == '"')
          state = State::c;
        else
          type += c;
        break;
      // Waiting for ="xxx".
      case State::u:
        if (c == '=')
          state = State::ue;
        break;
      case State::ue:
        if (c == '"')
          state = State::ued;
        else
          state = State::start;
        break;
      case State::ued:
        if (c == '"')
          state = State::c;
        break;
      // Waiting for <v> or for <t>
      case State::next:
        if (c == '<')
          state = State::nlt;
        break;
      case State::nlt:
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
} // namespace fd_read_xlsx
#endif // FD_READ_XLSX_HPP
