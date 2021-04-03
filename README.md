# fd-read-xlsx
A fast and dirty C++ library to get the contents of a xlsx Microsoft workbook.

This is an header only library : put the `fd-read-xlsx-header-only.hpp` file in a appropriate place and use it as
```C++
#include <fd-read-xlsx-header-only.hpp>

int
main()
{

  // Get the active sheet of the “test.xlsx” workbook.
  auto const sheet{ fd_read_xlsx::read("test.xlsx") };

  // “sheet” is a std::vector of std::vectors of cells.
  // “cell” is a std::variant<std::string, int64_t, double>.
  for (auto const& row : sheet) {
    std::cout << '|';
    for (auto const& cell : row)
      std::cout << fd_read_xlsx::to_string(cell) << '|';
    std::cout << '\n';
  }

  return 0;
}
```

This library depends on the libzip library: https://libzip.org/.

This library does not cope with xml comments and xml CDATA sections.
