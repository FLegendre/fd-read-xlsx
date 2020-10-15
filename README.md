# fd-read-xlsx
A fast and dirty C++ library to get the contents of a xlsx Microsoft workbook.

This is an header only library : put the `fd-read-xlsx.hpp` file in a appropriate place and use it as
```C++
#include <fd-read-xlsx.hpp>

int
main()
{

  // Get the active sheet of the “test.xlsx” workbook.
  auto const table{ fd_read_xlsx::read("test.xlsx") };

  // “table” is a std::vector of std::vectors of cells.
  // “cells” are a std::variant<std::string, int64_t, double>.
  for (auto const& row : table) {
    std::cout << '|';
    for (auto const& cell : row)
      std::cout << fd_read_xlsx::to_string(cell) << '|';
    std::cout << '\n';
  }

  return 0;
}
```

This library depends on the libzip library : https://libzip.org/.

This library do not cope with xml comments and xml CDATA sections.
