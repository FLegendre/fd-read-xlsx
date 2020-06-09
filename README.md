# fd-read-xlsx
A fast and dirty C++ library to get the contents of a xlsx Microsoft workbook.

This is an header only library : put the `fd-read-xlsx.hpp` file in a appropriate place and use it as
```C++
#include <iostream>
#include <fd-read-xlsx.hpp>

int
main()
{

  auto const table{ fd_read_xlsx::read("test.xlsx") };

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
