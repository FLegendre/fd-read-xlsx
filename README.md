# fd-read-xlsx
A fast and dirty C++ library to get the xlsx worksheet contents.

This is an header only library : put the `fd-read-xlsx.hpp` file in a appropriate place and use it as
```C++
#include <fd-read-xlsx.hpp>

int
main()
{

  auto const table{ fd_read_xlsx::read("test.xlsx") };

  for (auto const& row : table) {
    std::cout << '|';
    for (auto const& cell : row)
      std::visit([](auto&& arg) { std::cout << arg << '|'; }, cell);
    std::cout << '\n';
  }

  return 0;
}
```

This libray depends on the libzip library : https://libzip.org/.
