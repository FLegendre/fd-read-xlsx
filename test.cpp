#include "fd-read-xlsx.hpp"

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
