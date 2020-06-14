#include "fd-read-xlsx.hpp"
#include <cassert>

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
  // The namespace defines “cell_t” as the cell type.
  fd_read_xlsx::cell_t v{ "a" };
  // The namespace defines “compare” using the type of the second argument.
  assert(fd_read_xlsx::compare(v, std::string{ "a" }));
  // Syntactic sugar...
  assert(fd_read_xlsx::holds_string(table[0][0]));
  assert(fd_read_xlsx::holds_int(table[1][0]));
  assert(fd_read_xlsx::holds_double(table[2][0]));
  // The namespace defines “holds_num” for int or double.
  assert(fd_read_xlsx::holds_num(table[1][0]) &&
         fd_read_xlsx::holds_num(table[2][0]));
  // The namespace defines “get_num" which returns a double.
  assert(fd_read_xlsx::get_num(table[1][0]) == 1.);

  return 0;
}
