all : test-header-only fd-read-xlsx.a test format 

test-header-only : fd-read-xlsx-header-only.hpp test-header-only.cpp
	g++ -std=c++17 -Wall -g test-header-only.cpp -lzip --output test-header-only

# Make a static library.
fd-read-xlsx.a : fd-read-xlsx.cpp
	g++ -std=c++17 -Wall -o9 -c fd-read-xlsx.cpp --output fd-read-xlsx.a

# Test with static library.
test : fd-read-xlsx.hpp test.cpp
	g++ -std=c++17 -Wall -g test.cpp fd-read-xlsx.a -lzip --output test

format :
	clang-format -i fd-read-xlsx-header-only.hpp fd-read-xlsx.hpp fd-read-xlsx.cpp test-header-only.cpp
