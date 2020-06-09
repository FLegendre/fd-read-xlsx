all : test format 

test : test.cpp fd-read-xlsx.hpp
	g++ -std=c++17 -Wall -g test.cpp -lzip --output test

format :
	clang-format -i --style=Mozilla fd-read-xlsx.hpp test.cpp
