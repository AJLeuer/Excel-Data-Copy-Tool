
#pragma once

#include <iostream>
#include <string>

#include "../BasicExcel/BasicExcel.hpp"


using namespace std;
using namespace YExcel ;

class DreamSheet : public BasicExcelWorksheet {

public:

	using BasicExcelWorksheet::BasicExcelWorksheet ;

	DreamSheet(const BasicExcelWorksheet & sheet) : BasicExcelWorksheet(sheet) {}

	wstring getRowName(unsigned rowNumber) ;

	wstring getColumnName(unsigned columnNumber) ;

} ;

