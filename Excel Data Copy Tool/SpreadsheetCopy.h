
#pragma once

#include <iostream>
#include <string>

#include "../BasicExcel/BasicExcel.hpp"

#include ""



using namespace std;
using namespace YExcel ;

void copySpreadsheetData(BasicExcel & source, BasicExcel & destination) ;

vector<YarnEntry> * retrieveSpreadsheetData(BasicExcel & source) ;

void copySpreadsheetData(vector<YarnEntry> * source, BasicExcel & destination) ;


wstring getRowName(BasicExcelWorksheet * sheet, unsigned rowNumber) { return wstring(sheet->Cell(rowNumber, 0)->GetWString()) ; }
wstring getColumnName(BasicExcelWorksheet * sheet, unsigned columnNumber) { return wstring(sheet->Cell(0, columnNumber)->GetWString()) ; }
