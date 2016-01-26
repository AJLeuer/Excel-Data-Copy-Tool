
#pragma once

#include <vector>
#include <iostream>
#include <string>

#include "../BasicExcel/BasicExcel.hpp"
//#include "../BasicExcel/ExcelFormat.h"

#include "Util.hpp"
#include "YarnEntry.h"



using namespace std;
using namespace YExcel ;

void copySpreadsheetData(BasicExcel & source, BasicExcel & destination) ;

vector<const YarnEntry> * retrieveSpreadsheetData(BasicExcel & source) ;

void copySpreadsheetData(vector<const YarnEntry> * source, BasicExcel & destination) ;

/**
 * Returns true if there is a string to label the start of this row, false otherwise
 */
bool rowExists(BasicExcelWorksheet * sheet, unsigned rowNumber) ;
/**
 * Returns true if there is a string to label the start of this column, false otherwise
 */
bool colummExists(BasicExcelWorksheet * sheet, unsigned columnNumber) ;

string getRowName(BasicExcelWorksheet * sheet, unsigned rowNumber) ;
string getColumnName(BasicExcelWorksheet * sheet, unsigned columnNumber) ;

void setRowName(BasicExcelWorksheet * sheet, unsigned rowNumber, string name) ;
void setColumnName(BasicExcelWorksheet * sheet, unsigned columnNumber, string name) ;

/**
 * For a BasicExcelCell whose contents are unknown, this function extracts a string if
 * it contains either a string or wstring, else otherwise throws an exception
 */
string extractStringFromUnknownCell(BasicExcelCell * cell) ;
