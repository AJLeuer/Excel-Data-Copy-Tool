
#pragma once

#include <vector>
#include <iostream>
#include <string>

#include "../BasicExcel/BasicExcel.hpp"

#include "Util.hpp"
#include "YarnEntry.h"



using namespace std;
using namespace YExcel ;

void copySpreadsheetData(BasicExcel & source, BasicExcel & destination) ;

vector<YarnEntry> * retrieveSpreadsheetData(BasicExcel & source) ;

void copySpreadsheetData(vector<YarnEntry> * source, BasicExcel & destination) ;

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
