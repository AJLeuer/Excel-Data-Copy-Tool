
#include <iostream>
#include <string>


#include "SpreadsheetCopy.h"

using namespace YExcel ;
using namespace std;


int main(int argc, char ** argv) {

	//open source and destination workbooks
	BasicExcel sourceBook("Input.xls") ;

	BasicExcel destinationBook("Output.xls") ;

	copySpreadsheetData(sourceBook, destinationBook) ;
	
	destinationBook.Save() ;








	return 0 ;
}


/*
	Todo: check if value in each cell is string or wstring, and match accordingly.
	Use int BasicExcelCell::Type() const to check
*/