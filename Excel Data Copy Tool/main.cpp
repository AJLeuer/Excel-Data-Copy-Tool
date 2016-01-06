
#include <iostream>
#include <string>

#include "SpreadsheetCopy.h"

using namespace YExcel ;
using namespace std;


int main(int argc, char ** argv) {

	//open source and destination workbooks
	
	BasicExcel sourceBook("Sheet.xls") ;
	
	BasicExcel destinationBook("Sheet.xls") ;

	copySpreadsheetData(sourceBook, destinationBook) ;
	
	 






	return 0 ;
}