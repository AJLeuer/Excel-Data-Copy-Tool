
#include <iostream>
#include <string>

#include "SpreadsheetCopy.h"

using namespace libxl;
using namespace std;


int main(int argc, char ** argv) {

	//open source and destination workbooks
	Book * sourceBook = xlCreateXMLBook() ;
	
	Book * destinationBook = xlCreateXMLBook() ;
	
	bool noErrorsLoading = (sourceBook->load(L"Transactions.xlsx") && destinationBook->load(L"Dye List.xlsx")) ;

	if (noErrorsLoading == false) {
		cerr << "Problem loading Excel file" << endl ;
		throw exception() ;
	}

	
	copySpreadsheetData(sourceBook, destinationBook) ;







	return 0 ;
}