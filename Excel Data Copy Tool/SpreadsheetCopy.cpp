#include "SpreadsheetCopy.h"


void copySpreadsheetData(Book * sourceBook, Book * destinationBook) {

	//We expect the list of transactions (the source) to be sheet at index 1 (the second sheet)
	Sheet * sourceSheet = sourceBook->getSheet(1) ;
	
	wstring test = sourceSheet->readStr(4, 5) ;

	//
	



}



	