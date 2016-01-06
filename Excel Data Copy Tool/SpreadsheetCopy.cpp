#include "SpreadsheetCopy.h"


void copySpreadsheetData(BasicExcel & source, BasicExcel & destination) {

	//We expect the list of transactions (the source) to be sheet at index 1 (the second sheet)
	BasicExcelWorksheet * sourceSheet = source.GetWorksheet(1) ;
	
	wchar_t * emptyString = new wchar_t() ;
	bool test = sourceSheet->Cell(4,3)->Get(emptyString) ;

	//
	



}



	