#include "SpreadsheetCopy.h"


void copySpreadsheetData(BasicExcel & source, BasicExcel & destination) {

	vector<YarnEntry> * yarnEntries = retrieveSpreadsheetData(source) ;

	copySpreadsheetData(yarnEntries, destination) ;
}

vector<YarnEntry> * retrieveSpreadsheetData(BasicExcel & source) {

	 vector<YarnEntry> * yarnEntries = new vector<YarnEntry>() ;

	 //We expect the list of transactions (the source) to be sheet at index 1 (the second sheet)
	 BasicExcelWorksheet * sourceSheet = source.GetWorksheet(1) ;

	 //the actual data on the source sheet starts at row 2 (i.e. 3rd row, BasicExcel counts rows starting at 0)
	 for (unsigned row = 2 ; row < sourceSheet->GetTotalRows() ; row++) {
		 const wchar_t * code = sourceSheet->Cell(row, 3)->GetWString()  ;
		 const wchar_t * color = sourceSheet->Cell(row, 4)->GetWString() ;
		 const unsigned quantity = sourceSheet->Cell(row, 2)->GetInteger() ;

		 yarnEntries->push_back(YarnEntry(wstring(code), wstring(color), quantity)) ;
	 }

	 return yarnEntries ;
}

void copySpreadsheetData(vector<YarnEntry> * source, BasicExcel & destination) {

	//We expect the sheet we're outputting to be sheet at index 0 (the first sheet)
	BasicExcelWorksheet * destinationSheet = destination.GetWorksheet(0) ;


	for (size_t i = 0 ; i < source->size() ; i++) {

		YarnEntry & currentEntry = source->at(i) ;

		//the first row in our destination sheet is "Absolute Magenta," at index 2
		for (unsigned row = 2 ; row < destinationSheet->GetTotalRows()  ; row++) {

			if (currentEntry.color == getRowName(destinationSheet, row)) {

				for (unsigned column = 2 ; column < destinationSheet->GetTotalRows()  ; column++) {

					if (currentEntry.twoDigitCode == getColumnName(destinationSheet, column)) {

						BasicExcelCell * destinationCell = destinationSheet->Cell(row, column) ;

						int currentCellValue = destinationCell->GetInteger() ;

						currentCellValue += currentEntry.quantity ;

						destinationCell->SetInteger(currentCellValue) ;
					}

				}

			}

		}




	}






}
