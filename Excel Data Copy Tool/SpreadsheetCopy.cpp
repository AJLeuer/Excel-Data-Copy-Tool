#include "SpreadsheetCopy.h"


void copySpreadsheetData(BasicExcel & source, BasicExcel & destination) {

	vector<YarnEntry> * yarnEntries = retrieveSpreadsheetData(source) ;

	copySpreadsheetData(yarnEntries, destination) ;
}

vector<YarnEntry> * retrieveSpreadsheetData(BasicExcel & source) {

	 vector<YarnEntry> * yarnEntries = new vector<YarnEntry>() ;

	 //We expect the list of transactions (the source) to be sheet at index 1 (the second sheet)
	 BasicExcelWorksheet * sourceSheet = source.GetWorksheet(size_t(1)) ;

	 //the actual data on the source sheet starts at row 2 (i.e. 3rd row, BasicExcel counts rows starting at 0)
	 for (unsigned row = 1 ; row < sourceSheet->GetTotalRows() ; row++) {
		 const char * code = sourceSheet->Cell(row, 2)->GetString()  ;
		 const char * color = sourceSheet->Cell(row, 3)->GetString() ;
		 const unsigned quantity = sourceSheet->Cell(row, 1)->GetInteger() ;
		 
		 if ((code == nullptr) || (color == nullptr)) {
			 cout << "Bad data on row" << row << " of input spreadsheet, skipping." << endl ;
			 continue ;
		 }

		 yarnEntries->push_back(YarnEntry(string(code), string(color), quantity)) ;
	 }

	 return yarnEntries ;
}

void copySpreadsheetData(vector<YarnEntry> * source, BasicExcel & destination) {

	//We expect the sheet we're outputting to be sheet at index 0 (the first sheet)
	BasicExcelWorksheet * destinationSheet = destination.GetWorksheet(size_t(0)) ;


	for (size_t i = 0 ; i < source->size() ; i++) {

		YarnEntry & currentEntry = source->at(i) ;
		bool currentEntryWrittenToCell = false ;

		//the first row in our destination sheet is "Absolute Magenta," at index 2
		for (unsigned row = 2 ; row < destinationSheet->GetTotalRows()  ; row++) {
			
			string rowName = getRowName(destinationSheet, row) ;
			
			if (currentEntry.color == rowName) {

				for (unsigned column = 2 ; column < destinationSheet->GetTotalCols()  ; column++) {
					
					if (colummExists(destinationSheet, column) == false) {
						break ;
					}
					
					string columnName = getColumnName(destinationSheet, column) ;
					
					if (currentEntry.code == columnName) {

						BasicExcelCell * destinationCell = destinationSheet->Cell(row, column) ;

						int currentCellValue = destinationCell->GetInteger() ;

						currentCellValue += currentEntry.quantity ;

						destinationCell->SetInteger(currentCellValue) ;
						
						currentEntryWrittenToCell = true ;
						break ;
					}

				}

			}
			if (currentEntryWrittenToCell) {
				break ;
			}

		}




	}


}

bool rowExists(BasicExcelWorksheet * sheet, unsigned rowNumber) {
	char * empty = new char ;
	bool exists =  sheet->Cell(rowNumber, 0)->Get(empty) ;
	delete empty ;
	return exists ;
}

bool colummExists(BasicExcelWorksheet * sheet, unsigned columnNumber) {
	char * empty = new char ;
	bool exists = sheet->Cell(0, columnNumber)->Get(empty) ;
	delete empty ;
	return exists ;
}



string getRowName(BasicExcelWorksheet * sheet, unsigned rowNumber) {
	string row = string(sheet->Cell(rowNumber, 0)->GetString()) ;
	return clean(row) ;
}
string getColumnName(BasicExcelWorksheet * sheet, unsigned columnNumber) {
	string column = string(sheet->Cell(0, columnNumber)->GetString()) ;
	return clean(column) ;
}





