#include "SpreadsheetCopy.h"


void copySpreadsheetData(BasicExcel & source, BasicExcel & destination) {

	vector<YarnEntry> * yarnEntries = retrieveSpreadsheetData(source) ;

	copySpreadsheetData(yarnEntries, destination) ;
}

vector<YarnEntry> * retrieveSpreadsheetData(BasicExcel & source) {

	 vector<YarnEntry> * yarnEntries = new vector<YarnEntry>() ;

	 //We expect the list of transactions (the source) to be sheet at index 1 (the second sheet)
	 BasicExcelWorksheet * sourceSheet = source.GetWorksheet(size_t(1)) ;

	 //the actual data on the source sheet starts at row 2 (i.e. the 3rd row, BasicExcel counts rows starting at 0)
	 for (unsigned row = 1 ; row < sourceSheet->GetTotalRows() ; row++) {

		 string code = extractStringFromUnknownCell(sourceSheet->Cell(row, 2)) ;
		 string color = extractStringFromUnknownCell(sourceSheet->Cell(row, 3)) ;
		 const unsigned quantity = sourceSheet->Cell(row, 1)->GetInteger() ;

		 if ((code == nullptr) || (color == nullptr)) {
			 cout << "Bad data on row" << row << " of input spreadsheet, skipping." << endl ;
			 continue ;
		 }

		 yarnEntries->push_back(YarnEntry(code, color, quantity)) ;
	 }

	 return yarnEntries ;
}

void copySpreadsheetData(vector<YarnEntry> * source, BasicExcel & destination) {

	//We expect the sheet we're outputting to be sheet at index 0 (the first sheet)
	BasicExcelWorksheet * destinationSheet = destination.GetWorksheet(size_t(0)) ;


	for (size_t i = 0 ; i < source->size() ; i++) {

		YarnEntry & currentEntry = source->at(i) ;
		bool currentEntryWrittenToCell = false ;

		/* the first row in our destination sheet is currently "Absolute Magenta," at index 2
		 hopefully, if this program works as intended, adding new rows or columns will not cause
		 any issues
		 */
		for (unsigned row = 2 ; row < destinationSheet->GetTotalRows()  ; row++) {

			if (rowExists(destinationSheet, row) == false) {
				continue ; //then skip this row
			}

			string rowName = getRowName(destinationSheet, row) ;

			if (currentEntry.color == rowName) {

				for (unsigned column = 2 ; column < destinationSheet->GetTotalCols()  ; column++) {

					if (colummExists(destinationSheet, column) == false) {
						continue ; //then skip this column
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
	BasicExcelCell * rowStart  =  sheet->Cell(rowNumber, 0) ;
	int type = rowStart->Type() ;
	return (type != BasicExcelCell::UNDEFINED) ;
}

bool colummExists(BasicExcelWorksheet * sheet, unsigned columnNumber) {
	BasicExcelCell * columnStart = sheet->Cell(0, columnNumber) ;
	int type = columnStart->Type() ;
	return (type != BasicExcelCell::UNDEFINED) ;
}



string getRowName(BasicExcelWorksheet * sheet, unsigned rowNumber) {
	BasicExcelCell * rowStart = sheet->Cell(rowNumber, 0) ;
	string row = extractStringFromUnknownCell(rowStart) ;
	return clean(row) ;
}

string getColumnName(BasicExcelWorksheet * sheet, unsigned columnNumber) {
	BasicExcelCell * columnStart = sheet->Cell(0, columnNumber) ;
	string column = extractStringFromUnknownCell(columnStart) ;
	return clean(column) ;
}

string extractStringFromUnknownCell(BasicExcelCell * cell) {
	int type = cell->Type() ;

	if (type == BasicExcelCell::WSTRING) {
		wstring tempWString = wstring(cell->GetWString()) ;
		row = convertToString(tempWString) ;
	}
	else if (type == BasicExcelCell::STRING) {
		row = string(cell->GetString()) ;
	}
	else {
		cerr << "Non-string data in cell where string was expected."  << endl ;
		throw exception() ;
	}
}
