#include "SpreadsheetCopy.h"


void copySpreadsheetData(BasicExcel & source, BasicExcel & destination) {

	vector<const YarnEntry> * yarnEntries = retrieveSpreadsheetData(source) ;

	copySpreadsheetData(yarnEntries, destination) ;
}

vector<const YarnEntry> * retrieveSpreadsheetData(BasicExcel & source) {

	 vector<const YarnEntry> * yarnEntries = new vector<const YarnEntry>() ;

	 //We expect the list of transactions (the source) to be sheet at index 0 (the first sheet)
	 BasicExcelWorksheet * sourceSheet = source.GetWorksheet(size_t(0)) ;

	 //the actual data on the source sheet starts at row 2 (i.e. the 3rd row, BasicExcel counts rows starting at 0)
	 for (unsigned row = 1 ; row < sourceSheet->GetTotalRows() ; row++) {

		 string code = extractStringFromUnknownCell(sourceSheet->Cell(row, 2)) ;
		 string color = extractStringFromUnknownCell(sourceSheet->Cell(row, 3)) ;
		 const unsigned quantity = sourceSheet->Cell(row, 1)->GetInteger() ;
		 
		 const YarnEntry entry(code, color, quantity) ;
		 
		 if (entry.checkValidity() == false) {
			 cout << "Warning: bad data in cell(s) on row " << row + 1 << endl ;
			 continue ;
		 }

		 yarnEntries->push_back(entry) ;
	 }
	 
	 yarnEntries.sort(yarnEntries.begin(), yarnEntries.end()) ;

	 return yarnEntries ;
}

void copySpreadsheetData(vector<const YarnEntry> * source, BasicExcel & destination) {

	//We expect the sheet we're outputting to be sheet at index 0 (the first sheet)
	BasicExcelWorksheet * destinationSheet = destination.GetWorksheet(size_t(0)) ;

	
	for (size_t i = 0, row = 2 ; i < source->size() ; i++) {

		const YarnEntry & currentEntry = source->at(i) ;

		//check to make sure there's a labelled row here
		if (rowExists(destinationSheet, row)) {
			if (currentEntry.color != getRowName(destinationSheet, row)) { //then we need to go to the next row and label it
				row++ ;
				setRowName(destinationSheet, row, currentEntry.color) ;
			}
		}
		else { //otherwise label the row
			setRowName(destinationSheet, row, currentEntry.color) ;	
		}

		for (unsigned column = 2 ; column < destinationSheet->GetTotalCols()  ; column++) {

			//we've reached a column that hasn't been labelled yet, so name it and put our entry here
			if (colummExists(destinationSheet, column) == false) {
				setColumnName(destinationSheet, columnn, currentEntry.code) ;
			}
			
			//if this column name matches our entry, write to the cell in that column
			if (currentEntry.code == getColumnName(destinationSheet, column)) {
				BasicExcelCell * destinationCell = destinationSheet->Cell(row, column) ;

				int currentCellValue = destinationCell->GetInteger() ;

				currentCellValue += currentEntry.quantity ;

				destinationCell->SetInteger(currentCellValue) ;
				
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

void setRowName(BasicExcelWorksheet * sheet, unsigned rowNumber, string name) {
	BasicExcelCell * rowStart = sheet->Cell(rowNumber, 0) ;
	rowStart->SetString(name.c_str()) ;
}

void setColumnName(BasicExcelWorksheet * sheet, unsigned columnNumber, string name) {
	BasicExcelCell * columnStart = sheet->Cell(0, columnNumber) ;
	columnStart->SetString(name.c_str()) ;
}

string extractStringFromUnknownCell(BasicExcelCell * cell) {
	int type = cell->Type() ;
	
	string str = "";

	if (type == BasicExcelCell::WSTRING) {
		wstring tempWString = wstring(cell->GetWString()) ;
		str = convertToString(tempWString) ;
	}
	else if (type == BasicExcelCell::STRING) {
		str = string(cell->GetString()) ;
	}
	else {
		cerr << "Missing or non-string data in cell where string was expected."  << endl ;
	}
	return str ;
}
