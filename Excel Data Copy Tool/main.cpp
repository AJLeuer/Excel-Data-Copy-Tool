
#include <iostream>
#include <string>

#include <unistd.h>

#include "SpreadsheetCopy.h"

using namespace YExcel ;
using namespace std;


int main(int argc, const char *argv[], const char *env[], const char *path[]) {
	
	string currentWorkingDirectory = string(getcwd(nullptr, 1024));
	
	string inputEnclosingDirectoryFilePath ;
	string inputFileName ;
	string inputFilePath ;
	
	string outputEnclosingDirectoryFilePath ;
	string outputFileName ;
	string outputFilePath ;
	
	cout << "Current directory: " << currentWorkingDirectory << '\n' << endl ;
	
	cout << "Please note: this program is only compatible with Microsoft Excel 97-2004 (.xls) files. Other formats, such " <<
		"as the more recent Office Open XML Workbook (.xlsx) file type, are not supported." << '\n' << endl ;
	
	cout << "Please enter the path to the folder containing the Excel (.xls) file to be used as input:" << '\n' << endl ;
	getline(cin, inputEnclosingDirectoryFilePath) ; //currentWorkingDirectory + "/Input.xls" ;
	
	cout << endl ;
	cout << "Please enter the name (including extension) of the Excel (.xls) file to be used as input:" << '\n' << endl ;
	getline(cin, inputFileName) ;
	
	cout << endl ;
	cout << "Please enter the path to the folder containing the Excel (.xls) file to be used as output:" << '\n' << endl ;
	getline(cin, outputEnclosingDirectoryFilePath) ; //currentWorkingDirectory + "/Output.xls" ;
	
	cout << endl ;
	cout << "Please enter the name (including extension) of the Excel (.xls) file to be used as output:" << '\n' << endl ;
	getline(cin, outputFileName) ;
	
	cout << endl ;
	cout << "Working..." << endl ;
	
	inputFilePath = inputEnclosingDirectoryFilePath + '/' + inputFileName ;
	outputFilePath = outputEnclosingDirectoryFilePath + '/' + outputFileName ;
	
	//open source and destination workbooks
	BasicExcel sourceBook(inputFilePath.c_str()) ;
	BasicExcel destinationBook(outputFilePath.c_str()) ;

	copySpreadsheetData(sourceBook, destinationBook) ;
	
	destinationBook.Save() ;
	
	cout << "Data transfer complete." << endl ;

	return 0 ;
}


/*
	Todo: check if value in each cell is string or wstring, and match accordingly.
	Use int BasicExcelCell::Type() const to check
*/