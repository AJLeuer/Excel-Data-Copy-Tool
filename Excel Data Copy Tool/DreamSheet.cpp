
#include "DreamSheet.h"



wstring DreamSheet::getRowName(unsigned rowNumber) {

	const wchar_t * rowName = readStr(rowNumber, 0) ;
	return wstring(rowName) ;
}


wstring DreamSheet::getColumnName(unsigned columnNumber) {
	const wchar_t * columnName = readStr(0, columnNumber) ;
	return wstring(columnName) ;
}
