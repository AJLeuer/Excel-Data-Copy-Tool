
#include <iostream>
#include <wstring>

struct YarnEntry {

	unsigned quantity ;

	wstring type ;

	wstring color ;

	YarnEntry(unsigned quantity, wstring type, wstring color) :
		quantity(quantity), type(type), color(color) {}

}
	
	