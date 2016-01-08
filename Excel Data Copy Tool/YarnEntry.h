
#pragma once

#include <iostream>
#include <string>

#include "Util.hpp"

struct YarnEntry {

	string code = "" ;

	string color = "" ;

	unsigned quantity = 0 ;

	YarnEntry(string code, string color, unsigned quantity) :
		code(clean(code)), color(clean(color)), quantity(quantity) {}
	
	/**
	 * Returns true if all data members are initialized with non-default values,
	 * false otherwise
	 */
	bool checkValidity() const { return ((code != "") && (color != "") && (quantity != 0)) ; }

} ;
