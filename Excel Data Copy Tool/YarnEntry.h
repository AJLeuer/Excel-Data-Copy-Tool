
#pragma once

#include <iostream>
#include <string>

#include "Util.hpp"

struct YarnEntry {

	string code ;

	string color ;

	unsigned quantity ;

	YarnEntry(string code, string color, unsigned quantity) :
		code(clean(code)), color(clean(color)), quantity(quantity) {}

} ;
