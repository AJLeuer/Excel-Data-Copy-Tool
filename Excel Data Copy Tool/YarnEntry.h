
#pragma once

#include <iostream>
#include <wstring>

struct YarnEntry {

	wstring twoDigitCode ;

	wstring color ;

	unsigned quantity ;

	YarnEntry(wstring twoDigitCode, wstring color, unsigned quantity) :
		twoDigitCode(twoDigitCode), color(color), quantity(quantity) {}

}
