//
//  Util.hpp
//  Excel Data Copy Tool
//
//  Created by Adam James Leuer on 1/6/16.
//  Copyright (c) 2016 Adam James Leuer. All rights reserved.
//

#ifndef Excel_Data_Copy_Tool_Util_hpp
#define Excel_Data_Copy_Tool_Util_hpp

#include <iostream>
#include <string>
#include <locale>
#include <codecvt>

using namespace std ;

static wstring_convert<std::codecvt_utf8<wchar_t>> stringConverter ;

/**
 * Converts a std::wstring to a std::string
 */
static string convertToString(const wstring & wide_string) {
	return stringConverter.to_bytes(wide_string) ;
}

/**
 * Removes all whitespace in str
 * Code credit stackoverflow
 */
static void removeWhitespace(string & str) {
	 str.erase(std::remove (str.begin(), str.end(), ' '), str.end()) ;
}

/**
 * Replaces all upper case chars with lower case equivalents
 */
static void toLowerCase(string & str) {
	std::transform(str.begin(), str.end(), str.begin(), ::tolower) ;
}

/**
 * Calls removeWhitespace() and toLowerCase() on str, and returns
 * a reference to it
 */
static string & clean(string & str) {
	removeWhitespace(str) ;
	toLowerCase(str) ;

	return str ;
}


#endif
