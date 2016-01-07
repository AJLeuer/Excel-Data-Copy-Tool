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

using namespace std ;

/* Code credit stackoverflow
 */
static void removeWhitespace(string & str) {
	 str.erase(std::remove (str.begin(), str.end(), ' '), str.end()) ;
}

static void toLowerCase(string & str) {
	std::transform(str.begin(), str.end(), str.begin(), ::tolower) ;
}

static string & clean(string & str) {
	removeWhitespace(str) ;
	toLowerCase(str) ;
	
	return str ;
}


#endif
