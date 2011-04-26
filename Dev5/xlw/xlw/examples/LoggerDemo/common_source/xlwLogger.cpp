/*

 Copyright (C) 2011 Narinder S Claire

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#ifdef WIN32
#include "xlwLogger.h"

#include<ctime>

xlwLogger::xlwLogger(){
	theConsoleHandel= GetConsoleWindow();
	if(theConsoleHandel)FreeConsole();

	BOOL flag = AllocConsole();
    theConsoleHandel= GetConsoleWindow();
	theScreenHandel = GetStdHandle(STD_OUTPUT_HANDLE);
	char titlePtr[]=" xlwLogger Window \0";
	
	SetConsoleTitle(titlePtr);

	time(&theTime[0]);
	theInnerStream << " xlwLogger Initiated    " 
				   << ctime(&theTime[0])
				   << " \n Code Compiled " 
				   << __TIME__  << "  "
				   << __DATE__;
	Display();
	
	theTimeIndex=0;
	
}


void xlwLogger::Display(){
	WriteConsole(theScreenHandel,
				 theInnerStream.str().c_str(),
				 theInnerStream.str().size(),
				 &CharsWritten, 0);
	theInnerStream.str(EmptyString);
	
}
void xlwLogger::Display(const std::string& theLog ){
	WriteConsole(theScreenHandel,
				 theLog.c_str(),
				 theLog.size(),
				 &CharsWritten, 0);
}
double xlwLogger::GetTau() {
	theTimeIndex = 1-theTimeIndex;
	time(&theTime[theTimeIndex]);
	return (theTime[theTimeIndex]-theTime[1-theTimeIndex]);
}

const char* const xlwLogger::GetTime() {
	time_t tem(time(&theTime[2]));
	char * timePtr = asctime(gmtime(&tem));
	for(unsigned long i(0);i<8;i++)timeTemp[i]=char(timePtr[i+11]);
	timeTemp[8]=0;
	return timeTemp;
}


#endif
