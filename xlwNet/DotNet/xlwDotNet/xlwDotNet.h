/*
 Copyright (C) 2008 Narinder Claire

 This file is part of XLW.NET 

 XLW.NET is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#ifndef XLW_DOT_NET_H
#define XLW_DOT_NET_H



#include<xlw/port.h>
#include<xlw/MyContainers.h>
#include<xlw/CellMatrix.h>
#include<xlw/ArgList.h>
#include<string>
#include <xlw/XlfException.h>

#define DLLEXPORT __declspec(dllexport)




namespace {
	std::string errMessage;
	xlw::CellMatrix errCells;
}

#define DOT_NET_EXCEL_BEGIN  try {
#define DOT_NET_EXCEL_END \
} catch (xlwDotNet::cellMatrixException^ theError ) { \
   errCells = *(xlw::CellMatrix*)((theError->inner).ToPointer());\
   throw(errCells);\
} catch (XlfException& theError) { \
    throw(theError); \
} catch (std::runtime_error& theError){\
    throw(theError); \
} catch (std::string& theError){\
    throw(theError); \
} catch (const char* theError){\
    throw(theError); \
} catch (System::Exception^ theError ) { \
    String^ ErrorType = theError->GetType()->ToString() + " : ";\
	String^ ErrorMessage = theError->Message;\
	errMessage = (const char*)(Marshal::StringToHGlobalAnsi(ErrorType+ErrorMessage).ToPointer());\
    throw(errMessage.c_str());\
} 


using namespace xlw;

#endif
