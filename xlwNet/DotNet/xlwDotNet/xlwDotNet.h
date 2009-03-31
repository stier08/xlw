#ifndef XLW_DOT_NET_H
#define XLW_DOT_NET_H

// Shouldn't really need this file, but I think theres an 
// oversigt by the xlw people.

#include<xlw/port.h>
#include<xlw/MyContainers.h>
#include<xlw/CellMatrix.h>
#include<xlw/ArgList.h>
#include<string>
#include<xlw/XlfException.h>



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
