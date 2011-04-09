
#ifndef NATIVE_H
#define NATIVE_H

#include <xlw/MyContainers.h>
#include <xlw/CellMatrix.h>
#include <xlw/DoubleOrNothing.h>
#include <xlw/ArgList.h>
#include <string>

using namespace xlw;

//<xlw:libraryname=cpp_clr_Template

CellMatrix // Obtains historial market data from yahoo 
GetHistoricDataFromYahoo(
	                      std::string  symbol // Yahoo Symbol 
                         ,double beginDate // Begin Date
                         ,double endDate //End Date
						 );





#endif
