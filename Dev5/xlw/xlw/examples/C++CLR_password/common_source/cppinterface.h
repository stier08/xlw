
#ifndef NATIVE_H
#define NATIVE_H
            
#include <xlw/MyContainers.h>
#include <xlw/CellMatrix.h>
#include <xlw/DoubleOrNothing.h>
#include <xlw/ArgList.h>
#include <string>

using namespace System;
             
using namespace xlw;
  
//<xlw:libraryname=password_example

//<xlw:onopen(authenticate)

void // Authenticate
authenticate();

double // Echoes Date
//<xlw:time
EchoDate(
		 DateTime date // the Date
		 );


#endif
