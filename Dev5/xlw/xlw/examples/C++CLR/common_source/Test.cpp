#include "cppinterface.h"
using namespace System;



double // Echoes Date
EchoDate(
		 DateTime date // the Date
		 )
{
	return date.ToOADate();
}