#include "cppinterface.h"
using namespace System;


bool IsAuthenticated();
void authenticate();

void checkAuthenticated()
{
	if(!IsAuthenticated())
	{
		throw("NOT AUTHENTICATED TO USE THIS XLL");
	}
}

double // Echoes Date
EchoDate(
		 DateTime date // the Date
		 )
{
	checkAuthenticated();

	return date.ToOADate();
}

