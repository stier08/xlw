
using namespace System;
using namespace Runtime::InteropServices;

#include<string>

std::string CLR2CPP(String^ clrString)
{
	std::string result =  (const char*)(Marshal::StringToHGlobalAnsi(clrString)).ToPointer();
	return result;
}
