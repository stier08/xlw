//// 
//// created by xlwplus
////

#include <xlw/pragmas.h>
#include <xlw/MyContainers.h>
#include <xlw/CellMatrix.h>
#include "DLL.h"
#include <xlw/xlw.h>
#include <xlw/XlFunctionRegistration.h>
#include <stdexcept>
#include <xlw/XlOpenClose.h>
#include <ctime>
namespace {
const char* LibraryName = "NonPassive";
};

// dummy function to force linkage
namespace {
void DummyFunction()
{
xlAutoOpen();
xlAutoClose();
}
}

// registrations start here


namespace
{
XLRegistration::Arg
MyAddressArgs[]=
{
{ "x"," x ","B"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerMyAddress("xlMyAddress",
"MyAddress",
"Gets the Address of the calling cell ",
LibraryName,
MyAddressArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlMyAddress(
double x)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);


std::wstring result(
	MyAddress(
		x)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
MessageInStatusBarArgs[]=
{
{ "Message"," The String ","XLW_WSTR"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerMessageInStatusBar("xlMessageInStatusBar",
"MessageInStatusBar",
"Writes string to status bar ",
LibraryName,
MessageInStatusBarArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlMessageInStatusBar(
XLWSTR Messagea)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

std::wstring Message(
	voidToWstr(Messagea));

std::wstring result(
	MessageInStatusBar(
		Message)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
GetPidArgs[]=
{
 { "","" } 
};
  XLRegistration::XLFunctionRegistrationHelper
registerGetPid("xlGetPid",
"GetPid",
" ",
LibraryName,
GetPidArgs,
0
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlGetPid(
)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

double result(
	GetPid());
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////
