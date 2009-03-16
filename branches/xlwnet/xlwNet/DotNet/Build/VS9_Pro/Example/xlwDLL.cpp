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
#include <xlw/ArgList.h>

namespace {
const char* LibraryName = "DLL";
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
EmptyArgFunctionArgs[]=
{
 { "","" } 
};
  XLRegistration::XLFunctionRegistrationHelper
registerEmptyArgFunction("xlEmptyArgFunction",
"EmptyArgFunction",
"tests empty args ",
LibraryName,
EmptyArgFunctionArgs,
0
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEmptyArgFunction(
)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

std::wstring result(
	EmptyArgFunction());
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
EchoShortArgs[]=
{
{ "x"," number to be echoed ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoShort("xlEchoShort",
"EchoShort",
"echoes a short ",
LibraryName,
EchoShortArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoShort(
LPXLFOPER xa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper xb(
	(xa));
short x(
	xb.AsShort("x"));

short result(
	EchoShort(
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
EchoMatArgs[]=
{
{ "Echoee"," argument to be echoed ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoMat("xlEchoMat",
"EchoMat",
"echoes a matrix ",
LibraryName,
EchoMatArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoMat(
LPXLFOPER Echoeea)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper Echoeeb(
	(Echoeea));
MyMatrix Echoee(
	Echoeeb.AsMatrix("Echoee"));

MyMatrix result(
	EchoMat(
		Echoee)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
EchoArrayArgs[]=
{
{ "Echoee","  argument to be echoed ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoArray("xlEchoArray",
"EchoArray",
"echoes an array ",
LibraryName,
EchoArrayArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoArray(
LPXLFOPER Echoeea)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper Echoeeb(
	(Echoeea));
MyArray Echoee(
	Echoeeb.AsArray("Echoee"));

MyArray result(
	EchoArray(
		Echoee)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
EchoCellsArgs[]=
{
{ "Echoee","  argument to be echoed ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoCells("xlEchoCells",
"EchoCells",
"echoes a  CellMatrix ",
LibraryName,
EchoCellsArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoCells(
LPXLFOPER Echoeea)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper Echoeeb(
	(Echoeea));
CellMatrix Echoee(
	Echoeeb.AsCellMatrix("Echoee"));

CellMatrix result(
	EchoCells(
		Echoee)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
CircArgs[]=
{
{ "Diameter"," the circle's diameter ","B"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerCirc("xlCirc",
"Circ",
"computes the circumference of a circle ",
LibraryName,
CircArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlCirc(
double Diameter)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);


double result(
	Circ(
		Diameter)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
ConcatArgs[]=
{
{ "str1"," first string ","XLW_WSTR"},
{ "str2"," second string ","XLW_WSTR"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerConcat("xlConcat",
"Concat",
"Concatenates two strings ",
LibraryName,
ConcatArgs,
2
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlConcat(
XLWSTR str1a,
XLWSTR str2a)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

std::wstring str1(
	voidToWstr(str1a));

std::wstring str2(
	voidToWstr(str2a));

std::wstring result(
	Concat(
		str1,
		str2)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
statsArgs[]=
{
{ "data"," input for computation ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerstats("xlstats",
"stats",
"computes mean and variance of a range ",
LibraryName,
statsArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlstats(
LPXLFOPER dataa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper datab(
	(dataa));
MyArray data(
	datab.AsArray("data"));

MyArray result(
	stats(
		data)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
ExceptionArgs[]=
{
 { "","" } 
};
  XLRegistration::XLFunctionRegistrationHelper
registerException("xlException",
"Exception",
"throws an exception ",
LibraryName,
ExceptionArgs,
0
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlException(
)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

double result(
	Exception());
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
HelloWorldAgainArgs[]=
{
{ "name"," name to be echoed ","XLW_WSTR"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerHelloWorldAgain("xlHelloWorldAgain",
"HelloWorldAgain",
"says hello name ",
LibraryName,
HelloWorldAgainArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlHelloWorldAgain(
XLWSTR namea)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

std::wstring name(
	voidToWstr(namea));

std::wstring result(
	HelloWorldAgain(
		name)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
EchoArgListArgs[]=
{
{ "args","  arguments to echo ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoArgList("xlEchoArgList",
"EchoArgList",
"echoes arg list ",
LibraryName,
EchoArgListArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoArgList(
LPXLFOPER argsa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper argsb(
	(argsa));
CellMatrix argsc(
	argsb.AsCellMatrix("argsc"));
ArgumentList args(
	ArgumentList(argsc,"args"));

CellMatrix result(
	EchoArgList(
		args)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
CastToCSArrayArgs[]=
{
{ "csarray","  double Array ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerCastToCSArray("xlCastToCSArray",
"CastToCSArray",
"takes double[] ",
LibraryName,
CastToCSArrayArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlCastToCSArray(
LPXLFOPER csarraya)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper csarrayb(
	(csarraya));
MyArray csarray(
	csarrayb.AsArray("csarray"));

MyArray result(
	CastToCSArray(
		csarray)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
CastToCSArrayTwiceArgs[]=
{
{ "csarray","  double Array ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerCastToCSArrayTwice("xlCastToCSArrayTwice",
"CastToCSArrayTwice",
"takes double[] ",
LibraryName,
CastToCSArrayTwiceArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlCastToCSArrayTwice(
LPXLFOPER csarraya)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper csarrayb(
	(csarraya));
MyArray csarray(
	csarrayb.AsArray("csarray"));

MyArray result(
	CastToCSArrayTwice(
		csarray)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
CastToCSMatrixArgs[]=
{
{ "csmatrix","  double Array ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerCastToCSMatrix("xlCastToCSMatrix",
"CastToCSMatrix",
"takes double[,] ",
LibraryName,
CastToCSMatrixArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlCastToCSMatrix(
LPXLFOPER csmatrixa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper csmatrixb(
	(csmatrixa));
MyMatrix csmatrix(
	csmatrixb.AsMatrix("csmatrix"));

MyMatrix result(
	CastToCSMatrix(
		csmatrix)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
CastToCSMatrixTwiceArgs[]=
{
{ "csmatrix","  double Array ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerCastToCSMatrixTwice("xlCastToCSMatrixTwice",
"CastToCSMatrixTwice",
"takes double[,] ",
LibraryName,
CastToCSMatrixTwiceArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlCastToCSMatrixTwice(
LPXLFOPER csmatrixa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper csmatrixb(
	(csmatrixa));
MyMatrix csmatrix(
	csmatrixb.AsMatrix("csmatrix"));

MyMatrix result(
	CastToCSMatrixTwice(
		csmatrix)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
throwStringArgs[]=
{
{ "err"," Just any random string ","XLW_WSTR"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerthrowString("xlthrowString",
"throwString",
"throws an exception of type ArgumentNullException ",
LibraryName,
throwStringArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlthrowString(
XLWSTR erra)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

std::wstring err(
	voidToWstr(erra));

double result(
	throwString(
		err)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
throwCellMatrixArgs[]=
{
{ "err"," Just any random string ","XLW_WSTR"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerthrowCellMatrix("xlthrowCellMatrix",
"throwCellMatrix",
"throws an exception of type cellMatrixException ",
LibraryName,
throwCellMatrixArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlthrowCellMatrix(
XLWSTR erra)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

std::wstring err(
	voidToWstr(erra));

double result(
	throwCellMatrix(
		err)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
throwCErrorArgs[]=
{
{ "err"," Just any random string ","XLW_WSTR"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerthrowCError("xlthrowCError",
"throwCError",
"makes the C Runtime throw an exception ",
LibraryName,
throwCErrorArgs,
1
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlthrowCError(
XLWSTR erra)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

std::wstring err(
	voidToWstr(erra));

double result(
	throwCError(
		err)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

