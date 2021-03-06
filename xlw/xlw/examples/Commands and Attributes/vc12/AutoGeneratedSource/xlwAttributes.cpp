//// 
//// Autogenerated by xlw 
//// Do not edit this file, it will be overwritten 
//// by InterfaceGenerator 
////

#include "xlw/MyContainers.h"
#include <xlw/CellMatrix.h>
#include "Attributes.h"
#include <xlw/xlw.h>
#include <xlw/XlFunctionRegistration.h>
#include <stdexcept>
#include <xlw/XlOpenClose.h>
#include <xlw/HiResTimer.h>
#include <xlw/ArgList.h>

#include <xlw/xlarray.h>

using namespace xlw;

namespace {
const char* LibraryName = "MyTestLibrary";
};


// registrations start here


namespace
{
  XLRegistration::XLFunctionRegistrationHelper
registerEmptyArgFunction("xlEmptyArgFunction",
"EmptyArgFunction",
" tests empty args ",
LibraryName,
0,
0
,false
,true
,""
,""
,false
,false
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

std::string result(
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
" echoes a short ",
LibraryName,
EchoShortArgs,
1
,false
,true
,""
,""
,false
,false
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
" echoes a matrix ",
LibraryName,
EchoMatArgs,
1
,false
,true
,""
,""
,false
,false
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
EchoMatrixArgs[]=
{
{ "Echoee"," argument to be echoed ","XLW_FP"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoMatrix("xlEchoMatrix",
"EchoMatrix",
" echoes a matrix ",
LibraryName,
EchoMatrixArgs,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoMatrix(
LPXLARRAY Echoeea)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

NEMatrix Echoee(
	GetMatrix(Echoeea));

MyMatrix result(
	EchoMatrix(
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
{ "Echoee"," argument to be echoed ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoArray("xlEchoArray",
"EchoArray",
" echoes an array ",
LibraryName,
EchoArrayArgs,
1
,false
,true
,""
,""
,false
,false
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
{ "Echoee"," argument to be echoed ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoCells("xlEchoCells",
"EchoCells",
" echoes a cell matrix ",
LibraryName,
EchoCellsArgs,
1
,false
,true
,""
,""
,false
,false
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

 HiResTimer t;
CellMatrix result(
	EchoCells(
		Echoee)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
CircArgs[]=
{
{ "Diameter","the circle's diameter ","B"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerCirc("xlCirc",
"Circ",
" computes the circumference of a circle ",
LibraryName,
CircArgs,
1
,false
,true
,""
,""
,false
,false
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


 HiResTimer t;
double result(
	Circ(
		Diameter)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
ConcatArgs[]=
{
{ "str1"," first string ","XLF_OPER"},
{ "str2","second string ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerConcat("xlConcat",
"Concat",
" Concatenates two strings ",
LibraryName,
ConcatArgs,
2
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlConcat(
LPXLFOPER str1a,
LPXLFOPER str2a)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper str1b(
	(str1a));
std::wstring str1(
	str1b.AsWstring("str1"));

XlfOper str2b(
	(str2a));
std::wstring str2(
	str2b.AsWstring("str2"));

 HiResTimer t;
std::wstring result(
	Concat(
		str1,
		str2)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
StatsArgs[]=
{
{ "data"," input for computation ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerStats("xlStats",
"Stats",
" computes mean and variance of a range ",
LibraryName,
StatsArgs,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlStats(
LPXLFOPER dataa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper datab(
	(dataa));
MyArray data(
	datab.AsArray("data"));

 HiResTimer t;
MyArray result(
	Stats(
		data)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
HelloWorldAgainArgs[]=
{
{ "name"," name to be echoed ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerHelloWorldAgain("xlHelloWorldAgain",
"HelloWorldAgain",
" says hello name ",
LibraryName,
HelloWorldAgainArgs,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlHelloWorldAgain(
LPXLFOPER namea)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper nameb(
	(namea));
std::string name(
	nameb.AsString("name"));

 HiResTimer t;
std::string result(
	HelloWorldAgain(
		name)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
EchoULArgs[]=
{
{ "b"," number to echo ","B"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoUL("xlEchoUL",
"EchoUL",
" echoes an unsigned int ",
LibraryName,
EchoULArgs,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoUL(
double ba)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

unsigned long b(
	static_cast<unsigned long>(ba));

unsigned long result(
	EchoUL(
		b)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
EchoIntArgs[]=
{
{ "b"," number to echo ","B"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoInt("xlEchoInt",
"EchoInt",
" echoes an int ",
LibraryName,
EchoIntArgs,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoInt(
double ba)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

int b(
	static_cast<int>(ba));

int result(
	EchoInt(
		b)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
EchoDoubleOrNothingArgs[]=
{
{ "x"," value to specify ","XLF_OPER"},
{ "defaultValue"," value to use if not specified ","B"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoDoubleOrNothing("xlEchoDoubleOrNothing",
"EchoDoubleOrNothing",
" tests DoubleOrNothingType ",
LibraryName,
EchoDoubleOrNothingArgs,
2
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlEchoDoubleOrNothing(
LPXLFOPER xa,
double defaultValue)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper xb(
	(xa));
CellMatrix xc(
	xb.AsCellMatrix("xc"));
DoubleOrNothing x(
	DoubleOrNothing(xc,"x"));


 HiResTimer t;
double result(
	EchoDoubleOrNothing(
		x,
		defaultValue)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
EchoArgListArgs[]=
{
{ "args"," arguments to echo ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerEchoArgList("xlEchoArgList",
"EchoArgList",
" echoes arg list ",
LibraryName,
EchoArgListArgs,
1
,false
,true
,""
,""
,false
,false
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

 HiResTimer t;
CellMatrix result(
	EchoArgList(
		args)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
SystemTimeArgs[]=
{
{ "ticksPerSecond"," number to divide by ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerSystemTime("xlSystemTime",
"SystemTime",
" system clock ",
LibraryName,
SystemTimeArgs,
1
,true
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlSystemTime(
LPXLFOPER ticksPerSeconda)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper ticksPerSecondb(
	(ticksPerSeconda));
CellMatrix ticksPerSecondc(
	ticksPerSecondb.AsCellMatrix("ticksPerSecondc"));
DoubleOrNothing ticksPerSecond(
	DoubleOrNothing(ticksPerSecondc,"ticksPerSecond"));

 HiResTimer t;
double result(
	SystemTime(
		ticksPerSecond)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
ContainsErrorArgs[]=
{
{ "input2"," data to check for errors ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerContainsError("xlContainsError",
"ContainsError",
" checks to see if there's an error ",
LibraryName,
ContainsErrorArgs,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlContainsError(
LPXLFOPER input2a)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper input2b(
	(input2a));
CellMatrix input2(
	input2b.AsCellMatrix("input2"));

bool result(
	ContainsError(
		input2)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
ContainsDivByZeroArgs[]=
{
{ "input"," data to check for errors ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registerContainsDivByZero("xlContainsDivByZero",
"ContainsDivByZero",
" checks to see if there's a div by zero ",
LibraryName,
ContainsDivByZeroArgs,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlContainsDivByZero(
LPXLFOPER inputa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper inputb(
	(inputa));
CellMatrix input(
	inputb.AsCellMatrix("input"));

bool result(
	ContainsDivByZero(
		input)
	);
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
  XLRegistration::XLFunctionRegistrationHelper
registerGetThreadId("xlGetThreadId",
"GetThreadId",
" returns ID of current execution thread ",
LibraryName,
0,
0
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlGetThreadId(
)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

double result(
	GetThreadId());
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
typeStringArgs[]=
{
{ "input"," value on which to perform type check ","XLF_OPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registertypeString("xltypeString",
"typeString",
" return a string indicating datatype of OPER/OPER12 input ",
LibraryName,
typeStringArgs,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xltypeString(
LPXLFOPER inputa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

XlfOper input(
	(inputa));

 HiResTimer t;
std::string result(
	typeString(
		input)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
XLRegistration::Arg
typeString2Args[]=
{
{ "input"," value on which to perform type check ","XLF_XLOPER"}
};
  XLRegistration::XLFunctionRegistrationHelper
registertypeString2("xltypeString2",
"typeString2",
" return a string indicating datatype of XLOPER/XLOPER12 input ",
LibraryName,
typeString2Args,
1
,false
,true
,""
,""
,false
,false
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xltypeString2(
LPXLFOPER inputa)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

reftest input(
	(inputa));

 HiResTimer t;
std::string result(
	typeString2(
		input)
	);
CellMatrix resultCells(result);
CellMatrix time(1,2);
time(0,0) = "time taken";
time(0,1) = t.elapsed();
resultCells.PushBottom(time);
return XlfOper(resultCells);
EXCEL_END
}
}



//////////////////////////

namespace
{
  XLRegistration::XLFunctionRegistrationHelper
registerGetNote("xlGetNote",
"GetNote",
" the text of the note attached to the calling cell ",
LibraryName,
0,
0
,false
,false
,""
,""
,false
,true
,false
);
}



extern "C"
{
LPXLFOPER EXCEL_EXPORT
xlGetNote(
)
{
EXCEL_BEGIN;

	if (XlfExcel::Instance().IsCalledByFuncWiz())
		return XlfOper(true);

std::string result(
	GetNote());
return XlfOper(result);
EXCEL_END
}
}



//////////////////////////

//////////////////////////
// Methods that will get registered to execute in AutoOpen
//////////////////////////

//////////////////////////
// Methods that will get registered to execute in AutoClose
//////////////////////////

