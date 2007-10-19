#include <fstream>
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>
#include <xlfoperimpl4.hpp>
#include <xlfoperimpl12.hpp>

#define DLLEXPORT extern "C" __declspec(dllexport)

bool excel12() {
    XLOPER xRet1, xRet2;
    Excel(xlfGetWorkspace, &xRet1, 1, TempInt(2));
    Excel(xlCoerce, &xRet2, 2, &xRet1, TempInt(xltypeInt));
    Excel(xlFree, 0, 1, &xRet1);
    //Excel(xlcAlert, 0, 1, &xRet2);
    return (xRet2.val.w == 12);
}

DLLEXPORT int xlAutoOpen() {
    XLOPER xDll;
    Excel(xlGetName, &xDll, 0);

    if (excel12()) {

        static XlfOperImpl12 xlfOperImpl12;

        Excel(xlfRegister, 0, 7, &xDll,
            TempStrConst("echo"),       // function code name
            TempStrConst("QQ"),         // parameter codes
            TempStrConst("echo"),       // function display name
            TempStrConst(""),           // comma-delimited list of parameters
            TempStrConst("1"),          // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
            TempStrConst("PROTOTYPE")); // function category

        Excel(xlfRegister, 0, 7, &xDll,
            TempStrConst("echo4"),      // function code name
            TempStrConst("PP"),         // parameter codes
            TempStrConst("echo4"),      // function display name
            TempStrConst(""),           // comma-delimited list of parameters
            TempStrConst("1"),          // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
            TempStrConst("PROTOTYPE")); // function category

        Excel(xlfRegister, 0, 7, &xDll,
            TempStrConst("echo12"),     // function code name
            TempStrConst("QQ"),         // parameter codes
            TempStrConst("echo12"),     // function display name
            TempStrConst(""),           // comma-delimited list of parameters
            TempStrConst("1"),          // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
            TempStrConst("PROTOTYPE")); // function category

    } else {

        static XlfOperImpl4 xlfOperImpl4;

        Excel(xlfRegister, 0, 7, &xDll,
            TempStrConst("echo"),       // function code name
            TempStrConst("PP"),         // parameter codes
            TempStrConst("echo"),       // function display name
            TempStrConst(""),           // comma-delimited list of parameters
            TempStrConst("1"),          // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
            TempStrConst("PROTOTYPE")); // function category

        Excel(xlfRegister, 0, 7, &xDll,
            TempStrConst("echo4"),      // function code name
            TempStrConst("PP"),         // parameter codes
            TempStrConst("echo4"),      // function display name
            TempStrConst(""),           // comma-delimited list of parameters
            TempStrConst("1"),          // function type (0 = hidden function, 1 = worksheet function, 2 = command macro)
            TempStrConst("PROTOTYPE")); // function category

    }
                                
    Excel(xlFree, 0, 1, &xDll);
    return 1;
}
