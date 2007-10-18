
#include <xlfoperimpl12.hpp>
#include <xlfoper.hpp>
#include <framewrk.h>

long XlfOperImpl12::strlen(const XlfOper &xlfOper) const {

    XLOPER12 xStr;
    Excel12f(xlCoerce, &xStr, 2, xlfOper.lpxloper12_, TempInt12(xltypeStr));
    long ret = xStr.val.str[0];
    Excel12f(xlFree, 0, 1, &xStr);
    return ret;

}

