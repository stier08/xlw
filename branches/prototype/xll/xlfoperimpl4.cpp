
#include <xlfoperimpl4.hpp>
#include <xlfoper.hpp>
#include <framewrk.h>

long XlfOperImpl4::strlen(const XlfOper &xlfOper) const {

    XLOPER xStr;
    Excel(xlCoerce, &xStr, 2, xlfOper.lpxloper4_, TempInt(xltypeStr));
    long ret = xStr.val.str[0];
    Excel(xlFree, 0, 1, &xStr);
    return ret;

}
