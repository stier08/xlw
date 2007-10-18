
#include <xlfoperimpl4.hpp>
#include <xlfoper.hpp>
#include <framewrk.h>

long XlfOperImpl4::strlen(const XlfOper &xlfOper) const {

    XLOPER xStr;
    long ret;
    Excel(xlCoerce, &xStr, 2, xlfOper.lpxloper4_, TempInt(xltypeStr));
    if (xStr.xltype == xltypeStr) {
        ret = (BYTE)xStr.val.str[0];
        Excel(xlFree, 0, 1, &xStr);
    } else {
        ret = 0;
    }
    return ret;

}
