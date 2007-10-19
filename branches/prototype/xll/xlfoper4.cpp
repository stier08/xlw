#include <xlfoper4.hpp>
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>

LPXLOPER XlfOper4::echo() const {

    static XLOPER ret;
    Excel(xlCoerce, &ret, 1, lpxloper4_);
    ret.xltype |= xlbitXLFree;
    return &ret;

}
