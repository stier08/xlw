#include <xlfoper12.hpp>
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>

LPXLOPER12 XlfOper12::echo() const {

    static XLOPER12 ret;
    Excel12f(xlCoerce, &ret, 1, lpxloper12_);
    ret.xltype |= xlbitXLFree;
    return &ret;

}
