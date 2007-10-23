
#include <xlfoperimpl12.hpp>
#include <xlfoper.hpp>
#include <framewrk.h>

LPXLFOPER XlfOperImpl12::as_LPXLFOPER(const XlfOper &xlfOper) const {
    return xlfOper.lpxloper12_;
}

LPXLOPER XlfOperImpl12::as_LPXLOPER(const XlfOper &xlfOper) const {
    throw("operation not implemented");
}

LPXLOPER12 XlfOperImpl12::as_LPXLOPER12(const XlfOper &xlfOper) const {
    return xlfOper.lpxloper12_;
}

XlfOper XlfOperImpl12::echo(const XlfOper &xlfOper) const {

    static XLOPER12 x;
    Excel12f(xlCoerce, &x, 1, xlfOper.lpxloper12_);
    x.xltype |= xlbitXLFree;
    XlfOper ret(&x);
    return ret;

}
