
#include <xlfoperimpl4.hpp>
#include <xlfoper.hpp>
#include <framewrk.h>

LPXLFOPER XlfOperImpl4::as_LPXLFOPER(const XlfOper &xlfOper) const {
    return xlfOper.lpxloper4_;
}

LPXLOPER XlfOperImpl4::as_LPXLOPER(const XlfOper &xlfOper) const {
    return xlfOper.lpxloper4_;
}

LPXLOPER12 XlfOperImpl4::as_LPXLOPER12(const XlfOper &xlfOper) const {
    throw("operation not implemented");
}

XlfOper XlfOperImpl4::echo(const XlfOper &xlfOper) const {

    static XLOPER x;
    Excel(xlCoerce, &x, 1, xlfOper.lpxloper4_);
    x.xltype |= xlbitXLFree;
    XlfOper ret(&x);
    return ret;

}

