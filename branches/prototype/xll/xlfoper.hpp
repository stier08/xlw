
#ifndef xlfoper_hpp
#define xlfoper_hpp

#include <xlfoperimpl.hpp>

class XlfOperImpl4;
class XlfOperImpl12;

// Class XlfOper
// Can share the memory footprint with either LPXLOPER or LPXLOPER12.
// Relies on XlfOperImpl to switch between the two at runtime.

class XlfOper {
    friend class XlfOperImpl4;
    friend class XlfOperImpl12;
    union {
        LPXLOPER lpxloper4_;
        LPXLOPER12 lpxloper12_;
    };
public:
    XlfOper(LPXLFOPER lpxloper) : lpxloper4_(reinterpret_cast<LPXLOPER>(lpxloper)) {}
    XlfOper(LPXLOPER lpxloper4) : lpxloper4_(lpxloper4) {}
    XlfOper(LPXLOPER12 lpxloper12) : lpxloper12_(lpxloper12) {}
    XlfOper echo() const { return XlfOperImpl::instance().echo(*this); }
    operator LPXLFOPER() const { return XlfOperImpl::instance().as_LPXLFOPER(*this); }
    operator LPXLOPER() const { return XlfOperImpl::instance().as_LPXLOPER(*this); }
    operator LPXLOPER12() const { return XlfOperImpl::instance().as_LPXLOPER12(*this); }
};

#endif

