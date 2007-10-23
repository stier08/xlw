
#ifndef xlfoperimpl12_hpp
#define xlfoperimpl12_hpp

#include <xlfoperimpl.hpp>

class XlfOperImpl12 : public XlfOperImpl {
    virtual LPXLFOPER as_LPXLFOPER(const XlfOper &xlfOper) const;
    virtual LPXLOPER as_LPXLOPER(const XlfOper &xlfOper) const;
    virtual LPXLOPER12 as_LPXLOPER12(const XlfOper &xlfOper) const;
    virtual XlfOper echo(const XlfOper &xlfOper) const;
};

#endif


