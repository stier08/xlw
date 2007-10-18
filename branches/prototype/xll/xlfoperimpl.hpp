
#ifndef xlfoperimpl_hpp
#define xlfoperimpl_hpp

class XlfOper;

// Class XlfOperImpl
// Abstract base class encapsulating XlfOper XLOPER logic.
// Implemented as a polymorphic Singleton.  At startup, the core xlw code
// must determine under which version of Excel we're running, and instantiate
// the corresponding concrete derived class of XlfOperImpl.

class XlfOperImpl {
    static XlfOperImpl *instance_;
public:
    static const XlfOperImpl &instance() { return *instance_; }
    XlfOperImpl() { instance_ = this; }
    virtual long strlen(const XlfOper &xlfOper) const = 0;
};

#endif


