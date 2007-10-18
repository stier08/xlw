
#ifndef xlfoper_hpp
#define xlfoper_hpp

#include <windows.h>
#include <xlcall.h>

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
    // XlfOper default ctor declared but not implemented.  This emphasizes 
    // the fact that when Excel passes LPXLOPER values to an xlw Addin function 
    // with parms of type XlfOper, the XlfOper ctor is not called.
    XlfOper();
public:
    long strlen() const;
};

#endif

