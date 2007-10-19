
#ifndef xlfoper4_hpp
#define xlfoper4_hpp

#include <windows.h>
#include <xlcall.h>

// Class XlfOper4
// Can share the memory footprint with LPXLOPER.

class XlfOper4 {
    LPXLOPER lpxloper4_;
public:
    LPXLOPER echo() const;
};

#endif

