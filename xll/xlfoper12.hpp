
#ifndef xlfoper12_hpp
#define xlfoper12_hpp

#include <windows.h>
#include <xlcall.h>

// Class XlfOper12
// Can share the memory footprint with LPXLOPER12.

class XlfOper12 {
    LPXLOPER12 lpxloper12_;
public:
    LPXLOPER12 echo() const;
};

#endif

