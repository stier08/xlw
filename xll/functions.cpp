
#include <xlfoper.hpp>

#define DLLEXPORT extern "C" __declspec(dllexport)

DLLEXPORT long *STRLEN(XlfOper in) {
    static long ret;
    ret = in.strlen();
    return &ret;
}

