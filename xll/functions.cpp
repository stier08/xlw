
#include <xlfoper.hpp>
#include <xlfoper4.hpp>
#include <xlfoper12.hpp>
#include <xlfoperimpl.hpp>

#define DLLEXPORT extern "C" __declspec(dllexport)

DLLEXPORT LPXLFOPER xlw_echo(LPXLFOPER in) {
    XlfOper in2(in);
    return in2.echo();
}

DLLEXPORT LPXLOPER xlw_echo4(XlfOper4 in) {
    return in.echo();
}

DLLEXPORT LPXLOPER12 xlw_echo12(XlfOper12 in) {
    return in.echo();
}
