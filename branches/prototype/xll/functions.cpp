
#include <xlfoper.hpp>
#include <xlfoper4.hpp>
#include <xlfoper12.hpp>
#include <xlfoperimpl.hpp>

#define DLLEXPORT extern "C" __declspec(dllexport)

DLLEXPORT void *echo(void *in) {
    XlfOper in2 = XlfOperImpl::instance().create(in);
    return in2.echo();
}

DLLEXPORT LPXLOPER echo4(XlfOper4 in) {
    return in.echo();
}

DLLEXPORT LPXLOPER12 echo12(XlfOper12 in) {
    return in.echo();
}
