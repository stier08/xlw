#include <xlfoper.hpp>
#include <xlfoperimpl.hpp>

XlfOper XlfOper::echo() const {
    return XlfOperImpl::instance().echo(*this);
}

XlfOper::operator void*() const {
    return XlfOperImpl::instance().as_void(*this);
}

XlfOper::operator LPXLOPER() const {
    return XlfOperImpl::instance().as_LPXLOPER(*this);
}

XlfOper::operator LPXLOPER12() const {
    return XlfOperImpl::instance().as_LPXLOPER12(*this);
}
