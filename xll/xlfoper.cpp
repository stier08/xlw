#include <xlfoper.hpp>
#include <xlfoperimpl.hpp>

long XlfOper::strlen() const {
    return XlfOperImpl::instance().strlen(*this);
}
