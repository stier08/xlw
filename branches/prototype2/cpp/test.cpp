
// Prototype xlw upgrade to Excel 2007

#include <iostream>
#include <exception>

#include <windows.h>
#include "xlcall.h"

using std::cout;
using std::endl;

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
    virtual double AsDouble(const XlfOper &xlfOper) const = 0;
    virtual void InitXlfOper(XlfOper &xlfOper, const double &d) const = 0;
    virtual LPXLOPER operatorLPXLOPER(const XlfOper &xlfOper) const = 0;
};

XlfOperImpl *XlfOperImpl::instance_ = 0;

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
    // XlfOper default ctor declared but not implemented.  In this demo no
    // XlfOper ctor is called and that is also true in the case where Excel
    // passes LPXLOPER values to an xlw Addin function with parms of type XlfOper.
    XlfOper();
public:
    XlfOper(double d) {
        XlfOperImpl::instance().InitXlfOper(*this, d);
    }
    double AsDouble() const {
        return XlfOperImpl::instance().AsDouble(*this);
    }
    operator LPXLOPER() const {
        return XlfOperImpl::instance().operatorLPXLOPER(*this);
    }
};

// Class XlfOperImpl4
// Implementation of XlfOperImpl for Excel 4,
// dereferences the LPXLOPER in XlfOper.

class XlfOperImpl4 : public XlfOperImpl {

    void InitXlfOper(XlfOper &xlfOper, const double &d) const {
        // FIXME at this point xlw allocates memory from a buffer.
        // For purposes of this demo we just take the address of a static variable.
        static XLOPER x;
        xlfOper.lpxloper4_ = &x;
        xlfOper.lpxloper4_->xltype = xltypeNum;
        xlfOper.lpxloper4_->val.num = d;
    }

    virtual double AsDouble(const XlfOper &xlfOper) const {
        return xlfOper.lpxloper4_->val.num;
    }

    virtual LPXLOPER operatorLPXLOPER(const XlfOper &xlfOper) const {
        return xlfOper.lpxloper4_;
    }

};

// Class XlfOperImpl12
// Implementation of XlfOperImpl for Excel 12,
// dereferences the LPXLOPER12 in XlfOper.

class XlfOperImpl12 : public XlfOperImpl {

    void InitXlfOper(XlfOper &xlfOper, const double &d) const {
        // FIXME at this point xlw allocates memory from a buffer.
        // For purposes of this demo we just take the address of a static variable.
        static XLOPER12 x;
        xlfOper.lpxloper12_ = &x;
        xlfOper.lpxloper12_->xltype = xltypeNum;
        xlfOper.lpxloper12_->val.num = d;
    }

    virtual double AsDouble(const XlfOper &xlfOper) const {
        return xlfOper.lpxloper12_->val.num;
    }

    virtual LPXLOPER operatorLPXLOPER(const XlfOper &xlfOper) const {
        // FIXME need to convert from LPXLOPER12 to LPXLOPER.
        // For purposes of this demo we skip the conversion and rely on the fact
        // that double has the same offset in either structure.

        //LPXLOPER ret = LPXLOPER12_to_4(xlfOper.lpxloper12_);
        //return ret;
        return xlfOper.lpxloper4_;
    }

};

// Example of an Addin function built using xlw (taken from xlw documentation).
// This Addin function declares its parameters as type XlfOper
// and relies on that class to encapsulate either LPXLOPER or LPXLOPER12
// depending on the version of Excel detected at startup.

// The following function takes the diameter of a circle as an argument
// and computes the circumference of this circle.

LPXLOPER xlCirc(XlfOper xlDiam) {
    // Converts xlDiam to a double.
    double ret=xlDiam.AsDouble();
    // Multiplies it.
    ret *= 3.14159;
    // Returns the result as a XlfOper.
    return XlfOper(ret);
}

// The Addin registers with Excel functions having parameters of type LPXLOPER.
// The actual Addin function instead declares parameters of type XlfOper.
// Excel passes values of type LPXLOPER which are received by parameters of type XlfOper.
// No conversion is performed and the XlfOper ctor is not called, xlw relies on bitwise
// equivalence of LPXLOPER and XlfOper.  This function simulates the effect.

XlfOper &hack_type(void *x) {
    static char buf[4];
    memcpy(buf, &x, 4);
    XlfOper *ret = reinterpret_cast<XlfOper*>(buf);
    return *ret;
}

void main() {

    cout << "start" << endl;

    try {

        {   // Pretend we're running under Excel 4

            // Instantiate the XlfOperImpl singleton for Excel 4
            XlfOperImpl4 xlfOperImpl4;

            // Register with Excel Addin function xlCirc accepting argument of type LPXLOPER
            // ...

            // Simulate Excel preparing to call our Addin function.
            // Initialize the input parameter as a double.
            XLOPER x;
            x.xltype = xltypeNum;
            x.val.num = 10.;

            // Simulate Excel calling our Addin function
            LPXLOPER ret = xlCirc(hack_type(&x));
            // The return value displayed by Excel in the calling cell
            cout << ret->val.num << endl;
        }

        {   // Pretend we're running under Excel 12

            // Instantiate the XlfOperImpl singleton for Excel 12
            XlfOperImpl12 xlfOperImpl12;

            // Register with Excel Addin function xlCirc accepting argument of type LPXLOPER12
            // ...

            // Simulate Excel preparing to call our Addin function.
            // Initialize the input parameter as a double.
            XLOPER12 x;
            x.xltype = xltypeNum;
            x.val.num = 10.;

            // Simulate Excel calling our Addin function
            LPXLOPER ret = xlCirc(hack_type(&x));
            // The return value displayed by Excel in the calling cell
            cout << ret->val.num << endl;
        }

    } catch(const std::exception &e) {
        cout << "error: " << e.what() << endl;
    }
    cout << "end" << endl;
}

