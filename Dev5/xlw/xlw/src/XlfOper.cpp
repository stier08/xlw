
/*
 Copyright (C) 1998, 1999, 2001, 2002 Jérôme Lecomte
 Copyright (C) 2007 Tim Brunne
 Copyright (C) 2007, 2008 Eric Ehlers

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

/*!
\file XlfOper.cpp
\brief Implements the XlfOper class.
*/

// $Id$

#include <xlw/XlfOper.h>
#include <xlw/XlfException.h>
#include <xlw/XlfRef.h>
#include <xlw/macros.h>
#include <cassert>
#include <iostream>
#include <stdexcept>
#include <xlw/CellMatrix.h>
#include <algorithm>

// Stop header precompilation
#ifdef _MSC_VER
#pragma hdrstop
#endif

#ifndef NDEBUG
#include <xlw/XlfOper.inl>
#endif

/*!
This bit is currently unused by Microsoft Excel. We set it
to indicate that the LPXLOPER (passed by Excel) holds some extra
memory to be freed.

This bit is controled in ~XlfOper to know if the DLL should release
auxiliary memory or not by a call to FreeAuxiliaryMemory.
*/
int xlw::XlfOper::xlbitFreeAuxMem = 0x8000;

/*!
Attempts to convert the implict object to a double. If pxlret is not null
the method won't throw and the Excel return code will be returned in this
variable.

\sa xlw::XlfOper::ConvertToDouble.
*/
double xlw::XlfOper::AsDouble(const char* ErrorId, int *pxlret) const
{
    double d;
    int xlret = ConvertToDouble(d);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError( xlret, ErrorId, "Conversion to double");
    return d;
};

/*!
Attempts to convert the implict object to an array.  Does this by calling
AsDoubleVector.  If pxlret is not null the method won't throw and the Excel
return code will be returned in this variable.
*/
xlw::MyArray xlw::XlfOper::AsArray(const char* ErrorId, XlfOperImpl::DoubleVectorConvPolicy policy, int *pxlret ) const
{
    std::vector<double> tmp(AsDoubleVector(ErrorId,policy,pxlret));
    MyArray result(tmp.size());
    for (unsigned long i=0; i < result.size(); i++)
        result[i] = tmp[i];

    return result;
}

/*!
Attempts to convert the implict object to a vector of double. If pxlret is
not null the method won't throw and the Excel return code will be returned
in this variable.

\sa xlw::XlfOper::ConvertToDoubleVector.
*/
std::vector<double> xlw::XlfOper::AsDoubleVector(const char* ErrorId, XlfOperImpl::DoubleVectorConvPolicy policy, int *pxlret) const
{
    std::vector<double> v;
    int xlret = ConvertToDoubleVector(v, policy);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError(xlret, ErrorId, "conversion to double vector");
    return v;
}

/*!
Attempts to convert the implict object to a short. If pxlret is not null the
method won't throw and the Excel return code will be returned in this variable.

\sa xlw::XlfOper::ConvertToShort.
*/
short xlw::XlfOper::AsShort(const char* ErrorId, int *pxlret) const
{
    short s;
    int xlret = ConvertToShort(s);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError(xlret,  ErrorId, "conversion to short failed");
    return s;
};

/*!
Attempts to convert the implict object to a bool. If pxlret is not null the
method won't throw and the Excel return code will be returned in this variable.

\sa xlw::XlfOper::ConvertToBool.
*/
bool xlw::XlfOper::AsBool(const char* ErrorId, int *pxlret) const
{
    bool b;
    int xlret = ConvertToBool(b);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError(xlret, ErrorId, "conversion to bool failed");
    return b;
};

/*!
Attempts to convert the implict object to a matrix. If pxlret is not null the
method won't throw and the Excel return code will be returned in this variable.

\sa xlw::XlfOper::ConvertToMatrix.
*/
xlw::MyMatrix xlw::XlfOper::AsMatrix(const char* ErrorId, int *pxlret) const
{
    MyMatrix output; // will be resized anyway
    int xlret = ConvertToMatrix(output);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError(xlret, ErrorId, "conversion to matrix failed");
    return output;
}

/*!
Attempts to convert the implict object to a cell matrix. If pxlret is not null
the method won't throw and the Excel return code will be returned in this
variable.

\sa xlw::XlfOper::ConvertToCellMatrix.
*/
xlw::CellMatrix xlw::XlfOper::AsCellMatrix(const char* ErrorId, int *pxlret) const
{
    CellMatrix output(1,1); // will be resized anyway
    int xlret = ConvertToCellMatrix(output);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError(xlret, ErrorId, "conversion to cell matrix failed");
    return output;
}

/*!
Attempts to convert the implict object to a char string. If pxlret is not
null the method won't throw and the Excel return code will be returned in
this variable.

\sa xlw::XlfOper::ConvertToString.

The XLL allocates the memory on its own buffer. This buffer is automatically
freed when a function of the XLL is called again. Note that coerce to
a char string is the slowest cast of all.
*/
char *xlw::XlfOper::AsString(const char* ErrorId, int *pxlret) const
{
    char * s;
    int xlret = ConvertToString(s);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError(xlret, ErrorId, "conversion to char* failed");
    return s;
};

std::wstring xlw::XlfOper::AsWstring(const char* ErrorId, int *pxlret) const
{
    std::wstring s;
    int xlret = ConvertToWstring(s);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError(xlret);
    return s;
};

/*!
Attempts to convert the implict object to an XlfRef. If pxlret is not null
the method won't throw and the Excel return code will be returned in this
variable.

\sa xlw::XlfOper::ConvertToRef.
*/
xlw::XlfRef xlw::XlfOper::AsRef(const char* ErrorId, int *pxlret) const
{
    XlfRef r;
    int xlret = ConvertToRef(r);
    if (pxlret)
        *pxlret=xlret;
    else
        ThrowOnError(xlret);
    return r;
}

xlw::XlfOper& xlw::XlfOper::Set(const MyMatrix& values)
{
    if (values.size1() ==0 || values.size2() ==0)
    {
        return *this;
    }

    CellMatrix tmp(values.size1(), values.size2());
    for (unsigned long i=0; i < values.size1(); i++)
        for (unsigned long j=0; j < values.size2(); j++)
            tmp(i,j) = Element(values,i,j);
    return Set(tmp);
}

xlw::XlfOper& xlw::XlfOper::Set(const MyArray& values)
{
    if (values.size() ==0)
    {
        return *this;
    }
    CellMatrix tmp(static_cast<unsigned long>(values.size()), 1UL);
    for (unsigned long i=0; i < values.size(); i++)
            tmp(i,0) = values[i];
    return Set(tmp);
}

/*!
Throws an exception if the argument is anything other than xlretSuccess.

Events that require an immediate return to excel (uncalculated cell, abort,
stack overflow and invalid OPER (potential memory exhaustion)) throw an
XlfException.

Other events throw std::runtime_error. The Identifier is tagged on to the error
message to help track down problems.
*/
int xlw::XlfOper::ThrowOnError(int xlret, const char* ErrorId, const char* Identifier) const
{
    if (xlret == xlretSuccess)
        return xlret;

    if (xlret & xlretUncalced)
        throw XlfExceptionUncalculated();
    if (xlret & xlretAbort)
        throw XlfExceptionAbort();
    if (xlret & xlretStackOvfl)
        throw XlfExceptionStackOverflow();
    if (xlret & xlretInvXloper) {
        std::string msg("invalid OPER structure (memory could be exhausted)");
        if(ErrorId)
        {
            msg += ", ";
            msg += ErrorId;
        }
        if(Identifier)
        {
            msg += " ";
            msg += Identifier;
        }
        throw XlfException(msg);
    }
    if (xlret & xlretFailed) {
        std::string msg("command failed");
        if(ErrorId)
        {
            msg += ",";
            msg += ErrorId;
        }
        if(Identifier)
        {
            msg += " ";
            msg += Identifier;
        }
        throw std::runtime_error(msg);
    }
    if (xlret & xlretInvCount) {
        std::string msg("invalid number of arguments");
        if(ErrorId)
        {
            msg += ",";
            msg += ErrorId;
        }
        if(Identifier)
        {
            msg += " ";
            msg += Identifier;
        }
        throw std::runtime_error(msg);
    }
    if (xlret & xlretInvXlfn) {
        std::string msg("invalid function number");
        if(ErrorId)
        {
            msg += ",";
            msg += ErrorId;
        }
        if(Identifier)
        {
            msg += " ";
            msg += Identifier;
        }
        throw std::runtime_error(msg);
    }
    // should never get here.
    assert(0);
    return xlret;
}

std::string xlw::XlfOper::xltypeName() const {
    DWORD type = xltype();
    if (type == xltypeNum)
        return "xltypeNum";
    else if (type == xltypeStr)
        return "xltypeStr";
    else if (type == xltypeBool)
        return "xltypeBool";
    else if (type == xltypeRef)
        return "xltypeRef";
    else if (type == xltypeErr)
        return "xltypeErr";
    else if (type == xltypeMulti)
        return "xltypeMulti";
    else if (type == xltypeMissing)
        return "xltypeMissing";
    else if (type == xltypeNil)
        return "xltypeNil";
    else if (type == xltypeSRef)
        return "xltypeSRef";
    else if (type == xltypeInt)
        return "xltypeInt";
    else
        return "unknown";
}

