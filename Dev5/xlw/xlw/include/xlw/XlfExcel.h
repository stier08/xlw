
/*
 Copyright (C) 1998, 1999, 2001, 2002, 2003, 2004 Jérôme Lecomte
 Copyright (C) 2007, 2008 Eric Ehlers
 Copyright (C) 2011 John Adcock
 Copyright (C) 2011 Narinder S Claire

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/


#ifndef INC_XlfExcel_H
#define INC_XlfExcel_H

/*!
\file XlfExcel.h
\brief Declares class XlfExcel.
*/

// $Id$

#include <xlw/EXCEL32_API.h>
#include <xlw/xlcall32.h>
#include <list>
#include <map>
#include <string>
#include <xlw/TempMemory.h>
#include <xlw/XlfServices.h>

#if defined(_MSC_VER)
#pragma once
#endif

namespace xlw {

    //! Interpreted as LPXLOPER (Excel 4) or LPXLOPER12 (Excel 12).
    typedef struct
    {
        union
        {
            XLOPER oper4;
            XLOPER12 oper12;
        };
    } XLFOPER;

    typedef XLFOPER* LPXLFOPER;


    //! Interface between excel and the framework.
    /*!
    Implemented as a singleton (see \ref DP). You can't access the
    ctors (private) and should access the class via the static
    Instance() method.
    */
    class EXCEL32_API XlfExcel
    {
    public:
        //! \name Structors and static members
        //@{
        //! Used to obtain instance on XlfExcel object.
        static XlfExcel& Instance();
        //! Used to delete instance in xlAutoClose.
        static void DeleteInstance();
       
        //@}

        //! \name Inspectors
        //@{
        //! Gets XLL name
        std::string GetName() const;
        //@}

        /** \name Wrappers for Excel Call function
        The Call() functions accept up to 4 arguments
        the Callv() functions accept an argument array.
        */
        //@{
        //! Wrap the Callv4 or Callv12 depending on which version of Excel is running that take no parameters
        int Call(int xlfn, LPXLFOPER pxResult) const;
        //! Wrap the Callv4 or Callv12 depending on which version of Excel is running that takes 1 parameter
        int Call(int xlfn, LPXLFOPER pxResult, const LPXLFOPER& param1) const;
        //! Wrap the Callv4 or Callv12 depending on which version of Excel is running that takes 2 parameters
        int Call(int xlfn, LPXLFOPER pxResult, const LPXLFOPER& param1, const LPXLFOPER& param2) const;
        //! Wrap the Callv4 or Callv12 depending on which version of Excel is running that takes 3 parameters
        int Call(int xlfn, LPXLFOPER pxResult, const LPXLFOPER& param1, const LPXLFOPER& param2, const LPXLFOPER& param3) const;
        //! Wrap the Callv4 or Callv12 depending on which version of Excel is running that takes 4 parameters
        int Call(int xlfn, LPXLFOPER pxResult, const LPXLFOPER& param1, const LPXLFOPER& param2, const LPXLFOPER& param3, const LPXLFOPER& param4) const;
        //! Wrap the Call function for Excel version 4 functions that take no parameters
        int Call4(int xlfn, LPXLOPER pxResult) const;
        //! Wrap the Call function for Excel version 4 functions that takes 1 parameter
        int Call4(int xlfn, LPXLOPER pxResult, const LPXLOPER& param1) const;
        //! Wrap the Call function for Excel version 4 functions that takes 2 parameters
        int Call4(int xlfn, LPXLOPER pxResult, const LPXLOPER& param1, const LPXLOPER& param2) const;
        //! Wrap the Call function for Excel version 4 functions that takes 3 parameters
        int Call4(int xlfn, LPXLOPER pxResult, const LPXLOPER& param1, const LPXLOPER& param2, const LPXLOPER& param3) const;
        //! Wrap the Call function for Excel version 4 functions that takes 4 parameters
        int Call4(int xlfn, LPXLOPER pxResult, const LPXLOPER& param1, const LPXLOPER& param2, const LPXLOPER& param3, const LPXLOPER& param4) const;
        //! Wrap the Call function for Excel version 12 functions that take no parameters
        int Call12(int xlfn, LPXLOPER12 pxResult) const;
        //! Wrap the Call function for Excel version 12 functions that takes 1 parameter
        int Call12(int xlfn, LPXLOPER12 pxResult, const LPXLOPER12& param1) const;
        //! Wrap the Call function for Excel version 12 functions that takes 2 parameters
        int Call12(int xlfn, LPXLOPER12 pxResult, const LPXLOPER12& param1, const LPXLOPER12& param2) const;
        //! Wrap the Call function for Excel version 12 functions that takes 3 parameters
        int Call12(int xlfn, LPXLOPER12 pxResult, const LPXLOPER12& param1, const LPXLOPER12& param2, const LPXLOPER12& param3) const;
        //! Wrap the Call function for Excel version 12 functions that takes 4 parameters
        int Call12(int xlfn, LPXLOPER12 pxResult, const LPXLOPER12& param1, const LPXLOPER12& param2, const LPXLOPER12& param3, const LPXLOPER12& param4) const;
        //! Wrap the Callv function for Excel version 4.
        int Call4v(int xlfn, LPXLOPER pxResult, int count, LPXLOPER pxdata[]) const;
        //! Wrap the Callv function for Excel version 12.
        int Call12v(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 pxdata[]) const;
        //! Invoke Callv4 or Callv12 depending on which version of Excel is running.
        int Callv(int xlfn, LPXLFOPER pxResult, int count, LPXLFOPER pxdata[]) const;
        //@}

        //! \name Wrappers for selected Excel operations
        //@{
        /*!
        Wrapped functions that are often needed and/or painful to code
        */

        //! Was the Esc key pressed ?
        bool IsEscPressed() const;
        //! Is the function being calculated currently called by the Function Wizard ?
        bool IsCalledByFuncWiz() const;
        //! Gets the HWND of excel's main window
        HWND GetMainWindow();
        //! Gets the instance of Excel we are running under
        HINSTANCE GetExcelInstance();
        //@}

        //! \name Information about the running version of Excel
        //@{
        //! Boolean differentiating Excel 12 (2007) and above from previous versions
        bool excel12() const { return (excelVersion_ >= 12); }
        //! Boolean differentiating Excel 14 (2010) from previous versions
        bool excel14() const { return (excelVersion_ >= 14); }
        //! Returns Excel version number e.g. 12 is 2007
        int excelVersion() const { return excelVersion_; }
        //! The OPER type in use by this version of Excel
        const std::string & xlfOperType() const { return xlfOperType_; }
        //! The XLOPER type in use by this version of Excel
        const std::string & xlfXloperType() const { return xlfXloperType_; }
        //! The string type in use by this version of Excel
        const std::string & wStrType() const { return wStrType_; }
        //! The double array type use by this version of Excel
        const std::string & fpType() const { return fpArrayType_; }
        //@}

        //Services 
        //These are services that are for delivery for
        //the end user as opposed to for prdeominantly internal use
        //for xlw
        StatusBar_t status_bar;
        Reflection_t reflection;

    private:
        //! Static pointer to the unique instance of XlfExcel object.
        static XlfExcel *this_;

        //! Pointer to internal implementation (pimpl idiom, see \ref HS).
        struct XlfExcelImpl * impl_;

        //! Ctor.
        XlfExcel();
        //! Dtor.
        ~XlfExcel();
        //! Copy ctor is not defined.
        XlfExcel(const XlfExcel&);
        //! Assignment operator is not defined.
        XlfExcel& operator=(const XlfExcel&);
        //! Initialize the C++ framework.
        void InitLibrary();
        //! Create a new static buffer and add it to the free list.
        void PushNewBuffer(size_t);

        int excelVersion_;
        std::string xlfOperType_;
        std::string xlfXloperType_;
        std::string wStrType_;
        std::string fpArrayType_;

        
    };

}
#endif

