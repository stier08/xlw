
/*
 Copyright (C) 1998, 1999, 2001, 2002, 2003, 2004 J�r�me Lecomte
 Copyright (C) 2007, 2008 Eric Ehlers
 Copyright (C) 2009 Narinder S Claire
 Copyright (C) 2011 John Adcock

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
\file XlfExcel.cpp
\brief Implements the classes XlfExcel.
*/

// $Id$

#include <xlw/XlfExcel.h>
#include <cstdio>
#include <iostream>
#include <stdexcept>
#include <cstdlib>
#include <xlw/XlfOper.h>
#include <xlw/macros.h>
#include <xlw/TempMemory.h>
#include <sys/stat.h>
#include <assert.h>

extern "C"
{
    //! Main API function to Excel.
    static int (__cdecl *Excel4_)(int xlfn, LPXLOPER operRes, int count, ...);
    //! Main API function to Excel, passing the argument as an array.
    static int (__stdcall *Excel4v_)(int xlfn, LPXLOPER operRes, int count, LPXLOPER far opers[]);
}

namespace
{
    // wrap up CRT way of checking for file existance
    bool doesFileExist(const std::string& fileName)
    {
        struct _stat st;
        return (_stat(fileName.c_str(), &st) == 0);
    }
}

xlw::XlfExcel *xlw::XlfExcel::this_ = 0;

//! Internal implementation of XlfExcel.
struct xlw::XlfExcelImpl {
    //! Ctor.
    XlfExcelImpl(): handle_(0) {}
    //! Handle to the DLL module.
    HINSTANCE handle_;
};

/*!
You always have the choice with the singleton in returning a pointer or
a reference. By returning a reference and declaring the copy ctor and the
assignment otor private, we limit the risk of a wrong use of XlfExcel
(typically duplication).
*/
xlw::XlfExcel& xlw::XlfExcel::Instance() {
    if (!this_) {
        this_ = new XlfExcel;
        // intialize library first because log displays
        // XLL name in header of log window
        this_->InitLibrary();
    }
    return *this_;
}

/*!
Keep destructor private and ensure that we can be
resurected cleanly
*/
void xlw::XlfExcel::DeleteInstance() {
    delete this_;
    this_ = 0;
}



bool xlw::XlfExcel::IsEscPressed() const {
    XlfOper ret;
    Call(xlAbort, ret, 1, XlfOper(false));
    return ret.AsBool();
}

// classes and structs needed for search for window with Excel 4
namespace
{
    typedef struct
    {
        HWND hWnd;
        unsigned short loWord;
    } GetMainWindowStruct;

    BOOL CALLBACK GetMainWindowProc(HWND hWnd, LPARAM lParam)
    {
        GetMainWindowStruct* pEnum = (GetMainWindowStruct*)lParam;
        // first check the loword of the window handle
        // as this will be fast
        if (LOWORD(hWnd) == pEnum->loWord)
        {
            // then check the class of the window. Must be "XLMAIN".
            char className[7];
            if(GetClassName(hWnd, className, 7) != 0)
            {
                if (!lstrcmpi(className, "XLMAIN"))
                {
                    // We've found it stop enum.
                    pEnum->hWnd = hWnd;
                    return FALSE;
                }
            }
        }
        // carry on
        return TRUE;
    }
}


HWND xlw::XlfExcel::GetMainWindow()
{
    // on Excel4 we get LOWORD of handle back
    // so we have to faff about to find the real handle
    // On excel 12 we do get an int back but this doesn't help
    // with 64 bit so always do the search
    XLOPER ret;
    if(Call4(xlGetHwnd, &ret, 0) == xlretSuccess)
    {
        GetMainWindowStruct getMainWindowStruct = { NULL, ret.val.w};
        EnumWindows(GetMainWindowProc, (LPARAM) &getMainWindowStruct);

        if (getMainWindowStruct.hWnd != NULL)
        {
            return getMainWindowStruct.hWnd;
        }
        else
        {
            THROW_XLW("xlGetHwnd no match for partial handle");
        }
    }
    else
    {
        THROW_XLW("xlGetHwnd call failed");
    }
}

// tell VS7 to shut up
#if defined(_MSC_VER) && _MSC_VER < 1400
#pragma warning(push)
#pragma warning(disable:4312)
#endif

HINSTANCE xlw::XlfExcel::GetExcelInstance()
{
    return (HINSTANCE)GetWindowLongPtr(GetMainWindow(), GWLP_HINSTANCE);
}

#if defined(_MSC_VER) && _MSC_VER < 1400
#pragma warning(pop)
#endif

xlw::XlfExcel::XlfExcel(): impl_(0) {
    impl_ = new XlfExcelImpl();
    return;
}

xlw::XlfExcel::~XlfExcel() {
    delete impl_;
    this_ = 0;
    return;
}

int get_excel_version() {
    int version(10);
    XLOPER xRet1, xRet2, xTemp1, xTemp2;
    xTemp1.xltype = xTemp2.xltype = xltypeInt;
    xTemp1.val.w = 2;
    xTemp2.val.w = xltypeInt;
    Excel4_(xlfGetWorkspace, &xRet1, 1, &xTemp1);
    Excel4_(xlCoerce, &xRet2, 2, &xRet1, &xTemp2);
    // need to check for errors here
    // French excel doesn't like the decimal point
    // in the version string, fall back to the unsafe looking
    // atoi, not too bad as Excel always seems to NULL terminate the
    // version string, if we can't get anything make sure we don't
    // return a version greater than 12
    if(xRet2.xltype == xltypeInt)
    {
        version = xRet2.val.w;
    }
    else if(xRet1.xltype == xltypeStr)
    {
        version = atoi(xRet1.val.str + 1);
    }
    Excel4_(xlFree, 0, 1, &xRet1);
    return version;
}

/*!
Load \c XlfCALL32.DLL to interface excel (this library is shipped with Excel)
and link it to the XLL.
*/
void xlw::XlfExcel::InitLibrary() {
    HINSTANCE handle = LoadLibrary("XLCALL32.DLL");
    if (handle == 0)
        THROW_XLW("Could not load library XLCALL32.DLL");
    Excel4_ = (int (__cdecl *)(int, struct xloper *, int, ...))GetProcAddress(handle, "Excel4");
    if (Excel4_ == 0)
        THROW_XLW("Could not get address of Excel4 callback");
    Excel4v_ = (int (__stdcall *)(int, struct xloper *, int, struct xloper *[]))GetProcAddress(handle, "Excel4v");
    if (Excel4v_ == 0)
        THROW_XLW("Could not get address of Excel4v callback");

    excelVersion_ = get_excel_version();
    impl::XlfOperProperties<LPXLFOPER>::setExcel12(excel12());
    if (excel12()) {
        xlfOperType_ = "Q";
        xlfXloperType_ = "U";
        wStrType_ = "C%";
        fpArrayType_ = "K%";
    } else {
        xlfOperType_ = "P";
        xlfXloperType_ = "R";
        wStrType_ = "C";
        fpArrayType_ = "K";
    }

    impl_->handle_ = handle;


    // get the file name
    XlfOper xName;
    int err = Call(xlGetName, (LPXLFOPER)xName, 0);
    if (err == xlretSuccess)
    {
        xllFileName_ = xName.AsString();
    }
    else
    {
        std::cerr << XLW__HERE__ << "Could not get DLL name" << std::endl;
    }

    LookForHelp();
}

const std::string& xlw::XlfExcel::GetName() const {
    return xllFileName_;
}

const std::string& xlw::XlfExcel::GetHelpName() const {
    return helpFileName_;
}

std::string xlw::XlfExcel::GetXllDirectory() const {
    // find the last slash in the xll file name
    size_t slashPos(xllFileName_.find_last_of("\\/"));
    if(slashPos == std::string::npos)
    {
        return ".";
    }
    else
    {
        return xllFileName_.substr(0, slashPos);
    }
}

void xlw::XlfExcel::LookForHelp() {
    helpFileName_.clear();
    // first look for the file with the extension chm 
    // this will work as long as xll has extension .???
    size_t nameLen(xllFileName_.length());
    if(nameLen < 5 || xllFileName_[nameLen - 4] != '.')
    {
        return;
    }
    std::string testFile = xllFileName_;
    testFile[nameLen - 3] = 'c';
    testFile[nameLen - 2] = 'h';
    testFile[nameLen - 1] = 'm';

    if(doesFileExist(testFile))
    {
        helpFileName_ = testFile;
        return;
    }

    // try the directory one up
    // by inserting .. into the path
    size_t slashPos(testFile.find_last_of("\\/"));
    if(slashPos == std::string::npos)
    {
        return;
    }

    testFile = testFile.substr(0, slashPos) + "\\.." + testFile.substr(slashPos);
    if(doesFileExist(testFile))
    {
        helpFileName_ = testFile;
    }
}

int xlw::XlfExcel::Call4(int xlfn, LPXLOPER pxResult, int count) const
{
    assert(count == 0);
    return Call4v(xlfn, pxResult, 0, 0);
}

int xlw::XlfExcel::Call4(int xlfn, LPXLOPER pxResult, int count, LPXLOPER param1) const
{
    assert(count == 1);
    return Call4v(xlfn, pxResult, 1, &param1);
}

int xlw::XlfExcel::Call4(int xlfn, LPXLOPER pxResult, int count, LPXLOPER param1, LPXLOPER param2) const
{
    assert(count == 2);
    LPXLOPER paramArray[2] = {param1, param2};
    return Call4v(xlfn, pxResult, 2, paramArray);
}

int xlw::XlfExcel::Call4(int xlfn, LPXLOPER pxResult, int count, LPXLOPER param1, LPXLOPER param2, LPXLOPER param3) const
{
    assert(count == 3);
    LPXLOPER paramArray[3] = {param1, param2, param3};
    return Call4v(xlfn, pxResult, 3, paramArray);
}

int xlw::XlfExcel::Call4(int xlfn, LPXLOPER pxResult, int count, LPXLOPER param1, LPXLOPER param2, LPXLOPER param3, LPXLOPER param4) const
{
    assert(count == 4);
    LPXLOPER paramArray[4] = {param1, param2, param3, param4};
    return Call4v(xlfn, pxResult, 4, paramArray);
}

int xlw::XlfExcel::Call4(int xlfn, LPXLOPER pxResult, int count, LPXLOPER param1, LPXLOPER param2, LPXLOPER param3, LPXLOPER param4, 
         LPXLOPER param5, LPXLOPER param6, LPXLOPER param7, LPXLOPER param8, LPXLOPER param9, LPXLOPER param10) const
{
    assert(count >= 5 && count <= 10);
    LPXLOPER paramArray[10] = {param1, param2, param3, param4, param5, param6, param7, param8, param9, param10};
    return Call4v(xlfn, pxResult, count, paramArray);
}

int xlw::XlfExcel::Call12(int xlfn, LPXLOPER12 pxResult, int count) const
{
    assert(count == 0);
    return Call12v(xlfn, pxResult, 0, 0);
}

int xlw::XlfExcel::Call12(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 param1) const
{
    assert(count == 1);
    return Call12v(xlfn, pxResult, 1, &param1);
}

int xlw::XlfExcel::Call12(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 param1, LPXLOPER12 param2) const
{
    assert(count == 2);
    LPXLOPER12 paramArray[2] = {param1, param2};
    return Call12v(xlfn, pxResult, 2, paramArray);
}

int xlw::XlfExcel::Call12(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 param1, LPXLOPER12 param2, LPXLOPER12 param3) const
{
    assert(count == 3);
    LPXLOPER12 paramArray[3] = {param1, param2, param3};
    return Call12v(xlfn, pxResult, 3, paramArray);
}

int xlw::XlfExcel::Call12(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 param1, LPXLOPER12 param2, LPXLOPER12 param3, LPXLOPER12 param4) const
{
    assert(count == 4);
    LPXLOPER12 paramArray[4] = {param1, param2, param3, param4};
    return Call12v(xlfn, pxResult, 4, paramArray);
}

int xlw::XlfExcel::Call12(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 param1, LPXLOPER12 param2, LPXLOPER12 param3, LPXLOPER12 param4, 
         LPXLOPER12 param5, LPXLOPER12 param6, LPXLOPER12 param7, LPXLOPER12 param8, LPXLOPER12 param9, LPXLOPER12 param10) const
{
    assert(count >= 5 && count <= 10);
    LPXLOPER12 paramArray[10] = {param1, param2, param3, param4, param5, param6, param7, param8, param9, param10};
    return Call12v(xlfn, pxResult, count, paramArray);
}

int xlw::XlfExcel::Call(int xlfn, LPXLFOPER pxResult, int count) const
{
    assert(count == 0);
    return Callv(xlfn, pxResult, 0, 0);
}

int xlw::XlfExcel::Call(int xlfn, LPXLFOPER pxResult, int count, LPXLFOPER param1) const
{
    assert(count == 1);
    return Callv(xlfn, pxResult, 1, &param1);
}

int xlw::XlfExcel::Call(int xlfn, LPXLFOPER pxResult, int count, LPXLFOPER param1, LPXLFOPER param2) const
{
    assert(count == 2);
    LPXLFOPER paramArray[2] = {param1, param2};
    return Callv(xlfn, pxResult, 2, paramArray);
}

int xlw::XlfExcel::Call(int xlfn, LPXLFOPER pxResult, int count, LPXLFOPER param1, LPXLFOPER param2, LPXLFOPER param3) const
{
    assert(count == 3);
    LPXLFOPER paramArray[3] = {param1, param2, param3};
    return Callv(xlfn, pxResult, 3, paramArray);
}

int xlw::XlfExcel::Call(int xlfn, LPXLFOPER pxResult, int count, LPXLFOPER param1, LPXLFOPER param2, LPXLFOPER param3, LPXLFOPER param4) const
{
    assert(count == 4);
    LPXLFOPER paramArray[4] = {param1, param2, param3, param4};
    return Callv(xlfn, pxResult, 4, paramArray);
}

int xlw::XlfExcel::Call(int xlfn, LPXLFOPER pxResult, int count, LPXLFOPER param1, LPXLFOPER param2, LPXLFOPER param3, LPXLFOPER param4, 
         LPXLFOPER param5, LPXLFOPER param6, LPXLFOPER param7, LPXLFOPER param8, LPXLFOPER param9, LPXLFOPER param10) const
{
    assert(count >= 5 && count <= 10);
    LPXLFOPER paramArray[10] = {param1, param2, param3, param4, param5, param6, param7, param8, param9, param10};
    return Callv(xlfn, pxResult, count, paramArray);
}

/*!
If one (or more) cells referred as argument is(are) uncalculated, the framework
throws an exception and returns immediately to Excel.

If \c pxResult is not 0 and has auxilliary memory, flags it for deletion
with XlfOper::xlbitCallFreeAuxMem.

\sa XlfOper::~XlfOper
*/
int xlw::XlfExcel::Callv(int xlfn, LPXLFOPER pxResult, int count, LPXLFOPER pxdata[]) const {
    if (excel12())
        return Call12v(xlfn, (LPXLOPER12)pxResult, count, (LPXLOPER12*)pxdata);
    else
        return Call4v(xlfn, (LPXLOPER)pxResult, count, (LPXLOPER*)pxdata);
}

int xlw::XlfExcel::Call4v(int xlfn, LPXLOPER pxResult, int count, LPXLOPER pxdata[]) const {
#ifndef NDEBUG
    for (int i = 0; i<count;++i)
    if (!pxdata[i]) {
        if (xlfn & xlCommand)
            std::cerr << XLW__HERE__ << "xlCommand | " << (xlfn & 0x0FFF) << std::endl;
        if (xlfn & xlSpecial)
            std::cerr << "xlSpecial | " << (xlfn & 0x0FFF) << std::endl;
        if (xlfn & xlIntl)
            std::cerr << "xlIntl | " << (xlfn & 0x0FFF) << std::endl;
        if (xlfn & xlPrompt)
            std::cerr << "xlPrompt | " << (xlfn & 0x0FFF) << std::endl;
        std::cerr << "0 pointer passed as argument #" << i << std::endl;
    }
#endif
    int xlret = Excel4v_(xlfn, pxResult, count, pxdata);
    if (pxResult) {
        int type = pxResult->xltype;

        bool hasAuxMem = (type & xltypeStr ||
                        type & xltypeRef ||
                        type & xltypeMulti ||
                        type & xltypeBigData);
        if (hasAuxMem)
            pxResult->xltype |= XlfOperImpl::xlbitFreeAuxMem;
    }
    return xlret;
}

int xlw::XlfExcel::Call12v(int xlfn, LPXLOPER12 pxResult, int count, LPXLOPER12 pxdata[]) const {
#ifndef NDEBUG
    for (int i = 0; i<count; ++i)
        if (!pxdata[i]) {
            if (xlfn & xlCommand)
                std::cerr << XLW__HERE__ << "xlCommand | " << (xlfn & 0x0FFF) << std::endl;
            if (xlfn & xlSpecial)
                std::cerr << "xlSpecial | " << (xlfn & 0x0FFF) << std::endl;
            if (xlfn & xlIntl)
                std::cerr << "xlIntl | " << (xlfn & 0x0FFF) << std::endl;
            if (xlfn & xlPrompt)
                std::cerr << "xlPrompt | " << (xlfn & 0x0FFF) << std::endl;
            std::cerr << "0 pointer passed as argument #" << i << std::endl;
        }
#endif
    int xlret = Excel12v(xlfn, pxResult, count, pxdata);
    if (pxResult) {
        int type = pxResult->xltype;

        bool hasAuxMem = (type & xltypeStr ||
                          type & xltypeRef ||
                          type & xltypeMulti ||
                          type & xltypeBigData);
        if (hasAuxMem)
            pxResult->xltype |= XlfOperImpl::xlbitFreeAuxMem;
    }
    return xlret;
}

namespace {

//! Needed by IsCalledByFuncWiz.
typedef struct _EnumStruct {
    bool bFuncWiz;
    short hwndXLMain;
}
EnumStruct, FAR * LPEnumStruct;

//! Needed by IsCalledByFuncWiz.
bool CALLBACK EnumProc(HWND hwnd, LPEnumStruct pEnum) {
    const size_t CLASS_NAME_BUFFER = 50;

    // first check the class of the window.  Will be szXLDialogClass
    // if function wizard dialog is up in Excel
    char rgsz[CLASS_NAME_BUFFER];
    GetClassName(hwnd, (LPSTR)rgsz, CLASS_NAME_BUFFER);
    if (2 == CompareString(MAKELCID(MAKELANGID(LANG_ENGLISH,
        SUBLANG_ENGLISH_US),SORT_DEFAULT), NORM_IGNORECASE,
        (LPSTR)rgsz,  (lstrlen((LPSTR)rgsz)>lstrlen("bosa_sdm_XL"))
        ? lstrlen("bosa_sdm_XL"):-1, "bosa_sdm_XL", -1)) {

        if(LOWORD( GetParent(hwnd)) == pEnum->hwndXLMain) {
            pEnum->bFuncWiz = TRUE;
            return false;
        }
    }
    // no luck - continue the enumeration
    return true;
}

} // empty namespace

bool xlw::XlfExcel::IsCalledByFuncWiz() const {
    XLOPER xHwndMain;
    EnumStruct enm;

    if (Excel4_(xlGetHwnd, &xHwndMain, 0) == xlretSuccess) {
        enm.bFuncWiz = false;
        enm.hwndXLMain = xHwndMain.val.w;
        EnumWindows((WNDENUMPROC) EnumProc,
            (LPARAM) ((LPEnumStruct)  &enm));
        return enm.bFuncWiz;
    }
    return false;    //safe case: Return false if not sure
}

