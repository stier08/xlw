/*
Copyright (C) 1998, 1999, 2001, 2002, 2003, 2004 Jérôme Lecomte
Copyright (C) 2007, 2008 Eric Ehlers
Copyright (C) 2009 2011 Narinder S Claire
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

#include <xlw/XlfServices.h>
#include <xlw/XlfExcel.h>
#include <xlw/XlfOper.h>
#include <string>
#include <cstdio>

namespace xlw
{

    /// I am not sure whether we need the following at all 

    /*!
    If no title is specified, the message is assumed to be an error log
    */
    void MsgBox(const char *errmsg, const char *title) {
        LPVOID lpMsgBuf;
        // retrieve message error from system err code
        if (!title) {
            DWORD err = GetLastError();
            FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                NULL,
                err,
                MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
                (LPTSTR) &lpMsgBuf,
                0,
                NULL);
            // Process any inserts in lpMsgBuf.
            char completeMessage[255];
            std::sprintf(completeMessage,"%s due to error %lu :\n%s", errmsg, err, (LPCSTR)lpMsgBuf);
            MessageBox(NULL, completeMessage,"XLL Error", MB_OK | MB_ICONINFORMATION);
            // Free the buffer.
            LocalFree(lpMsgBuf);
        } else {
            MessageBox(NULL, errmsg, title, MB_OK | MB_ICONINFORMATION);
        }
        return;
    }


    StatusBar_t & StatusBar_t::operator=(const std::string &message)
    {
        XlfExcel::Instance().Call(xlcMessage, 0, XlfOper(true), XlfOper(message));
        return *this;
    }

    StatusBar_t & StatusBar_t::operator=(const std::wstring &message)
    {
        XlfExcel::Instance().Call(xlcMessage, 0, XlfOper(true), XlfOper(message));
        return *this;
    }
    void StatusBar_t::clear()
    {
        XlfExcel::Instance().  Call(xlcMessage, 0, XlfOper(false));
    }

    std::string Reflection_t::GetNote() {
        XlfOper callingCell;
        XlfOper startChars(short(1));
        XlfOper numChars(short(255));
        XlfOper result;

        XlfExcel::Instance().Call(xlfCaller, callingCell, 0);
        XlfExcel::Instance().Call(xlfGetNote, result, callingCell, startChars, numChars);

        return result.AsString();
    }


}