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
#include <stdexcept>

namespace xlw
{
    struct Services_t XlfServices;

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

    namespace
    {
        inline XlfOper CallFunction(int xlfn, const char* errorString)
        {
            XlfOper result;
            int err = XlfExcel::Instance().Call(xlfn, result);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
            return result;
        }

        inline XlfOper CallFunction(int xlfn, const XlfOper& param1, const char* errorString)
        {
            XlfOper result;
            int err = XlfExcel::Instance().Call(xlfn, result, param1);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
            return result;
        }

        inline XlfOper CallFunction(int xlfn, const XlfOper& param1, const XlfOper& param2, const char* errorString)
        {
            XlfOper result;
            int err = XlfExcel::Instance().Call(xlfn, result, param1, param2);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
            return result;
        }

        inline XlfOper CallFunction(int xlfn, const XlfOper& param1, const XlfOper& param2, const XlfOper& param3, const char* errorString)
        {
            XlfOper result;
            int err = XlfExcel::Instance().Call(xlfn, result, param1, param2, param3);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
            return result;
        }

        inline XlfOper CallFunction(int xlfn, const XlfOper& param1, const XlfOper& param2, const XlfOper& param3, const XlfOper& param4, const char* errorString)
        {
            XlfOper result;
            int err = XlfExcel::Instance().Call(xlfn, result, param1, param2, param3, param4);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
            return result;
        }

        inline void CallCommand(int xlcmd, const char* errorString)
        {
            int err = XlfExcel::Instance().Call(xlcmd, 0);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
        }

        inline void CallCommand(int xlcmd, const XlfOper& param1, const char* errorString)
        {
            int err = XlfExcel::Instance().Call(xlcmd, 0, param1);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
        }

        inline void CallCommand(int xlcmd, const XlfOper& param1, const XlfOper& param2, const char* errorString)
        {
            int err = XlfExcel::Instance().Call(xlcmd, 0, param1, param2);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
        }

        inline void CallCommand(int xlcmd, const XlfOper& param1, const XlfOper& param2, const XlfOper& param3, const char* errorString)
        {
            int err = XlfExcel::Instance().Call(xlcmd, 0, param1, param2, param3);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
        }

        inline void CallCommand(int xlcmd, const XlfOper& param1, const XlfOper& param2, const XlfOper& param3, const XlfOper& param4, const char* errorString)
        {
            int err = XlfExcel::Instance().Call(xlcmd, 0, param1, param2, param3, param4);
            if(err != xlretSuccess)
            {
                throw std::logic_error(errorString);
            }
        }
    }


    StatusBar_t & StatusBar_t::operator=(const std::string &message)
    {
        CallCommand(xlcMessage, true, message, "Set Status Bar failed");
        return *this;
    }

    void StatusBar_t::clear()
    {
        CallCommand(xlcMessage, 0, false, "Clear Status Bar Failed");
    }

    XlfOper Notes_t::GetNote(const XlfOper& cellRef, int startCharacter, int numChars)
    {
        if(numChars == 0)
        {
            return CallFunction(xlfGetNote, cellRef, startCharacter + 1, "Get Note Failed");
        }
        else
        {
            return CallFunction(xlfGetNote, cellRef, startCharacter + 1, numChars, "Get Note Failed");
        }
    }

    void Notes_t::SetNote(const XlfOper& cellRef, const std::string& note)
    {
        CallCommand(xlcNote, note, cellRef, "Set Note failed");
    }
    
    void Notes_t::ClearNote(const XlfOper& cellRef)
    {
        CallCommand(xlcNote, XlfOper(), cellRef, "Clear Note failed");
    }

    XlfOper Information_t::GetCallingCell()
    {
        return CallFunction(xlfCaller, "Calling Cell Failed");
    }

    XlfOper Information_t::GetActiveCell()
    {
        return CallFunction(xlfActiveCell, "Get Active Cell failed");
    }

    std::string Information_t::GetFormula(const XlfOper& cellRef)
    {
        return CallFunction(xlfGetFormula, cellRef, "Get Formula failed").AsString();
    }

    std::string Information_t::ConvertA1FormulaToR1C1(std::string a1Formula)
    {
        return CallFunction(xlfFormulaConvert, a1Formula, true, false, 1, "Formula Convert failed").AsString();
    }

    std::string Information_t::ConvertR1C1FormulaToA1(std::string r1c1Formula, bool fixRows, bool fixColums)
    {
        int toRefType(4 - fixRows?2:0 - fixColums?1:0);
        return CallFunction(xlfFormulaConvert, r1c1Formula, false, true, toRefType, "Formula Convert failed").AsString();
    }

    XlfOper Information_t::GetCellRefA1(std::string a1Location)
    {
        return CallFunction(xlfTextref, a1Location, true, "Text Ref failed");
    }

    XlfOper Information_t::GetCellRefR1C1(std::string r1c1Location)
    {
        return CallFunction(xlfTextref, r1c1Location, false, "Text Ref failed");
    }

    XlfOper Information_t::GetCellRefR1C1(XlfOper referenceCell, std::string r1c1RelativeLocation)
    {
        return CallFunction(xlfAbsref, r1c1RelativeLocation, referenceCell, "Abs Ref failed");
    }

    std::string Information_t::GetRefTextA1(const XlfOper& ref)
    {
        return CallFunction(xlfReftext, ref, true, "Ref Text failed").AsString();
    }
    
    std::string Information_t::GetRefTextR1C1(const XlfOper& ref)
    {
        return CallFunction(xlfReftext, ref, false, "Ref Text failed").AsString();
    }

    std::string Information_t::GetSheetName(const XlfOper& ref)
    {
        return CallFunction(xlSheetNm, ref, "Get Sheet Name failed").AsString();
    }
    void Commands_t::Alert(const std::string& message)
    {
        CallCommand(xlcAlert, message, "Alert Command failed");
    }

    std::string Commands_t::InputFormula(const std::string& message, const std::string& title)
    {
        return CallFunction(xlfInput, message, 0, title, "Input failed").AsString();
    }

    double Commands_t::InputNumber(const std::string& message, const std::string& title)
    {
        return CallFunction(xlfInput, message, 1, title, "Input failed").AsDouble();
    }

    std::string Commands_t::InputText(const std::string& message, const std::string& title)
    {
        return CallFunction(xlfInput, message, 2, title, "Input failed").AsString();
    }

    bool Commands_t::InputBool(const std::string& message, const std::string& title)
    {
        return CallFunction(xlfInput, message, 4, title, "Input failed").AsBool();
    }

    XlfOper Commands_t::InputReference(const std::string& message, const std::string& title)
    {
        return CallFunction(xlfInput, message, 8, title, "Input failed");
    }

    XlfOper Commands_t::InputArray(const std::string& message, const std::string& title)
    {
        return CallFunction(xlfInput, message, 64, title, "Input failed");
    }
}
