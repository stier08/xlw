
/*
 Copyright (C) 2011  John Adcock

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#ifndef INC_PascalStringConversions_H
#define INC_PascalStringConversions_H

/*!
\file PascalStringConversions.h
\brief defines for conversion functions used when dealing with Excel strings
*/

// $Id: XlfOper.h 920 2011-04-12 12:00:51Z adcockj $

#include <string>

#if defined(_MSC_VER)
#pragma once
#endif

namespace xlw {

    class PascalStringConversions
    {
    public:
        static std::string PascalStringToString(char* pascalString);
        static std::wstring PascalStringToWString(char* pascalString);
        static char* StringToPascalString(const std::string& cString);
        static char* WStringToPascalString(const std::wstring& cString);
        static std::string WPascalStringToString(wchar_t* pascalString);
        static std::wstring WPascalStringToWString(wchar_t* pascalString);
        static wchar_t* StringToWPascalString(const std::string& cString);
        static wchar_t* WStringToWPascalString(const std::wstring& cString);
    };
}

#endif

