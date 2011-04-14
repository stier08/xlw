/*
Copyright (C) 2011 John Adcock

This file is part of xlw, a free-software/open-source C++ wrapper of the
Excel C API - http://xlw.sourceforge.net/

xlw is free software: you can redistribute it and/or modify it under the
terms of the xlw license.  You should have received a copy of the
license along with this program; if not, please email xlw-users@lists.sf.net

This program is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#include <xlw/PascalStringConversions.h>
#include <xlw/TempMemory.h>
#include <xlw/macros.h>
#include <iostream>

#ifndef WC_NO_BEST_FIT_CHARS
#define WC_NO_BEST_FIT_CHARS 0x00000400
#endif


std::string xlw::PascalStringConversions::PascalStringToString(char* pascalString)
{
    // Must use datatype unsigned char (BYTE) to process 0th byte
    // otherwise numbers greater than 128 are incorrect
    size_t n = static_cast<BYTE>(pascalString[0]);
    std::string result(pascalString + 1, pascalString + 1 + n);
    return result;
}

std::wstring xlw::PascalStringConversions::PascalStringToWString(char* pascalString)
{
    // Must use datatype unsigned char (BYTE) to process 0th byte
    // otherwise numbers greater than 128 are incorrect
    size_t n = static_cast<BYTE>(pascalString[0]);
    std::wstring result(n, L'\0');
    MultiByteToWideChar(CP_ACP, 0, pascalString + 1, (int)n, &result[0], (int)n);
    return result;
}

char* xlw::PascalStringConversions::StringToPascalString(const std::string& cString)
{
    size_t n(cString.length());

    if (n > 255)
    {
        std::cerr << XLW__HERE__ << "String truncated to 255 bytes" << std::endl;
        n = 255;
    }
    // One byte more for the string length (convention used by Excel)
    // and another so that the string is null terminated so that the
    // debugger sees it correctly
    LPSTR result = TempMemory::GetMemory<char>(n + 2);
    strncpy(result + 1, cString.c_str(), n);
    result[n + 1] = 0;
    result[0] = static_cast<BYTE>(n);
    return result;
}

char* xlw::PascalStringConversions::WStringToPascalString(const std::wstring& cString)
{
    size_t n(cString.length());

    if (n > 255)
    {
        std::cerr << XLW__HERE__ << "String truncated to 255 bytes" << std::endl;
        n = 255;
    }
    // One byte more for the string length (convention used by Excel)
    // and another so that the string is null terminated so that the
    // debugger sees it correctly
    LPSTR result = TempMemory::GetMemory<char>(n + 2);
    WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, cString.c_str(), (int)n, result + 1, (int)n, NULL, NULL);
    result[n + 1] = 0;
    result[0] = static_cast<BYTE>(n);
    return result;
}

std::string xlw::PascalStringConversions::WPascalStringToString(wchar_t* pascalString)
{
    size_t n = pascalString[0];
    if(n > 0)
    {
        std::string result(n, '\0');
        WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, pascalString + 1, (int)n, &result[0], (int)n, NULL, NULL);
        return result;
    }
    else
    {
        return std::string();
    }
}

std::wstring xlw::PascalStringConversions::WPascalStringToWString(wchar_t* pascalString)
{
    // Must use datatype unsigned char (BYTE) to process 0th byte
    // otherwise numbers greater than 128 are incorrect
    size_t n = pascalString[0];
    std::wstring result(pascalString + 1, pascalString + 1 + n);
    return result;
}

wchar_t* xlw::PascalStringConversions::StringToWPascalString(const std::string& cString)
{
    size_t n(cString.length());

    if (n > 32766) {
        std::cerr << XLW__HERE__ << "String truncated to 32766 bytes" << std::endl;
        n = 32766;
    }

    // One byte more for the string length (convention used by Excel)
    // and another so that the string is null terminated so that the
    // debugger sees it correctly
    wchar_t* result  = TempMemory::GetMemory<wchar_t>(n+2);
    MultiByteToWideChar(CP_ACP, 0, cString.c_str(), (int)n, result + 1, (int)n);
    result[n + 1] = 0;
    result[0] = static_cast<XCHAR>(n);
    return result;
}

wchar_t* xlw::PascalStringConversions::WStringToWPascalString(const std::wstring& cString)
{
    size_t n(cString.length());

    if (n > 32766)
    {
        std::cerr << XLW__HERE__ << "String truncated to 32766 bytes" << std::endl;
        n = 32766;
    }
    // One byte more for the string length (convention used by Excel)
    // and another so that the string is null terminated so that the
    // debugger sees it correctly
    wchar_t* result = TempMemory::GetMemory<wchar_t>(n + 2);
    wcsncpy(result + 1, cString.c_str(), n);
    result[n + 1] = 0;
    result[0] = static_cast<wchar_t>(n);
    return result;
}
