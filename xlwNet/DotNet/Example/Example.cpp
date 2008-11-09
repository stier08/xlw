
/*
 Copyright (C) 2008 Narinder Claire

 This file is part of XLW.NET 

 XLW.NET is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#include <xlwDotNet.h>
using namespace System;
using namespace Runtime::InteropServices;
using namespace xlwDotNet;

std::wstring //tests empty args 
 DLLEXPORT EmptyArgFunction()
{
DOT_NET_EXCEL_BEGIN 
        return (const wchar_t*)(Marshal::StringToHGlobalUni(Example::Class1::EmptyArgFunction(
        )).ToPointer());
DOT_NET_EXCEL_END
}
short //echoes a short 
 DLLEXPORT EchoShort(short x)
{
DOT_NET_EXCEL_BEGIN 
        return Example::Class1::EchoShort(
                x
        );
DOT_NET_EXCEL_END
}
MyMatrix //echoes a matrix
 DLLEXPORT EchoMat(const MyMatrix&  Echoee)
{
DOT_NET_EXCEL_BEGIN 
        return *(MyMatrix*)(xlwTypes::MyMatrix::getInner(Example::Class1::EchoMat(
                gcnew xlwTypes::MyMatrix(IntPtr((void*)&Echoee))
        )));
DOT_NET_EXCEL_END
}
MyArray //echoes an array
 DLLEXPORT EchoArray(const MyArray&  Echoee)
{
DOT_NET_EXCEL_BEGIN 
        return *(MyArray*)(xlwTypes::MyArray::getInner(Example::Class1::EchoArray(
                gcnew xlwTypes::MyArray(IntPtr((void*)&Echoee))
        )));
DOT_NET_EXCEL_END
}
CellMatrix //echoes a  CellMatrix
 DLLEXPORT EchoCells(const CellMatrix&  Echoee)
{
DOT_NET_EXCEL_BEGIN 
        return *(CellMatrix*)(xlwTypes::CellMatrix::getInner(Example::Class1::EchoCells(
                gcnew xlwTypes::CellMatrix(IntPtr((void*)&Echoee))
        )));
DOT_NET_EXCEL_END
}
double //computes the circumference of a circle 
 DLLEXPORT Circ(double Diameter)
{
DOT_NET_EXCEL_BEGIN 
        return Example::Class1::Circ(
                Diameter
        );
DOT_NET_EXCEL_END
}
std::wstring //Concatenates two strings
 DLLEXPORT Concat(const std::wstring&  str1,const std::wstring&  str2)
{
DOT_NET_EXCEL_BEGIN 
        return (const wchar_t*)(Marshal::StringToHGlobalUni(Example::Class1::Concat(
                 gcnew String(str1.c_str()),
                 gcnew String(str2.c_str())
        )).ToPointer());
DOT_NET_EXCEL_END
}
MyArray //computes mean and variance of a range 
 DLLEXPORT stats(const MyArray&  data)
{
DOT_NET_EXCEL_BEGIN 
        return *(MyArray*)(xlwTypes::MyArray::getInner(Example::Class1::stats(
                gcnew xlwTypes::MyArray(IntPtr((void*)&data))
        )));
DOT_NET_EXCEL_END
}
double //throws an exception
 DLLEXPORT Exception()
{
DOT_NET_EXCEL_BEGIN 
        return Example::Class1::Exception(
        );
DOT_NET_EXCEL_END
}
std::wstring //says hello name
 DLLEXPORT HelloWorldAgain(const std::wstring&  name)
{
DOT_NET_EXCEL_BEGIN 
        return (const wchar_t*)(Marshal::StringToHGlobalUni(Example::Class1::HelloWorldAgain(
                 gcnew String(name.c_str())
        )).ToPointer());
DOT_NET_EXCEL_END
}
CellMatrix //echoes arg list
 DLLEXPORT EchoArgList(const ArgumentList&  args)
{
DOT_NET_EXCEL_BEGIN 
        return *(CellMatrix*)(xlwTypes::CellMatrix::getInner(Example::Class1::EchoArgList(
                gcnew xlwTypes::ArgumentList(IntPtr((void*)&args))
        )));
DOT_NET_EXCEL_END
}
MyArray //takes double[]
 DLLEXPORT CastToCSArray(const MyArray&  csarray)
{
DOT_NET_EXCEL_BEGIN 
        return *(MyArray*)(xlwTypes::MyArray::getInner(Example::Class1::CastToCSArray(
                 xlwTypes::MyArray(IntPtr((void*)&csarray))
        )));
DOT_NET_EXCEL_END
}
MyArray //takes double[]
 DLLEXPORT CastToCSArrayTwice(const MyArray&  csarray)
{
DOT_NET_EXCEL_BEGIN 
        return *(MyArray*)(xlwTypes::MyArray::getInner(Example::Class1::CastToCSArrayTwice(
                 xlwTypes::MyArray(IntPtr((void*)&csarray))
        )));
DOT_NET_EXCEL_END
}
MyMatrix //takes double[,]
 DLLEXPORT CastToCSMatrix(const MyMatrix&  csmatrix)
{
DOT_NET_EXCEL_BEGIN 
        return *(MyMatrix*)(xlwTypes::MyMatrix::getInner(Example::Class1::CastToCSMatrix(
                 xlwTypes::MyMatrix(IntPtr((void*)&csmatrix))
        )));
DOT_NET_EXCEL_END
}
MyMatrix //takes double[,]
 DLLEXPORT CastToCSMatrixTwice(const MyMatrix&  csmatrix)
{
DOT_NET_EXCEL_BEGIN 
        return *(MyMatrix*)(xlwTypes::MyMatrix::getInner(Example::Class1::CastToCSMatrixTwice(
                 xlwTypes::MyMatrix(IntPtr((void*)&csmatrix))
        )));
DOT_NET_EXCEL_END
}
double //throws an exception of type ArgumentNullException
 DLLEXPORT throwString(const std::wstring&  err)
{
DOT_NET_EXCEL_BEGIN 
        return Example::Class1::throwString(
                 gcnew String(err.c_str())
        );
DOT_NET_EXCEL_END
}
double //throws an exception of type cellMatrixException
 DLLEXPORT throwCellMatrix(const std::wstring&  err)
{
DOT_NET_EXCEL_BEGIN 
        return Example::Class1::throwCellMatrix(
                 gcnew String(err.c_str())
        );
DOT_NET_EXCEL_END
}
double //makes the C Runtime throw an exception
 DLLEXPORT throwCError(const std::wstring&  err)
{
DOT_NET_EXCEL_BEGIN 
        return Example::Class1::throwCError(
                 gcnew String(err.c_str())
        );
DOT_NET_EXCEL_END
}
