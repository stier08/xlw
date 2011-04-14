/*
 Copyright (C) 2008 2009  Narinder S Claire

 This file is part of XLWDOTNET, a free-software/open-source C# wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLWDOTNET is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#ifndef  XLW_DOT_NET_CELLSTRUCTURE_H
#define  XLW_DOT_NET_CELLSTRUCTURE_H


using namespace System;
using namespace Runtime::InteropServices;
#include<xlw/CellMatrix.h>
#include"xlwTypeBaseClass.h"

#include"managedMyMatrix.h"
#include"managedMyArray.h"

namespace xlwDotNet
{
    namespace xlwTypes
    {
        public ref class CellValue :public xlwTypebaseClass<xlw::CellValue>
        {

        public:
            CellValue(IntPtr theRealThing):
              xlwTypebaseClass<xlw::CellValue>(theRealThing,false){}

            CellValue(String^ theString):
                xlwTypebaseClass<xlw::CellValue>
                    ( new xlw::CellValue(std::wstring((const wchar_t*)(Marshal::StringToHGlobalUni(theString)).ToPointer())) ,true){}

            CellValue(double Number):
                xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue(Number) ,true){}

        //    CellValue(int Code): //Error =  error code
        //        xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue(Code,false) ,true){}

            CellValue(int Code, bool Error) :
                xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue(Code,Error) ,true){}

            CellValue(bool TrueFalse):
                xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue(true) ,true){}

            CellValue(int i):
                xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue(i) ,true){}

            CellValue():
                xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue() ,true){}

            CellValue(const CellValue^ value):
                xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue(*(value->theInner)) ,true){}

           /////      What is it
            property bool IsAString
            {
                bool get()
                {
                    return (theInner->IsAWstring() || theInner->IsAString());
                }
            }

            property bool IsANumber
            {
                bool get()
                {
                    return theInner->IsANumber();
                }
            }
            property bool IsBoolean
            {
                bool get()
                {
                    return theInner->IsBoolean();
                }
            }
            property bool IsXlfOper
            {
                bool get()
                {
                    return theInner->IsXlfOper();
                }
            }

            property bool IsError
            {
                bool get()
                {
                    return theInner->IsError();
                }
            }

            property bool IsEmpty
            {
                bool get()
                {
                    return theInner->IsEmpty();
                }
            }


             /////   Get the Value
            String^ StringValue(){return gcnew String((theInner->StringValue()).c_str())  ;}
            double NumericValue(){return theInner->NumericValue();}
            bool BooleanValue(){return theInner->BooleanValue();}
            UInt64 ErrorValue(){return theInner->ErrorValue();}

            void clear(){theInner->clear();}


        };



        ////// CellMatrix
        public ref class CellMatrix:public xlwTypebaseClass<xlw::CellMatrix>
        {

            public:
            CellMatrix(IntPtr theRealThing):
              xlwTypebaseClass<xlw::CellMatrix>(theRealThing,false){}

            CellMatrix(String^ theString):
                xlwTypebaseClass<xlw::CellMatrix>
                ( new xlw::CellMatrix((const wchar_t*)(Marshal::StringToHGlobalUni(theString)).ToPointer()) ,true){}

            CellMatrix(double Number):
                xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(Number) ,true){}

            //CellMatrix(int Code): //Error =  error code
            //    xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(Code,false) ,true){}

            CellMatrix(xlwTypes::MyArray^ theArray):
                xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(*(theArray->theInner)) ,true){}

            CellMatrix(xlwTypes::MyMatrix^ theMatrix):
                xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(*(theMatrix->theInner)) ,true){}

            CellMatrix(int i):
                xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(i) ,true){}

            CellMatrix():
                xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix() ,true){}

            CellMatrix(int rows, int columns):
                xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(rows, columns) ,true){}

            CellMatrix(CellMatrix^ theCellsMatrix):
                xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(*theCellsMatrix->theInner) ,true){}

            property CellValue^ default[int,int]
            {
                CellValue^ get(int i,int j)
                {
                    CellValue^ result =  gcnew CellValue(IntPtr((void*)(&theInner->operator()(i,j))));
                    owned = true;
                    return result;
                }
                void set(int i, int j, CellValue^ value)
                {
                    theInner->operator()(i,j) = *(value->theInner) ;
                }
            }

        property int RowsInStructure
        {
            int get(){return (int)theInner->RowsInStructure();}
        }

        property int ColumnsInStructure
        {
            int get(){return (int)theInner->ColumnsInStructure();}
        }

        void PushBottom(CellMatrix^ newRows)
        {
            theInner->PushBottom(*(newRows->theInner));
        }


        static void *getInner (CellMatrix^ theArray){return theArray->theInner;}

        };


    }
}

#endif
