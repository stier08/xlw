/*
 Copyright (C) 2008 Narinder Claire

 This file is part of XLW.NET 

 XLW.NET is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
 */#
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
				( new xlw::CellValue((const wchar_t*)(Marshal::StringToHGlobalUni(theString)).ToPointer()) ,true){}

			CellValue(double Number):
				xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue(Number) ,true){}

			CellValue(UInt32 Code): //Error =  error code
				xlwTypebaseClass<xlw::CellValue>( new xlw::CellValue(Code,false) ,true){}

			CellValue(UInt32 Code, bool Error) :
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

			CellMatrix(UInt32 Code): //Error =  error code
				xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(Code,false) ,true){}

			CellMatrix(xlwTypes::MyArray^ theArray):
				xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(*(theArray->theInner)) ,true){}

			CellMatrix(xlwTypes::MyMatrix^ theMatrix):
				xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(*(theMatrix->theInner)) ,true){}

			CellMatrix(int i):
				xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(i) ,true){}

			CellMatrix():
				xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix() ,true){}

			CellMatrix(UInt32 rows, UInt32 columns):
				xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(rows, columns) ,true){}

			CellMatrix(CellMatrix^ theCellsMatrix):
				xlwTypebaseClass<xlw::CellMatrix>( new xlw::CellMatrix(*theCellsMatrix->theInner) ,true){}

			property CellValue^ default[UInt32,UInt32]
			{
				CellValue^ get(UInt32 i,UInt32 j) 
				{
					CellValue^ result =  gcnew CellValue(IntPtr((void*)(&theInner->operator()(i,j))));
					owned = true;
					return result;
				}
				void set(UInt32 i, UInt32 j, CellValue^ value)
				{
					theInner->operator()(i,j) = *(value->theInner) ;
				}
			}

        property UInt32 RowsInStructure
		{
			UInt32 get(){return theInner->RowsInStructure();}
		}

        property UInt32 ColumnsInStructure
		{
			UInt32 get(){return theInner->ColumnsInStructure();}
		}

	/*
			operator array<object^,2>^()
		   {
			   array<object^,2>^ theCSMatrix =  gcnew array<object^,2>(theInner->rows(),theInner->columns());
			   for(int i(0);i<theInner->rows();++i)
			   {
				   for(int j(0);j<theInner->columns();++j)
						theCSMatrix[i,j]=theInner->operator[](i)[j];
			   }
			   return theCSMatrix;
		   }
*/
		  // static operator MyMatrix^(array<double,2>^ theCSMatrix)
		   //{
			//   MyMatrix^ theXLWMatrix =  gcnew MyMatrix(theCSMatrix->GetLength(0),theCSMatrix->GetLength(1));
			//    for(int i(0);i<theCSMatrix->GetLength(0);++i)
			//   {
			//	   for(int j(0);j<theCSMatrix->GetLength(1);++j)
			//		   theXLWMatrix->theInner->operator[](i)[j]=theCSMatrix[i,j];
			//   }
			//	return theXLWMatrix;

		  // }
				
			static void *getInner (CellMatrix^ theArray){return theArray->theInner;}



		};


	}
}

#endif
