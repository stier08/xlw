#ifndef  XLW_DOT_NET_MYMATRIX_H
#define XLW_DOT_NET_MYMATRIX_H 




using namespace System;
#include<xlw/MyContainers.h>
#include"xlwTypeBaseClass.h"


namespace xlwDotNet 
{
	namespace xlwTypes 
	{
		public ref class MyMatrix :public xlwTypebaseClass<xlw::MyMatrix>
		{

		public:
			MyMatrix(IntPtr theRealThing):
			  xlwTypebaseClass<xlw::MyMatrix>(theRealThing,false){}

			MyMatrix(const MyMatrix^ theOther):
			    xlwTypebaseClass<xlw::MyMatrix>(new xlw::MyMatrix(*(theOther->theInner)),true){}

			MyMatrix(int rows_, int cols_):
				xlwTypebaseClass<xlw::MyMatrix>(new xlw::MyMatrix(rows_,cols_),true){}


			property double default[int,int]
			{
				double get(int i,int j) 
				{
					return theInner->operator[](i)[j];
				}
				void set(int i,int j,double val) 
				{
					theInner->operator[](i)[j]=val;
				}
			}
			
		   property int rows
			{
				int get() 
				{
					return theInner->rows();
				}
			}
		   property int columns
			{
				int get() 
				{
					return theInner->columns();
				}
			}

		   operator array<double,2>^()
		   {
			   array<double,2>^ theCSMatrix =  gcnew array<double,2>(theInner->rows(),theInner->columns());
			   for(size_t i(0);i<theInner->rows();++i)
			   {
				   for(size_t j(0);j<theInner->columns();++j)
						theCSMatrix[i,j]=theInner->operator[](i)[j];
			   }
			   return theCSMatrix;
		   }

		   static operator MyMatrix^(array<double,2>^ theCSMatrix)
		   {
			   MyMatrix^ theXLWMatrix =  gcnew MyMatrix(theCSMatrix->GetLength(0),theCSMatrix->GetLength(1));
			    for(int i(0);i<theCSMatrix->GetLength(0);++i)
			   {
				   for(int j(0);j<theCSMatrix->GetLength(1);++j)
					   theXLWMatrix->theInner->operator[](i)[j]=theCSMatrix[i,j];
			   }
				return theXLWMatrix;

		   }
			static void *getInner (MyMatrix^ theArray){return theArray->theInner;}
		};

	}
}

#endif 

