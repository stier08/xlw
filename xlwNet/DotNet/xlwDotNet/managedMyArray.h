#ifndef  XLW_DOT_NET_MYARRAY_H
#define XLW_DOT_NET_MYARRAY_H 


using namespace System;
#include<xlw/MyContainers.h>
#include"xlwTypeBaseClass.h"
namespace xlwDotNet 
{
	namespace xlwTypes 
	{
		public ref class MyArray :public xlwTypebaseClass<xlw::MyArray>
		{

		public:
			MyArray(IntPtr theRealThing):
			  xlwTypebaseClass<xlw::MyArray>(theRealThing,false){}

			MyArray(const MyArray^ theOther):
			    xlwTypebaseClass<xlw::MyArray>(new xlw::MyArray(*(theOther->theInner)),true){}

			MyArray(int size_):
				xlwTypebaseClass<xlw::MyArray>(new xlw::MyArray(size_),true){}

			void clear(){theInner->clear();}

			property double default[int]
			{
				double get(int i) 
				{
					return theInner->operator[](i);
				}
				void set(int i,double val) 
				{
					theInner->operator[](i)=val;
				}
			}
			
		   property int size
			{
				int get() 
				{
					return theInner->size();
				}
			}	

		   operator array<double>^()
		   {
			   array<double>^ theCSArray =  gcnew array<double>(theInner->size());
			   for(size_t i(0);i<theInner->size();++i)theCSArray[i]=theInner->operator[](i);
			   return theCSArray;
		   }
		   static operator MyArray^(array<double>^ theCSArray)
		   {
			   MyArray^ theXLWArray =  gcnew MyArray(theCSArray->Length);
			   for(int i(0);i<theCSArray->Length;++i)theXLWArray->theInner->operator[](i)=theCSArray[i];
			   return theXLWArray;

		   }
		   static void *getInner (MyArray^ theArray){return theArray->theInner;}
		};

	}
}
#endif