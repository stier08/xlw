
#ifndef  XLW_TYPE_BASE_CLASS_H
#define XLW_TYPE_BASE_CLASS_H 


using namespace System;
namespace xlwDotNet 
{
	namespace xlwTypes 
	{
		template<class T>
		public ref class xlwTypebaseClass
		{
		public :
			xlwTypebaseClass(IntPtr theManagedPointer, bool owned_):
				theInner((T *)(theManagedPointer.ToPointer())),owned(owned_){}


			xlwTypebaseClass(const T * theObject, bool owned_):
				theInner((T *)(theObject)),owned(owned_){}

		    property void * inner
			{
				void * get() 
				{
					return (void *)theInner;
				}
			}
			 ~xlwTypebaseClass()
			{
				if(owned)delete theInner;
				theInner = 0;owned = false;
			}
			T * theInner; 
			bool owned;


		};
	}
}

#endif 