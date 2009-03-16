using namespace System;
#include"managedCellmatrix.h"

namespace xlwDotNet 
{
	public ref class cellMatrixException :public System::Exception
	{
	public:
		cellMatrixException(String^ message, xlwTypes::CellMatrix ^ matrix):
		  System::Exception(message), theMatrix(matrix->theInner){}

	    property IntPtr inner
		{
			IntPtr get()
			{
				return IntPtr(theMatrix);
			}


		}

		xlw::CellMatrix *theMatrix;

	};
}
