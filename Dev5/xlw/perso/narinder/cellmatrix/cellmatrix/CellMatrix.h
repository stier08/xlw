
#ifndef CELL_MATRIX_H
#define CELL_MATRIX_H

#include "CellMatrixPimpl.h"
#include "CellValue.h"
#include <xlw/MyContainers.h>
#include <string>
#include <vector>

namespace xlw {

   
    class CellMatrix
    {
    public:
		
		//Copy Constructor
		CellMatrix(const CellMatrix &theOther):pimpl(theOther.pimpl.copy()){}


		CellMatrix():pimpl(new MJCellMatrix()){}
		CellMatrix(size_t rows, size_t columns ):pimpl(new CellMatrixImpl(rows,cols)){}

        explicit CellMatrix(double x):pimpl(new CellMatrixImpl(1,1))
		{	
			
		}
        explicit CellMatrix(const std::string & x):pimpl(new MJCellMatrix(x)){}
        explicit CellMatrix(const std::wstring & x):pimpl(new MJCellMatrix(x)){}
        explicit CellMatrix(const char* x):pimpl(new MJCellMatrix(std::string (x))){}
        explicit CellMatrix(const MyArray& data):pimpl(new MJCellMatrix(data)){}
        explicit CellMatrix(const MyMatrix& data):pimpl(new MJCellMatrix(data)){}
        explicit CellMatrix(unsigned long i):pimpl(new MJCellMatrix(i)){}
        explicit CellMatrix(int i):pimpl(new MJCellMatrix(i)){}



        const CellValue& operator()(size_t i, size_t j) const
		{
			return pimpl->operator()(i,j);
		}
		CellValue& operator()(size_t i, size_t j) 
		{	
			return pimpl->operator()(i,j);
		}

        size_t RowsInStructure() const
		{
			return pimpl->RowsInStructure();
		}
        size_t ColumnsInStructure() const
		{
			return pimpl->ColumnsInStructure();
		}

        //void PushBottom(const CellMatrix& newRows);
		
	private:
		CellMatrix_pimpl_abstract::pimpl pimpl;


    };

}

#endif
