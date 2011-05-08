/*
Copyright (C) 2006 Mark Joshi
Copyright (C) 2007, 2008 Eric Ehlers
Copyright (C) 2009 2011 Narinder S Claire
Copyright (C) 2011 John Adcock



This file is part of XLW, a free-software/open-source C++ wrapper of the
Excel C API - http://xlw.sourceforge.net/

XLW is free software: you can redistribute it and/or modify it under the
terms of the XLW license.  You should have received a copy of the
license along with this program; if not, please email xlw-users@lists.sf.net

This program is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
FOR A PARTICULAR PURPOSE.  See the license for more details.
*/
#ifndef CELL_MATRIX_H
#define CELL_MATRIX_H


#include "CellValue.h"
#include "MJCellMatrix.h"
#include <xlw/MyContainers.h>
#include <string>
#include <vector>

namespace xlw {

	typedef impl::MJCellMatrix defaultGenericCellMatrixImpl;

	template<class inner_impl = impl::MJCellMatrix>
	class GenericCellMatrix
	{
	public:

		//Copy Constructor
		GenericCellMatrix(const GenericCellMatrix &theOther):pimpl(theOther.pimpl.copy()){}

		template<class T>
		GenericCellMatrix(const GenericCellMatrix<T> &theOther):pimpl(theOther.pimpl.copy()){}


		GenericCellMatrix(size_t rows, size_t columns):pimpl(new inner_impl(rows, columns)){}

		GenericCellMatrix():pimpl(new inner_impl()){}


		GenericCellMatrix(double data):pimpl(new inner_impl(1,1))
		{
			(*pimpl)(0,0)=data;
		}

		GenericCellMatrix(const std::string &  data):pimpl(new inner_impl(1,1))
		{
			(*pimpl)(0,0)=data;
		}

		GenericCellMatrix(const std::wstring &  data):pimpl(new inner_impl(1,1))
		{
			(*pimpl)(0,0)=data;
		}
		
		GenericCellMatrix(const char* data):pimpl(new inner_impl(1,1))
		{
			(*pimpl)(0,0)=std::string(data);
		}
		
		GenericCellMatrix(const MyArray& data):
		pimpl(new defaultGenericCellMatrixImpl(ArrayTraits<MyArray>::size(data),1))
		{
			for(size_t i(0); i < ArrayTraits<MyArray>::size(data); ++i)
			{
				(*pimpl)(i,0) = ArrayTraits<MyArray>::getAt(data,i);
			}
		}
		
		GenericCellMatrix(const MyMatrix& data):
		pimpl(new defaultGenericCellMatrixImpl(MatrixTraits<MyMatrix>::rows(data),MatrixTraits<MyMatrix>::columns(data)))
		{

			for(size_t i(0); i < MatrixTraits<MyMatrix>::rows(data); ++i)
			{
				for(size_t j(0); j < MatrixTraits<MyMatrix>::columns(data); ++j)
				{
					(*pimpl)(i,j) = MatrixTraits<MyMatrix>::getAt(data,i,j);
				}
			}

		}
		
		GenericCellMatrix(unsigned long data):pimpl(new inner_impl(1,1))
		{
			(*pimpl)(0,0)=data;
		}
		
		GenericCellMatrix(int data):pimpl(new inner_impl(1,1))
		{
			(*pimpl)(0,0)=data;
		}

		GenericCellMatrix & operator=(const GenericCellMatrix<inner_impl> &theOther)
		{
            GenericCellMatrix<inner_impl> temp(theOther);
			temp.swap(*this);
			return *this;
		}

		template<class T>
	    GenericCellMatrix & operator=(const GenericCellMatrix<T> &theOther)
		{
			GenericCellMatrix<T> temp(theOther);
			temp.swap(*this);
			return *this;
		}

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

		template<class T>
		void PushBottom(const GenericCellMatrix<T> & newRows)
		{
			GenericCellMatrix temp(*this);
			temp.pimpl->PushBottom(*(newRows.pimpl));
			temp.swap(*this);

		}

		template<class T>
		void swap(GenericCellMatrix<T> &theOther)
		{
			pimpl.swap(theOther.pimpl);
		}

	private:
		eshared_ptr<CellMatrix_pimpl_abstract> pimpl;

	};

	template<class T, class U>
	GenericCellMatrix<T> MergeCellMatrices(const GenericCellMatrix<T>& Top, const GenericCellMatrix<U>& Bottom)
	{
		GenericCellMatrix<T> temp(Top);
		temp.PushBottom(Bottom);
		return temp;
	}

	typedef GenericCellMatrix<defaultGenericCellMatrixImpl> CellMatrix;

}

#endif
