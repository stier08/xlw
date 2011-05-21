
/*
Copyright (C) 2011 John Adcock
Copyright (C) 2011 Narunder S Claire

This file is part of XLW, a free-software/open-source C++ wrapper of the
Excel C API - http://xlw.sourceforge.net/

XLW is free software: you can redistribute it and/or modify it under the
terms of the XLW license.  You should have received a copy of the
license along with this program; if not, please email xlw-users@lists.sf.net

This program is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#include "Test.h"


typedef MatrixTraits<MyMatrix> mtraits;
typedef ArrayTraits<MyArray> atraits;

MyMatrix & choelsky(const MyMatrix  &sym, MyMatrix  &upp){
	size_t i,j,sum_iter;
	double sum;

	size_t rows(mtraits::rows(sym));

	mtraits::setAt(upp,0,0,sqrt(mtraits::getAt(sym,0,0))); 

	for ( j=1;j<rows;j++) 
	{
		mtraits::setAt(upp,0,j, mtraits::getAt(sym,0,j)/mtraits::getAt(upp,0,0));
	}

	for (i=1;i<rows;i++) {  // the outer loop, working along coloms iof L^T
		sum=0;
		for (sum_iter=0;sum_iter<i;sum_iter++)
		{
			sum=sum+mtraits::getAt(upp,sum_iter,i)*mtraits::getAt(upp,sum_iter,i);
		}
		mtraits::setAt(upp,i,i, sqrt(mtraits::getAt(sym,i,i)-sum));

		for(j=i+1;j<rows;j++) 
		{
			sum=0.0;

			for(sum_iter=0;sum_iter<i;sum_iter++)
			{
				sum=sum+mtraits::getAt(upp,sum_iter,i)*mtraits::getAt(upp,sum_iter,j);
			}
			mtraits::setAt(upp,i,j, (mtraits::getAt(sym,i,j)-sum)/mtraits::getAt(upp,i,i));
		}
	}
	return upp;
}

MyMatrix // Returns the Cholesky Decomposition of the matrix
	cholesky(const MyMatrix& inMat) // matrix to decompose
{
	MyMatrix upperMatrix = mtraits::create(mtraits::rows(inMat),mtraits::rows(inMat));
	return choelsky(inMat,upperMatrix);
}

double // computes the inner product of two vectors
	inner_product(const MyArray &x, const MyArray &y)
{
	double sum(0.0);

	for(size_t i(0); i<atraits::size(x); ++i)
	{
		sum += (atraits::getAt(x,i) * atraits::getAt(y,i));
	}
	return sum;
}
