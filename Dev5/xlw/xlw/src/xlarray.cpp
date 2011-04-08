/*
 Copyright (C) 2006 Mark Joshi
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
#include <xlw/xlarray.h>
#include <xlw/xlfExcel.h>

xlw::NEMatrix xlw::GetMatrix(LPXLARRAY input)
{
    if(xlw::XlfExcel::Instance().excel12())
    {
        int rows = input->fp12.rows;
        int cols = input->fp12.columns;

        NEMatrix result(rows,cols);
        for (int i=0; i < rows; ++i)
            for (int j=0; j < cols; ++j)
            {
                int k = i*cols+j;
                double val = input->fp12.array[k];
                result(i,j)= val;
            }
        return result;
    }
    else
    {
        int rows = input->fp.rows;
        int cols = input->fp.columns;

        NEMatrix result(rows,cols);
        for (int i=0; i < rows; ++i)
            for (int j=0; j < cols; ++j)
            {
                int k = i*cols+j;
                double val = input->fp.array[k];
                result(i,j)= val;
            }
        return result;
    }
}
