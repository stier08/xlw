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
#ifndef XL_ARRAY_H
#define XL_ARRAY_H

#include "xlcall32.h"
#include "MyContainers.h"

namespace xlw {

    //! union of the 2 FP types
    struct xlarray
    {
        union
        {
            FP fp;
            FP12 fp12;
        };
    };


    //! Interpreted as FP (Excel 4) or FP12 (Excel 12).
    typedef xlarray* LPXLARRAY;

    NEMatrix GetMatrix(LPXLARRAY);

}

#endif

