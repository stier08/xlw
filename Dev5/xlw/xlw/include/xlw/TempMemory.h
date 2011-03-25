
/*
 Copyright (C) 1998, 1999, 2001, 2002, 2003, 2004 Jérôme Lecomte
 Copyright (C) 2007, 2008 Eric Ehlers
 Copyright (C) 2009 Narinder S Claire
 Copyright (C) 2010 John Adcock

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/


#ifndef INC_TempMemory_H
#define INC_TempMemory_H

/*!
\file TempMemory.h
\brief Declares class TempMemory.
*/

// $Id: XlfExcel.h 478 2008-03-06 11:43:12Z ericehlers $

#include <list>

#if defined(_MSC_VER)
#pragma once
#endif

namespace xlw {

    class TempMemory
    {
    public:
        //! \name Memory management
        //@{
        //! Allocates memory in the framework temporary buffer
        template<typename TYPE>
        static TYPE* GetMemory(size_t numItems = 1)
        {
            return reinterpret_cast<TYPE*>(GetBytes(numItems * sizeof(TYPE)));
        }
        //! Frees temporary memory used by the XLL
        static void FreeMemory();

        //! \name Functions called from DllMain
        //@{
        //! Another thread has started so create the TLS
        static void ThreadAttach();
        //! A thread/process has stoppped so clean up
        static void ThreadDetach();
        //@}

    private:
        //! \name Structors and static members
        //@{
        //! Ctor.
        TempMemory();
        //! Dtor.
        ~TempMemory();
        //@}

        //! Allocates memory in the framework temporary buffer
        static char* GetBytes(size_t bytes);

        char* InternalGetMemory(size_t bytes);
        //! Frees temporary memory used by the XLL
        void InternalFreeMemory(bool finished=false);

        //! Memory buffer used to store data that are passed to Excel
        /*!
        When we pass XLOPER from the XLL towards Excel we should not use
        XLL local variables as it would be freed before MSExcel could get
        the data.

        Therefore the framework C library that comes with \ref XLSDK97 "Excel
        97 developer's kit" suggests to use a static area where the XlfOper
        are stored (see XlfExcel::GetMemory) and still available to
        Excel when we exit the XLL routine. This array is then reset
        by a call to XlfExcel::FreeMemory at the begining of each new
        call of one of the XLL functions on each thread.

        \sa XlfExcel::GetMemory, XlfExcel::FreeMemory
        */
        struct XlfBuffer
        {
            //! Size of the buffer.
            size_t size;
            //! Start address.
            char * start;
        };

        //! A list of buffers.
        typedef std::list<XlfBuffer> BufferList;
        //! Internal memory buffer holding memory to be referenced by Excel (excluded from the pimpl to allow inlining).
        BufferList freeList_;
        //! Pointer to next free area (excluded from the pimpl to allow inlining).
        size_t offset_;

        //! Create a new static buffer and add it to the free list.
        void PushNewBuffer(size_t);
    };
}

#endif

