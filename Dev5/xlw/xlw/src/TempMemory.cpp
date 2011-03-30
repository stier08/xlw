// $Id: XlfExcel.inl 507 2008-03-27 13:33:20Z ericehlers $
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

#include "xlw/TempMemory.h"
#include <iostream>
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

static DWORD tlsIndex = TLS_OUT_OF_INDEXES;

namespace xlw {

    char* TempMemory::GetBytes(size_t bytes)
    {
#ifndef NDEBUG
        if(tlsIndex == TLS_OUT_OF_INDEXES)
        {
            throw("DllMainTls not called from DllMain, use the XLW_DLLMAIN_IMPL macro in your xll");
        }
#endif
        TempMemory* threadStorage = (TempMemory*)TlsGetValue(tlsIndex);
        if(!threadStorage)
        {
            threadStorage = new TempMemory;
            TlsSetValue(tlsIndex, (void*)threadStorage);
        }
        return threadStorage->InternalGetMemory(bytes);
    }

    void TempMemory::FreeMemory() {
        TempMemory* threadStorage = (TempMemory*)TlsGetValue(tlsIndex);
        if(threadStorage)
        {
            threadStorage->InternalFreeMemory(false);
        }
    }

    TempMemory::TempMemory() :
        offset_(0) {
    }

    TempMemory::~TempMemory() {
        InternalFreeMemory(true);
    }

    void TempMemory::InternalFreeMemory(bool finished) {
        size_t nbBuffersToKeep = 1;
        if (finished)
            nbBuffersToKeep = 0;
        while (freeList_.size() > nbBuffersToKeep) {
            delete[] freeList_.back().start;
            freeList_.pop_back();
        }
        offset_ = 0;
    }

    void TempMemory::PushNewBuffer(size_t size) {
        XlfBuffer newBuffer;
        newBuffer.size = size;
        newBuffer.start = new char[size];
        freeList_.push_front(newBuffer);
        offset_=0;
    #if !defined(NDEBUG)
        std::cerr << "xlw is allocating a new buffer of " << static_cast<unsigned int>(size) << " bytes" << std::endl;
    #endif
        return;
    }

    char* TempMemory::InternalGetMemory(size_t bytes) {
        if (freeList_.empty())
            PushNewBuffer(8192);
        XlfBuffer& buffer = freeList_.front();
        if (offset_ + bytes > buffer.size) {
            // if we need more space allocate either 50% more than last time
            // or enough to hold all the of data used on previous buffer and this data and a bit of spare
            // space, whichever is greater
            PushNewBuffer(std::max((buffer.size * 3) / 2, (offset_ + bytes) + 4096));
            offset_ = bytes;
            return freeList_.front().start;
        }
        else
        {
            size_t temp = offset_;
            offset_ += bytes;
            return buffer.start + temp;
        }
    }

    void TempMemory::ThreadAttach() {
        TlsSetValue(tlsIndex, 0);
    }

    void TempMemory::ThreadDetach() {
        TempMemory* threadStorage = (TempMemory*)TlsGetValue(tlsIndex);
        if(threadStorage)
        {
            delete threadStorage;
            TlsSetValue(tlsIndex, (void*)0);
        }
    }

}

BOOL DllMainTls(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved) {
    switch (fdwReason) {
        case DLL_PROCESS_ATTACH:
            tlsIndex = TlsAlloc();
            if (tlsIndex == TLS_OUT_OF_INDEXES) {
                return FALSE;
            }
            xlw::TempMemory::ThreadAttach();
            break;

        case DLL_THREAD_ATTACH:
            xlw::TempMemory::ThreadAttach();
            break;

        case DLL_THREAD_DETACH:
            xlw::TempMemory::ThreadDetach();
            break;

        case DLL_PROCESS_DETACH:
            xlw::TempMemory::ThreadDetach();
            TlsFree(tlsIndex);
            break;

        default:
            break;
    }
    return TRUE;
}
