
/*
 Copyright (C) 1998, 1999, 2001, 2002, 2003, 2004 J�r�me Lecomte
 Copyright (C) 2007 Eric Ehlers

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

/*!
\file XlfExcel.inl
\brief Implements inline methods of XlfExcel.
*/

// $Id$

#ifdef NDEBUG
#define INLINE inline
#else
#define INLINE
#endif
#include <iostream>

/*!
The macro EXCEL_BEGIN includes a call to XlfExcel::FreeMemory.

FreeMemory frees *all* memory previously allocated by the
framework. Keeps the biggest buffer allocated (front one) so far
for subsequent calls.

\sa XlfBuffer.
*/
//INLINE void XlfExcel::FreeMemory(bool finished)
//{
//  size_t nbBuffersToKeep = 1;
//  if (finished)
//    nbBuffersToKeep = 0;
//  while (freeList_.size() > nbBuffersToKeep)
//  {
//    delete[] freeList_.back().start;
//    freeList_.pop_back();
//  }
//  offset_ = 0;
//}
// Fixes for thread safety
INLINE void XlfExcel::FreeBuffer(DWORD threadId, size_t nbBuffersToKeep)
{
  BufferList &bufferList = freeList_[threadId];
  while (bufferList.size() > nbBuffersToKeep)
  {
    delete[] bufferList.back().start;
    bufferList.pop_back();
  }
  offset_[threadId] = 0;
}
INLINE void XlfExcel::FreeMemory(bool finished)
{
    if (finished) {
        for (BufferMap::const_iterator i = freeList_.begin(); i != freeList_.end(); ++i) {
            FreeBuffer(i->first, 0);
        }
    } else {
        FreeBuffer(GetCurrentThreadId(), 1);
    }
}

/*!
Allocates a \c new[] buffer and pushes it in front of the list of buffers.
\arg size Size of the new buffer in bytes.
*/
//INLINE void XlfExcel::PushNewBuffer(size_t size)
//{
//  XlfBuffer newBuffer;
//  newBuffer.size = size;
//  newBuffer.start = new char[size];
//  freeList_.push_front(newBuffer);
//  offset_=0;
//#if !defined(NDEBUG)
//    std::cerr << "xlw is allocating a new buffer of " << size << " bytes" << std::endl;
//#endif
//  return;
//}
// Fixes for thread safety
INLINE void XlfExcel::PushNewBuffer(size_t size, BufferList &bufferList)
{
  XlfBuffer newBuffer;
  newBuffer.size = size;
  newBuffer.start = new char[size];
  bufferList.push_front(newBuffer);
  //offset_=0;
#if !defined(NDEBUG)
    std::cerr << "xlw is allocating a new buffer of " << size << " bytes" << std::endl;
#endif
  return;
}

/*!
\param bytes is the size of the chunk required in bytes.
\return the address of the chunk.

Check if the buffer has enough memory, and move the offset of the static
buffer of the ammount of memory requested. If the buffer is full, a new
buffer is allocated whose size is 150% of the one we just filled.

\sa XlfBuffer, XlfExcel::PushNewBuffer(size_t)
*/
//INLINE LPSTR XlfExcel::GetMemory(size_t bytes)
//{
//  if (freeList_.empty())
//    PushNewBuffer(8192);
//  while (1)
//  {
//    XlfBuffer& buffer = freeList_.front();
//    if (offset_ + bytes < buffer.size)
//    {
//      int temp = static_cast<int>(offset_);
//      offset_ += bytes;
//      return buffer.start + temp;
//    }
//    else
//      PushNewBuffer(size_t(buffer.size*1.5));
//  }
//  // should never get to this point...
//  return 0;
//}
// Fixes for thread safety
INLINE LPSTR XlfExcel::GetMemory(size_t bytes)
{
  DWORD threadId = GetCurrentThreadId();
  BufferList &bufferList = freeList_[threadId];
  if (bufferList.empty()) {
    PushNewBuffer(8192, bufferList);
    offset_[threadId]=0;
  }
  while (1)
  {
    XlfBuffer& buffer = bufferList.front();
    if (offset_[threadId] + bytes < buffer.size)
    {
      int temp = static_cast<int>(offset_[threadId]);
      offset_[threadId] += bytes;
      return buffer.start + temp;
    }
    else {
      PushNewBuffer(size_t(buffer.size*1.5), bufferList);
      offset_[threadId]=0;
    }
  }
  // should never get to this point...
  return 0;
}
