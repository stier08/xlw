/*
 Copyright (C) 1998, 1999, 2001, 2002, 2003, 2004 J�r�me Lecomte

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/


#ifndef INC_XlfRef_H
#define INC_XlfRef_H

/*!
\file XlfRef.h
\brief Declares the XlfRef class.
*/

// $Id$

#include <xlw/EXCEL32_API.h>
#include <xlw/xlcall32.h>

#if defined(_MSC_VER)
#pragma once
#endif

// Forward declaration.
//! Wrapper around XLOPER Excel data type.
class EXCEL32_API XlfOper;

//! Encapsulate a range of cells.
/*!
A range is actually a reference to a range of cells in the spreadsheet.
This range is stored as an absolute reference even if you can access
the elements relatively to the upper left corner of the range (starting
at 0 to number of row/column minus 1).

The dtor, copy ctor and assignment otor are generated by the compiler.

\note The Excel API is limited to the first 256 columns. This class holds 
a reference to a single range in a single spreadsheet. It is intended to 
be a helper class to refer range through XlfOper.

\note It is currently not possible for XlfRef to handle union of range
in the same way other Excel function can. This feature remains to be 
implemented.

*/
class EXCEL32_API XlfRef
{
public:
  //! Default ctor.
  XlfRef();
  //! Absolute reference ctor.
  XlfRef(WORD top, BYTE left, WORD bottom, BYTE right, DWORD sheetId = 0);
  //! Absolute reference ctor, to a single cell.
  XlfRef(WORD row, BYTE col, DWORD sheetId = 0);
  //! Gets the first row of the range (0 based).
  WORD GetRowBegin() const;
  //! Gets passed the last row of the range (0 based).
  WORD GetRowEnd() const;
  //! Gets the first column of the range (0 based).
  BYTE GetColBegin() const;
  //! Gets passed the last column of the range (0 based).
  BYTE GetColEnd() const;
  //! Gets the number of columns.
  BYTE GetNbCols() const;
  //! Gets the number of rows.
  WORD GetNbRows() const;
  //! Gets MS Excel sheet identifier of the range.
  DWORD GetSheetId() const;
  //! Sets the first row of the range.
  void SetRowBegin(WORD rowbegin);
  //! Sets passed the last row of the range.
  void SetRowEnd(WORD rowend);
  //! Sets the first column of the range.
  void SetColBegin(BYTE colbegin);
  //! Sets passed the last column of the range.
  void SetColEnd(BYTE colend);
  //! Sets MS Excel sheet identifier of the range.
  void SetSheetId(DWORD);
  //! Access operator
  XlfOper operator()(WORD relativerow, BYTE relativecol) const;

private:
  //! Index of the top row of the range reference.
  WORD rowbegin_;
  //! Index of one past the last row of the range reference.
  WORD rowend_;
  //! Index of the left most column of the range reference.
  BYTE colbegin_;
  //! Index of one past the right most column of the range reference.
  BYTE colend_;
  //! Index of the sheet the reference is pointing to.
  DWORD sheetId_;
};

#ifdef NDEBUG
#include <xlw/XlfRef.inl>
#endif

#endif