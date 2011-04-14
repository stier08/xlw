/*
 Copyright (C) 2011  John Adcock

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

namespace xlw { namespace Impl {

    template <>
    struct XlfOperProperties<LPXLOPER12>
    {
        typedef int ErrorType;
        typedef int IntType;
        typedef RW MultiRowType;
        typedef COL MultiColType;
        typedef DWORD XlTypeType;
        typedef XLOPER12 OperType;

        static XlTypeType getXlType(LPXLOPER12 oper)
        {
            return oper->xltype;
        }
        static void setXlType(LPXLOPER12 oper, XlTypeType newValue)
        {
            oper->xltype = newValue;
        }
        static double getDouble(LPXLOPER12 oper)
        {
            return oper->val.num;
        }
        static void setDouble(LPXLOPER12 oper, double newValue)
        {
            oper->val.num = newValue;
            oper->xltype = xltypeNum;
        }
        static IntType getInt(LPXLOPER12 oper)
        {
            return oper->val.w;
        }
        static void setInt(LPXLOPER12 oper, IntType newValue)
        {
            oper->val.w = newValue;
            oper->xltype = xltypeInt;
        }
        static bool getBool(LPXLOPER12 oper)
        {
            return !!oper->val.xbool;
        }
        static void setBool(LPXLOPER12 oper, bool newValue)
        {
            oper->val.xbool = newValue;
            oper->xltype = xltypeBool;
        }
        static ErrorType getError(LPXLOPER12 oper)
        {
            return oper->val.err;
        }
        static void setError(LPXLOPER12 oper, ErrorType newValue)
        {
            oper->val.err = newValue;
            oper->xltype = xltypeErr;
        }
        static RW getRows(LPXLOPER12 oper)
        {
            switch(oper->xltype & 0xFFF)
            {
            case xltypeMulti:
                return oper->val.array.rows;
                break;

            case xltypeSRef:
                return oper->val.sref.ref.rwLast - oper->val.sref.ref.rwFirst + 1;
                break;

            case xltypeRef:
                if(oper->val.mref.lpmref->count == 1)
                {
                    return oper->val.mref.lpmref->reftbl[0].rwLast - oper->val.mref.lpmref->reftbl[0].rwFirst + 1;
                }
                else
                {
                    XlfException("No implementation for multiple Refs");
                }
                break;

            case xltypeNum:
            case xltypeStr:
            case xltypeBool:
            case xltypeInt:
                return 1;
                break;

            case xltypeErr:
            case xltypeMissing:
            case xltypeNil:
                return 0;
                break;
            default:
                break;
            }
            throw XlfException("No implementation on XlfOper rows");
        }
        static void setArraySize(LPXLOPER12 oper, RW rows, COL cols)
        {
            if(rows > 0 && cols > 0)
            {
                if(cols > 16384)
                {
                    std::cerr << "Truncating columns to 16384" << std::endl;
                    cols = 16384;
                }
                if(rows > 1048576)
                {
                    std::cerr << "Truncating rows to 1048576" << std::endl;
                    rows = 1048576;
                }

                oper->val.array.lparray = TempMemory::GetMemory<XLOPER12>(rows * cols);
                oper->val.array.rows = rows;
                oper->val.array.columns = cols;
                oper->xltype = xltypeMulti;
            }
            else
            {
                oper->xltype = xltypeMissing;
            }
        }
        static COL getCols(LPXLOPER12 oper)
        {
            switch(oper->xltype & 0xFFF)
            {
            case xltypeMulti:
                return oper->val.array.columns;
                break;

            case xltypeSRef:
                return oper->val.sref.ref.colLast - oper->val.sref.ref.colFirst + 1;
                break;

            case xltypeRef:
                if(oper->val.mref.lpmref->count == 1)
                {
                    return oper->val.mref.lpmref->reftbl[0].colLast - oper->val.mref.lpmref->reftbl[0].colFirst +1;
                }
                else
                {
                    XlfException("No implementation for multiple Refs");
                }
                break;

            case xltypeNum:
            case xltypeStr:
            case xltypeBool:
            case xltypeInt:
                return 1;
                break;

            case xltypeErr:
            case xltypeMissing:
            case xltypeNil:
                return 0;
                break;
            default:
                break;
            }
            throw XlfException("No implementation on XlfOper rows");
        }
        static LPXLOPER12 getElement(LPXLOPER12 oper, RW row, COL column)
        {
#ifndef NDEBUG
            // fasten seat belts when not in release mode
            if(row < 0 || row >= getRows(oper) || column < 0 || column >= getCols(oper))
            {
                throw XlfOutOfBounds();
            }
#endif
            switch(oper->xltype & 0xFFF)
            {
            case xltypeMulti:
                return oper->val.array.lparray + row * oper->val.array.columns + column;
                break;

            case xltypeSRef:
                {
                    LPXLOPER12 result = TempMemory::GetMemory<XLOPER12>();
                    *result = *oper;
                    result->val.sref.ref.rwFirst += row;
                    result->val.sref.ref.rwLast = result->val.sref.ref.rwFirst;
                    result->val.sref.ref.colFirst += column;
                    result->val.sref.ref.colLast = result->val.sref.ref.colFirst;
                }
                break;

            case xltypeRef:
                if(oper->val.mref.lpmref->count == 1)
                {
                    LPXLOPER12 result = TempMemory::GetMemory<XLOPER12>();
                    *result = *oper;
                    result->val.mref.lpmref->reftbl[0].rwFirst += row;
                    result->val.mref.lpmref->reftbl[0].rwLast = result->val.mref.lpmref->reftbl[0].rwFirst;
                    result->val.mref.lpmref->reftbl[0].colFirst += column;
                    result->val.mref.lpmref->reftbl[0].colLast = result->val.mref.lpmref->reftbl[0].colFirst;
                }
                else
                {
                    XlfException("No implementation for multiple Refs");
                }
                break;

            case xltypeNum:
            case xltypeStr:
            case xltypeBool:
            case xltypeInt:
                return oper;
                break;

            default:
                break;
            }
            throw XlfException("Wrong type for element by element access ");
        }
        static std::string getString(LPXLOPER12 oper)
        {
            return PascalStringConversions::WPascalStringToString(oper->val.str);
        }
        static void setString(LPXLOPER12 oper, const std::string& newValue)
        {
            oper->val.str = PascalStringConversions::StringToWPascalString(newValue);
            oper->xltype = xltypeStr;
        }
        static std::wstring getWString(LPXLOPER12 oper)
        {
            return PascalStringConversions::WPascalStringToWString(oper->val.str);
        }
        static void setWString(LPXLOPER12 oper, const std::wstring& newValue)
        {
            oper->val.str = PascalStringConversions::WStringToWPascalString(newValue);
            oper->xltype = xltypeStr;
        }
        static XlfRef getRef(LPXLOPER12 oper)
        {
            const XLREF12& ref = oper->val.mref.lpmref->reftbl[0];
            return XlfRef (ref.rwFirst,  // top
                            ref.colFirst, // left
                            ref.rwLast,   // bottom
                            ref.colLast,  // right
                            oper->val.mref.idSheet); // sheet id
            return XlfRef();
        }
        static int coerce(LPXLOPER12 fromOper, XlTypeType toType, LPXLOPER12 toOper)
        {
            XLOPER12 typeOper;
            typeOper.val.w = toType;
            typeOper.xltype = xltypeInt;
            return XlfExcel::Instance().Call12(xlCoerce, toOper, 2, fromOper, &typeOper);
        }
        static void XlFree(LPXLOPER12 oper)
        {
            XlfExcel::Instance().Call12(xlFree, 0, 1, oper);
        }
        static void copy(LPXLOPER12 fromOper, LPXLOPER12 toOper)
        {
            *toOper = *fromOper;
        }
    };

} }