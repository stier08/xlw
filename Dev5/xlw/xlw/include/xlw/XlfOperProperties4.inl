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
    struct XlfOperProperties<LPXLOPER>
    {
        typedef WORD ErrorType;
        typedef short int IntType;
        typedef WORD MultiRowType;
        typedef WORD MultiColType;
        typedef WORD XlTypeType;
        typedef XLOPER OperType;

        static XlTypeType getXlType(LPXLOPER oper)
        {
            return oper->xltype;
        }
        static void setXlType(LPXLOPER oper, XlTypeType newValue)
        {
            oper->xltype = newValue;
        }
        static double getDouble(LPXLOPER oper)
        {
            return oper->val.num;
        }
        static void setDouble(LPXLOPER oper, double newValue)
        {
            oper->val.num = newValue;
            oper->xltype = xltypeNum;
        }
        static IntType getInt(LPXLOPER oper)
        {
            return oper->val.w;
        }
        static void setInt(LPXLOPER oper, IntType newValue)
        {
            oper->val.w = newValue;
            oper->xltype = xltypeInt;
        }
        static bool getBool(LPXLOPER oper)
        {
            return !!oper->val.xbool;
        }
        static void setBool(LPXLOPER oper, bool newValue)
        {
            oper->val.xbool = newValue;
            oper->xltype = xltypeBool;
        }
        static ErrorType getError(LPXLOPER oper)
        {
            return oper->val.err;
        }
        static void setError(LPXLOPER oper, ErrorType newValue)
        {
            oper->val.err = newValue;
            oper->xltype = xltypeErr;
        }
        static RW getRows(LPXLOPER oper)
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
        static void setArraySize(LPXLOPER oper, RW rows, COL cols)
        {
            if(rows > 0 && cols > 0)
            {
                if(cols > 255)
                {
                    std::cerr << "Truncating columns to 255" << std::endl;
                    cols = 255;
                }
                if(rows > 65536)
                {
                    std::cerr << "Truncating rows to 65536" << std::endl;
                    rows = 65536;
                }
                oper->val.array.lparray = TempMemory::GetMemory<XLOPER>(rows * cols);
                oper->val.array.rows = rows;
                oper->val.array.columns = cols;
                oper->xltype = xltypeMulti;
            }
            else
            {
                oper->xltype = xltypeMissing;
            }
        }
        static COL getCols(LPXLOPER oper)
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
                    return oper->val.mref.lpmref->reftbl[0].colLast - oper->val.mref.lpmref->reftbl[0].colFirst + 1;
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
        static LPXLOPER getElement(LPXLOPER oper, RW row, COL column)
        {
            switch(oper->xltype & 0xFFF)
            {
            case xltypeMulti:
                return oper->val.array.lparray + row * oper->val.array.columns + column;
                break;

            case xltypeSRef:
                {
                    LPXLOPER result = TempMemory::GetMemory<XLOPER>();
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
                    LPXLOPER result = TempMemory::GetMemory<XLOPER>();
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
        static std::string getString(LPXLOPER oper)
        {
            return PascalStringConversions::PascalStringToString(oper->val.str);
        }
        static void setString(LPXLOPER oper, const std::string& newValue)
        {
            oper->val.str = PascalStringConversions::StringToPascalString(newValue);
            oper->xltype = xltypeStr;
        }
        static std::wstring getWString(LPXLOPER oper)
        {
            return PascalStringConversions::PascalStringToWString(oper->val.str);
        }
        static void setWString(LPXLOPER oper, const std::wstring& newValue)
        {
            oper->val.str = PascalStringConversions::WStringToPascalString(newValue);
            oper->xltype = xltypeStr;
        }
        static XlfRef getRef(LPXLOPER oper)
        {
            const XLREF& ref = oper->val.mref.lpmref->reftbl[0];
            return XlfRef (ref.rwFirst,  // top
                            ref.colFirst, // left
                            ref.rwLast,   // bottom
                            ref.colLast,  // right
                            oper->val.mref.idSheet); // sheet id
        }
        static int coerce(LPXLOPER fromOper, XlTypeType toType, LPXLOPER toOper)
        {
            XLOPER typeOper;
            typeOper.val.w = toType;
            typeOper.xltype = xltypeInt;
            return XlfExcel::Instance().Call4(xlCoerce, toOper, 2, fromOper, &typeOper);
        }
        static void XlFree(LPXLOPER oper)
        {
            XlfExcel::Instance().Call4(xlFree, 0, 1, oper);
        }
        static void copy(LPXLOPER fromOper, LPXLOPER toOper)
        {
            *toOper = *fromOper;
        }
    };

} }
