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

namespace xlw { namespace impl {

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
            THROW_XLW("No implementation on XlfOper rows");
        }
        static void setArraySize(LPXLOPER oper, RW rows, COL cols)
        {
            if(rows > 0 && cols > 0)
            {
                if(cols > 256)
                {
                    std::cerr << "Truncating columns to 256" << std::endl;
                    cols = 256;
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
                    result->xltype = xltypeRef;
                    result->val.mref.idSheet = oper->val.mref.idSheet;
                    result->val.mref.lpmref = TempMemory::GetMemory<XLMREF>();
                    result->val.mref.lpmref->count = 1;
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
        static void setRef(LPXLOPER oper, const XlfRef& newValue)
        {
            oper->xltype = xltypeRef;
            XLMREF* pmRef = TempMemory::GetMemory<XLMREF>();
            pmRef->count=1;
            pmRef->reftbl[0].rwFirst = newValue.GetRowBegin();
            pmRef->reftbl[0].rwLast = newValue.GetRowEnd()-1;
            pmRef->reftbl[0].colFirst = newValue.GetColBegin();
            pmRef->reftbl[0].colLast = newValue.GetColEnd()-1;
            oper->val.mref.lpmref = pmRef;
            oper->val.mref.idSheet = newValue.GetSheetId();
        }
        static int coerce(LPXLOPER fromOper, XlTypeType toType, LPXLOPER toOper)
        {
            XLOPER typeOper;
            typeOper.val.w = toType;
            typeOper.xltype = xltypeInt;
            return XlfExcel::Instance().Call4(xlCoerce, toOper, fromOper, &typeOper);
        }
        static void XlFree(LPXLOPER oper)
        {
            XlfExcel::Instance().Call4(xlFree, 0, oper);
        }
        static void copy(LPXLOPER fromOper, LPXLOPER toOper)
        {
            switch(fromOper->xltype & 0xFFF)
            {
                case xltypeMulti:
                    // need to do a deep copy of each element
                    toOper->xltype = xltypeMulti;
                    toOper->val.array.lparray = TempMemory::GetMemory<XLOPER>(fromOper->val.array.rows * fromOper->val.array.columns);
                    for(size_t item(0) ; item < (size_t)(fromOper->val.array.rows * fromOper->val.array.columns); ++item)
                    {
                        copy(toOper->val.array.lparray + item, toOper->val.array.lparray + item);
                    }
                    toOper->val.array.rows = fromOper->val.array.rows;
                    toOper->val.array.columns = fromOper->val.array.columns;
                    break;
                case xltypeRef:
                    {
                        toOper->xltype = xltypeRef;
                        toOper->val.mref.idSheet = fromOper->val.mref.idSheet;
                        size_t bytes(sizeof(XLMREF) + (fromOper->val.mref.lpmref->count - 1) * sizeof(XLREF));
                        toOper->val.mref.lpmref = (XLMREF*)TempMemory::GetMemory<BYTE>(bytes);
                        memcpy(toOper->val.mref.lpmref, fromOper->val.mref.lpmref, bytes);
                    }
                    break;
                case xltypeStr:
                    toOper->xltype = xltypeStr;
                    toOper->val.str = PascalStringConversions::PascalStringCopy(fromOper->val.str);
                    break;
                case xltypeBigData:
                    THROW_XLW("can't copy a big data oper");
                    break;
                default:
                    // just straight copy is fine
                    *toOper =*fromOper;
                    break;
            }
        }
    };

} }
