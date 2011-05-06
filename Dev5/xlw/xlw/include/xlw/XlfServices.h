/*
Copyright (C) 2011 Narinder S Claire
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
#ifndef XLFSERVICES_HEADER_GUARD
#define XLFSERVICES_HEADER_GUARD 
#include<string>
#include <xlw/XlfOper.h>

namespace xlw
{
    struct StatusBar_t
    {
        StatusBar_t & operator=(const std::string &message);
        void clear();
    };

    struct Notes_t
    {
        //! gets the text of the note attached to the cell from startCharacter (zero-based)
        XlfOper GetNote(const XlfOper& cellRef, int startCharacter = 0, int numChars = 0);
        //! sets the text of the note attached to the cell
        void SetNote(const XlfOper& cellRef, const std::string& note);
        //! clears the note attached to the cell
        void ClearNote(const XlfOper& cellRef);
    };

    struct Information_t
    {
        //! gets the reference of the calling cell
        XlfOper GetCallingCell();
        //! gets the reference of the active cell
        XlfOper GetActiveCell();
        //! gets the formula in the supplied cell ref
        std::string GetFormula(const XlfOper& cellRef);
        //! convert a formula from A1 to R1C1 notation
        std::string ConvertA1FormulaToR1C1(std::string a1Formula);
        //! convert a formula from R1C1 to A1 notation
        std::string ConvertR1C1FormulaToA1(std::string r1c1Formula, bool fixRows = false, bool fixColums = false);
        //! convert A1 style string to a reference
        XlfOper GetCellRefA1(std::string a1Location);
        //! convert R1C1 string to a reference
        XlfOper GetCellRefR1C1(std::string r1c1Location);
        //! convert R[-1]C[1] string to a reference given another cell
        XlfOper GetCellRefR1C1(XlfOper referenceCell, std::string r1c1RelativeLocation);
        //! convert a reference to A1 style text
        std::string GetRefTextA1(const XlfOper& ref);
        //! convert a reference to R1C1 style text
        std::string GetRefTextR1C1(const XlfOper& ref);
        //! get sheet name for a reference
        std::string GetSheetName(const XlfOper& ref);
    };

    struct Commands_t
    {
        //! Sends an Excel alert message box
        void Alert(const std::string& message);
        //! Asks the user to input a forumla
        std::string InputFormula(const std::string& message, const std::string& title);
        //! Asks the user to input a number
        double InputNumber(const std::string& message, const std::string& title);
        //! Asks the user to input a string
        std::string InputText(const std::string& message, const std::string& title);
        //! Asks the user to input a boolean
        bool InputBool(const std::string& message, const std::string& title);
        //! Asks the user to input a reference
        XlfOper InputReference(const std::string& message, const std::string& title);
        //! Asks the user to input an array
        XlfOper InputArray(const std::string& message, const std::string& title);
    };

    struct Services_t
    {
        StatusBar_t StatusBar;
        Notes_t Notes;
        Information_t Information;
        Commands_t Commands;
    };

    //! Macro functions
    //! These are wrappers for Excel4 macro functions that
    //! are callable from commands or functions defined as 
    //! macros sheet functions
    extern struct Services_t XlfServices;
}

#endif //  XLFSERVICES_HEADER_GUARD

