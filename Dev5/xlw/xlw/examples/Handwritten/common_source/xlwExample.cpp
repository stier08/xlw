
/*
 Copyright (C) 1998, 1999, 2001, 2002 Jérôme Lecomte
 Copyright (C) 2007, 2008 Eric Ehlers
 Copyright (C) 2011 John Adcock
 Copyright (C) 2011 Narinder S Claire

 This file is part of xlw, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 xlw is free software: you can redistribute it and/or modify it under the
 terms of the xlw license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

// $Id$

#include <xlw/xlw.h>
#include <sstream>
#include <vector>
#include <xlw/XlOpenClose.h>
#include <xlw/XlfServices.h>
using namespace xlw;


void welcome()
{
    XlfServices.StatusBar = "This is the XLW Handwritten Project, Enjoy !!       " 
                            "Email any questions to xlw-users@lists.sourceforge.net. " ;
}

// Registers welcome() to be executed by xlAutoOpen
// method must have return type void and take no parameters.
// You can register as many functions as you like.
namespace 
{
    MacroCache<Open>::MacroRegistra welcome_registra("welcome","welcome",welcome);
}

void goodbye()
{
    xlw::XlfServices.Commands.Alert("Thanks for choosing XLW. \n"
                                    "Email any questions to xlw-users@lists.sourceforge.net. ");
}
// Registers goodbye() to be executed by xlAutoClose
// method must have return type void and take no parameters.
// You can register as many functions as you like.
namespace 
{
    MacroCache<Close>::MacroRegistra goodbye_registra("goodbye","goodbye",goodbye);
}




extern "C" {

    LPXLFOPER EXCEL_EXPORT xlCirc(LPXLFOPER inDiam) {
        EXCEL_BEGIN;
        XlfOper xlDiam(inDiam);

        // Converts d to a double.
        double ret = xlDiam.AsDouble();
        // Multiplies it.
        ret *= 3.14159;
        // Returns the result as an XlfOper.
        return XlfOper(ret);
        EXCEL_END;
    }

    LPXLFOPER EXCEL_EXPORT xlConcat(LPXLFOPER inStr1, LPXLFOPER inStr2) {
        EXCEL_BEGIN;
        XlfOper xlStr1(inStr1);
        XlfOper xlStr2(inStr2);

        // Converts the 2 strings.
        std::wstring str1 = xlStr1.AsWstring();
        std::wstring str2 = xlStr2.AsWstring();
        // Returns the concatenation of the 2 string as an XlfOper.
        std::wstring ret = str1 + str2;
        return XlfOper(ret);
        EXCEL_END;
    }

    LPXLOPER EXCEL_EXPORT xlConcat4(LPXLOPER inStr1, LPXLOPER inStr2) {
        EXCEL_BEGIN;
        XlfOper4 xlStr1(inStr1);
        XlfOper4 xlStr2(inStr2);

        // Converts the 2 strings.
        std::string str1 = xlStr1.AsString();
        std::string str2 = xlStr2.AsString();
        // Returns the concatenation of the 2 string as an XlfOper.
        std::string ret = str1+str2;
        return XlfOper4(ret);
        EXCEL_END_4;
    }

    LPXLOPER12 EXCEL_EXPORT xlConcat12(LPXLOPER12 inStr1, LPXLOPER12 inStr2) {
        EXCEL_BEGIN;
        XlfOper12 xlStr1(inStr1);
        XlfOper12 xlStr2(inStr2);

        // Converts the 2 strings.
        std::wstring str1 = xlStr1.AsWstring();
        std::wstring str2 = xlStr2.AsWstring();
        // Returns the concatenation of the 2 string as an XlfOper.
        std::wstring ret = str1+str2;
        return XlfOper12(ret);
        EXCEL_END_12;
    }

    LPXLFOPER EXCEL_EXPORT xlStats(LPXLFOPER inTargetRange) {
        EXCEL_BEGIN;
        XlfOper xlTargetRange(inTargetRange);

        // Temporary variables.
        double averageTmp = 0.0;
        double varianceTmp = 0.0;

        // Iterate over the cells in the incoming matrix.
        for (RW i = 0; i < xlTargetRange.rows(); ++i)
        {
            for (RW j = 0; j < xlTargetRange.columns(); ++j)
            {
                // sums the values.
                double value(xlTargetRange(i,j).AsDouble());
                averageTmp += value;
                // sums the squared values.
                varianceTmp += value * value;
            }
        }
        size_t popSize = xlTargetRange.rows() * xlTargetRange.columns();

        // avoid divide by zero
        if(popSize == 0)
        {
            throw("Can't calculate stats on empty range");
        }

        // Initialization of the results Array oper.
        XlfOper result(1, 2);
        // compute average.
        double average = averageTmp / popSize;
        result(0, 0) = average;
        // compute variance
        result(0, 1) = varianceTmp / popSize - average * average;
        return result;
        EXCEL_END;
    }

    /*!
     * Demonstrates how to detect that the function is called by
     * the function wizard, and how to retrieve the coordinates
     * of the caller cell
     */
    LPXLFOPER EXCEL_EXPORT xlIsInWiz() {
        EXCEL_BEGIN;

        // Checks if called from the function wizard
        if (XlfExcel::Instance().IsCalledByFuncWiz())
            return XlfOper(true);

        // Gets reference of the called cell
        XlfRef ref = XlfServices.Information.GetCallingCell().AsRef();

        // Returns the reference in A1 format
        std::ostringstream ostr;
        COL col = ref.GetColBegin();
        COL colLeft = col;
        if(col > 26*26)
        {
            char colChar = 'A' + colLeft / (26 * 26) - 1;
            colLeft %= (26 * 26);
            ostr << colChar;
        }
        if(col > 26)
        {
            char colChar = 'A' + colLeft / 26 - 1;
            colLeft %= 26;
            ostr << colChar;
        }
        {
            char colChar = 'A' + colLeft;
            ostr << colChar;
        }
        ostr << ref.GetRowBegin() + 1 << std::ends;
        std::string address = ostr.str();

        return XlfOper(address.c_str());

        EXCEL_END;
    }

    /*!
     * Registered as volatile to demonstrate how functions can be
     * recalculated automatically even if none of the arguments
     * has changed.
     *
     * \return the number of times the function has been called.
     */
    LPXLFOPER EXCEL_EXPORT xlNbCalls() {
        EXCEL_BEGIN;

        static short nbCalls = 0;
        ++nbCalls;
        return XlfOper(nbCalls);

        EXCEL_END;
    }

    /*!
     * Demonstrates calling Excel Service functions from within a
     * a function, for this to work you must set the MacroSheetEquivalent
     * to true when registering.
     * This function gets the formula in the current cell when calculated
     */
    LPXLFOPER EXCEL_EXPORT xlCurrentFormula() {
        EXCEL_BEGIN;
        XlfOper activeCell(XlfServices.Information.GetActiveCell());
        return XlfOper(XlfServices.Information.GetFormula(activeCell));
        EXCEL_END;
    }

    LPXLFOPER EXCEL_EXPORT xlMatrixTest(LPXLFOPER inRows, LPXLFOPER inCols) {
        EXCEL_BEGIN;
        XlfOper xlRows(inRows);
        XlfOper xlCols(inCols);

        RW nbRows(xlRows.AsInt());
        RW nbCols(xlCols.AsInt());
        XlfOper result(nbRows, nbCols);

        for(RW row(0); row < nbRows; ++row)
        {
            for(RW col(0); col < nbCols; ++col)
            {
                if(row == col)
                {
                    result(row, col) = 1.0;
                }
                else
                {
                    result(row, col) = 0.0;
                }
            }
        }

        return result;
        EXCEL_END;
    }

    LPXLOPER EXCEL_EXPORT xlMatrixTest4(LPXLOPER inRows, LPXLOPER inCols) {
        EXCEL_BEGIN;
        XlfOper4 xlRows(inRows);
        XlfOper4 xlCols(inCols);

        RW nbRows(xlRows.AsInt());
        RW nbCols(xlCols.AsInt());
        XlfOper4 result(nbRows, nbCols);
        // may be truncated
        nbRows = result.rows();
        nbCols = result.columns();

        for(RW row(0); row < nbRows; ++row)
        {
            for(RW col(0); col < nbCols; ++col)
            {
                if(row == col)
                {
                    result(row, col) = 1.0;
                }
                else
                {
                    result(row, col) = 0.0;
                }
            }
        }

        return result;
        EXCEL_END_4;
    }

    LPXLOPER12 EXCEL_EXPORT xlMatrixTest12(LPXLOPER12 inRows, LPXLOPER12 inCols) {
        EXCEL_BEGIN;
        XlfOper12 xlRows(inRows);
        XlfOper12 xlCols(inCols);

        RW nbRows(xlRows.AsInt());
        RW nbCols(xlCols.AsInt());
        XlfOper12 result(nbRows, nbCols);
        // may be truncated
        nbRows = result.rows();
        nbCols = result.columns();

        for(RW row(0); row < nbRows; ++row)
        {
            for(RW col(0); col < nbCols; ++col)
            {
                if(row == col)
                {
                    result(row, col) = 1.0;
                }
                else
                {
                    result(row, col) = 0.0;
                }
            }
        }

        return result;
        EXCEL_END_12;
    }

    int EXCEL_EXPORT xlTestCmd() {
        EXCEL_BEGIN;
        xlw::XlfServices.Commands.Alert("This is the TestCmd function\n"
                                        "Email any questions to xlw-users@lists.sourceforge.net. ");
        EXCEL_END_CMD;
    }
}

namespace {

    // Register the function Circ.

    XLRegistration::Arg CircArgs[] = {
        { "diameter", "Diameter of the circle", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerCirc(
        "xlCirc", "Circ", "Compute the circumference of a circle",
        "xlw Example", CircArgs, 1);

    // Register the function Concat.

    XLRegistration::Arg ConcatArgs[] = {
        { "string1", "First string", "XLF_OPER" },
        { "string2", "Second string", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerConcat(
        "xlConcat", "Concat", "Concatenate two strings",
        "xlw Example", ConcatArgs, 2);

    // Register the function Concat4.

    XLRegistration::Arg Concat4Args[] = {
        { "string1", "First string", "P" },
        { "string2", "Second string", "P" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerConcat4(
        "xlConcat4", "Concat4", "Concatenate two strings",
        "xlw Example", Concat4Args, 2, false, false, "P");

    // Register the function Concat12.

    XLRegistration::Arg Concat12Args[] = {
        { "string1", "First string", "Q" },
        { "string2", "Second string", "Q" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerConcat12(
        "xlConcat12", "Concat12", "Concatenate two strings",
        "xlw Example", Concat12Args, 2, false, false, "Q");

    // Register the function Stats.

    XLRegistration::Arg StatsArgs[] = {
        { "Population", "Target range containing the population", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerStats(
        "xlStats", "Stats", "Return a 1x2 range containing "
        "the average and the variance of a numeric population",
        "xlw Example", StatsArgs, 1);

    // Register the function IsInWiz.

    XLRegistration::XLFunctionRegistrationHelper registerIsInWiz(
        "xlIsInWiz", "IsInWiz", "Return true if the function is called "
        "from the function wizard, and the address of the caller otherwise",
        "xlw Example");

    // Register the function NbCalls as volatile.

    XLRegistration::XLFunctionRegistrationHelper registerNbCalls(
        "xlNbCalls", "NbCalls", "Return the number of times the function "
        "has been calculated since the xll was loaded (volatile)",
        "xlw Example", 0, 0, true);

    // Register the function CurrentFormula as volatile and MacroSheetEquivalent

    XLRegistration::XLFunctionRegistrationHelper registerCurrentFormula(
        "xlCurrentFormula", "CurrentFormula", "Returns the formula in the current cell "
        "by calling Excel Macro functions",
        "xlw Example", 0, 0, true, false, "", "", false, true);

    XLRegistration::Arg MatrixTestArgs[] = {
        { "rows", "Number of rows", "XLF_OPER" },
        { "cols", "Number of columns", "XLF_OPER" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerMatrixTest(
        "xlMatrixTest", "MatrixTest", "Generate an identity matrix",
        "xlw Example", MatrixTestArgs, 2);

    XLRegistration::Arg MatrixTest4Args[] = {
        { "rows", "Number of rows", "P" },
        { "cols", "Number of columns", "P" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerMatrixTest4(
        "xlMatrixTest4", "MatrixTest4", "Generate an identity matrix",
        "xlw Example", MatrixTest4Args, 2, false, false, "P");

    XLRegistration::Arg MatrixTest12Args[] = {
        { "rows", "Number of rows", "Q" },
        { "cols", "Number of columns", "Q" }
    };

    XLRegistration::XLFunctionRegistrationHelper registerMatrixTest12(
        "xlMatrixTest12", "MatrixTest12", "Generate an identity matrix",
        "xlw Example", MatrixTest12Args, 2, false, false, "Q");

    XLRegistration::XLCommandRegistrationHelper registerTestCmd(
        "xlTestCmd", "TestCmd", "A Test Command",
        "Handwritten", "Test Command");
}
