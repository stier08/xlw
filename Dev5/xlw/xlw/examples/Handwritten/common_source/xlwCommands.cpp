
/*
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


extern "C" 
{
    int EXCEL_EXPORT xlTestCmd() 
    {
        EXCEL_BEGIN;
        xlw::XlfServices.Commands.Alert("This is the TestCmd function\n"
                                        "Email any questions to xlw-users@lists.sourceforge.net. ");
        EXCEL_END_CMD;
    }
}

namespace 
{

    XLRegistration::XLCommandRegistrationHelper registerTestCmd(
        "xlTestCmd", "TestCmd", "A Test Command",
        "Handwritten", "Test Command");

}
