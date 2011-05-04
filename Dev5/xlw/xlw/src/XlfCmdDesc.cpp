
/*
 Copyright (C) 1998, 1999, 2001, 2002, 2003, 2004 Jérôme Lecomte
 Copyright (C) 2007, 2008 Eric Ehlers
 Copyright (C) 2009 2011 Narinder S Claire


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
\file XlfCmdDesc.cpp
\brief Implements the XlfCmdDesc class.
*/

// $Id$

#include <xlw/XlfCmdDesc.h>
#include <xlw/XlfOper4.h>
#include <xlw/XlfException.h>
#include <iostream>
#include <xlw/macros.h>
#include <xlw/xlwshared_ptr.h>
// Stop header precompilation
#ifdef _MSC_VER
#pragma hdrstop
#endif

/*! \e see XlfAbstractCmdDesc::XlfAbstractCmdDesc(const std::string&, const std::string&, const std::string&)
*/
xlw::XlfCmdDesc::XlfCmdDesc(const std::string& name,
                            const std::string& alias,
                            const std::string& comment,
                            const std::string& menu,
                            const std::string& menuText,
                            bool hidden) :
       XlfAbstractCmdDesc(name, alias, comment), 
       menu_(menu),
       text_(menuText),
       hidden_(hidden)
{}

xlw::XlfCmdDesc::~XlfCmdDesc()
{}

bool xlw::XlfCmdDesc::IsAddedToMenuBar()
{
  return !menu_.empty();
}

/// This function is using a naked XLOPER
int xlw::XlfCmdDesc::AddToMenuBar(const char* menu, const char* text)
{
    // allow user to override stored values
    if(menu)
    {
        menu_ = menu;
    }
    if(text)
    {
        text_ = text;
    }

    // we can only proceed if we have both
    if(menu_.empty() || text_.empty())
    {
        return 0;
    }

    //first check that the menu exists
    XlfOper barNum(10);
    XlfOper menuLocation;
    int err = XlfExcel::Instance().Call(xlfGetBar, (LPXLFOPER)menuLocation, 3, (LPXLFOPER)barNum, (LPXLFOPER)XlfOper(menu_), (LPXLFOPER)XlfOper(0));
    if (err || menuLocation.IsError())
    {
        XlfOper menuDesc(1,5);
        menuDesc(0,0) = menu_;
        menuDesc(0,1) = "";
        menuDesc(0,2) = "";
        menuDesc(0,3) = "";
        menuDesc(0,4) = "";
        err = XlfExcel::Instance().Call(xlfAddMenu, (LPXLFOPER)menuLocation, 2, (LPXLFOPER)barNum, (LPXLFOPER)menuDesc);
        if(err != xlretSuccess)
            std::cerr << XLW__HERE__ << "Add Menu " <<  menu_.c_str() << " failed" << std::endl;
    }

    XlfOper command(1,5);
    command(0,0) = text_;
    command(0,1) = GetAlias();
    command(0,2) = "";
    command(0,3) = GetComment();
    command(0,4) = "";

    err = XlfExcel::Instance().Call(xlfAddCommand, 0, 3, (LPXLFOPER)barNum, (LPXLFOPER)menuLocation, (LPXLFOPER)command);
    if (err != xlretSuccess)
        std::cerr << XLW__HERE__ << "Add command " << GetName().c_str() << " to " << menu_.c_str() << " failed" << std::endl;
    return err;
}

int xlw::XlfCmdDesc::Check(bool ERR_CHECK) const
{
    if (menu_.empty())
    {
        std::cerr << XLW__HERE__ << "No menu specified for the command \"" << GetName().c_str() << "\"" << std::endl;
        return xlretFailed;
    }
    int err = XlfExcel::Instance().Call(xlfCheckCommand, 0, 4, (LPXLFOPER)XlfOper(10), (LPXLFOPER)XlfOper(menu_), (LPXLFOPER)XlfOper(text_), (LPXLFOPER)XlfOper(ERR_CHECK));
    if (err != xlretSuccess)
    {
        std::cerr << XLW__HERE__ << "Registration of " << GetAlias().c_str() << " failed" << std::endl;
        return err;
    }
    return xlretSuccess;
}

void xlw::XlfCmdDesc::RemoveFromMenuBar()
{
    // first check that the menu exists and then delete this command
    XlfOper barNum(10);
    XlfOper menu(menu_);
    XlfOper menuLocation;
    int err = XlfExcel::Instance().Call(xlfGetBar, (LPXLFOPER)menuLocation, 3, (LPXLFOPER)barNum, (LPXLFOPER)menu, (LPXLFOPER)XlfOper(0));
    if (!err && !menuLocation.IsError())
    {
        err = XlfExcel::Instance().Call(xlfDeleteCommand, 0, 3, (LPXLFOPER)barNum, (LPXLFOPER)menu, (LPXLFOPER)XlfOper(text_));
        if(err != xlretSuccess) 
            std::cerr << XLW__HERE__ << "Delete Command " << GetName().c_str() << " from " << menu_.c_str() <<  " failed" << std::endl;

        // check if the menu is now empty, if it is then delete it
        // if it is empty then the first item won't exist
        XlfOper firstItemLocation;
        err = XlfExcel::Instance().Call(xlfGetBar, (LPXLFOPER)firstItemLocation, 3, (LPXLFOPER)XlfOper(10), (LPXLFOPER)menu, (LPXLFOPER)XlfOper(1));
        if(!err && firstItemLocation.IsError())
        {
            err = XlfExcel::Instance().Call(xlfDeleteMenu, 0, 2, (LPXLFOPER)barNum, (LPXLFOPER)menu);
            if(err != xlretSuccess)
                std::cerr << XLW__HERE__ << "Delete Menu " << menu_.c_str() <<  " failed" << std::endl;
        }
    }
}


/*!
Registers the command as a macro in excel.
\sa XlfExcel, XlfFuncDesc.
*/
int xlw::XlfCmdDesc::DoRegister(const std::string& dllName) const
{

    XlfArgDescList arguments = GetArguments();

    // 2 = normal macro, 0 = hidden command
    double type = hidden_ ? 0 : 2;

    /*
    //ERR_LOG("Registering command \"" << alias_.c_str() << "\" from \"" << name_.c_str() << "\" in \"" << dllname.c_str() << "\"");
    int err = XlfExcel::Instance().Call(
        xlfRegister, NULL, 7,
        (LPXLOPER)XlfOper(dllName.c_str()),
        (LPXLOPER)XlfOper(GetName().c_str()),
        (LPXLOPER)XlfOper("A"),
        (LPXLOPER)XlfOper(GetAlias().c_str()),
        (LPXLOPER)XlfOper(""),
        (LPXLOPER)XlfOper(type),
        (LPXLOPER)XlfOper(""));
    int err = XlfExcel::Instance().Call(
        xlfRegister, NULL, 7,
        (LPXLFOPER)XlfOper(dllName.c_str()),
        (LPXLFOPER)XlfOper(GetName().c_str()),
        (LPXLFOPER)XlfOper("A"),
        (LPXLFOPER)XlfOper(GetAlias().c_str()),
        (LPXLFOPER)XlfOper(""),
        (LPXLFOPER)XlfOper(type),
        (LPXLFOPER)XlfOper(""));
    return err;
    */

    int nbargs = static_cast<int>(arguments.size());
    std::string args("A");
    std::string argnames;

    XlfArgDescList::const_iterator it = arguments.begin();
    while (it != arguments.end())
    {
        argnames += (*it).GetName();
        args += (*it).GetType();
        ++it;
        if (it != arguments.end())
            argnames+=", ";
    }

	xlw_tr1::shared_ptr<LPXLFOPER> smart_px(new LPXLFOPER[10 + nbargs],CustomArrayDeleter<LPXLFOPER>());
	LPXLFOPER *rgx = smart_px.get();
    LPXLFOPER *px = rgx;

    (*px++) = XlfOper(dllName);
    (*px++) = XlfOper(GetName());
    (*px++) = XlfOper(args);
    (*px++) = XlfOper(GetAlias());
    (*px++) = XlfOper(argnames);
    (*px++) = XlfOper(type);
    (*px++) = XlfOper("");
    (*px++) = XlfOper("");
    (*px++) = XlfOper("");
    (*px++) = XlfOper(GetComment());
    for (it = arguments.begin(); it != arguments.end(); ++it)
    {
        (*px++) = XlfOper((*it).GetComment());
    }

    int err = static_cast<int>(XlfExcel::Instance().Callv(xlfRegister, NULL, 10 + nbargs, rgx));
    // delete[] rgx; no thankyou .. not anymore 
    return err;

}

int xlw::XlfCmdDesc::DoUnregister(const std::string& dllName) const
{
    return xlretSuccess;
}

