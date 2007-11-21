
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
\file XlfFuncDesc.cpp
\brief Implements the XlfFuncDesc class.
*/

// $Id$

#include <xlw/XlfFuncDesc.h>
#include <xlw/XlfArgDescList.h>
#include <xlw/XlfOper.h>

// Stop header precompilation
#ifdef _MSC_VER
#pragma hdrstop
#endif

//! Internal implementation of XlfFuncDesc.
struct XlfFuncDescImpl
{
  //! Ctor.
  XlfFuncDescImpl(XlfFuncDesc::RecalcPolicy recalcPolicy, bool Threadsafe,
	  const std::string& category): recalcPolicy_(recalcPolicy), Threadsafe_(Threadsafe), category_(category)
  {}
  //! Recalculation policy
  XlfFuncDesc::RecalcPolicy recalcPolicy_;
  //! Category where the function is displayed in the function wizard.
  std::string category_;
  //! List of the argument descriptions of the function.
  XlfArgDescList arguments_;
  //! Flag indicating whether this function is threadsafe in Excel 2007.
  bool Threadsafe_;
};

/*!
\param name
\param alias
\param comment The first 3 argument are directly passed to
XlfAbstractCmdDesc::XlfAbstractCmdDesc.
\param category Category in which the function should appear.
\param recalcPolicy Policy to recalculate the cell.
*/
XlfFuncDesc::XlfFuncDesc(const std::string& name, const std::string& alias,
						 const std::string& comment, const std::string& category,
						 RecalcPolicy recalcPolicy, bool Threadsafe)
    :XlfAbstractCmdDesc(name, alias, comment), impl_(0)
{
  impl_ = new XlfFuncDescImpl(recalcPolicy,Threadsafe,category);
}

XlfFuncDesc::~XlfFuncDesc()
{
  delete impl_;
}

/*!
The new arguments replace all the old one (if any set). You can not
push back the arguments one by one.
*/
void XlfFuncDesc::SetArguments(const XlfArgDescList& arguments)
{
  impl_->arguments_ = arguments;
}

int XlfFuncDesc::DoRegister4(const std::string& dllName) const
{
  // alias arguments
  XlfArgDescList& arguments = impl_->arguments_;

	size_t nbargs = arguments.size();
	//std::string args("R");
	std::string args("P");
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
	if (impl_->recalcPolicy_ == XlfFuncDesc::Volatile)
	{
		args+="!";
        args+=" "; // make string longer so next line doesn't cause memory violation
		args[nbargs + 2] = 0;
	}
	else {
        args+=" "; // make string longer so next line doesn't cause memory violation
		args[nbargs + 1] = 0;
    }
	LPXLOPER *rgx = new LPXLOPER[10 + nbargs];
	LPXLOPER *px = rgx;
	(*px++) = XlfOper(dllName.c_str());
	(*px++) = XlfOper(GetName().c_str());
	(*px++) = XlfOper(args.c_str());
	(*px++) = XlfOper(GetAlias().c_str());
	(*px++) = XlfOper(argnames.c_str());
	(*px++) = XlfOper(1.0);
	(*px++) = XlfOper(impl_->category_.c_str());
	(*px++) = XlfOper("");
	(*px++) = XlfOper("");
	(*px++) = XlfOper(GetComment().c_str());
	for (it = arguments.begin(); it != arguments.end(); ++it)
  {
		(*px++) = XlfOper((*it).GetComment().c_str());
  }
	int err = static_cast<int>(XlfExcel::Instance().Call4v(xlfRegister, NULL, 10 + nbargs, rgx));
	delete[] rgx;
	return err;
}

int XlfFuncDesc::DoRegister12(const std::string& dllName) const
{
  // alias arguments
  XlfArgDescList& arguments = impl_->arguments_;

	size_t nbargs = arguments.size();
	std::string args("Q");
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
    int offset = 1;
	if (impl_->recalcPolicy_ == XlfFuncDesc::Volatile)
	{
		args+="!";
        offset++;
	}
	if (impl_->Threadsafe_)
	{
		args+="$";
        offset++;
	}

    args+=" "; // make string longer so next line doesn't cause memory violation
	args[nbargs + offset] = 0;

	LPXLOPER12 *rgx = new LPXLOPER12[10 + nbargs];
	LPXLOPER12 *px = rgx;
    // FIXME - Excel 2007 support - temporary hack - to be replaced by proper solution
    for (unsigned int i=0; i<args.length(); i++) {
        if (args[i] == 'P') args[i] = 'Q';
        if (args[i] == 'R') args[i] = 'U';
    }
	(*px++) = XlfOper(dllName.c_str());
	(*px++) = XlfOper(GetName().c_str());
	(*px++) = XlfOper(args.c_str());
	(*px++) = XlfOper(GetAlias().c_str());
	(*px++) = XlfOper(argnames.c_str());
	(*px++) = XlfOper(1.0);
	(*px++) = XlfOper(impl_->category_.c_str());
	(*px++) = XlfOper("");
	(*px++) = XlfOper("");
	(*px++) = XlfOper(GetComment().c_str());
	for (it = arguments.begin(); it != arguments.end(); ++it)
  {
		(*px++) = XlfOper((*it).GetComment().c_str());
  }
	int err = static_cast<int>(XlfExcel::Instance().Call12v(xlfRegister, NULL, 10 + nbargs, rgx));
	delete[] rgx;
	return err;
}

/*!
Registers the function as a function in excel.
\sa XlfExcel, XlfCmdDesc.
*/
int XlfFuncDesc::DoRegister(const std::string& dllName) const
{
    // FIXME temporary hack in support of Excel 12 - to be cleaned up.
    if (XlfExcel::Instance().excel12()) {
        return DoRegister12(dllName);
    } else {  // Excel 4
        return DoRegister4(dllName);
    }
}

