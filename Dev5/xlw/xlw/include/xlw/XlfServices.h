/*
 Copyright (C) 2011 Narinder S Claire

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

namespace xlw
{
	 //! Sends an Excel message box
    void MsgBox(const char *, const char *title = 0);

	struct StatusBar_t
	{
		StatusBar_t & operator=(const std::string &message);
		StatusBar_t & operator=(const std::wstring &message);
		void clear();

	};
}

#endif //  XLFSERVICES_HEADER_GUARD

