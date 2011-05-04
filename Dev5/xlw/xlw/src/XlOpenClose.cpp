
/*
Copyright (C) 1998, 1999, 2001, 2002 Jérôme Lecomte
Copyright (C) 2006 Mark Joshi
Copyright (C) 2009 2011 Narinder S Claire
Copyright (C) 2011 John Adcock

This file is part of xlw, a free-software/open-source C++ wrapper of the
Excel C API - http://xlw.sourceforge.net/

xlw is free software: you can redistribute it and/or modify it under the
terms of the xlw license.  You should have received a copy of the
license along with this program; if not, please email xlw-users@lists.sf.net

This program is distributed in the hope that it will be useful, but WITHOUT
ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
FOR A PARTICULAR PURPOSE.  See the license for more details.
*/
#include <xlw/XlOpenClose.h>
#include <vector>
#include <xlw/Win32StreamBuf.h>
#include <xlw/XlFunctionRegistration.h>
#include <xlw/CellMatrix.h>
#include <xlw/TempMemory.h>
#include "PathUpdater.h"
#include<memory>
#include<string>

// redirect std::cerr at a global level to avoid
// losing the debugging info when xlAutoClose is called
// but Excel still can call the functions
static xlw::CerrBufferRedirector redirectCerr;

namespace xlw
{
	
	void executer(const std::list<eshared_ptr<IMacro> > & m_macros)
	{
		std::list<eshared_ptr<IMacro> >::const_iterator theIterator;
		for(theIterator = m_macros.begin(); theIterator!=m_macros.end(); ++theIterator)
		{
			theIterator->get()->operator()();
		}

	}

	template<>
	void MacroCache<xlw::Open>::ExecuteMacros()
	{
	   executer(m_macros);
	}
	template<>
	void MacroCache<xlw::Close>::ExecuteMacros()
	{
	   executer(m_macros);
	}

}



extern "C"
{

    long EXCEL_EXPORT xlAutoOpen()
    {
        try
        {
            // ensure temporary memory can be created
            xlw::TempMemory::InitializeProcess();

            // when we load any dll's dynamically
            // we want to make sure we look in the
            // current directory
            // static so that this is only done once per process
            static xlw::PathUpdater updatePath;

            // Displays a message in the status bar.
            xlw::XlfExcel::Instance().SendMessage("Registering library...");

            xlw::XLRegistration::ExcelFunctionRegistrationRegistry::Instance().DoTheRegistrations();

			xlw::MacroCache<xlw::Open>::Instance().ExecuteMacros();

            // Clears the status bar.
            xlw::XlfExcel::Instance().SendMessage();
            return 1;
        }
        catch(...)
        {
            return 0;
        }
    }

    long EXCEL_EXPORT xlAutoClose()
    {
        std::cerr << XLW__HERE__ << "Releasing resources" << std::endl;
		xlw::MacroCache<xlw::Close>::Instance().ExecuteMacros();
        // note that we don't unregister the functions here
        // excel has some strange behaviour when exiting and can
        // call xlAutoClose before the user has been asked about the close
        // if the user then cancels the close then we need to ensure we
        // have enough state to come back to life
        xlw::XlfExcel::DeleteInstance();

		
        // clear up any temporary memory used
        // but keep enough alive so that exel can still use
        // the functions
        xlw::TempMemory::TerminateProcess();
        return 1;
    }

    long EXCEL_EXPORT xlAutoRemove()
    {
        std::cerr << XLW__HERE__ << "Addin being unloaded" << std::endl;

        xlw::MacroCache<xlw::Close>::Instance().ExecuteMacros();

        // we can safely unregister the functions here as the user has unloaded the
        // xll and so won't expect to be able to use the functions
        xlw::XLRegistration::ExcelFunctionRegistrationRegistry::Instance().DoTheDeregistrations();

        xlw::XlfExcel::DeleteInstance();
		
        // clear up any temporary memory used
        xlw::TempMemory::TerminateProcess();
        return 1;
    }

}


