/*
 Copyright (C) 2006 Mark Joshi
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

#ifndef XL_OPEN_CLOSE
#define XL_OPEN_CLOSE
#include<xlw/xlw.h>
#include<xlw/Singleton.h>
#include<xlw/eshared_ptr.h>
#include<list>

extern "C"
{
  long EXCEL_EXPORT xlAutoOpen();
  long EXCEL_EXPORT xlAutoClose();
}


namespace xlw
{
	struct IMacro
	{
		IMacro(const std::string &name_,const std::string & description_):
	           m_name(name_), m_description(description_){}
		virtual void operator()()const =0;
		const std::string &GetName()const{return m_name;}
		const std::string &GetDescription()const{return m_description;}
		virtual ~IMacro(){}
	private:
		std::string m_name;
		std::string m_description;
	};

	struct MacroFromFPtr: public IMacro
	{
		typedef void (*MacroFPtr)();
		MacroFromFPtr(const std::string &name_,const std::string & description_, const MacroFPtr fptr_):
	            IMacro(name_,description_),m_fptr(fptr_){}
		
		void operator()()const{m_fptr();}

	private:
		MacroFPtr m_fptr;
	};

	struct Open{};
	struct Close{};

	template<class policy>
	class MacroCache : public singleton<MacroCache<policy> >
	{
		friend class singleton<MacroCache<policy> >;
	public:
		struct MacroRegistra
		{
			typedef void (*MacroFPtr)();
			MacroRegistra(const std::string &name_,const std::string & description_, const MacroFPtr fptr_)
			{
				eshared_ptr<IMacro> theMacro(new MacroFromFPtr(name_,description_,fptr_));
				MacroCache<policy>::Instance().RegisterMacro(theMacro);
			}
		};

		void RegisterMacro(eshared_ptr<IMacro> theMacro);
		void ExecuteMacros();
	protected:
		MacroCache(){}
	private:
		std::list<eshared_ptr<IMacro> > m_macros;
	};


}



#endif

