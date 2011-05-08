/*
Copyright (C) 2006 Mark Joshi
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

#ifndef CELLVALUE_HEADER_GUARD
#define CELLVALUE_HEADER_GUARD

#include <xlw/MyContainers.h>
#include <string>
#include <vector>

namespace xlw {

	    class CellValue
		{

		public:
			struct error_type
			{
				error_type(unsigned long v_):value(v_){}
				unsigned long value;
			};
			virtual bool IsAString() const=0;
			virtual bool IsAWstring() const=0;
			virtual bool IsANumber() const=0;
			virtual bool IsBoolean() const=0;
			virtual bool IsEmpty() const=0;
			virtual bool IsError() const=0;

			virtual const std::string & StringValue() const=0;
			virtual const std::wstring& WstringValue() const=0;
			virtual double NumericValue() const=0;
			virtual bool BooleanValue() const=0;
			virtual unsigned long ErrorValue() const=0;

			virtual operator std::string() const=0;
			virtual operator std::wstring() const=0;
			virtual operator bool() const=0;
			virtual operator double() const=0;
			virtual operator unsigned long() const=0;

			CellValue & operator=(const CellValue &value){return assign(value);}
			CellValue & operator=(const std::string &value){return assign(value);}
			CellValue & operator=(const char * value){return assign(std::string(value));}
			CellValue & operator=(const std::wstring &value){return assign(value);}
			CellValue & operator=(double value){return assign(value);}
			CellValue & operator=(unsigned long  value){return assign(value);}
			CellValue & operator=(bool value){return assign(value);}
			CellValue & operator=(int value){return assign(value);}
			CellValue & operator=(error_type value){return assign(value);}

			virtual void clear()=0;

			virtual ~CellValue(){}

		private:
			// Note assignment operators are not inherited
			CellValue & assign(const CellValue & data)
			{
				if(data.IsAString())
				{
					return assign(data.StringValue());
				}
				else if(data.IsAWstring())
				{
					return assign(data.WstringValue());
				}
				else if(data.IsANumber())
				{

					return assign(data.NumericValue());
				}
				else if(data.IsBoolean())
				{
					return assign(data.BooleanValue());
				}
				else if(data.IsEmpty())
				{
					clear();
					return *this;
				}
				else if(data.IsError())
				{
                    error_type e(data.ErrorValue());
					return assign(e);
				}
				return *this;
			}
			virtual CellValue & assign(const std::string&)=0;
			virtual CellValue & assign(const std::wstring&)=0;
			virtual CellValue & assign(double Number)=0;
			virtual CellValue & assign(unsigned long Code)=0; 
			virtual CellValue & assign(bool TrueFalse)=0;
			virtual CellValue & assign(int i)=0;
			virtual CellValue & assign(error_type)=0;

		};
}

#endif // CELLVALUE_HEADER_GUARD


