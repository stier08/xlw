/*
 Copyright (C) 2008 Narinder Claire

 This file is part of XLW.NET 

 XLW.NET is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/


using namespace System;

namespace xlwDotNet {
	[AttributeUsage(AttributeTargets::Method)]
public ref class ExcelExportAttribute: Attribute
{
public:
	String^ description;
	ExcelExportAttribute(String^ description_)
	{
		description = description_;
	}
};
[AttributeUsage(AttributeTargets::Parameter)]
public ref class ParameterAttribute: Attribute
{
public:
	String^ description;
	ParameterAttribute(String^ description_)
	{
		description = description_;
	}
};
}
