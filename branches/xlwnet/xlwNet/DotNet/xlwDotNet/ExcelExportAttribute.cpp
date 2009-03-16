
using namespace System;

namespace xlwDotNet {
	[AttributeUsage(AttributeTargets::Method)]
public ref class ExcelExportAttribute: Attribute
{
public:
	String^ description;
	bool volatileFlag;
	ExcelExportAttribute(String^ description_)
	{
		description = description_;
		volatileFlag = false;

	}
	ExcelExportAttribute(String^ description_, bool volatileFlag_)
	{
		description = description_;
		volatileFlag = volatileFlag_;

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
