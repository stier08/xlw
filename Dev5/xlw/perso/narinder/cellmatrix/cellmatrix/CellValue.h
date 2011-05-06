
#ifndef CELLVALUE_HEADER_GUARD
#define CELLVALUE_HEADER_GUARD

#include <xlw/MyContainers.h>
#include <string>
#include <vector>

namespace xlw {

	    class CellValue
		{

		public:
			virtual bool IsAString() const=0;
			virtual bool IsAWstring() const=0;
			virtual bool IsANumber() const=0;
			virtual bool IsBoolean() const=0;
			virtual bool IsEmpty() const=0;

			virtual const std::string & StringValue() const=0;
			virtual const std::wstring& WstringValue() const=0;
			virtual double NumericValue() const=0;
			virtual bool BooleanValue() const=0;
			virtual unsigned long ErrorValue() const=0;

			virtual std::string StringValueLowerCase() const=0;
			virtual std::wstring WstringValueLowerCase() const=0;

			virtual operator std::string() const=0;
			virtual operator std::wstring() const=0;
			virtual operator bool() const=0;
			virtual operator double() const=0;
			virtual operator unsigned long() const=0;

			virtual void clear()=0;

			virtual ~CellValue(){}

		};
}

#endif // CELLVALUE_HEADER_GUARD


