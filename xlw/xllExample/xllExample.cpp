/*! $Id$
 *********************************************************************
 *              
 *  \file       xllExample.cpp
 *              
 *  \brief      main module for MS Excel example add-in
 *              
 *  \project    libXLL - support for MS Excel Add-Ins
 *              
 *  \author     Jens Thiel <Jens.Thiel@stochastix.de>
 *              
 *  \date       17.03.2002
 *              
 ***************************************************** \begincopyright
 *
 *  Copyright (c) 2001-2002 by Jens Thiel. All rights reserved.
 *
 *  You should have received a printed copy of the license agreement.
 *
 *  Otherwise, please contact
 *
 *      Jens Thiel
 *      Stochastix GmbH
 *      Rathausallee 10
 *      D-53757 Sankt Augustin
 *
 *      Fon	  +49 (2241) 1484-200
 *      Fax   +49 (2241) 1484-222
 *
 *      Email Jens.Thiel@stochastix.de
 *
 *  or lookup our web page at:
 *
 *      http://www.stochastix.de/
 *
 ******************************************************* \endcopyright
 *
 *  \internal
 *
 *  $Log$
 *  Revision 1.1.2.1  2003/02/20 16:44:47  nando
 *  libXLL added
 *
 *
 *********************************************************************
 */

// QuantLib does not like it -- should define templates here instead
#define NOMINMAX 

/////////////////////////////////////////////////////////////////////////////
// This has been compiled from the corresponding xllExample.idl
// and includes the functions definitions as well as references to the
// necessary Excel types
#include "xllExample.h"

/////////////////////////////////////////////////////////////////////////////
// define a long name for the add-in
// used by xlw/xllMain_c.h to supply a value to Excel
#ifdef NDEBUG
#define XLL_LONG_NAME "My Add-In (Release)"
#else
#define XLL_LONG_NAME  "My Add-In (" __DATE__ " " __TIME__ " Debug)"
#endif

// this is the do-all-the-add-in-stuff boilerplate header file.
// it must be included in one of your .cpp files 
#include <xlw/xllMain_c.h>

// used for the xllIsNaN example
#include <limits>

// comment out if you have no QuantLib installed
// #include <ql\quantlib.hpp>

// comment out if you don't have Intel MKL installed
// #include <mkl.h>

/////////////////////////////////////////////////////////////////////////////
// Now we implement the functions
// You can look up the definition at the end of xllExample.h, which has
// been compiled from the corresponding xllExample.idl


// notice the following statement - it is important ;-)
extern "C" {


void XLL_EXPORT xllSayHello()
{
	if (!XlfExcel::Instance().IsCalledByFuncWiz()) {
		XlfExcel::MsgBox(
			"Hello ;o)",
			XLL_LONG_NAME
		);
	}
}

double XLL_EXPORT xllCirc(double diameter)
{
	return diameter * 3.14159;
}

VARIANT_BOOL XLL_EXPORT xllIsNaN(double number)
{
	return number == std::numeric_limits<double>::signaling_NaN();
}

double XLL_EXPORT xllStdEuroOptValue(
    QlOptionType optionType, 
	double underlying, double strike, double dividendYield,
    double riskFreeRate, double startTime, double endTime,
    double volatility)
{

#ifndef QL_VERSION
	return std::numeric_limits<double>::signaling_NaN();
#else
	
	using namespace QuantLib;

	QuantLib::DayCounters::Actual360 dayCounter;
	Time maturity = dayCounter.yearFraction(
		Date((long)startTime), Date((long)endTime)
	);

	Pricers::EuropeanOption option(
				(Option::Type)(optionType-1), 
				underlying, 
				strike, 
				(Rate)dividendYield,
				(Rate)riskFreeRate,
				maturity,
				volatility
			);

	return option.value();

#endif

}

void XLL_EXPORT xllMINV( FP *Matrix )
{
#ifdef _MKL_H_	
	
	int n = Matrix->columns;
	if( (n != Matrix->rows) || (n <= 0)  ) {
		Matrix = 0;
		return;
	}

	std::vector<int> ipiv;
	ipiv.reserve(n);
	int info;
	DGETRF( &n, &n, Matrix->array, &n, ipiv.begin(), &info );
	int lwork = n*64;
	std::vector<double> work;
	work.reserve(lwork);
	DGETRI( &n, Matrix->array, &n, ipiv.begin(), work.begin(), &lwork, &info);
	// double optLwork = work[1];

#else
	Matrix = 0;	// signal error
#endif
}




} // extern "C"

