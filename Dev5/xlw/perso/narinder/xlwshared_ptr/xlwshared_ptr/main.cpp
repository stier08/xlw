//          Copyright Narinder S Claire 2011.
// Distributed under the Boost Software License, Version 1.0.
//    (See accompanying file LICENSE_1_0.txt or copy at
//          http://www.boost.org/LICENSE_1_0.txt)


#include<vector>
#include<iostream>
#include<cctype>

#include"xlwshared_ptr.h"


using xlw::impl::shared_ptr;

#include"test_classes.h"

const int B = 0;
const int D = 1;
const int DS = 2;
const int DC = 3;


int main()
{
	int bc = 0;
	int dc = 0; 
	int dsc = 0; 
	int dcc = 0; 
	try
	{
		///////////////////////////////
		shared_ptr<Base> the_ptr_B(new Derived<'B'>);

		++bc;
		++dc;
		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_B,D);

		///////////////////////////////
		shared_ptr<Base> the_ptr_B2(the_ptr_B);

		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_B2,D);

		///////////////////////////////
		shared_ptr<Derived<'B'> > the_ptr_D1(new Derived<'B'>);

	    ++bc;
		++dc;
		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_D1,D);

		///////////////////////////////
		shared_ptr<Base> the_ptr_B3(the_ptr_D1);


		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_B3,D);

		///////////////////////////////
		shared_ptr<DerivedSquared<'B'> > the_ptr_DS1(new DerivedSquared<'B'>);

	    ++bc;
		++dc;
		++dsc;
		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_DS1,DS);

		///////////////////////////////
		shared_ptr<Derived<'B'> > the_ptr_D2(the_ptr_DS1);

		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_D2,DS);

		///////////////////////////////
		shared_ptr<DerivedSquared<'B'> >the_ptr_DS2(new DerivedSquared<'B'>);

		++bc;
		++dc;
		++dsc;
		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_DS2,DS);

		///////////////////////////////
		shared_ptr<Derived<'B'> > the_ptr_D3;

	
		the_ptr_D3 = the_ptr_DS2;


		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_D3,DS);

		///////////////////////////////
		shared_ptr<Base> the_ptr_B4(the_ptr_D3);

		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_B4,DS);

		///////////////////////////////
		the_ptr_B2 = the_ptr_D3;

		test_all_counts<'B'>(bc,dc,dsc,dcc);
		check_splice(the_ptr_B2,DS);


		throw(int(0));
	}
	catch(...)
	{}


	char ch;
	std::cin >> ch;
	return 0;
}

