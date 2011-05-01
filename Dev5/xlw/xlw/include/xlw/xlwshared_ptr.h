/* Copyright (C) 2011 Narinder S Claire

 This file is part of XLW, a free-software/open-source C++ wrapper of the
 Excel C API - http://xlw.sourceforge.net/

 XLW is free software: you can redistribute it and/or modify it under the
 terms of the XLW license.  You should have received a copy of the
 license along with this program; if not, please email xlw-users@lists.sf.net

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/
#include<xlw/xlwshared_ptr_details.h>

namespace xlw
{
	namespace impl
	{
		template<class T>

		class shared_ptr 
		{
			template<class Y>
			friend class shared_ptr;

			typedef details::control_block_policy policy;

			mutable policy::control control;

		public:

			typedef T element_type;

			///    Constructors

			shared_ptr():control(policy::controlled_alloc()){}

			template<class Y> explicit shared_ptr(Y * p): 
			control(policy::controlled_alloc(p)){} // MAY throw

			template<class Y, class D> shared_ptr(Y * p, D d):
			control( policy::controlled_alloc(p,d)){} // MAY throw

			// template<class Y, class D, class A> shared_ptr(Y * p, D d, A a);

			shared_ptr(shared_ptr const & r):
						control(policy::increment(r.control)){} // never throws
				

			template<class Y> shared_ptr(shared_ptr<Y> const & r):
						control(policy::increment(r.control)){} // never throws

			//template<class Y> shared_ptr(shared_ptr<Y> const & r, T * p); // never throws
			//template<class Y> explicit shared_ptr(weak_ptr<Y> const & r);

			template<class Y> explicit shared_ptr(std::auto_ptr<Y> & r):
			control( policy::controlled_alloc(r)){} // MAY throw but if it does r still owns p

			///    Destructor
			~shared_ptr(){ policy::decrement(control);}


			shared_ptr & operator=(shared_ptr const & r) // never throws
			{
				shared_ptr(r).swap(*this);
				return *this;
			}
			template<class Y> shared_ptr & operator=(shared_ptr<Y> const & r) // never throws
			{
				shared_ptr(r).swap(*this);
				return *this;
			}
			template<class Y> shared_ptr & operator=(std::auto_ptr<Y> & r) // could throw
			{
				shared_ptr(r).swap(*this);
				return *this;
			}

			void reset() // never throws
			{ shared_ptr().swap(*this);}

			template<class Y> void reset(Y * p) // constructor of shared_ptr may throw
			{ shared_ptr(p).swap(*this);}

			template<class Y, class D> void reset(Y * p, D d) // constructor of shared_ptr may throw
			{ shared_ptr(p,d).swap(*this);}

			//template<class Y, class D, class A> void reset(Y * p, D d, A a)// constructor of shared_ptr may throw

			//template<class Y> void reset(shared_ptr<Y> const & r, T * p); // never throws

			T * get() const {return static_cast<T*>(policy::get(control));} // never throws

			typename impl::details::non_void_type<T*>::type &
			operator*() const {return *get();}// never throws
			T * operator->() const {return get();} // never throws



			long use_count() const{return control->use_count();} // never throws
			bool unique() const{return use_count()==1;} // never throws

			operator bool() const{return  get();}// never throws

			void swap(shared_ptr &p) // CANNOT THROW !
			{
				policy::swap(control,p.control);
			}


		};


	template<class T, class U>
	shared_ptr<T> static_pointer_cast(shared_ptr<U> const & r)
	{
		return shared_ptr<T>(r);
	}




	}

}

namespace xlw_tr1 = xlw::impl;


