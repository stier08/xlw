
#include<list>
#include<memory>

namespace xlw
{
	namespace impl
	{ 
		namespace details
		{
			struct control_block
			{

				long use_count()const{return count;}// Never throws 
				void * get()const{return p;}// Never throws 

			private :
				friend void decrement(control_block *p);// CANNOT throw;
			    friend control_block * increment(control_block *p);// CANNOT throw;
				

				virtual void release ()throw()=0; // CANNOT throw
				virtual control_block * operator++()throw() // CANNOT throw
				{
					++count;
					return this;
				}
				virtual long operator--() throw() // CANNOT throw
				{
					if (count==1)
					{	
						release();
					}
					return --count;
				}
		
				
			protected :
				control_block(void * p_):p(p_), count(1){}
				virtual ~control_block(){};
				void *p;
				long count;
			};

			struct null_ptr :public control_block
			{
				null_ptr(): control_block(0){}
				virtual void release ()throw(){}
		
			};


			template<class T>
			struct default_control_block :public control_block
			{
				default_control_block(T * p):control_block(p){}
				void release ()throw() // CANNOT throw
				{
					delete static_cast<T*>(p);
				}
			};
			template<class T, class D>
			struct destroyer_control_block :public destroyer_base
			{
				// the copy constructor of D is not allowed to throw
				destroyer_control_block(T* p, D d_):control_block(p),d(d_){}
				void release ()throw() // CANNOT throw
				{
					 (static_cast<T*>(p));
				}
				D d;

			};

			template<class T>
			// may throw BUT will clean up p too if it does 
			control_block * controlled_alloc( T * p)
			{
				std::auto_ptr<T> temp_holder(p);

				// if the following line throws we clean up p too
				control_block *cb_ptr(new default_control_block<T>(p)) ;
				temp_holder.release();
				return cb_ptr;

			}

			template<class T>
			// may throw BUT will clean up p too if it does 
			control_block * controlled_alloc(std::auto_ptr<T> &r)
			{

				// if the following line throws we ensue
				// r still owns p
				control_block *cb_ptr(new default_control_block<T>(r.get())) ;
				r.release();
				return cb_ptr;

			}

			control_block * controlled_alloc()
			{

				return new null_ptr();
			}

			template<class T,class D>
			// may throw BUT will clean up p too if it does
			// uses d to clean p with d(p). d(p) CANNOT throw
			control_block * controlled_alloc( T * p, D d)
			{
				control_block *cb_ptr = 0;
				try
				{

					// if the following line throws we clean up p too
					cb_ptr =  new default_control_block(p) ;
					return cb_ptr;

				}
				catch(...)
				{
					d(p);
					throw;
				}
			}

			void decrement(control_block *p)throw() // CANNOT throw;
			{
				if(p)
				{
					--(*p);
					if(!(p->use_count()))
					{
						delete p;
					}
				}
			}

			control_block * increment(control_block *p)throw() // CANNOT throw;
			{
				return ++(*p);
			}

		}
	}
}