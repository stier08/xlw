/*! $Id$
 *********************************************************************
 *              
 *  \file       xll.h
 *              
 *  \brief      include file for libXLL
 *              
 *  \project    libXLL - support for MS Excel Add-Ins
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
 */

#ifndef xll_h
#define xll_h

#include <xlw/xldata32.h>

#define XLL_EXPORT __stdcall

#if defined(__midl)

	// do NOT change these constants unless really necessary!!

	#define XLL_ARG_HELP	98FD2C27-7584-4292-BBFD-9F8A0C41791D
	#define XLL_ARG_DEFAULT	5F09994E-316D-45f5-AB45-795FAA420B17
	#define XLL_CATEGORY	6AD914E0-4B91-4d61-A151-89F75B3298CB
	#define XLL_FUNC_TEXT	989B8E17-09ED-4a71-8384-CCB4FD4D3EEE
	#define XLL_SHORTCUT	6B15F3BC-EF74-45a3-A4E3-B1D96F4FA17E
	#define XLL_MACRO_TYPE	F1115E50-BEF3-43c1-92B3-6F5F35D826D1
	#define XLL_VOLATILE	2237869E-4A85-413a-9543-988EFB01F36C

	#pragma midl_echo("#include <xlw/xll.h>") 

#else	// defined(__midl)

	namespace xll {

		/*
		** custom properties for use in IDL
		**
		** do NOT change these constants unless really necessary!!
		** 
		*/

		static const GUID XLL_ARG_HELP = 
		{ 0x98fd2c27, 0x7584, 0x4292, 
			{ 0xbb, 0xfd, 0x9f, 0x8a, 0xc, 0x41, 0x79, 0x1d } 
		};

		static const GUID XLL_ARG_DEFAULT = 
		{ 0x5f09994e, 0x316d, 0x45f5, 
			{ 0xab, 0x45, 0x79, 0x5f, 0xaa, 0x42, 0xb, 0x17 } 
		};

		static const GUID XLL_CATEGORY = 
		{ 0x6ad914e0, 0x4b91, 0x4d61, 
			{ 0xa1, 0x51, 0x89, 0xf7, 0x5b, 0x32, 0x98, 0xcb } 
		};

		static const GUID XLL_FUNC_TEXT = 
		{ 0x989b8e17, 0x9ed, 0x4a71, 
			{ 0x83, 0x84, 0xcc, 0xb4, 0xfd, 0x4d, 0x3e, 0xee } 
		};

		static const GUID XLL_SHORTCUT = 
		{ 0x6b15f3bc, 0xef74, 0x45a3, 
			{ 0xa4, 0xe3, 0xb1, 0xd9, 0x6f, 0x4f, 0xa1, 0x7e } 
		};

		static const GUID XLL_MACRO_TYPE = 
		{ 0xf1115e50, 0xbef3, 0x43c1, 
			{ 0x92, 0xb3, 0x6f, 0x5f, 0x35, 0xd8, 0x26, 0xd1 } 
		};

		static const GUID XLL_VOLATILE = 
		{ 0x2237869e, 0x4a85, 0x413a, 
			{ 0x95, 0x43, 0x98, 0x8e, 0xfb, 0x1, 0xf3, 0x6c } 
		};


	} // namespace xll

#endif	// defined(__midl)

#endif	// xll_h
