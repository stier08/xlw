/*! $Id$
 *********************************************************************
 *              
 *  \file       XllTypeLib.h
 *              
 *  \brief      parse typelib and export as Excel add-in functions
 *
 *  \seealso	TypelibRegistry.h
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

#ifndef XllTypeLib_h
#define XllTypeLib_h

#include <atlbase.h>
#include <atlconv.h>
#include "XlfOper.h"
#include <vector>
#include <string>

// http://msdn.microsoft.com/library/en-us/automat/htm_hh2/chap9_4ib7.asp

namespace xll {
	
	typedef std::vector<LPXLOPER> xllCallStack;

	//! typelib parser and excel add-in registry
	/*! this was hacked to tohether using some examples from MS and 
		should not serve as an example for OO programming style...
	 */
	class xllTypeLib {
	public:		
		xllTypeLib(LPCWSTR tlbFile);
		xllTypeLib(LPCSTR tlbFile);
		xllTypeLib(HINSTANCE hIns);
		// xllTypeLib( /* libname to search in reg */ );
		bool EnumerateItems(void);
		void registerModule(LPCWSTR modName = 0, LPCWSTR funcName = 0);

	private:
		void InitLibInfo(LPCWSTR tlbFile);
		void enumModuleFunctions(UINT uiIndex, LPCWSTR funcName);
		int registerModuleFunction(UINT uiIndex);
		void AddArguments( const FUNCDESC *pFuncDesc, const std::string &functionText, xllCallStack &pxl);
		void AddArgumentHelp(const FUNCDESC *pFuncDesc, UINT uiIndex, xllCallStack &pxl);
		std::string ExcelUserDefinedType( HREFTYPE hRefType, bool bInOut, bool bPtr );
		std::string ExcelVariantType( VARTYPE vt, bool bInOut, bool bPtr);
		std::string ExcelTypeDesc( const TYPEDESC *pTypeDesc, bool bInOut, bool bPtr = false);

		std::string m_tlbFile;
		CComPtr<ITypeLib> m_pTLib;
		CComPtr<ITypeInfo> m_pTInfo;
		CComPtr<ITypeInfo2> m_pTInfo2;
		WORD m_wMajorVerNum;
		WORD m_wMinorVerNum;
		std::string m_libName;
		std::string m_libDocString;
		std::string m_libHelpFile;
		DWORD m_libHelpContext;
		std::string m_modName;
		std::string m_modDocString;
		std::string m_modHelpFile;
		std::string m_modCategory;
		DWORD m_modHelpContext;
		
		std::string StringifyInvokeKind( INVOKEKIND invokeKind );
		std::string ComposeFuncAttr( const FUNCDESC *pFuncDesc, bool bDisp );
		std::string StringifyUserDefinedType( HREFTYPE hRefType );
		std::string StringifyFuncDesc( const FUNCDESC *pFuncDesc, bool bDisp );
		std::string StringifyFuncParams( const FUNCDESC *pFuncDesc );
		std::string StringifyVarDesc( const VARDESC *pVarDesc );
		std::string StringifyTypeDesc( const TYPEDESC *pTypeDesc );
		std::string StringifyVariantType( VARTYPE vt );
		std::string StringifyCallConv( CALLCONV callConv );
		std::string StringifyParamFlags( PARAMDESC *pParamDesc );
		std::string GetTypeString(TYPEKIND typeKind);
		std::string GetBaseInterface( const TYPEATTR *pTAttr );
		// void GetTypeFlagStrings( WORD wTypeFlags, list<string>& listStrings );
		bool IsValidDispid( DISPID id );
		void EnumStructTypes( const TYPEATTR *pTAttr );
		void EnumInterfaceTypes( const TYPEATTR *pTAttr );
		void EnumModuleTypes( const TYPEATTR *pTAttr );
		void EnumAliasTypes( const TYPEATTR *pTAttr );
		void EnumCoClassTypes( const TYPEATTR *pTAttr );
		void MoreInfo(UINT uiIndex);

	};
	
} // namespace xll

#endif // XllTypeLib_h
