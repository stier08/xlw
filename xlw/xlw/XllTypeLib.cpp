/*! $Id$
 *********************************************************************
 *              
 *  \file       XllTypeLib.cpp
 *              
 *  \brief      parse typelib and export as Excel add-in functions
 *              
 *  \project    libXLL - support for MS Excel Add-Ins
 *              
 *  \author     Jens Thiel <Jens.Thiel@stochastix.de>
 *              
 *  \date       03.10.2001
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
 *  Revision 1.1.2.1  2003/02/20 16:41:58  nando
 *  libXLL added
 *
 *  Revision 1.5  2002/03/17 16:45:47  jens
 *  initial semi-public release
 *
 *
 *********************************************************************
 */

#define NOMINMAX

#include <COMDEF.H>
#include "XllTypeLib.h"
#include "xll.h"
#include "XlfExcel.h"	// use XlfExcel::Instance().Callv()
#include <cassert>
#include <limits>
#include <exception>
#include <stdexcept>
#include <iostream>


namespace xll {
	
	std::string xllTypeLib::StringifyInvokeKind( INVOKEKIND invokeKind )
	{
		switch( invokeKind )
		{
		case INVOKE_PROPERTYGET:
			return "propget";
			
		case INVOKE_PROPERTYPUT:
			return "propput";
			
		case INVOKE_PROPERTYPUTREF:
			return "propputref";
		}
		
		return "";
	}
	
	
	std::string xllTypeLib::StringifyVariantType( VARTYPE vt )
	{
		switch( vt )
		{
		case VT_I1:
			return "char";
			
		case VT_I2:
			return "short";
			
		case VT_I4:
			return "long";
			
		case VT_R4:
			return "float";
			
		case VT_R8:
			return "double";
			
		case VT_CY:
			return "CY";
			
		case VT_DATE:
			return "DATE";
			
		case VT_BSTR:
			return "BSTR";
			
		case VT_DISPATCH:
			return "IDispatch";
			
		case VT_ERROR:
			return "SCODE";
			
		case VT_BOOL:
			return "VARIANT_BOOL";
			
		case VT_VARIANT:
			return "VARIANT";
			
		case VT_UNKNOWN:
			return "IUnknown";
			
		case VT_UI1:
			return "BYTE";
			
		case VT_DECIMAL:
			return "DECIMAL";
			
		case VT_UI2:
			return "USHORT";
			
		case VT_UI4:
			return "ULONG";
			
		case VT_I8:
			return "__int64";
			
		case VT_UI8:
			return "unsigned __int64";
			
		case VT_INT:
			return "int";
			
		case VT_UINT:
			return "UINT";
			
		case VT_HRESULT:
			return "HRESULT";
			
		case VT_VOID:
			return "void";
			
		case VT_LPSTR:
			return "char *";
			
		case VT_LPWSTR:
			return "wchar_t *";
		}
		
		return "ERROR!";
	}
	
	std::string xllTypeLib::StringifyCallConv( CALLCONV callConv )
	{
		switch( callConv )
		{
		case CC_FASTCALL:
			return "__fastcall";
			
		case CC_CDECL:
			return "__cdecl";
			
		case CC_MSCPASCAL:
			return "__pascal";
			
		case CC_MACPASCAL:
			return "MACPASCAL";
			
		case CC_STDCALL:
			return "__stdcall";
			
		case CC_FPFASTCALL:
			return "FPFFASTCALL";
			
		case CC_SYSCALL:
			return "SYSCALL";
			
		case CC_MPWCDECL:
			return "MPWCDECL";
			
		case CC_MPWPASCAL:
			return "MPWPASCAL";
		}
		
		return "ERROR!";
	}
	
	std::string xllTypeLib::GetTypeString(TYPEKIND typeKind)
	{
		std::string strKind = "UNKNOWN";
		
		switch( typeKind )
		{
		case TKIND_ENUM:
			strKind = "ENUM";
			break;
			
		case TKIND_RECORD:
			strKind = "STRUCT";
			break;
			
		case TKIND_MODULE:
			strKind = "MODULE";
			break;
			
		case TKIND_INTERFACE:
			strKind = "INTERFACE";
			break;
			
		case TKIND_DISPATCH:
			strKind = "DISPATCH";
			break;
			
		case TKIND_COCLASS:
			strKind = "COCLASS";
			break;
			
		case TKIND_ALIAS:
			strKind = "TYPEDEF";
			break;
			
		case TKIND_UNION:
			strKind = "UNION";
			break;
		}
		
		return strKind;
	}
	
	// void GetTypeFlagStrings( WORD wTypeFlags, list<string>& listStrings )
	//{
	//}
	
	std::string xllTypeLib::StringifyUserDefinedType( HREFTYPE hRefType )
	{
		std::string  strText = "ERROR!";
		CComPtr<ITypeInfo> pTInfoCust;
		CComBSTR bstrName;
		HRESULT hr;
		USES_CONVERSION;
		
		hr = m_pTInfo->GetRefTypeInfo( hRefType, &pTInfoCust );
		if( SUCCEEDED( hr ) )
		{
			hr = pTInfoCust->GetDocumentation( -1, &bstrName, NULL, NULL, NULL );
			if( SUCCEEDED( hr ) )
				strText = W2CA( bstrName.m_str );
		}
		
		return strText;
	}
	
	std::string xllTypeLib::StringifyTypeDesc( const TYPEDESC *pTypeDesc )
	{
		std::string strText = _T( "ERROR! ERROR! ERROR!" );
		
		switch( pTypeDesc->vt )
		{
		case VT_PTR:
			strText = StringifyTypeDesc( pTypeDesc->lptdesc );
			strText += "*";
			return strText;
			
		case VT_SAFEARRAY:
			strText = "SAFEARRAY[";
			strText += StringifyTypeDesc( pTypeDesc->lptdesc );
			strText += "]";
			return strText;
			
		case VT_CARRAY:
			{
				ARRAYDESC *pArrDesc = pTypeDesc->lpadesc;
				std::string strAttr;
				char s[64];
				
				strText = StringifyTypeDesc( &( pArrDesc->tdescElem ) );
				for( int i = 0 ; i < pArrDesc->cDims ; ++i )
				{
					sprintf(s, "[%d]", pArrDesc->rgbounds[i].cElements - 1 );
					strText.append(s);
				}
				return strText;
			}
			
		case VT_USERDEFINED:
			return StringifyUserDefinedType( pTypeDesc->hreftype );
		}
		
		return StringifyVariantType( pTypeDesc->vt );
	}
	
	std::string xllTypeLib::StringifyVarDesc( const VARDESC *pVarDesc )
	{
		std::string strText;
		HRESULT hr;
		UINT uiNames;
		CComBSTR bstrName( _T( "ERROR!" ) );
		USES_CONVERSION;
		
		if( pVarDesc->varkind == VAR_CONST )
			strText = _T( "const " );
		
		strText += StringifyTypeDesc( &( pVarDesc->elemdescVar.tdesc ) );
		strText += _T( " " );
		hr = m_pTInfo->GetNames( pVarDesc->memid, &bstrName, 1, &uiNames );
		if( SUCCEEDED( hr ) )
			strText += W2CA( bstrName.m_str );
		
		if( pVarDesc->varkind == VAR_CONST )
		{
			CComVariant var( *( pVarDesc->lpvarValue ) );
			var.ChangeType( VT_BSTR );
			
			strText += " = ";
			strText += W2CA( var.bstrVal );
		}
		
		strText += ";";
		return strText;
	}
	
	std::string xllTypeLib::ComposeFuncAttr( const FUNCDESC *pFuncDesc, bool bDisp )
	{
		std::string strAttr = "";
		std::string strInvokeKind;
		
		if( bDisp )	//is it a dispinterface? 
		{
			char s[64];
			sprintf(s,"[id(0x%lx)", pFuncDesc->memid );
			strAttr.append(s);
		}
		
		strInvokeKind = StringifyInvokeKind( pFuncDesc->invkind );
		if( !strInvokeKind.empty() )
		{
			if( bDisp )
				strAttr += ", ";
			else
				strAttr = "[";
			
			strAttr += strInvokeKind;
		}
		
		if( !strAttr.empty() )
			strAttr += "]";
		
		return strAttr;
	}
	
	
	std::string xllTypeLib::StringifyParamFlags( PARAMDESC *pParamDesc )
	{
		std::string strFlags( "" );
		
		if( pParamDesc->wParamFlags & PARAMFLAG_FIN )
			strFlags += "in";
		if( pParamDesc->wParamFlags & PARAMFLAG_FOUT )
			strFlags += ( strFlags.empty() ) ? "out" : ", out";
		if( pParamDesc->wParamFlags & PARAMFLAG_FLCID )
			strFlags += ( strFlags.empty() ) ? "lcid" : ", lcid";
		if( pParamDesc->wParamFlags & PARAMFLAG_FRETVAL )
			strFlags += ( strFlags.empty() ) ? "retval" : ", retval";
		if( pParamDesc->wParamFlags & PARAMFLAG_FOPT )
			strFlags += ( strFlags.empty() ) ? "optional" : ", optional";
		if( pParamDesc->wParamFlags & PARAMFLAG_FHASDEFAULT )
		{
			CComVariant var( pParamDesc->pparamdescex->varDefaultValue );
			USES_CONVERSION;
			
			var.ChangeType( VT_BSTR );
			strFlags += ( strFlags.empty() ) ? "defaultvalue(" : ", defaultvalue(";
			strFlags += W2CA( var.bstrVal );
			strFlags += ")";
		}
		
		if( !strFlags.empty() )
		{
			strFlags.insert( 0, "\n        [" ).append("]");
		}
		
		return strFlags;
	}
	
	std::string xllTypeLib::StringifyFuncParams( const FUNCDESC *pFuncDesc )
	{
		std::string strAllParams = "";
		std::string strParam;
		
		for( int i = 0 ; i < pFuncDesc->cParams ; ++i )
		{
			strParam = StringifyParamFlags(
				&pFuncDesc->lprgelemdescParam[i].paramdesc );
			strParam += StringifyTypeDesc(
				&pFuncDesc->lprgelemdescParam[i].tdesc );
			
			strParam.append(" [").append(
				ExcelTypeDesc( &pFuncDesc->lprgelemdescParam[i].tdesc,
				0 != (pFuncDesc->lprgelemdescParam[i].paramdesc.wParamFlags & PARAMFLAG_FOUT),
				false)
				).append("] ");
			
			//is there another arg?
			if( ( i + 1 ) < pFuncDesc->cParams )
				strParam += ", ";
			
			strAllParams += strParam;
		}
		
		return strAllParams;
	}
	
	std::string xllTypeLib::StringifyFuncDesc( const FUNCDESC *pFuncDesc, bool bDisp )
	{
		std::string strText;
		CComBSTR bstrName;
		UINT uiNames;
		HRESULT hr;
		USES_CONVERSION;
		
		//compose function attributes
		if( bDisp && pFuncDesc->memid != DISPID_UNKNOWN )
		{
			strText = ComposeFuncAttr( pFuncDesc, bDisp );
			if( !strText.empty() )
				strText += " ";
		}
		
		//specify function return type
		strText += StringifyTypeDesc( &pFuncDesc->elemdescFunc.tdesc );
		strText += " ";
		
		//specify function calling convention
		strText += StringifyCallConv( pFuncDesc->callconv );
		strText += " ";
		
		//specify function name
		hr = m_pTInfo->GetNames( pFuncDesc->memid, &bstrName, 1 /* max names */ , &uiNames );
		if( SUCCEEDED( hr ) )
			strText += W2CA( bstrName.m_str );
		else
			strText += "ERROR!";
		strText += "( ";
		
		//specify all the function params
		std::string strParams = StringifyFuncParams( pFuncDesc );  // pFuncDesc.cParams[Opt]
		if( strParams.empty() ) //no args?
			strText += "void );";
		else
			strText += strParams.append("\n  );");
		return strText;
	}
				
	
	
	
	bool xllTypeLib::IsValidDispid( DISPID id )
	{
		static DISPID invalid[] = { 0x60000000, 0x60000001, 0x60000002,
			0x60010000, 0x60010001, 0x60010002,
			0x60010003 };
		
		for( int i = 0 ; i < sizeof( invalid ) ; ++i )
			if( id == invalid[i] )
				return false;
			
			return true;
	}
	
	void xllTypeLib::EnumInterfaceTypes( const TYPEATTR *pTAttr )
	{
		HRESULT hr;
		FUNCDESC *pFuncDesc;
		bool bDisp = ( pTAttr->typekind == TKIND_DISPATCH );
		
		for( int i = 0 ; i < pTAttr->cFuncs ; ++i )
		{
			hr = m_pTInfo->GetFuncDesc( i, &pFuncDesc );
			if( SUCCEEDED( hr ) )
			{
				if( IsValidDispid( pFuncDesc->memid ) ) // comment out for all methods
					std::cout << "  " << StringifyFuncDesc( pFuncDesc, bDisp ).c_str() << "\n";
				m_pTInfo->ReleaseFuncDesc( pFuncDesc );
			}
		}
	}
	
	void xllTypeLib::EnumModuleTypes( const TYPEATTR *pTAttr )
	{
		HRESULT hr;
		FUNCDESC *pFuncDesc;
		
		for( int i = 0 ; i < pTAttr->cFuncs ; ++i )
		{
			hr = m_pTInfo->GetFuncDesc( i, &pFuncDesc );
			if( SUCCEEDED( hr ) )
			{
				std::cout << "  " << StringifyFuncDesc( pFuncDesc, false ).c_str() << "\n";
				m_pTInfo->ReleaseFuncDesc( pFuncDesc );
			}
		}
	}
	
	void xllTypeLib::EnumStructTypes( const TYPEATTR *pTAttr )
	{
		HRESULT hr;
		VARDESC *pVarDesc;
		
		for( int i = 0 ; i < pTAttr->cVars ; ++i )
		{
			hr = m_pTInfo->GetVarDesc( i, &pVarDesc );
			if( SUCCEEDED( hr ) )
			{
				std::cout << "  " << StringifyVarDesc( pVarDesc ).c_str() << "\n";
				m_pTInfo->ReleaseVarDesc( pVarDesc );
			}
		}
	}
	
	void xllTypeLib::EnumAliasTypes( const TYPEATTR *pTAttr )
	{
		std::string strText("typedef ");
		CComBSTR bstrName;
		HRESULT hr;
		USES_CONVERSION;
		
		strText.append(StringifyTypeDesc( &pTAttr->tdescAlias ));
		strText.append( " " );
		hr = m_pTInfo->GetDocumentation( -1, &bstrName, NULL, NULL, NULL );
		if( SUCCEEDED( hr ) )
			strText += ( W2CA( bstrName.m_str ) );
		
		strText += ";";
		std::cout << "  " << strText.c_str() << "\n";
	}
	
	void xllTypeLib::EnumCoClassTypes( const TYPEATTR *pTAttr )
	{
		HREFTYPE hRefType;
		HRESULT hr;
		USES_CONVERSION;
		
		for( int i = 0 ; i < pTAttr->cImplTypes ; ++i )
		{
			hr = m_pTInfo->GetRefTypeOfImplType( i, &hRefType );
			if( SUCCEEDED( hr ) )
			{
				CComPtr<ITypeInfo> pTInfoImpl;
				hr = m_pTInfo->GetRefTypeInfo( hRefType, &pTInfoImpl );
				if( SUCCEEDED( hr ) )
				{
					CComBSTR bstrName;
					hr = pTInfoImpl->GetDocumentation( -1, &bstrName, NULL, NULL, NULL );
					std::cout << "  " << W2CA( bstrName.m_str ) << "\n";
				}
			}
		}
	}
	
	
	
	
	
	std::string xllTypeLib::GetBaseInterface( const TYPEATTR *pTAttr )
	{
		HREFTYPE hRefType;
		std::string strName;
		CComPtr<ITypeInfo> pTInfoBase;
		CComBSTR bstrName;
		USES_CONVERSION;
		
		//this should be atleast (and mostly, unless you derive your interface
		//from more than one other interface) one
		if( pTAttr->cImplTypes == 0 )	//this will be so only of IUnknown
			return "";					//which of course has no base interface
		
		HRESULT hr = m_pTInfo->GetRefTypeOfImplType( 0, &hRefType );
		if( SUCCEEDED( hr ) )
		{
			hr = m_pTInfo->GetRefTypeInfo( hRefType, &pTInfoBase );
			if( SUCCEEDED( hr ) )
			{
				hr = pTInfoBase->GetDocumentation( -1, &bstrName, NULL, NULL, NULL );
				if( SUCCEEDED( hr ) )
					strName = W2CA( bstrName.m_str );
			}
		}
		
		return strName;
	}
	
	
	
	
	void xllTypeLib::MoreInfo(UINT uiIndex) 
	{
		// CComPtr<ITypeInfo> pTInfo;
		std::string strTitle;
		
		
		//release any previous reference
		m_pTInfo.Release();
		
		HRESULT hr = m_pTLib->GetTypeInfo( uiIndex, &m_pTInfo );
		
		if( SUCCEEDED( hr ) )
		{
			TYPEATTR *pTAttr;
			HRESULT hr = m_pTInfo->GetTypeAttr( &pTAttr );
			if( SUCCEEDED( hr ) )
			{
				strTitle = GetTypeString( pTAttr->typekind );
				
				//is it a struct, union or enum?
				if( pTAttr->typekind == TKIND_RECORD ||
					pTAttr->typekind == TKIND_UNION ||
					pTAttr->typekind == TKIND_ENUM )
					EnumStructTypes( pTAttr );
				
				//is it a "typedef"?
				if( pTAttr->typekind == TKIND_ALIAS )
					EnumAliasTypes( pTAttr );
				
				//is it an interface?
				if( pTAttr->typekind == TKIND_INTERFACE ||
					pTAttr->typekind == TKIND_DISPATCH )
				{
					std::string strBaseInterface = GetBaseInterface( pTAttr );
					if( !strBaseInterface.empty() )
					{
						strTitle.append(" - ").append(strBaseInterface);
					}
					
					EnumInterfaceTypes( pTAttr );
				}
				
				//is it a module?
				if( pTAttr->typekind == TKIND_MODULE )
				{
					EnumModuleTypes( pTAttr );
				}
				
				//is it a coclass?
				if( pTAttr->typekind == TKIND_COCLASS )
					EnumCoClassTypes( pTAttr );
				
				m_pTInfo->ReleaseTypeAttr( pTAttr );
			}
		}
	}
	//}
	
	
	
	
	// Enumerate the libraries top level items (Interface, Enum, Module, ...)
	
	bool xllTypeLib::EnumerateItems()
	{
		std::string strType;
		TYPEKIND typeKind;
		HRESULT hr;
		
		std::cout << m_libName << " ("
			<< m_wMajorVerNum << "."
			<< m_wMinorVerNum 
			<< ")\n\n\n";
		
		// enumerate top level items
		UINT uiCount = m_pTLib->GetTypeInfoCount();
		
		for( UINT i = 0 ; i < uiCount ; ++i )
		{
			hr = m_pTLib->GetTypeInfoType( i, &typeKind );
			
			if( SUCCEEDED( hr ) )
			{
				//we put it here so that the string is released when
				//it goes out of scope
				CComBSTR bstrName, bstrHelpString, bstrHelpFile;
				DWORD dwHelpContext;
				
				//get the type in string format
				strType = GetTypeString( typeKind );
				
				//get the human readable name and description
				hr = m_pTLib->GetDocumentation( i, &bstrName, &bstrHelpString, &dwHelpContext, &bstrHelpFile );
				
				if( SUCCEEDED( hr ) )
				{
					std::cout << strType.c_str() << " ";
					std::wcout << bstrName.m_str;
					if (bstrHelpString.m_str) {
						std::wcout << L" {" << bstrHelpString.m_str << L"}";
					}
					std::wcout << L"\n";
					if (bstrHelpFile.m_str) {
						std::wcout << L"help context " << dwHelpContext << L" in file "
							<< bstrHelpFile.m_str << L"\n";
					}
					std::cout << "+----------------------------------+\n";
					MoreInfo(i);
					std::cout << "  ----------------------------------+\n\n";
				}
			}
		}
		
		return true;
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	xllTypeLib::xllTypeLib(LPCWSTR tlbFile)
	{
		InitLibInfo(tlbFile);
	}
	
	xllTypeLib::xllTypeLib(LPCSTR tlbFile)
	{
		USES_CONVERSION;
		InitLibInfo( A2CW(tlbFile) );
	}
	
	xllTypeLib::xllTypeLib(HINSTANCE hInst)
	{
		WCHAR szDllName[256];	// misused below as buffer for ANSI strings
		DWORD dwRet = GetModuleFileNameW(hInst, szDllName, 255);
		if( dwRet <= 0) {
			FormatMessageA( 
				FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL,
				GetLastError(),
				MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
				(LPSTR) szDllName,
				255,
				NULL
				);
			throw( std::exception( (LPSTR)szDllName ) );
		}
		
		InitLibInfo(szDllName);
	}
	
	
	void xllTypeLib::InitLibInfo(LPCWSTR tlbFile)
	{
		CComBSTR bstrLibName, bstrDocString, bstrHelpFile;
		TLIBATTR *pTLibAttr;
		HRESULT hr;
		USES_CONVERSION;
		
		hr = LoadTypeLib( tlbFile, &m_pTLib );
		if( FAILED( hr ) )
			throw( std::exception("LoadTypeLib() failed") );
		m_tlbFile = W2CA(tlbFile);
		
		m_wMajorVerNum = 0;
		m_wMinorVerNum = 0;
		m_libHelpContext = 0;
		m_modHelpContext = 0;
		
		hr = m_pTLib->GetLibAttr( &pTLibAttr );
		if( FAILED( hr ) )
			throw( std::exception("GetLibAttr() failed") ); 
		
		m_wMajorVerNum = pTLibAttr->wMajorVerNum;
		m_wMinorVerNum = pTLibAttr->wMinorVerNum;
		m_pTLib->ReleaseTLibAttr( pTLibAttr );
		
		//when you supply -1 as the index you get the library doc
		hr = m_pTLib->GetDocumentation( -1, &bstrLibName, &bstrDocString, &m_libHelpContext, &bstrHelpFile );
		if( FAILED( hr ) )
			throw( std::exception("GetDocumentation() for the libraray failed") ); 
		
		if( !!bstrLibName )
			m_libName = W2CA( bstrLibName.m_str );
		if( !!bstrDocString ) 
			m_libDocString = W2CA( bstrDocString.m_str );
		if( !!bstrHelpFile ) 
			m_libHelpFile = W2CA( bstrHelpFile.m_str );
	}
	
	
	void xllTypeLib::registerModule(LPCWSTR modName, LPCWSTR funcName)
	{
		std::string strType;
		USES_CONVERSION;
		CComBSTR _modName(modName);
		TYPEKIND typeKind;
		HRESULT hr;
		
		UINT uiCount = m_pTLib->GetTypeInfoCount(); // top level items
		
		for( UINT i = 0 ; i < uiCount ; ++i )
		{
			//we put it here so that the string is released when
			//it goes out of scope
			CComBSTR bstrName, bstrDocString, bstrHelpFile;
			
			hr = m_pTLib->GetTypeInfoType( i, &typeKind );
			if(	FAILED( hr ) || ( typeKind != TKIND_MODULE ) )
				continue;
			
			//get the human readable name and description
			hr = m_pTLib->GetDocumentation( i, &bstrName, &bstrDocString, &m_modHelpContext, &bstrHelpFile );
			if( FAILED( hr ) || ( modName && !(bstrName == _modName) ) )
				continue;
			
			m_modName = W2CA( bstrName.m_str );
			
			if( !bstrDocString ) {	// doc string
				m_modDocString = m_libDocString;	// defaults for use in
				m_modCategory = m_modName;			// the function wizard
			} else {
				m_modDocString = W2CA( bstrDocString.m_str );
				m_modCategory = m_modDocString; 
			}
			
			if( !bstrHelpFile )	// help file
				m_modHelpFile = m_libHelpFile;	// default
			else 
				m_modHelpFile = W2CA( bstrHelpFile.m_str );
			
			if( !m_modHelpContext )			// help context
				m_modHelpContext = m_libHelpContext;
			
			enumModuleFunctions(i, funcName);		// walk through this module
		}
	}
	
	void xllTypeLib::enumModuleFunctions(UINT uiIndex, LPCWSTR funcName) 
	{
		HRESULT hr;
		
		m_pTInfo.Release(); // release any previous reference
		hr = m_pTLib->GetTypeInfo( uiIndex, &m_pTInfo );
		if( FAILED( hr ) )
			throw( std::exception("GetTypeInfo() for module failed") );
		
		// QI for ITypeInfo2 (this may be done by casting ITypeInfo!?)
		m_pTInfo2.Release();
		hr = m_pTInfo->QueryInterface(&m_pTInfo2);
		if( FAILED( hr ) )
			throw( std::exception("QueryInterface for ITypeInfo2 failed") );
		
		// extract information from ITypeInfo2
		_variant_t v;
		hr = m_pTInfo2->GetCustData(XLL_CATEGORY, &v);
		if( SUCCEEDED( hr ) && (v.vt != VT_EMPTY) ) {
			_bstr_t data ( v );
			m_modCategory = (LPCSTR)data;
		}
		
		TYPEATTR *pTAttr;
		hr = m_pTInfo->GetTypeAttr( &pTAttr );
		if( FAILED( hr ) )
			throw( std::exception("GetTypeAttr() for module failed") );
		
		for( int i = 0 ; i < pTAttr->cFuncs ; ++i )
		{
			// TODO: check name
			registerModuleFunction(i);
			// TODO: check for exceptions and release pTAttr ???
		}
		
		m_pTInfo->ReleaseTypeAttr( pTAttr );
	}
	
	int xllTypeLib::registerModuleFunction(UINT uiIndex)
	{
		HRESULT hr;
		_variant_t v;
		USES_CONVERSION;
		
		FUNCDESC *pFuncDesc;
		hr = m_pTInfo->GetFuncDesc( uiIndex, &pFuncDesc );
		if( FAILED( hr ) )
			throw(std::exception("Failed to access function description"));
		
		//get the human readable name and description
		CComBSTR bstrName, bstrDocString, bstrHelpFile;
		DWORD dwHelpContext;
		hr = m_pTInfo->GetDocumentation(	pFuncDesc->memid, 
												&bstrName, 
												&bstrDocString, 
												&dwHelpContext, 
												&bstrHelpFile );
		if( FAILED( hr ) )
			throw(std::exception("GetDocumentation() for function failed"));
		
		if( !bstrName )
			throw(std::exception("GetDocumentation() did not return function name"));
		std::string fnName = W2CA( bstrName.m_str );
		
		std::string fnDocString;
		if( !bstrDocString ) 
			fnDocString = m_modDocString;  // default
		else 
			fnDocString = W2CA( bstrDocString.m_str );
		
		//			if( !bstrHelpFile ) 
		//				funcHelpFile  = m_modHelpFile.c_str();	// default
		//			else 
		//				funcHelpFile = W2CA( bstrHelpFile.m_str );
		
		// convert context to string
		if( dwHelpContext == 0)
			dwHelpContext = m_modHelpContext;		// default
		
		// extract category from ITypeInfo2 custom property
		std::string fnCategory;
		hr = m_pTInfo2->GetFuncCustData(uiIndex, XLL_CATEGORY, &v);
		if( SUCCEEDED( hr ) && (v.vt != VT_EMPTY) ) {
			_bstr_t data ( v );
			fnCategory = (LPCSTR)data;
		} else {
			fnCategory = m_modCategory;
		}
		
		// extract function text from ITypeInfo2 custom property
		std::string fnFunctionText;
		hr = m_pTInfo2->GetFuncCustData(uiIndex, XLL_FUNC_TEXT, &v);
		if( SUCCEEDED( hr ) && (v.vt != VT_EMPTY) ) {
			_bstr_t data ( v );
			fnFunctionText = (LPCSTR)data;
		} else {
			fnFunctionText = fnName;
		}

		// extract macro type from ITypeInfo2 custom property
		short macro_type = 1;
		hr = m_pTInfo2->GetFuncCustData(uiIndex, XLL_MACRO_TYPE, &v);
		if( SUCCEEDED( hr ) && (v.vt != VT_EMPTY) ) {
			macro_type = (short)v;
		}

		// extract short cut from ITypeInfo2 custom property
		std::string fnShortCut;
		hr = m_pTInfo2->GetFuncCustData(uiIndex, XLL_SHORTCUT, &v);
		if( SUCCEEDED( hr ) && (v.vt != VT_EMPTY) ) {
			_bstr_t data ( v );
			fnShortCut = (LPCSTR)data;
		}
		
		// prepare Excel REGISTER function stack and calculate size
		
		xllCallStack pxl;
		// 10 fixed params + arg help + 1 dummy element, see below
		pxl.reserve(10 + pFuncDesc->cParams + 1); 
		
		pxl.push_back(XlfOper(m_tlbFile));
		
		AddArguments(pFuncDesc, fnFunctionText, pxl);
		pxl.push_back(XlfOper(macro_type));
		pxl.push_back(XlfOper(fnCategory));
		pxl.push_back(XlfOper(fnShortCut));
		//   see original libXLL code where I convert to a string 
		pxl.push_back(XlfOper((double) dwHelpContext));	// jt: uh-oh...
		pxl.push_back(XlfOper(fnDocString));
		AddArgumentHelp(pFuncDesc, uiIndex, pxl);
		
		// add a dummy element to the end so Excel doesn't chop the last helpstring
		pxl.push_back(XlfOper((short)1));

		// call to Excel and register this function
		XLOPER xlRes;
		int xlret = XlfExcel::Instance().Callv(
			xlfRegister, &xlRes, pxl.size(), (LPXLOPER FAR *)&pxl.front()
		);
		
		// release all XlfOpers and manually free the global buffer
		pxl.clear();
		XlfExcel::Instance().FreeMemory();

		//			std::cout << "  ----------------------+" << std::endl << std::endl;
		m_pTInfo->ReleaseFuncDesc( pFuncDesc );
		
		return xlret;
	}
		
		
	// Make a first pass through all the arguments to
	//  - collect type information and
	//  - build the argument names string
	void xllTypeLib::AddArguments( 
		const FUNCDESC *pFuncDesc,
		const std::string &functionText, 
		xllCallStack &pxl)
	{
		// get the function ([0]) and argument ([1..]) names
		std::vector<BSTR> rgBstrNames;
		rgBstrNames.reserve(pFuncDesc->cParams+1);
		
		// BSTR rgBstrNames[20];
		// unsigned int cMaxNames = 20;
		unsigned int pcNames;
		HRESULT hr = m_pTInfo->GetNames(pFuncDesc->memid, 
			&rgBstrNames.front(), pFuncDesc->cParams+1, 
			&pcNames);
		if( FAILED( hr ) )
			throw(std::exception("Failed to access argument names"));
		
		// give control of data to _bstr_t
		// other array mebers are freed in the loop below
		_bstr_t bstrFuncName(rgBstrNames[0], false);
		std::string fnName((LPCSTR)bstrFuncName);
		
		// I am not sure of this...
		assert(pcNames == (pFuncDesc->cParams+1) );
		
		int modInPlace = 0;

		std::string strArgTypes, strArgNames;
		for( int i = 0 ; i < pFuncDesc->cParams ; ++i )
		{
			ELEMDESC Param = pFuncDesc->lprgelemdescParam[i];
			bool bInOut = (0 != (Param.paramdesc.wParamFlags & PARAMFLAG_FOUT));
			if( bInOut ) modInPlace = i+1;	// modified in place and returned
			std::string argType = ExcelTypeDesc( &Param.tdesc, bInOut );
			strArgTypes.append(argType);
			
			// give control of data to _bstr_t so it is freed
			// when leaving the scope
			_bstr_t argName(rgBstrNames[i+1], false);
			if( argType.size() > 0 ) {
				if( strArgNames.size() > 0) 
					strArgNames.append(", ");
				strArgNames.append((LPCSTR)argName);
			}
		}
		
		// specify function return type
		std::string argType = ExcelTypeDesc( &pFuncDesc->elemdescFunc.tdesc, false );
		if( modInPlace > 0 && argType.size() == 0 && modInPlace <=9 ) {
			if( modInPlace == 1) {
				argType = ">";
			} else {
				char s[4];
				sprintf(s,"%1d",modInPlace);
				argType = std::string(s);
			}
		}
		strArgTypes = argType.append(strArgTypes);
		
		// check function calling convention
		if( pFuncDesc->callconv != CC_STDCALL ) {
			throw( std::exception(fnName.append(": Excel can only register functions using __stdcall ").c_str() ));
		}
		
		pxl.push_back(XlfOper(fnName));
		pxl.push_back(XlfOper(strArgTypes));
		pxl.push_back(XlfOper(functionText));
		pxl.push_back(XlfOper(strArgNames));
	}
	
	
	
	// Make a second pass through all the arguments to collect
	//  - argument help strings and
	//  - default argument values
	void xllTypeLib::AddArgumentHelp(
		const FUNCDESC *pFuncDesc, 
		UINT uiIndex, 
		xllCallStack &pxl)
	{
		HRESULT hr;
		_variant_t v;
		for( int i = 0 ; i < pFuncDesc->cParams ; ++i )
		{
			std::string strArgHelp;
			// argument help string
			hr = m_pTInfo2->GetParamCustData(uiIndex, i, XLL_ARG_HELP, &v);
			if( SUCCEEDED( hr ) && (v.vt != VT_EMPTY) ) {
				_bstr_t data ( v );
				strArgHelp = (LPCSTR)data;
			}
			
			// argument default value
			hr = m_pTInfo2->GetParamCustData(uiIndex, i, XLL_ARG_DEFAULT, &v);
			if( SUCCEEDED( hr ) && (v.vt != VT_EMPTY) ) {
				_bstr_t data ( v );
				strArgHelp += " (default: ";
				strArgHelp += (LPCSTR)data;
				strArgHelp += ")";
			}
			pxl.push_back(XlfOper(strArgHelp));
		}
	}
	
	
	std::string xllTypeLib::ExcelTypeDesc( const TYPEDESC *pTypeDesc, bool bInOut, bool bPtr )
	{
		std::string strType = "?";
		
		switch( pTypeDesc->vt )
		{
		case VT_PTR:
			// flag as "by ref" values
			return ExcelTypeDesc( pTypeDesc->lptdesc, bInOut, true);
			
		case VT_SAFEARRAY:
			// I think Excel does not know a direct safearray.
			// type 'K' is a ptr to a {rows/cols/double array[]} struct FP
			throw( std::exception("SAFEARRAY not supported by Excel worksheet functions") );
			
		case VT_CARRAY:
			{
				// ARRAYDESC *pArrDesc = pTypeDesc->lpadesc;
				// std::string strAttr;
				// char s[64];
				// 
				// strText = StringifyTypeDesc( &( pArrDesc->tdescElem ) );
				// for( int i = 0 ; i < pArrDesc->cDims ; ++i )
				// {
				// 	sprintf(s, "[%d]", pArrDesc->rgbounds[i].cElements - 1 );
				// 	strText.append(s);
				// }
				// return strType;
				
				// Excel can only handle this with two leading row/cols
				// parameter as type 'O'
				// Until we implement this we throw an error
				
				throw( std::exception("CARRAY not (yet) supported by Excel worksheet functions") );
			}
			
		case VT_USERDEFINED:
			{
				return ExcelUserDefinedType( pTypeDesc->hreftype, bInOut, bPtr );
				
			}
		}
		
		return ExcelVariantType( pTypeDesc->vt, bInOut, bPtr );
	}
	
	
	std::string xllTypeLib::ExcelVariantType( VARTYPE vt, bool bInOut, bool bPtr )
	{
		switch( vt )
		{
		case VT_I2: // short: I by Val, M by Ref
			if (bPtr != bInOut) 
				throw( std::exception( StringifyVariantType(vt).append(": In/Out and Pointer mismatch").c_str() ) );
			return bPtr?"M":"I";
			
		case VT_I4: // long: J by Val, N by Ref
		case VT_HRESULT:
			if (bPtr != bInOut) 
				throw( std::exception( StringifyVariantType(vt).append(": In/Out and Pointer mismatch").c_str() ) );
			return bPtr?"N":"J";
			
		case VT_R8: // double: B by Val, E by Ref
		case VT_DATE: // double: B by Val, E by Ref
			if (bPtr != bInOut) 
				throw( std::exception( StringifyVariantType(vt).append(": In/Out and Pointer mismatch").c_str() ) );
			return bPtr?"E":"B";
			
		case VT_BOOL: // short: A by Val, L by Ref
			if (bPtr != bInOut) 
				throw( std::exception( StringifyVariantType(vt).append(": In/Out and Pointer mismatch").c_str() ) );
			return bPtr?"L":"A";
			
		case VT_UI2: // unsigned short: H by Val and used tohether with O type
			if( bPtr ) {
				// this maybe a hint for a coming O type, so we ignore it for now
				return "";
			}
			return "H";
			
		case VT_INT:
			throw( std::exception("machine INT is unsafe, use SHORT or LONG") );
			
		case VT_UINT:
			throw( std::exception("machine UINT is unsafe, use USHORT or ULONG") );
			
		case VT_VOID:
			return "";
			
		case VT_LPSTR: // char *
			return bInOut?"F":"C";
			
		}
		
		throw( std::exception( StringifyVariantType(vt).append(": not supported by Excel worksheet functions").c_str() ) );
	}
	
	
	std::string xllTypeLib::ExcelUserDefinedType( HREFTYPE hRefType, bool bInOut, bool bPtr  )
	{
		CComPtr<ITypeInfo> pTInfoCust;
		CComBSTR bstrName;
		HRESULT hr;
		USES_CONVERSION;
		
		hr = m_pTInfo->GetRefTypeInfo( hRefType, &pTInfoCust );
		if( SUCCEEDED( hr ) )
		{
			hr = pTInfoCust->GetDocumentation( -1, &bstrName, NULL, NULL, NULL );
			if( SUCCEEDED( hr ) ) {
				std::string strText = W2CA( bstrName.m_str );
				if( bPtr ) { 
					// handle ptr types
					if( 0 == strText.compare("_FP") ) {
						return "K";
					} else if ( 0 == strText.compare("xloper") ) {
						return "R";
					}
				} else {
					// usually an enum / long int
					// TODO: try to resolve name
					return "J";
				}
				throw( std::exception(strText.append(": can not convert this type").c_str() ) );
			}
			
		}
		
		throw( std::exception( "can not resolve user defined type" ) );
	}
		
		
		
} // namespace xll
