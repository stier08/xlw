/*! $Id$
 *********************************************************************
 *              
 *  \file       XllMain_c.h
 *              
 *  \brief      boilerplate module for libXLL
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
 *  Revision 1.1.2.1  2003/02/20 16:41:59  nando
 *  libXLL added
 *
 *  Revision 1.4  2002/03/17 16:45:47  jens
 *  initial semi-public release
 *
 *
 *********************************************************************
 */

#include "xll.h"
#include "XllTypeLib.h"
#include "macros.h"

// TODO: We can get this from the type lib

#ifndef XLL_LONG_NAME
#define XLL_LONG_NAME "Build with libXLL, http://jens-thiel.de/finance/libxll"
#endif

extern "C" {

/*!
 * xlAddInManagerInfo()
 *
 * This function is called by the Add-in Manager to find the long name
 * of the add-in. If xAction = 1, this function should return a string
 * containing the long name of this XLL, which the Add-in Manager will use
 * to describe this XLL. If xAction = 2 or 3, this function should return
 * #VALUE!. This is what the documentation says, but... The following info
 * combines some information from Excel 97 SDK to explain the undocumented
 * functionality of xlAddInManagerInfo()
 */

/*
	Defining Custom Menus and Custom Commands

	You can use the Init Menus and Init Commands subkeys to define custom
	menus and commands. These custom menus and commands appear every time
	you start Microsoft Excel and enable you to load add-ins or standard
	macro sheets after you choose the custom command.

		Note: The Add-In Manager (the Add-Ins command on the Tools menu) 
			reads and writes entries in both the Init Menus and Init 
			Commands subkeys.

	This is done on program exit after the very first installation of the
	Add-In by calling xlAddInManagerInfo() with actions 2..4. By returning
	the desired REG_SZ entries successively we can add as many entries as
	we want. 
	
	There seems to be a problem left here: I dont know what to return for the
	first action=2 call. Obviously this is not going to the registry, but is
	may be some indicator wether to add an OPEN value, and still continue
	to ask for reg keys.

	http://msdn.microsoft.com/library/en-us/office97/html/SFA92.asp
	(with additional notes by Jens Thiel <Jens@Thiel.de>)
*/

LPXLOPER __declspec(dllexport) xlAddInManagerInfo(XlfOper pxAction)
{
	static int flag2 = 0;
	static int flag3 = 0;

	switch( pxAction.AsInt() )
	{
	case 1:
		return (LPXLOPER)XlfOper(XLL_LONG_NAME);
		break;

	/*
		The Init Commands Subkey

			HKCU\Software\Microsoft\Office\<Version>\Excel\Init Commands\

		The syntax for Init Commands subkey entries is as follows (one line):

		  <Keyword>=<menu_bar_num>,<menu_name>,<command_name>,<macro>,
			  <command_position>,<macro_key>,<status_text>,<help_reference>

		Argument Description:

		<keyword>
		  A unique keyword, such as Scenario, that Microsoft Excel uses to
		  identify commands added by an INI file.
		  In case that you add this key by using xlAddInManagerInfo, you
		  do not supply a keyword. The Add-In manager automatically uses
		  the file name of your Add-In (with the full path and file extension
		  stripped off). If you supply  multiple commands, the Add-In manager 
		  adds numbers to the keyword, e.g. ADDIN, ADDIN1, ADDIN2, ...
 
		<menu_bar_num>
		  The number of the built-in menu bar to which you want to add the
		  command.
 
		<menu_name>
		  The name of the new menu. You can use an ampersand '&' to mark a
		  related key.
 
		<command_name>
		  The name of the new command.
 
		<macro>
		  A reference to a macro on an add-in workbook. Choosing the command
		  opens the add-in and runs the specified macro.
 
		<command_position>
		   The position of the command on the menu. This may be the name of
		   the command after which you want to place the new command or a
		   number indicating the command's position on the menu. If omitted, 
		   the command appears at the end of the menu.
 
		<macro_key>
           The key assigned to the macro, if any.
 
		<status_text>
		  A message to be displayed in the status bar when the command is
		  selected.
 
		<help_reference>
		  The file name and topic number for a custom Help topic relating to 
		  the command.
 
		The following example shows how an Init Commands entry specifies two
		commands to be added to built-in menu bar number 1.

		  Views=1,Window,&View...,'C:\EXCEL\LIBRARY\VIEWS.XLA'!STUB,
		        -,,Show or define a named view,EXCELHLP.HLP!1730
		  Solver=10,Extras,Sol&ver...,'C:\EXCEL\LIBRARY\SOLVER\SOLVER.XLA'!STUB,
			    ,,Find solution to worksheet model,EXCELHLP.HLP!1830

		http://msdn.microsoft.com/library/en-us/office97/html/SFA92.asp
		(with additional notes by Jens Thiel <Jens@Thiel.de>)
	*/

	case 2:
		if( flag2 < 2 ) {  
			flag2++;
			if(flag2==1)
				// TODO: I still do not know what magic Excel expects here...
				// returning a string results in the OPEN= key not being set
				// anything other stops the Add-In manager asking us about reg keys...
				return (LPXLOPER)XlfOper::Error(xlerrValue);
				// return (LPXLOPER)XlfOper("10,Extras,&Install...,'xlAutoOpen',,,,");
			else
				return (LPXLOPER)XlfOper("10,Extras,&Install...,'C:\\Dokumente und Einstellungen\\Administrator\\Eigene Dateien\\Visual Studio Projects\\xlw\\xllExample\\Debug\\xllExample.dll'!STUB,---,,Text,");
		}
		break;

	/*
		The Init Menus Subkey

			HKCU\Software\Microsoft\Office\<Version>\Excel\Init Menus\

		The syntax for Init Menus subkey entries is as follows:

		  <Keyword>=<menu_bar_num>,<menu_name>,<menu_position>[,<menu_parent>]

		Argument Description:

		<keyword>
		  A unique keyword, such as Custom1, that Microsoft Excel uses to
		  identify menus added by an Init Menus registry entry.
		  In case that you add this key by using xlAddInManagerInfo, you
		  do not supply a keyword. The Add-In manager automatically uses
		  the file name of your Add-In (with the full path and file extension
		  stripped off). If you supply  multiple menus, the Add-In manager 
		  adds numbers to the keyword, e.g. ADDIN, ADDIN1, ADDIN2, ...
 
		<menu_bar_num>
		  The number of the built-in menu bar to which you want to add the
		  menu or command.
 
		<menu_name>
		  The name of the new menu.
 
		<menu_position>
		  The position of the new menu on the menu bar. This may be the name
		  of the menu after which you want to place the new menu or a number 
		  indicating the menu's position from the left of the menu bar. If 
		  there is a menu_parent, then menu_position is the position of the 
		  new menu on the menu_parent.
 
		<menu_parent>
		  Optional. If defining a submenu, this is the menu name or number
		  on the menu bar that will contain this new submenu.
 
		Note: Microsoft Excel doesn't allow more than one new level of submenus.

		The following example shows how an Init Menus entry specifies a custom 
		menu to be added to the right of the Window menu.

		  Custom1=1,Work,Window

		http://msdn.microsoft.com/library/en-us/office97/html/SFA92.asp
		(with additional notes by Jens Thiel <Jens@Thiel.de>)
	*/

	case 3:
		if( flag3 < 1 ) {
			flag3++;
			return (LPXLOPER)XlfOper("10,&XLL Example,7");
		}
		break;

	/*	
		The Delete Commands Subkey

		  HKCU\Software\Microsoft\Office\<Version>\Excel\Delete Commands\

		This optional subkey allows you to delete built-in Microsoft Excel 
		commands. The syntax of the line to delete commands is:

			<Keyword>=<menu_bar_num>,<menu_name>,<command_position>

		The definitions of <keyword>, <menu_bar_num>, <menu_name>, and 
		<command_position> are the same as those given in the Init Commands
		subkey.

		  Caution: Don't delete the Exit command on the File menu unless 
		           you've included another way to quit Microsoft Excel!


		http://msdn.microsoft.com/library/en-us/office97/html/S10106.asp
		(with additional notes by Jens Thiel <Jens@Thiel.de>)
	*/

	case 4:

		// we do not want to delete menus, do we...?

	default:
		break;
	}
	return (LPXLOPER)XlfOper::Error(xlerrValue);
}

/*!
 * xlAutoAdd()
 *
 * This function is called by the Add-in Manager only. When you add a
 * DLL to the list of active add-ins, the Add-in Manager calls xlAutoAdd()
 * and then opens the XLL, which in turn calls xlAutoOpen.
 */
 
BOOL __declspec(dllexport) xlAutoAdd(void)
{
	// XlfExcel::MsgBox("Thank you for enabling our Add-In", XLL_LONG_NAME);
	return TRUE;
}

/*!
 * xlAutoRemove()
 *
 * This function is called by the Add-in Manager only. When you remove
 * an XLL from the list of active add-ins, the Add-in Manager calls
 * xlAutoRemove() and then UNREGISTER("EXAMPLE.XLL").
 * 
 * You can use this function to perform any special tasks that need to be
 * performed when you remove the XLL from the Add-in Manager's list
 * of active add-ins. For example, you may want to delete an
 * initialization file when the XLL is removed from the list.
 */

BOOL __declspec(dllexport) xlAutoRemove(void)
{
	// XlfExcel::MsgBox("Thank you for using our Add-In", XLL_LONG_NAME);
	return TRUE;
}

/*!
 * xlAutoOpen()
 *
 * Microsoft Excel uses xlAutoOpen to load XLL files. When you open an XLL file, 
 * the only action Microsoft Excel takes is to call the xlAutoOpen function.
 *
 * More specifically, xlAutoOpen is called:
 *
 *  - when you open this XLL file from the File menu,
 *  - when this XLL is in the XLSTART directory, and is
 *    automatically opened when Microsoft Excel starts,
 *  - when Microsoft Excel opens this XLL for any other reason, or
 *  - when a macro calls REGISTER(), with only one argument, which is the
 *    name of this XLL.
 *
 * xlAutoOpen is also called by the Add-in Manager when you add this XLL as an
 * add-in. The Add-in Manager first calls xlAutoAdd, then calls 
 * REGISTER("EXAMPLE.XLL"), which in turn calls xlAutoOpen.
 *
 * xlAutoOpen should:
 *
 * - register all the functions you want to make available while this
 *   XLL is open,
 * - add any menus or menu items that this XLL supports,
 * - perform any other initialization you need, and
 * - return 1 if successful, or return 0 if your XLL cannot be opened.
 *
 */

BOOL __declspec(dllexport) xlAutoOpen(void)
{
    XlfExcel::Instance().SendMessage("Registering library...");
	xll::xllTypeLib typelib(XlfExcel::Instance().GetName().c_str());
	typelib.registerModule();
    XlfExcel::Instance().SendMessage();
	return TRUE;
}

/*!
 * xlAutoClose()
 *
 * xlAutoClose is called by Microsoft Excel:
 * 
 *  - when you quit Microsoft Excel, or
 *  - when a macro sheet calls UNREGISTER(), giving a string argument
 *    which is the name of this XLL.
 *
 * xlAutoClose is called by the Add-in Manager when you remove this XLL from
 * the list of loaded add-ins. The Add-in Manager first calls xlAutoRemove,
 * then calls UNREGISTER("EXAMPLE.XLL"), which in turn calls xlAutoClose.
 * 
 * xlAutoClose should:
 * 
 *  - Remove any menus or menu items that were added in xlAutoOpen,
 *  - do any necessary global cleanup, and
 *  - delete any names that were added (names of exported functions, and 
 *    so on). Remember that registering functions may cause names to 
 *    be created.
 *
 *  xlAutoClose does NOT have to unregister the functions that were registered
 *  in xlAutoOpen. This is done automatically by Microsoft Excel after
 *  xlAutoClose returns.
 * 
 */

BOOL __declspec(dllexport) xlAutoClose(void)
{
	delete &XlfExcel::Instance();
	return TRUE;
}

}	// extern "C"

/*!
 * DllMain()
 *
 * Windows 95 and Windows NT call one function, DllMain, for both 
 * initialization and termination. It also makes calls on both a 
 * per-process and per-thread basis, so several initialization calls
 * can be made if a process is multithreaded.
 * This function is called when the DLL is first loaded, with a dwReason
 * of DLL_PROCESS_ATTACH. 
 */
 
BOOL APIENTRY DllMain( HANDLE hModule, DWORD dwReason, LPVOID)
{
    if( dwReason == DLL_PROCESS_ATTACH)
        XlfExcel::Instance().DllHandle(hModule);
    return TRUE;
}
