
#include<cppinterface.h>
#pragma warning (disable : 4996)


short // echoes a short
EchoShort(short x // number to be echoed
           )
{
    return x;
}

#include<xlw/XlfExcel.h>
#include<xlw/XlOpenClose.h>


void my_open_func()
{
	xlw::XlfExcel::MsgBox("Thanks for Opening me!","Welcome");
}



void my_close_func()
{
	xlw::XlfExcel::MsgBox("Bye-Bye ..see you next time!","Bye-Bye");
}


extern "C"
{
	int __declspec(dllexport) func()
	{
		xlw::XlfExcel::MsgBox("Thanks for calling the menu item","Menu");
		return 0;
	}
}
xlw::XLRegistration::XLCommandRegistrationHelper theItem("func","func","comment","File","func");