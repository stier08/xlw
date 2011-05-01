
#include<Test.h>
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
MacroCache<open>::MacroRegistra my_open_func_r("my_open_func","my_open_func",my_open_func);

void my_close_func()
{
	xlw::XlfExcel::MsgBox("Bye-Bye ..see you next time!","Bye-Bye");
}
MacroCache<close>::MacroRegistra my_close_func_r("my_close_func","my_close_func",my_close_func);

