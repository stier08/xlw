
Hello!

This is the second public release of libXLL. It contains the library and a 
working example of an add-in. 

Starting with this release, the libXLL code has incorporated the XLW library 
(http://xlw.sf.net). You can find the xlw ReadMes in xlwREADMEs and doxygen
files for the documentation in Docs. In the future, these two projects will
maybe merge into a single library.


Updated information from the first release follows:

The idea with this "yet-another-excel-helper" is to use IDL to supply worksheet 
functions. We have eg. the primary code implemented in COM/C++ (which uses IDL 
to define the necessary interfaces) and provide some functionality as worksheet 
functions, also defined in IDL. We package the COM and the add-in functionality 
in a single .DLL file, and the "libXLL-exported-functions" are also accessible 
by VBA or any other application that understands typelibs.

As I'm told, Excel 2002 directly supports COM worksheets functions -- but I have 
not tried 2002 yet. But this is, of course, not compatible with the countless 
1999/2000 installations still around. BTW, a third possibility to write 
worksheet functions is to write VBA functions, since these are also accessible 
from the worksheet.

Concerning speed, I think that C++ exported as worksheet functions should 
provide the best performance. Also, letting Excel "coerce" (that is, convert 
cell values to the type your function wants to work with) the arguments should 
be faster than doing it yourself. This means that you should prefer eg. "double" 
over "xloper" types, if it's a double you really need.

The documentation currently consists of the step-by-step introduction "Setting up 
a project with libXLL.pdf" and the resulting xllExample project. It contains a 
sample add-in with worksheet functions defined in an .IDL file. I also included 
xlfRegister.html in the archive; refer to this file for data types to be used as 
arguments. You can also generate the XLW documentation from the Docs directory.

The file layout of xllExample is:

  xllExample.cpp   Main DLL module
  xllExample.def   DLL exported functions, auto-hgenerated
  xllExample.idl   Worksheet functions to be exported 
  xllExample.xls   Worksheet using the xll_example add-in 
  
The .idl, .cpp and also the XllMain_c.h include files contain a lot of comments.
If you have QuantLib or MKL installed, you can uncomment the #includes in
xllExample.cpp; otherwise two of the examples will only return dummy values.



Bye for now,

Jens.

ps: My email is jens@stochastix.de

-- 
/* Jens Thiel * Stochastix GmbH, Germany * +49-700-STOCHASTIX */ 
