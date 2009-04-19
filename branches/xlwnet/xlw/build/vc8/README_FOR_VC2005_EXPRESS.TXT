

Visual C++ 2005 EXPRESS
=======================

Visual C++ 2005 Express is not shipped with the Microsoft Platform SDK.
This is a pre-requisite for building xlw.


It can be downloaded from 

	http://www.microsoft.com/downloads/details.aspx?FamilyId=A55B6B43-E24F-4EA3-A93E-40C0EC4F68E5&displaylang=en

If you are using the Express edition of Visual C++ 2005, you need to install this first.

The standard package will install to C:\Program Files\Microsoft SDKs\Windows.

You need to set the environment Variable PSDK to the path of the SDK. 
For me  the variable PSDK is set to 

                  C:\Program Files\Microsoft SDKs\Windows\v6.1

on one machine and set to 


                  C:\Program Files\Microsoft SDKs\Windows\v6.0A

on a different machine.


If you already had Visual Studio open when you set the environment variable PSDK
you have to close it and re-open it for the new environment variable to be recognised.


-------------------------------------------------------------------------------------

These extra few steps are only neccassary for the EXPRESS edition of VC 2005.
They are NOT neccassary for Visual C++ 2005 Professional.



