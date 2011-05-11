REM ********** Set path for .net framework2.0, sandcastle,hhc,hxcomp****************************

set PATH=%windir%\Microsoft.NET\Framework\v2.0.50727;%DXROOT%\ProductionTools;c:\program files (x86)\HTML Help Workshop;%PATH%

cd Debug

call %XLW%\xlw\build\%2\Debug\XlwDocGen.exe %1.xll

pause

if exist output rmdir output /s /q

call "%DXROOT%\Presentation\vs2005\copyOutput.bat"

copy %XLW%\xlw\build\sandcastle\*.maml .

BuildAssembler /config:%XLW%\xlw\build\sandcastle\conceptual.config .\toc.xml

if not exist chm mkdir chm
if not exist chm\html mkdir chm\html
if not exist chm\icons mkdir chm\icons
if not exist chm\scripts mkdir chm\scripts
if not exist chm\styles mkdir chm\styles
REM: if not exist chm\media mkdir chm\media

xcopy output\icons\* chm\icons\ /y /r
xcopy output\scripts\* chm\scripts\ /y /r
xcopy output\styles\* chm\styles\ /y /r

ChmBuilder.exe /project:%1 /html:Output\html /lcid:1033 /toc:.\toc.xml /out:Chm /config:.\ChmBuilder.config

DBCSFix.exe /d:Chm /l:1033 

copy %XLW%\xlw\build\sandcastle\alias.h chm\alias.h

hhc chm\%1.hhp

copy chm\%1.chm ..

cd ..