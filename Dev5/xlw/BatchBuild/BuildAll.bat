REM  CLEAN ALL

python CleanAll.py


REM BUILD THE LIBRRIES

REM CALL BuildOnlyDevCpp.bat 
CALL BuildOnlyGCCMake.bat 
CALL BuildOnlyVS7.bat 
CALL BuildOnlyVS8.bat 
CALL BuildOnlyVS9.bat 
CALL BuildOnlyVS10.bat 

REM BUILD XLW EXAMPLES

CALL BuildAllExamples.bat

REM BUILD XLW TEMPLATES

CALL BuildAllxlwTemplates.bat

REM --------------------------------- XLWDOTNET 
REM BUILD XLW TEMPLATES

CALL BuildAllxlwDotNetExamples.bat
CALL BuildAllxlwDotNetTemplates.bat









