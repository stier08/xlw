@SETLOCAL
@SET LOG_ROOT=%CD%\%2-VS8-
%VS8VSBUILD% %1 /rebuild "Debug|Win32"  >> "%LOG_ROOT%x86-Debug.log" 2>&1 
%VS8VSBUILD% %1 /rebuild "Release|Win32"  >> "%LOG_ROOT%x86-Release.log" 2>&1 
%VS8VSBUILD% %1 /rebuild "Debug|x64"  >> "%LOG_ROOT%x64-Debug.log" 2>&1 
%VS8VSBUILD% %1 /rebuild "Release|x64"  >> "%LOG_ROOT%x64-Release.log" 2>&1 
@ENDLOCAL