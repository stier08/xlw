
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Debug /t:rebuild  ..\xlw\build\vc10\xlw.sln  >> xlw10_Debug.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Release /t:rebuild  ..\xlw\build\vc10\xlw.sln  >> xlw10_Release.log 2>&1 
REM "C:\Program Files\Microsoft Visual Studio 10.0\VC\vcpackages\vcbuild.exe" /rebuild  ..\xlwDotNet\Build\VS9\xlw.Net.sln >> xlwDotNet9.log  2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Debug /t:rebuild  ..\xlwDotNet\Build\VS10\xlw.Net.sln  >> xlwDotNet10_Debug.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Release /t:rebuild  ..\xlwDotNet\Build\VS10\xlw.Net.sln  >> xlwDotNet10_Release.log 2>&1 
