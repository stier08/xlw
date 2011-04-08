set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\
set vcbuildtoolpath=C:\Program Files (x86)\Microsoft Visual Studio 10.0\vc\vcpackages


"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Debug /t:rebuild  ..\xlw\build\vc10\xlw.sln  >> xlw10_Debug.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Release /t:rebuild  ..\xlw\build\vc10\xlw.sln  >> xlw10_Release.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Debug /t:rebuild  ..\xlwDotNet\Build\VS10\xlw.Net.sln  >> xlwDotNet10_Debug.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Release /t:rebuild  ..\xlwDotNet\Build\VS10\xlw.Net.sln  >> xlwDotNet10_Release.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Debug;Platform=x64 /t:rebuild  ..\xlw\build\vc10\xlw.sln  >> xlw10_Debug_x64.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Release;Platform=x64 /t:rebuild  ..\xlw\build\vc10\xlw.sln  >> xlw10_Release_x64.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Debug;Platform=x64 /t:rebuild  ..\xlwDotNet\Build\VS10\xlw.Net.sln  >> xlwDotNet10_Debug.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\msbuild.exe" /p:Configuration=Release;Platform=x64 /t:rebuild  ..\xlwDotNet\Build\VS10\xlw.Net.sln  >> xlwDotNet10_Release.log 2>&1 

set DevEnvDir=
set vcbuildtoolpath=
