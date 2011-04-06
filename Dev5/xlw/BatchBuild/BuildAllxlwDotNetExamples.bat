
REM Debug

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\Example\VS8\Example.sln /t:rebuild   /property:Configuration=Debug    /property:Platform=x86 >> xlwDotNetExample_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS8\NonPassive.sln /t:rebuild /property:Configuration=Debug    /property:Platform=x86  >> xlwDotNetNonPassive_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS8\PythonDemo.sln /t:rebuild /property:Configuration=Debug    /property:Platform=x86  >> xlwDotNetPython_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS8\RTDExample.sln /t:rebuild /property:Configuration=Debug    /property:Platform=x86  >> xlwDotNetRTDExample_Debug_8.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\Example\VS8\Example.sln /t:rebuild   /property:Configuration=Debug    /property:Platform=x64 >> xlwDotNetExample_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS8\NonPassive.sln /t:rebuild /property:Configuration=Debug    /property:Platform=x64  >> xlwDotNetNonPassive_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS8\PythonDemo.sln /t:rebuild /property:Configuration=Debug    /property:Platform=x64  >> xlwDotNetPython_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS8\RTDExample.sln /t:rebuild /property:Configuration=Debug    /property:Platform=x64  >> xlwDotNetRTDExample_Debug_8.log 2>&1 


set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\IDE\

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\Example\VS9\Example.sln /t:rebuild  /property:Configuration=Debug    /property:Platform=x86 >> xlwDotNetExample_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS9\NonPassive.sln /t:rebuild /property:Configuration=Debug     /property:Platform=x86 >> xlwDotNetNonPassive_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS9\PythonDemo.sln /t:rebuild /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetPython_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS9\RTDExample.sln /t:rebuild  /property:Configuration=Debug   /property:Platform=x86  >> xlwDotNetRTDExample_Debug_9.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\Example\VS9\Example.sln /t:rebuild  /property:Configuration=Debug    /property:Platform=x64 >> xlwDotNetExample_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS9\NonPassive.sln /t:rebuild /property:Configuration=Debug     /property:Platform=x64 >> xlwDotNetNonPassive_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS9\PythonDemo.sln /t:rebuild /property:Configuration=Debug   /property:Platform=x64   >> xlwDotNetPython_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS9\RTDExample.sln /t:rebuild  /property:Configuration=Debug   /property:Platform=x64  >> xlwDotNetRTDExample_Debug_9.log 2>&1 

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\Example\VS10\Example.sln /t:rebuild  /property:Configuration=Debug    /property:Platform=x86 >> xlwDotNetExample_Debug_10.log 2>&1 
REM "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS10\NonPassive.sln /t:rebuild /property:Configuration=Debug     /property:Platform=x86 >> xlwDotNetNonPassive_Debug_10.log 2>&1 
REM "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS10\PythonDemo.sln /t:rebuild /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetPython_Debug_10.log 2>&1 
REM "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS10\RTDExample.sln /t:rebuild  /property:Configuration=Debug   /property:Platform=x86  >> xlwDotNetRTDExample_Debug_10.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\Example\VS10\Example.sln /t:rebuild  /property:Configuration=Debug    /property:Platform=x64 >> xlwDotNetExample_Debug_10.log 2>&1 


REM Release

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\Example\VS8\Example.sln /t:rebuild  /property:Configuration=Release   /property:Platform=x86  >> xlwDotNetExample_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS8\NonPassive.sln /t:rebuild /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetNonPassive_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS8\PythonDemo.sln /t:rebuild /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetPython_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS8\RTDExample.sln /t:rebuild  /property:Configuration=Release   /property:Platform=x86  >> xlwDotNetRTDExample_Release_8.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\Example\VS8\Example.sln /t:rebuild  /property:Configuration=Release   /property:Platform=x64  >> xlwDotNetExample_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS8\NonPassive.sln /t:rebuild /property:Configuration=Release   /property:Platform=x64   >> xlwDotNetNonPassive_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS8\PythonDemo.sln /t:rebuild /property:Configuration=Release   /property:Platform=x64   >> xlwDotNetPython_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS8\RTDExample.sln /t:rebuild  /property:Configuration=Release   /property:Platform=x64  >> xlwDotNetRTDExample_Release_8.log 2>&1 

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\IDE\

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\Example\VS9\Example.sln /t:rebuild /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetExample_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS9\NonPassive.sln /t:rebuild /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetNonPassive_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS9\PythonDemo.sln /t:rebuild /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetPython_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS9\RTDExample.sln /t:rebuild  /property:Configuration=Release   /property:Platform=x86  >> xlwDotNetRTDExample_Release_9.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\Example\VS9\Example.sln /t:rebuild /property:Configuration=Release   /property:Platform=x64   >> xlwDotNetExample_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS9\NonPassive.sln /t:rebuild /property:Configuration=Release   /property:Platform=x64   >> xlwDotNetNonPassive_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS9\PythonDemo.sln /t:rebuild /property:Configuration=Release   /property:Platform=x64   >> xlwDotNetPython_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS9\RTDExample.sln /t:rebuild  /property:Configuration=Release   /property:Platform=x64  >> xlwDotNetRTDExample_Release_9.log 2>&1 

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\Example\VS10\Example.sln /t:rebuild  /property:Configuration=Release    /property:Platform=x86 >> xlwDotNetExample_Release_10.log 2>&1 
REM "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\XtraExamples\NonPassive\VS10\NonPassive.sln /t:rebuild /property:Configuration=Release     /property:Platform=x86 >> xlwDotNetNonPassive_Release_10.log 2>&1 
REM "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\XtraExamples\Python\VS10\PythonDemo.sln /t:rebuild /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetPython_Release_10.log 2>&1 
REM "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\XtraExamples\RTDExample\VS10\RTDExample.sln /t:rebuild  /property:Configuration=Release   /property:Platform=x86  >> xlwDotNetRTDExample_Release_10.log 2>&1 

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   ..\xlwDotNet\Example\VS10\Example.sln /t:rebuild  /property:Configuration=Release    /property:Platform=x64 >> xlwDotNetExample_Release_10.log 2>&1 




