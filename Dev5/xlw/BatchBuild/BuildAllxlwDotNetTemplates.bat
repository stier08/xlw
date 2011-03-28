"C:\Python27\python.exe" prepareDotNetTemplateProject.py

REM Debug

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\
set vcbuildtoolpath=C:\Program Files (x86)\Microsoft Visual Studio 8\vc\vcpackages

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VS8\Template.sln"  /t:rebuild  /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetTemplate_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS8_PRO\HybridTemplate.sln"  /t:rebuild  /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetHybridTemplate_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VisualBasic\VB2005\Template.sln"  /t:rebuild  /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetTemplate_VB_Debug_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS8_PRO\HybridTemplate.sln"  /t:rebuild  /property:Configuration=Debug >> xlwDotNetTemplate_VB_Debug_8.log 2>&1 



set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\IDE\
set vcbuildtoolpath=C:\Program Files (x86)\Microsoft Visual Studio 9.0\vc\vcpackages

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VS9\Template.sln"  /t:rebuild  /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetTemplate_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS9_PRO\HybridTemplate.sln"  /t:rebuild  /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetHybridTemplate_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VisualBasic\VB2008\Template.sln"  /t:rebuild  /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetTemplate_VB_Debug_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS9_PRO\HybridTemplate.sln"  /t:rebuild  /property:Configuration=Debug >> xlwDotNetTemplate_VB_Debug_9.log 2>&1 

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\
set vcbuildtoolpath=C:\Program Files (x86)\Microsoft Visual Studio 10.0\vc\vcpackages

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VS10\Template.sln"  /t:rebuild  /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetTemplate_Debug_10.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS10_PRO\HybridTemplate.sln"  /t:clean  /property:Configuration=Debug >> xlwDotNetHybridTemplate_Debug_10.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS10_PRO\HybridTemplate.sln"  /t:build  /property:Configuration=Debug >> xlwDotNetHybridTemplate_Debug_10.log 2>&1 
REM "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VisualBasic\VB2010\Template.sln"  /t:rebuild  /property:Configuration=Debug   /property:Platform=x86   >> xlwDotNetTemplate_VB_Debug_10.log 2>&1 


REM Release

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\
set vcbuildtoolpath=C:\Program Files (x86)\Microsoft Visual Studio 8\vc\vcpackages

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VS8\Template.sln"  /t:rebuild  /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetTemplate_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS8_PRO\HybridTemplate.sln"  /t:rebuild  /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetHybridTemplate_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VisualBasic\VB2005\Template.sln"  /t:rebuild  /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetTemplate_VB_Release_8.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS8_PRO\HybridTemplate.sln"  /t:rebuild  /property:Configuration=Release >> xlwDotNetTemplate_VB_Release_8.log 2>&1 

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\IDE\
set vcbuildtoolpath=C:\Program Files (x86)\Microsoft Visual Studio 9.0\vc\vcpackages

"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VS9\Template.sln"  /t:rebuild  /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetTemplate_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS9_PRO\HybridTemplate.sln"  /t:rebuild  /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetHybridTemplate_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VisualBasic\VB2008\Template.sln"  /t:rebuild  /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetTemplate_VB_Release_9.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v3.5\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS9_PRO\HybridTemplate.sln"  /t:rebuild  /property:Configuration=Release >> xlwDotNetTemplate_VB_Release_9.log 2>&1 

set DevEnvDir=C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\
set vcbuildtoolpath=C:\Program Files (x86)\Microsoft Visual Studio 10.0\vc\vcpackages

"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VS10\Template.sln"  /t:rebuild  /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetTemplate_Release_10.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS10_PRO\HybridTemplate.sln"  /t:clean  /property:Configuration=Release  >> xlwDotNetHybridTemplate_Release_10.log 2>&1 
"C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\Hybrid_Cpp_CSharp_XLLs\VS10_PRO\HybridTemplate.sln"  /t:build  /property:Configuration=Release >> xlwDotNetHybridTemplate_Release_10.log 2>&1
REM "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild"   "C:\Temp\xlwDotNetTemplate Projects\VisualBasic\VB2010\Template.sln"  /t:rebuild  /property:Configuration=Release   /property:Platform=x86   >> xlwDotNetTemplate_VB_Release_10.log 2>&1 


