


"C:\Program Files\Microsoft Visual Studio .NET 2003\Common7\IDE\devenv.exe"   xlw\build\vc7\xlw.sln  /rebuild Debug    /out xlw7Debug.log
"C:\Program Files\Microsoft Visual Studio .NET 2003\Common7\IDE\devenv.exe"   xlw\build\vc7\xlw.sln  /rebuild Release  /out xlw7Release.log
"C:\Program Files\Microsoft Visual Studio 8\VC\vcpackages\vcbuild.exe" /rebuild    xlw\build\vc8\xlw.sln  >> xlw8.log 2>&1 
"C:\Program Files\Microsoft Visual Studio 9.0\VC\vcpackages\vcbuild.exe" /rebuild  xlw\build\vc9\xlw.sln  >> xlw9.log 2>&1 

"C:\Program Files\Microsoft Visual Studio 8\VC\vcpackages\vcbuild.exe"  /rebuild   xlwNet\DotNet\Build\VS8\xlw.Net.sln >> xlwDotNet8.log  2>&1 
"C:\Program Files\Microsoft Visual Studio 9.0\VC\vcpackages\vcbuild.exe" /rebuild  xlwNet\DotNet\Build\VS9\xlw.Net.sln >> xlwDotNet9.log  2>&1 
