# Microsoft Developer Studio Generated NMAKE File, Based on xllExample.dsp
!IF "$(CFG)" == ""
CFG=xllExample - Win32 Debug
!MESSAGE No configuration specified. Defaulting to xllExample - Win32 Debug.
!ENDIF 

!IF "$(CFG)" != "xllExample - Win32 Release" && "$(CFG)" != "xllExample - Win32 Debug"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "xllExample.mak" CFG="xllExample - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "xllExample - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "xllExample - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 

!IF  "$(CFG)" == "xllExample - Win32 Release"

OUTDIR=.\Release
INTDIR=.\Release
# Begin Custom Macros
OutDir=.\Release
# End Custom Macros

!IF "$(RECURSE)" == "0" 

ALL : "$(OUTDIR)\xllExample.dll"

!ELSE 

ALL : "xlw - Win32 Release" "$(OUTDIR)\xllExample.dll"

!ENDIF 

!IF "$(RECURSE)" == "1" 
CLEAN :"xlw - Win32 ReleaseCLEAN" 
!ELSE 
CLEAN :
!ENDIF 
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(INTDIR)\xllExample.obj"
	-@erase "$(INTDIR)\xllExample.res"
	-@erase "$(OUTDIR)\xllExample.dll"
	-@erase "$(OUTDIR)\xllExample.exp"
	-@erase "$(OUTDIR)\xllExample.lib"
	-@erase ".\xllExample.h"
	-@erase ".\xllExample.tlb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MD /W3 /GX /O2 /I ".." /I "../../stochastix/QuantLib" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "XLLEXAMPLE_EXPORTS" /Fp"$(INTDIR)\xllExample.pch" /YX /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /c 

.c{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/nologo /I ".." /D "NDEBUG" /win32 
RSC=rc.exe
RSC_PROJ=/l 0x407 /fo"$(INTDIR)\xllExample.res" /d "NDEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\xllExample.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /incremental:no /pdb:"$(OUTDIR)\xllExample.pdb" /machine:I386 /def:".\xllExample.def" /out:"$(OUTDIR)\xllExample.dll" /implib:"$(OUTDIR)\xllExample.lib" /libpath:"..\lib\Win32\VisualStudio" /libpath:"..\..\stochastix\QuantLib\lib\Win32\VisualStudio" 
DEF_FILE= \
	".\xllExample.def"
LINK32_OBJS= \
	"$(INTDIR)\xllExample.obj" \
	"$(INTDIR)\xllExample.res" \
	"..\lib\Win32\VisualStudio\xlw.lib"

"$(OUTDIR)\xllExample.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
   type xllExample.idl | ..\idl2def\idl2def > xllExample.def
	 $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

SOURCE="$(InputPath)"

!ELSEIF  "$(CFG)" == "xllExample - Win32 Debug"

OUTDIR=.\Debug
INTDIR=.\Debug
# Begin Custom Macros
OutDir=.\Debug
# End Custom Macros

!IF "$(RECURSE)" == "0" 

ALL : "$(OUTDIR)\xllExample.dll" ".\xllExample.tlb" ".\xllExample.h"

!ELSE 

ALL : "xlw - Win32 Debug" "$(OUTDIR)\xllExample.dll" ".\xllExample.tlb" ".\xllExample.h"

!ENDIF 

!IF "$(RECURSE)" == "1" 
CLEAN :"xlw - Win32 DebugCLEAN" 
!ELSE 
CLEAN :
!ENDIF 
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(INTDIR)\vc60.pdb"
	-@erase "$(INTDIR)\xllExample.obj"
	-@erase "$(INTDIR)\xllExample.res"
	-@erase "$(OUTDIR)\xllExample.dll"
	-@erase "$(OUTDIR)\xllExample.exp"
	-@erase "$(OUTDIR)\xllExample.ilk"
	-@erase "$(OUTDIR)\xllExample.lib"
	-@erase "$(OUTDIR)\xllExample.pdb"
	-@erase ".\xllExample.h"
	-@erase ".\xllExample.tlb"

"$(OUTDIR)" :
    if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

CPP=cl.exe
CPP_PROJ=/nologo /MDd /W3 /Gm /GR /GX /ZI /Od /I ".." /I "../../stochastix/QuantLib" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "XLLEXAMPLE_EXPORTS" /Fp"$(INTDIR)\xllExample.pch" /YX /Fo"$(INTDIR)\\" /Fd"$(INTDIR)\\" /FD /GZ /c 

.c{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx{$(INTDIR)}.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

MTL=midl.exe
MTL_PROJ=/nologo /I ".." /D "_DEBUG" /win32 
RSC=rc.exe
RSC_PROJ=/l 0x407 /fo"$(INTDIR)\xllExample.res" /d "_DEBUG" 
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\xllExample.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /incremental:yes /pdb:"$(OUTDIR)\xllExample.pdb" /debug /machine:I386 /def:".\xllExample.def" /out:"$(OUTDIR)\xllExample.dll" /implib:"$(OUTDIR)\xllExample.lib" /pdbtype:sept /libpath:"..\lib\Win32\VisualStudio" /libpath:"..\..\stochastix\QuantLib\lib\Win32\VisualStudio" 
DEF_FILE= \
	".\xllExample.def"
LINK32_OBJS= \
	"$(INTDIR)\xllExample.obj" \
	"$(INTDIR)\xllExample.res" \
	"..\lib\Win32\VisualStudio\xlwd.lib"

"$(OUTDIR)\xllExample.dll" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
   type xllExample.idl | ./idl2def > xllExample.def
	 $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<

SOURCE="$(InputPath)"

!ENDIF 


!IF "$(NO_EXTERNAL_DEPS)" != "1"
!IF EXISTS("xllExample.dep")
!INCLUDE "xllExample.dep"
!ELSE 
!MESSAGE Warning: cannot find "xllExample.dep"
!ENDIF 
!ENDIF 


!IF "$(CFG)" == "xllExample - Win32 Release" || "$(CFG)" == "xllExample - Win32 Debug"
SOURCE=.\xllExample.cpp

"$(INTDIR)\xllExample.obj" : $(SOURCE) "$(INTDIR)" ".\xllExample.h"


SOURCE=.\xllExample.idl

!IF  "$(CFG)" == "xllExample - Win32 Release"

MTL_SWITCHES=/nologo /I ".." /D "NDEBUG" /tlb "xllExample.tlb" /h "xllExample.h" /win32 

".\xllExample.tlb"	".\xllExample.h" : $(SOURCE) "$(INTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ELSEIF  "$(CFG)" == "xllExample - Win32 Debug"

MTL_SWITCHES=/nologo /I ".." /D "_DEBUG" /tlb "xllExample.tlb" /h "xllExample.h" /win32 

".\xllExample.tlb"	".\xllExample.h" : $(SOURCE) "$(INTDIR)"
	$(MTL) @<<
  $(MTL_SWITCHES) $(SOURCE)
<<


!ENDIF 

SOURCE=.\xllExample.rc

"$(INTDIR)\xllExample.res" : $(SOURCE) "$(INTDIR)" ".\xllExample.tlb"
	$(RSC) $(RSC_PROJ) $(SOURCE)


!IF  "$(CFG)" == "xllExample - Win32 Release"

"xlw - Win32 Release" : 
   cd "\Projects\XLW"
   $(MAKE) /$(MAKEFLAGS) /F .\xlw.mak CFG="xlw - Win32 Release" 
   cd ".\xllExample"

"xlw - Win32 ReleaseCLEAN" : 
   cd "\Projects\XLW"
   $(MAKE) /$(MAKEFLAGS) /F .\xlw.mak CFG="xlw - Win32 Release" RECURSE=1 CLEAN 
   cd ".\xllExample"

!ELSEIF  "$(CFG)" == "xllExample - Win32 Debug"

"xlw - Win32 Debug" : 
   cd "\Projects\XLW"
   $(MAKE) /$(MAKEFLAGS) /F .\xlw.mak CFG="xlw - Win32 Debug" 
   cd ".\xllExample"

"xlw - Win32 DebugCLEAN" : 
   cd "\Projects\XLW"
   $(MAKE) /$(MAKEFLAGS) /F .\xlw.mak CFG="xlw - Win32 Debug" RECURSE=1 CLEAN 
   cd ".\xllExample"

!ENDIF 


!ENDIF 

