# Microsoft Developer Studio Generated NMAKE File, Based on idl2def.dsp
!IF "$(CFG)" == ""
CFG=idl2def - Win32 Release
!MESSAGE No configuration specified. Defaulting to idl2def - Win32 Release.
!ENDIF 

!IF "$(CFG)" != "idl2def - Win32 Release"
!MESSAGE Invalid configuration "$(CFG)" specified.
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "idl2def.mak" CFG="idl2def - Win32 Release"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "idl2def - Win32 Release" (based on "Win32 (x86) Console Application")
!MESSAGE 
!ERROR An invalid configuration is specified.
!ENDIF 

!IF "$(OS)" == "Windows_NT"
NULL=
!ELSE 
NULL=nul
!ENDIF 

OUTDIR=.
INTDIR=.
# Begin Custom Macros
OutDir=.
# End Custom Macros

ALL : "$(OUTDIR)\idl2def.exe"


CLEAN :
	-@erase "$(INTDIR)\idl2def.obj"
	-@erase "$(INTDIR)\vc60.idb"
	-@erase "$(OUTDIR)\idl2def.exe"

CPP=cl.exe
CPP_PROJ=/nologo /ML /W3 /GX /O1 /D "WIN32" /D "NDEBUG" /D "_CONSOLE" /D "_MBCS" /Fp"$(INTDIR)\idl2def.pch" /Yu"stdafx.h" /FD /c 

.c.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx.obj::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.c.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cpp.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

.cxx.sbr::
   $(CPP) @<<
   $(CPP_PROJ) $< 
<<

RSC=rc.exe
BSC32=bscmake.exe
BSC32_FLAGS=/nologo /o"$(OUTDIR)\idl2def.bsc" 
BSC32_SBRS= \
	
LINK32=link.exe
LINK32_FLAGS=kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:console /incremental:no /pdb:"$(OUTDIR)\idl2def.pdb" /machine:I386 /out:"$(OUTDIR)\idl2def.exe" 
LINK32_OBJS= \
	"$(INTDIR)\idl2def.obj"

"$(OUTDIR)\idl2def.exe" : "$(OUTDIR)" $(DEF_FILE) $(LINK32_OBJS)
    $(LINK32) @<<
  $(LINK32_FLAGS) $(LINK32_OBJS)
<<


!IF "$(NO_EXTERNAL_DEPS)" != "1"
!IF EXISTS("idl2def.dep")
!INCLUDE "idl2def.dep"
!ELSE 
!MESSAGE Warning: cannot find "idl2def.dep"
!ENDIF 
!ENDIF 


!IF "$(CFG)" == "idl2def - Win32 Release"
SOURCE=.\idl2def.cpp
CPP_SWITCHES=/nologo /ML /W3 /GX /O1 /D "WIN32" /D "NDEBUG" /D "_CONSOLE" /D "_MBCS" /FD /c 

"$(INTDIR)\idl2def.obj" : $(SOURCE) "$(INTDIR)"
	$(CPP) @<<
  $(CPP_SWITCHES) $(SOURCE)
<<



!ENDIF 

