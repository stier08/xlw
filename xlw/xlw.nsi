
# to be used with NSIS 1.91 and up
#
# for a light installer (no source code, no documentation):
#       makensis /DLIGHT XLW.nsi
# for a full installer
#       makensis XLW.nsi


# $Id$


!define VER_NUMBER "1.2.0"

# HEADER CONFIGURATION COMMANDS
!ifdef LIGHT
    Name "XLW Light"
    Caption "XLW Light - Setup"
    #do not change the name below
    OutFile "..\xlw-${VER_NUMBER}-light-inst.exe"
    ComponentText "This will install XLW ${VER_NUMBER} Light on your computer.$\n A more complete version including documentation, examples, source code, etc. can be downloaded from http://xlw.sf.net"
!else
    Name "XLW"
    Caption "XLW - Setup"
    #do not change the name below
    OutFile "..\xlw-${VER_NUMBER}-full-inst.exe"

    InstType "Full (w/ Source Code)"
    InstType Typical
    InstType Minimal

    ComponentText "This will install XLW ${VER_NUMBER} on your computer"
!endif

SilentInstall normal
CRCCheck on
LicenseText "You must agree with the following license before installing:"
LicenseData License.txt
DirShow show
DirText "Please select a location to install XLW (or use the default):"
InstallDir $PROGRAMFILES\XLW
InstallDirRegKey HKEY_LOCAL_MACHINE SOFTWARE\XLW "Install_Dir"
AutoCloseWindow false
ShowInstDetails hide
SetDateSave on

# INSTALLATION EXECUTION COMMANDS



Section "-XLW"
SectionIn 1 2 3
# this directory must be created first, or the CreateShortCut will not work
    CreateDirectory "$SMPROGRAMS\XLW"
    SetOutPath $INSTDIR
    File "LICENSE.TXT"
    File "NEWS.TXT"
    File "README.TXT"
    File "Authors.txt"
    File "Contributors.txt"
    File "History.txt"
    File "TODO.txt"

    SetOutPath  $INSTDIR\xlw
    File /r "xlw\*.h"
    File /r "xlw\*.inl"

    SetOutPath $INSTDIR\lib\Win32\VisualStudio
    File "lib\Win32\VisualStudio\xlw.lib"
    #SetOutPath $INSTDIR\lib\Win32\Borland
    #File "lib\Win32\Borland\xlw.lib"

    WriteRegStr HKEY_LOCAL_MACHINE \
                "Software\Microsoft\Windows\CurrentVersion\Uninstall\XLW" \
                "DisplayName" "XLW (remove only)"
    WriteRegStr HKEY_LOCAL_MACHINE \
                "Software\Microsoft\Windows\CurrentVersion\Uninstall\XLW" \
                "UninstallString" '"XLWUninstall.exe"'
    WriteRegStr HKEY_LOCAL_MACHINE \
                "SOFTWARE\XLW" \
                "Install_Dir" "$INSTDIR"
    WriteRegStr HKEY_CURRENT_USER \
                "Environment" \
                "XLW_DIR" "$INSTDIR"
    CreateShortCut "$SMPROGRAMS\XLW\Uninstall XLW.lnk" \
                   "$INSTDIR\XLWUninstall.exe" \
                   "" "$INSTDIR\XLWUninstall.exe" 0
    CreateShortCut "$SMPROGRAMS\XLW\README.TXT.lnk" \
                   "$INSTDIR\README.TXT"
    CreateShortCut "$SMPROGRAMS\XLW\LICENSE.TXT.lnk" \
                   "$INSTDIR\LICENSE.TXT"
    CreateShortCut "$SMPROGRAMS\XLW\What's new.lnk" \
                   "$INSTDIR\NEWS.TXT"

    WriteUninstaller "XLWUninstall.exe"
SectionEnd



!ifndef LIGHT

Section "Source Code"
SectionIn 1
  SetOutPath $INSTDIR
  File "ChangeLog.txt"
  File makefile.mak
  File xlw.mak
  File xlw.dsp
  File xlw.dsw
  File xlw.nsi

  SetOutPath  $INSTDIR\xlw
  File /r "xlw\*.cpp"

  CreateShortCut "$SMPROGRAMS\XLW\XLW project workspace.lnk" \
                 "$INSTDIR\xlw.dsw"

SectionEnd



Section "Examples"
SectionIn 1 2
  SetOutPath  $INSTDIR\xlwExample
  File /r "xlwExample\*.cpp"
  File /r "xlwExample\*.dsp"
  File /r "xlwExample\*.dsw"
  File /r "xlwExample\*.mak"
  File /r "xlwExample\*.xls"
	File /r "xlwExample\*.h"
	File /r "xlwExample\*.inl"

  SetOutPath  $INSTDIR\xlwExample\xll
  File /r "xlwExample\xll\*.txt"

  SetOutPath  $INSTDIR\xlwExample\xll\Win32
  File /r "xlwExample\xll\Win32\*.txt"

  SetOutPath  $INSTDIR\xlwExample\xll\Win32\VisualStudio
  File /r "xlwExample\xll\Win32\VisualStudio\*.txt"
  File /r "xlwExample\xll\Win32\VisualStudio\*.xll"

  #SetOutPath  $INSTDIR\xlwExample\xll\Win32\Borland
  #File /r "xlwExample\xll\Win32\Borland\*.txt"
  #File /r "xlwExample\xll\Win32\Borland\*.xll"

  IfFileExists $SMPROGRAMS\QuantLib 0 NoSourceShortCuts
    CreateShortCut "$SMPROGRAMS\XLW\Demo XLL workspace.lnk" \
                 "$INSTDIR\xlwExample\xlwExample.dsw"
    CreateShortCut "$SMPROGRAMS\XLW\Demo spreadsheet.lnk" \
                 "$INSTDIR\xlwExample\xlwExample.xls"
  NoSourceShortCuts:
SectionEnd

SectionDivider

Section "HTML documentation"
SectionIn 1 2
  SetOutPath "$INSTDIR\Docs\html"
  File "Docs\html\*.*"
  CreateShortCut "$INSTDIR\refman.html.lnk" "$INSTDIR\Docs\html\index.html"
  CreateShortCut "$SMPROGRAMS\XLW\Documentation (HTML).lnk" \
                 "$INSTDIR\Docs\html\index.html"
SectionEnd

Section "WinHelp documentation"
SectionIn 1
  SetOutPath "$INSTDIR\Docs"
  File "Docs\html\index.chm"
  CreateShortCut "$SMPROGRAMS\XLW\Documentation (WinHelp).lnk" \
                 "$INSTDIR\Docs\index.chm"
SectionEnd

Section "PDF documentation"
SectionIn 1
  SetOutPath "$INSTDIR\Docs"
  File "Docs\latex\refman.pdf"
  CreateShortCut "$SMPROGRAMS\XLW\Documentation (PDF).lnk" \
                 "$INSTDIR\Docs\refman.pdf"
SectionEnd


SectionDivider

Section "Start Menu Group"
SectionIn 1 2 3
  SetOutPath $SMPROGRAMS\XLW

  WriteINIStr "$SMPROGRAMS\XLW\XLW Home Page.url" \
              "InternetShortcut" "URL" "http://xlw.sf.net"

  CreateShortCut "$SMPROGRAMS\XLW\XLW Directory.lnk" \
                 "$INSTDIR"
SectionEnd

!endif


UninstallText "This will uninstall XLW. Hit next to continue."


Section "Uninstall"
    DeleteRegKey HKEY_LOCAL_MACHINE \
        "Software\Microsoft\Windows\CurrentVersion\Uninstall\XLW"
    DeleteRegKey HKEY_LOCAL_MACHINE SOFTWARE\XLW
    DeleteRegValue HKEY_CURRENT_USER  "Environment" "XLW_DIR"
    Delete "$SMPROGRAMS\XLW\*.*"
    RMDir "$SMPROGRAMS\XLW"
    RMDir /r "$INSTDIR\xlwExample"
    RMDir /r "$INSTDIR\Docs\html"
    RMDir /r "$INSTDIR\lib"
    RMDir /r "$INSTDIR\xlw"
    RMDir /r "$INSTDIR"
SectionEnd
