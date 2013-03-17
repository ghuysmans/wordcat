[Setup]
AppName=WordCat
AppVerName=WordCat v0.9 r6
DefaultDirName={pf}\WordCat
DefaultGroupName=WordCat
VersionInfoCopyright=by Guillaume HUYSMANS, 2013
OutputDir=.
LicenseFile=..\license.txt
ShowLanguageDialog=no

[Files]
Source: DLL\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: DLL\COMCTL32.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: ..\WordCat.exe; DestDir: {app}
Source: ..\example\*.*; DestDir: {app}\example; Flags: recursesubdirs
Source: ..\lang\*.*; DestDir: {app}\lang; Flags: recursesubdirs
Source: ..\help\*.*; DestDir: {app}\help; Flags: recursesubdirs
Source: ..\license.txt; DestDir: {app}

[Icons]
Name: {group}\WordCat; Filename: {app}\WordCat.exe; WorkingDir: {app}
Name: {group}\Example; Filename: {app}\example
Name: {group}\License; Filename: {app}\license.txt
Name: {group}\Uninstall; Filename: {uninstallexe}
