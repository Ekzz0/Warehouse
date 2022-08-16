; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{2822D7B4-16F1-4D8B-9366-7C6FAA4082E8}
AppName=Pallet Counter
AppVersion=1.0
;AppVerName=Pallet Counter 1.0
AppPublisher=Borinskikh Semen, Inc.
DefaultDirName={pf}\Pallet Counter
DefaultGroupName=Pallet Counter
AllowNoIcons=yes
LicenseFile=C:\Users\S\PycharmProjects\Warehouse\output\License.txt
OutputDir=C:\Users\S\Desktop
OutputBaseFilename=setup
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\S\PycharmProjects\Warehouse\output\Pallet Counter.exe"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\Pallet Counter"; Filename: "{app}\Pallet Counter.exe"
Name: "{commondesktop}\Pallet Counter"; Filename: "{app}\Pallet Counter.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Pallet Counter.exe"; Description: "{cm:LaunchProgram,Pallet Counter}"; Flags: nowait postinstall skipifsilent

