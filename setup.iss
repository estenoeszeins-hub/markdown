[Setup]
AppName=Convertidor MD Zeins
AppVersion=1.0
DefaultDirName={autopf}\ConvertidorMDZeins
DefaultGroupName=Convertidor MD Zeins
OutputDir=dist
OutputBaseFilename=Instalador_Convertidor_Zeins
Compression=lzma
SolidCompression=yes
; "admin" pide los permisos al iniciar la instalación
PrivilegesRequired=admin
SetupIconFile=

[Files]
Source: "dist\Convertidor_Zeins_V1.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Convertidor MD Zeins"; Filename: "{app}\Convertidor_Zeins_V1.exe"
Name: "{autodesktop}\Convertidor MD Zeins"; Filename: "{app}\Convertidor_Zeins_V1.exe"

[Run]
Filename: "{app}\Convertidor_Zeins_V1.exe"; Description: "{cm:LaunchProgram,Convertidor MD Zeins}"; Flags: nowait postinstall skipifsilent
