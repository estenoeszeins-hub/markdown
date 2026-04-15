[Setup]
AppName=Convertidor MD Zeins
AppVersion=1.0
DefaultDirName={autopf}\ConvertidorMDZeins
DefaultGroupName=Convertidor MD Zeins
OutputDir=dist
OutputBaseFilename=Instalador_Convertidor_Zeins
Compression=lzma
SolidCompression=yes
; Pide permisos de administrador al iniciar la instalacion
PrivilegesRequired=admin

[Files]
; Asegurate de que el nombre del .exe coincida con el que genera PyInstaller
Source: "dist\Convertidor_Zeins_V1.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Crea el acceso directo en el menu Inicio
Name: "{autoprograms}\Convertidor MD Zeins"; Filename: "{app}\Convertidor_Zeins_V1.exe"
; Crea el acceso directo en el Escritorio
Name: "{autodesktop}\Convertidor MD Zeins"; Filename: "{app}\Convertidor_Zeins_V1.exe"

[Run]
; Opcion para ejecutar el programa justo despues de instalar
Filename: "{app}\Convertidor_Zeins_V1.exe"; Description: "Lanzar Convertidor MD"; Flags: nowait postinstall skipifsilent
