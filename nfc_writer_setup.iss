
[Setup]
AppName=NFC vCard Writer
AppVersion=1.0
DefaultDirName={pf}\NFC vCard Writer
DefaultGroupName=NFC vCard Writer
OutputDir=dist
OutputBaseFilename=nfc_vcard_writer_setup
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\nfc_gui_writer_multilang.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "contacts.xlsx"; DestDir: "{app}"; Flags: ignoreversion
Source: "nfc_excel_icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\NFC vCard Writer"; Filename: "{app}\nfc_gui_writer_multilang.exe"; IconFilename: "{app}\nfc_excel_icon.ico"
Name: "{commondesktop}\NFC vCard Writer"; Filename: "{app}\nfc_gui_writer_multilang.exe"; IconFilename: "{app}\nfc_excel_icon.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "建立桌面捷徑"; GroupDescription: "其他工作："

[Run]
Filename: "{app}\nfc_gui_writer_multilang.exe"; Description: "啟動 NFC vCard Writer"; Flags: nowait postinstall skipifsilent
