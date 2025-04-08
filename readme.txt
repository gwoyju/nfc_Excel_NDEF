打包指令（含 GUI 模式）
pyinstaller --onefile --noconsole nfc_gui_writer.py

打包指令（含多語系）
pyinstaller --onefile --noconsole --icon=nfc_excel_icon.ico nfc_gui_writer_multilang.py

安裝 Inno Setup
https://jrsoftware.org/isdl.php