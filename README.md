This is a very minimalistic utility that helps you extract and merge PDF pages. You can also use the pyinstaller module to create a portable executable file (*.exe) from this and use it on any Windows machine.

# PyInstaller Command
pyinstaller --icon ".\python.ico" -n PDF_Extract_Merge --version-file ".\file_version_info.txt" --onefile -w main.py