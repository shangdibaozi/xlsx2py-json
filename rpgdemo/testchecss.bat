call basePath.bat
@echo off
set pydatas=%ktpydatas%/d_chessConfig.py
set excel1=%ktexcels%/xlsxs/chessConfig.xlsx
echo on
python3 ../xlsx2py/xlsx2py.py %pydatas% %excel1%
if not defined ktall (ping -n 30 127.1>nul)

