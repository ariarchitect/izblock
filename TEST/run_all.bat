REM echo off
chcp 1251
set ACAD="C:\Program Files\Autodesk\AutoCAD 2019\acad.exe"
set SCR="C:\Users\Main\Downloads\run_TST25.scr"
set DWGDIR=D:\Home\YandexDisk\ПСО НАША ПАПКА\ПСО\Исходка\АР\DWG\АР1

%ACAD% /nologo "%DWGDIR%\240304-Лист - АР-09 - Cекция 1, секция 4- План 6-го этажа на отм- +21-350 М1_100.dwg" /b %SCR%
%ACAD% /nologo "%DWGDIR%\240304-Лист - АР-10 - Cекция 1, секция 4- План 7-го этажа на отм- +24-500 М1_100.dwg" /b %SCR%
%ACAD% /nologo "%DWGDIR%\240304-Лист - АР-11 - Cекция 1, секция 4- План 8-го этажа на отм- +27-650 М1_100.dwg" /b %SCR%
%ACAD% /nologo "%DWGDIR%\240304-Лист - АР-12 - Cекция 1, секция 4- План 9-го этажа на отм- +30-800 М1_100.dwg" /b %SCR%