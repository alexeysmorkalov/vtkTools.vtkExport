rem *********************************************
rem Usage:
rem   CopyDelphiDCU.bat <��⠫���㤠�����஢���> <�����Delphi���⮖��ன>
rem *********************************************

set dest=%1
set dver=%2

copy %prcat%\Source\*.dcu %dest%
copy %prcat%\Source\*.dfm %dest%
copy %prcat%\Source\*.inc %dest%
copy %prcat%\Source\prPackage%dver%.dpk %dest%
copy %prcat%\Source\pr_strings.pas %dest%

copy %prcat%\Source\*.res %dest%
del %dest%\prPackage*.res
copy %prcat%\Source\prPackage%dver%.res %dest%
