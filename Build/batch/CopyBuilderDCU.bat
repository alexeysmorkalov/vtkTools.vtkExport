rem *********************************************
rem Usage:
rem   CopyBuilderDCU.bat <��⠫���㤠�����஢���> <�����Builder���⮖��ன>
rem *********************************************

set dest=%1
set dver=%2

copy %prcat%\Source\*.dcu %dest%
copy %prcat%\Source\*.dfm %dest%
copy %prcat%\Source\*.hpp %dest%
copy %prcat%\Source\*.obj %dest%
copy %prcat%\Source\*.inc %dest%
copy %prcat%\Source\prPackageCB%dver%.bpk %dest%
copy %prcat%\Source\prPackageCB%dver%.cpp %dest%
copy %prcat%\Source\pr_strings.pas %dest%

copy %prcat%\Source\*.res %dest%
del %dest%\prPackage*.res
copy %prcat%\Source\prPackageCB%dver%.res %dest%
