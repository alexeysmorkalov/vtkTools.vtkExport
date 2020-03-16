@echo off

set prversion=%1
call Settings.bat

echo *********************************************************
echo
echo (bat file works only on NT/2000/XP):
echo
echo   %src%.zip - vtkExport sources
echo
echo *********************************************************
pause

md %dpup%
rmdir %dppr% /s /q
md %dppr%

rem *********************************************
rem Формируем каталог
rem *********************************************
md %dppr%\vtkExport
md %dppr%\vtkExport\Source
md %dppr%\vtkExport\Demos
md %dppr%\vtkExport\Help
md %dppr%\vtkExport\Help\rus
md %dppr%\vtkExport\Help\eng

rem *********************************************
rem Sources
rem *********************************************
copy %prcat%\source\*.pas %dppr%\vtkExport\Source
copy %prcat%\source\*.res %dppr%\vtkExport\Source
copy %prcat%\source\*.bpk %dppr%\vtkExport\Source
copy %prcat%\source\*.dfm %dppr%\vtkExport\Source
copy %prcat%\source\*.inc %dppr%\vtkExport\Source
copy %prcat%\source\*.dpk %dppr%\vtkExport\Source
copy %prcat%\source\*.cpp %dppr%\vtkExport\Source
copy %prcat%\source\*.inc %dppr%\vtkExport\Source

rem *********************************************
rem Demos
rem *********************************************
call batch\CopyDemos.bat

rem *********************************************
rem Addititional files
rem *********************************************
copy %prcat%\..\vtkExport_documentation\eng\vtkExport.chm %dppr%\vtkExport\Help\eng
copy %prcat%\..\vtkExport_documentation\rus\vtkExport.chm %dppr%\vtkExport\Help\rus

rem *********************************************
rem Создаем архив
rem *********************************************

%wz% -add -max -dir=relative %dpup%\%src%.zip %dppr%\*.*

rmdir %dppr% /q /s
