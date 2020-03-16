@echo off
rem Шаблон для имен файлов дистрибутива
rem параметр - трехзначный номер версии
set d4=vtkExport%prversion%D4
set d5=vtkExport%prversion%D5
set d6=vtkExport%prversion%D6
set d7=vtkExport%prversion%D7
set bcb5=vtkExport%prversion%BCB5
set bcb6=vtkExport%prversion%BCB6
set docrus=prdoc%prversion%rus
set doceng=prdoc%prversion%eng
set src=vtkExport%prversion%SRC

set prcat=%CD%\..

rem Каталог, куда помещать дистрибутив
set dpup=%prcat%\..\..\upload

rem Временный каталог для сборки дистрибутива
set dppr=%dpup%\Temp
set dplib=%dpup%\TempLib

rem Упаковщик 
set wz=%CD%\..\..\bin\pkzip25.exe

set prrescompile=%prcat%\RES\ENG

set commoncontrols=%prcat%\..\CommonControls\Source

set d4cat=C:\Program Files\Borland\Delphi4
set d5cat=C:\Program Files\Borland\Delphi5
set d6cat=C:\Program Files\Borland\Delphi6
set d7cat=C:\Program Files\Borland\Delphi7
set bcb5cat=C:\Program Files\Borland\CBuilder5
set bcb6cat=C:\Program Files\Borland\CBuilder6
