@echo off
rem ������ ��� ���� 䠩��� ����ਡ�⨢�
rem ��ࠬ��� - ��姭��� ����� ���ᨨ
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

rem ��⠫��, �㤠 ������� ����ਡ�⨢
set dpup=%prcat%\..\..\upload

rem �६���� ��⠫�� ��� ᡮન ����ਡ�⨢�
set dppr=%dpup%\Temp
set dplib=%dpup%\TempLib

rem ������騪 
set wz=%CD%\..\..\bin\pkzip25.exe

set prrescompile=%prcat%\RES\ENG

set commoncontrols=%prcat%\..\CommonControls\Source

set d4cat=C:\Program Files\Borland\Delphi4
set d5cat=C:\Program Files\Borland\Delphi5
set d6cat=C:\Program Files\Borland\Delphi6
set d7cat=C:\Program Files\Borland\Delphi7
set bcb5cat=C:\Program Files\Borland\CBuilder5
set bcb6cat=C:\Program Files\Borland\CBuilder6
