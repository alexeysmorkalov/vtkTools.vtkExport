
{******************************************}
{                                          }
{           vtk Export library             }
{            Example  program              }
{      Copyright (c) 2002 by vtkTools      }
{                                          }
{******************************************}

program Example;

uses
  Forms,
  Main in 'Main.pas' {FormMain};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
