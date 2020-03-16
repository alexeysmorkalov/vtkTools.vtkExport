//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USEFORM("Main.cpp", FormMain);
USEUNIT("..\..\Source\BIFF8_Types.pas");
USEUNIT("..\..\Source\vteConsts.pas");
USEUNIT("..\..\Source\vteExcelFormula_iftab.pas");
USEUNIT("..\..\Source\vteExcelFormula.pas");
USEUNIT("..\..\Source\vteExcel.pas");
USEUNIT("..\..\Source\vteExcelTypes.pas");
USEUNIT("..\..\Source\vteWriters.pas");
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->CreateForm(__classid(TFormMain), &FormMain);
                 Application->Run();
        }
        catch (Exception &exception)
        {
                 Application->ShowException(&exception);
        }
        return 0;
}
//---------------------------------------------------------------------------
