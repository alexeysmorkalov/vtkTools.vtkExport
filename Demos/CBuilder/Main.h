//---------------------------------------------------------------------------

#ifndef MainH
#define MainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Dialogs.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include <vteExcel.hpp>
#include <vteWriters.hpp>
#include <ImgList.hpp>
//---------------------------------------------------------------------------
class TFormMain : public TForm
{
__published:	// IDE-managed Components
        TPanel* Panel1;
        TPanel* Panel2;
        TPanel* Panel3;
        TLabel* Label1;
        TImage* Image1;
        TCheckBox* CBorders;
        TCheckBox* CFonts;
        TCheckBox* CMerges;
        TCheckBox* CBackgrounds;
        TCheckBox* CRotations;
        TCheckBox* CLargeText;
        TCheckBox* CFormats;
        TCheckBox* COpen;
        TButton* ButtonSave;
        TSaveDialog* SaveDialog;
        TImageList *ImageList;
        TCheckBox *CImages;
        TCheckBox *CFormulas;
        TCheckBox *CFormulasSheets;
        TCheckBox *CFormulasBooks;
        void __fastcall Image1Click(TObject *Sender);
        void __fastcall ButtonSaveClick(TObject *Sender);
        void __fastcall CFormulasClick(TObject *Sender);
private:	// User declarations
    void __fastcall Build(TvteXLSWorkbook* wb, TvteXLSWorkbook* wb_f, char* FileName);
    void __fastcall AddBordersSheet(TvteXLSWorkbook* wb);
    void __fastcall AddFontsSheet(TvteXLSWorkbook* wb);
    void __fastcall AddMergesSheet(TvteXLSWorkbook* wb);
    void __fastcall AddFillsSheet(TvteXLSWorkbook* wb);
    void __fastcall AddRotationsSheet(TvteXLSWorkbook* wb);
    void __fastcall AddLargeTextSheet(TvteXLSWorkbook* wb);
    void __fastcall AddNumberFormatsSheet(TvteXLSWorkbook* wb);
    void __fastcall AddPictures(TvteXLSWorkbook* wb);

    void __fastcall AddFormulas(TvteXLSWorkbook* wb, TvteXLSWorkbook* wb_f, char* FileName);
public:		// User declarations
        __fastcall TFormMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormMain *FormMain;
//---------------------------------------------------------------------------
#endif
