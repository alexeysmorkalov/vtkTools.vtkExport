//---------------------------------------------------------------------------

#include <vcl.h>
#include <typinfo.hpp>
#include <typeinfo.h>
#include <vteExcel.hpp>
#include <vteWriters.hpp>
#pragma hdrstop
#include "Main.h"
//---------------------------------------------------------------------------

#pragma package(smart_init)
#pragma resource "*.dfm"
TFormMain *FormMain;
//---------------------------------------------------------------------------
__fastcall TFormMain::TFormMain(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormMain::Image1Click(TObject *Sender)
{
  ShellExecute(0,"open","http://www.vtktools.ru","","",SW_SHOWNORMAL);
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::ButtonSaveClick(TObject *Sender)
{
AnsiString FileName;
TvteXLSWorkbook* wb;
TvteXLSWorkbook* wb_f;
TvteCustomWriter* Writer;

if (SaveDialog->Execute())
  {
    FileName = SaveDialog->FileName;
    if (ExtractFileExt(FileName).Length()==0)
      switch (SaveDialog->FilterIndex)
        {
           case 1: {FileName = FileName+".xls"; break;}
           case 2: {FileName = FileName+".htm"; break;}
        }
    wb = new TvteXLSWorkbook;
    if ((CFormulas->Checked) && (CFormulasBooks->Checked))
      wb_f = new TvteXLSWorkbook;
    else
      wb_f = NULL;
    try
      {
        Build(wb,wb_f,FileName.c_str());
        switch (SaveDialog->FilterIndex)
          {
            case 1: {Writer = (TvteCustomWriter*)new TvteExcelWriter; break;}
            case 2: {Writer = (TvteCustomWriter*)new TvteHTMLWriter; break;}
          }
        try
          {
            Writer->Save(wb,FileName);
            if ((CFormulas->Checked) && (CFormulasBooks->Checked))
              Writer->Save(wb_f,FileName+".xls");
            delete Writer;
          }
        catch (...)
          {
            delete Writer;
          }
      delete wb;
      if (wb_f!=NULL)
        delete wb_f;
      }
    catch (...)
      {
        delete wb;
        if (wb_f!=NULL)
          delete wb_f;
      }
    if (COpen->Checked)
      ShellExecute(0,"open",FileName.c_str(),"","",SW_SHOWNORMAL);
  }
}

void __fastcall TFormMain::Build(TvteXLSWorkbook* wb, TvteXLSWorkbook* wb_f, char* FileName)
{
  wb->Clear();
  if (CBorders->Checked) AddBordersSheet(wb);
  if (CFonts->Checked) AddFontsSheet(wb);
  if (CMerges->Checked) AddMergesSheet(wb);
  if (CBackgrounds->Checked) AddFillsSheet(wb);
  if (CRotations->Checked) AddRotationsSheet(wb);
  if (CLargeText->Checked) AddLargeTextSheet(wb);
  if (CFormats->Checked) AddNumberFormatsSheet(wb);
  if (CImages->Checked) AddPictures(wb);
  if (CFormulas->Checked) AddFormulas(wb,wb_f,FileName);
}

PTypeInfo tp(void)
{
  PTypeInfo ti = new TTypeInfo;
  ti->Name = "TvteXLSBorderType";
  ti->Kind = tkEnumeration;
  return ti;
}

void __fastcall TFormMain::AddBordersSheet(TvteXLSWorkbook* wb)
{
TvteXLSWorksheet* sh;
TvteXLSRange* r;
int i,j,Row,Col;
char s[255];
//AnsiString s;
char* aBordersCaptions[vtexlEdgeTop+1] = {"DiagonalDown","DiagonalUp","EdgeBottom","EdgeLeft","EdgeRight","EdgeTop"};
char* aLineStylesCaptions[vtelsSlantedDashDot+1] = {"None","Thin","Medium","Dashed","Dotted","Thick",
	"Double","Hair","MediumDashed","DashDot","MediumDashDot","DashDotDot","MediumDashDotDot",
	"SlantedDashDot"};

sh = wb->AddSheet();
sh->Title = "Borders";
for (j=(int)vtelsNone; j<=(int)vtelsSlantedDashDot; j++)
  for (i=(int)vtexlDiagonalDown; i<=(int)vtexlEdgeTop; i++)
    {
      Row = 1+j*2;
      Col = 1+((int)vtexlEdgeTop-i)*2;
      r = sh->Ranges[Col][Row][Col][Row];
      wsprintf(s,"BorderType : %s \nLineStyleType : %s",aBordersCaptions[i],aLineStylesCaptions[j]);
      r->Value = s;
      r->Borders->Borders[(TvteXLSBorderType)i]->LineStyle = (TvteXLSLineStyleType)j;
      r->Borders->Borders[(TvteXLSBorderType)i]->Color = (TColor)aDefaultColors[random(MaxDefaultColors-2)+1];
      r->WrapText = true;
    }
for (i=(int)vtexlDiagonalDown; i<=(int)vtexlEdgeTop; i++)
  sh->Cols[1+i*2]->Width = 30*256;
}

void __fastcall TFormMain::AddFontsSheet(TvteXLSWorkbook* wb)
{
#define MaxFontsVariants 100
#define MaxFontsNames 6
#define MaxColors 16
char* aFonts[MaxFontsNames] = {"Arial","Courier","Lucida console","Verdana","Tahoma","Times New Roman"};
TColor aColors[MaxColors] =
            {clWhite,clBlack,clSilver,clGray,
             clRed,clMaroon,clYellow,clOlive,
             clLime,clGreen,clAqua,clTeal,
             clBlue,clNavy,clFuchsia,clPurple};

TvteXLSWorksheet *sh;
TvteXLSRange *r;
int i;
char s1[255],s2[255];

sh = wb->AddSheet();
sh->Title = "Fonts";
for (i=0; i<MaxFontsVariants; i++)
  {
    r = sh->Ranges[1][i][1][i];
    r->Font->Name = aFonts[random(MaxFontsNames)];
    r->Font->Size = 6+random(40);
    s1[0] = 0;
    if (random(2)==1)
      {
        r->Font->Style = r->Font->Style << fsBold;
        strcat(s1,"Bold,");
      }
    if (random(2)==1)
      {
        r->Font->Style = r->Font->Style << fsItalic;
        strcat(s1,"Italic,");
      }
    if (random(2)==1)
      {
        r->Font->Style = r->Font->Style << fsUnderline;
        strcat(s1,"Underline,");
      }
    if (random(2)==1)
      {
        r->Font->Style = r->Font->Style << fsStrikeOut;
        strcat(s1,"StrikeOut,");
      }
    s1[strlen(s1)] = 0;
    r->Font->Color = aColors[random(MaxColors)];
    wsprintf(s2,"%s, %d, [%s]",r->Font->Name.c_str(),r->Font->Size,s1);
    r->Value = s2;
  }
}

void __fastcall TFormMain::AddMergesSheet(TvteXLSWorkbook* wb)
{
TvteXLSWorksheet *sh;
TvteXLSRange *r;
int i,x1,y1,x2,y2;
char s[255];

sh = wb->AddSheet();
sh->Title = "Merge";
for (i=0; i<=30; i++)
  sh->Cols[i]->Width = 256*15;
for (i=0; i<=10; i++)
  {
    x1 = random(10);
    x2 = x1+random(4);
    y1 = random(30);
    y2 = y1+random(5);

    r = sh->Ranges[x1][y1][x2][y2];
    r->VerticalAlignment = vtexlVAlignCenter;
    r->HorizontalAlignment = vtexlHAlignCenter;
    wsprintf(s,"%d,%d,%d,%d",x1,y1,x2,y2);
    r->Value = s;
    r->Borders->Borders[vtexlEdgeBottom]->LineStyle = vtelsThin;
    r->Borders->Borders[vtexlEdgeTop]->LineStyle = vtelsThin;
    r->Borders->Borders[vtexlEdgeLeft]->LineStyle = vtelsThin;
    r->Borders->Borders[vtexlEdgeRight]->LineStyle = vtelsThin;
  }
}

void __fastcall TFormMain::AddFillsSheet(TvteXLSWorkbook* wb)
{
TvteXLSWorksheet *sh;
TvteXLSRange *r;
int i,j;
char* aPatternCaptions[vtefpVertical+1] = {"None", "Automatic", "Checker", "CrissCross", "Down", "Gray8",
	"Gray16", "Gray25", "Gray50", "Gray75", "Grid", "Horizontal", "LightDown", "LightHorizontal",
	"LightUp", "LightVertical", "SemiGray75", "Solid", "Up", "Vertical"};

sh = wb->AddSheet();
sh->Title = "Fill";
for (j=vtefpNone; j<=vtefpVertical; j++)
  for (i=0; i<XLSMaxColorsInPalette; i++)
    {
      r = sh->Ranges[j][i][j][i];
      r->Value = aPatternCaptions[j];
      r->FillPattern = (TvteXLSFillPattern)j;
      r->ForegroundFillPatternColor = clWhite;
      r->BackgroundFillPatternColor = aDefaultColorPalette[i];
      r->VerticalAlignment = vtexlVAlignCenter;
      if (j==0)
        sh->Rows[i]->PixelHeight = 20;
    }
for (j=vtefpNone; j<=vtefpVertical; j++)
  sh->Cols[j]->Width = strlen(aPatternCaptions[j])*256;
}

void __fastcall TFormMain::AddRotationsSheet(TvteXLSWorkbook* wb)
{
TvteXLSWorksheet *sh;
TvteXLSRange *r;
int i;

sh = wb->AddSheet();
sh->Title = "Rotation";
for (i=0; i<=255; i++)
  {
    r = sh->Ranges[i%10][i/10][i%10][i/10];
    r->Value = i;
    r->Rotation = i;
    r->VerticalAlignment = vtexlVAlignCenter;
    r->HorizontalAlignment = vtexlHAlignCenter;
  }
}

void __fastcall TFormMain::AddLargeTextSheet(TvteXLSWorkbook* wb)
{
#define LargeText "Welcome to vtkExport Library\r\nVtkExport library intended for export from your programs in the Excel and HTML formats.\r\nThe shaping XLS of the file happens without use DDE, OLE, i.e. to receive XLS the file necessarily, that on the computer Excel was installed!.\r\nThe method of export is very simple - you form object TvteXLSWorkBook, which has properties and methods similar to properties and methods OLE of the Excel server and call  him method SaveAsXLS or SaveAsHTML.\r\n(c) VtkTools, 2002"



TvteXLSWorksheet *sh;
TvteXLSRange *r;

sh = wb->AddSheet();
sh->Title = "LargeText";
r = sh->Ranges[0][0][0][0];
r->Value = LargeText;
r->WrapText = true;
sh->Cols[0]->Width = 80*256;
}

void af(TvteXLSWorksheet* sh, double number, char* format, int* y)
{
TvteXLSRange *r;

r = sh->Ranges[0][*y][0][*y];
r->Value = number;
r = sh->Ranges[1][*y][1][*y];
r->Value = number;
r->Format = format;
r = sh->Ranges[2][*y][2][*y];
r->Value = format;
r->HorizontalAlignment = vtexlHAlignRight;
*y = *y+1;
}

void __fastcall TFormMain::AddNumberFormatsSheet(TvteXLSWorkbook* wb)
{
TvteXLSWorksheet *sh;
TvteXLSRange *r;
int y;

sh = wb->AddSheet();
sh->Title = "NumberFormats";

r = sh->Ranges[0][0][0][0];
r->Value = "Number";
r->FillPattern = vtefpSolid;
r->BackgroundFillPatternColor = clWhite;
r->ForegroundFillPatternColor = (TColor)0x00FF00;
sh->Cols[0]->Width = 20*256;

r = sh->Ranges[1][0][1][0];
r->Value = "Format";
r->FillPattern = vtefpSolid;
r->BackgroundFillPatternColor = clWhite;
r->ForegroundFillPatternColor = (TColor)0x00FF00;
sh->Cols[1]->Width = 20*256;

r = sh->Ranges[2][0][2][0];
r->Value = "Formatted number";
r->FillPattern = vtefpSolid;
r->BackgroundFillPatternColor = clWhite;
r->ForegroundFillPatternColor = (TColor)0x00FF00;
sh->Cols[2]->Width = 20*256;

y = 1;
af(sh,10123.456,"# ##0.00",&y);
af(sh,10123.456,"#,0.00",&y);
af(sh,10123.456,"0",&y);
}

void __fastcall TFormMain::AddPictures(TvteXLSWorkbook* wb)
{
TvteXLSWorksheet *sh;
TvteXLSRange *r;
TvteImage *img;
TPicture *Pic;
int i;

sh = wb->AddSheet();
sh->Title = "Pictures";

sh->Cols[0]->PixelWidth = ImageList->Width;
Pic = (TPicture*)new TPicture;

Pic->Bitmap->Width = ImageList->Width;
Pic->Bitmap->Height = ImageList->Height;

sh->Ranges[0][0][0][0]->Value = "Image stretched into cells:";
ImageList->GetBitmap(0,Pic->Bitmap);
sh->AddImage(4,0,7,2,Pic,true);

sh->Ranges[0][3][0][3]->Value = "Image with real size:";
ImageList->GetBitmap(1,Pic->Bitmap);
sh->AddImageScaled(4,0,3,0,100,100,Pic,true);

sh->Ranges[0][16][0][16]->Value = "Image scaled 50%:";
ImageList->GetBitmap(2,Pic->Bitmap);
sh->AddImageScaled(4,0,16,0,50,50,Pic,true);

sh->Ranges[0][23][0][23]->Value = "Scaled image with border:";
ImageList->GetBitmap(3,Pic->Bitmap);
img = sh->AddImageScaled(4,0,23,0,50,25,Pic,true);
img->BorderLineColor = clRed;
img->BorderLineStyle = vteblsDot;
img->BorderLineWeight = vteblwSingle;

sh->Ranges[0][28][0][28]->Value = "Image not aligned to cell:";
ImageList->GetBitmap(4,Pic->Bitmap);
sh->AddImageWithOffsets(4,512,28,128,8,256,33,128,Pic,true);

delete Pic;
}

void AddFormula(TvteXLSWorksheet* sh, int ARow, char* Description, char* AFormula)
{
TvteXLSRange *r;

sh->Rows[ARow]->PixelHeight = 60;

r = sh->Ranges[0][ARow][1][ARow];
r->Value = Description;
r->FillPattern = vtefpSolid;
r->ForegroundFillPatternColor = clGray;
r->Font->Style = r->Font->Style << fsBold;
r->VerticalAlignment = vtexlVAlignCenter;
r->WrapText = true;
r->Borders->SetAttributes(TvteXLSBorderTypes() << vtexlEdgeTop << vtexlEdgeBottom,clRed,vtelsThin,vtexlThin);

r = sh->Ranges[2][ARow][3][ARow];
r->Value = AFormula;
r->FillPattern = vtefpSolid;
r->ForegroundFillPatternColor = clWhite;
r->Font->Style = r->Font->Style << fsItalic;
r->VerticalAlignment = vtexlVAlignCenter;
r->Borders->SetAttributes(TvteXLSBorderTypes() << vtexlEdgeTop << vtexlEdgeBottom,clRed,vtelsThin,vtexlThin);

r = sh->Ranges[4][ARow][5][ARow];
r->Formula = AFormula;
r->FillPattern = vtefpSolid;
r->ForegroundFillPatternColor = clSilver;
r->Font->Style = r->Font->Style << fsBold;
r->VerticalAlignment = vtexlVAlignCenter;
r->Borders->SetAttributes(TvteXLSBorderTypes() << vtexlEdgeTop << vtexlEdgeBottom,clRed,vtelsThin,vtexlThin);
}

void __fastcall TFormMain::AddFormulas(TvteXLSWorkbook* wb, TvteXLSWorkbook* wb_f, char* FileName)
{
TvteXLSWorksheet *sh;
TvteXLSWorksheet *sh_f;
TvteXLSRange *r;
TvteImage *img;
TPicture *Pic;
int i,j;
char s[255];

sh = wb->AddSheet();
sh->Title = "Formulas";
sh->Cols[0]->PixelWidth = 100;
sh->Cols[1]->PixelWidth = 100;
sh->Cols[2]->PixelWidth = 100;
sh->Cols[3]->PixelWidth = 100;
sh->Cols[4]->PixelWidth = 100;
sh->Cols[5]->PixelWidth = 100;
r = sh->Ranges[1][1][1][1];
r->Value = "Example of using the formulas";
r->Font->Color = clNavy;
r->Font->Style = r->Font->Style << fsBold;
r->Font->Size = 18;

for (i=2; i<=5; i++)
  for (j=3; j<=5; j++)
    {
    r = sh->Ranges[i][j][i][j];
    r->FillPattern = vtefpSolid;
    r->ForegroundFillPatternColor = clSilver;
    r->Font->Style = r->Font->Style << fsBold;
    r->Value = i+j;
    }

r = sh->Ranges[0][3][0][5];
r->Value = "Data";
r->FillPattern = vtefpSolid;
r->ForegroundFillPatternColor = clGray;
r->Font->Style = r->Font->Style << fsBold;
r->VerticalAlignment = vtexlVAlignCenter;

AddFormula(sh,7,"Simple formula with operators and constants","((2+3)*2)-1.5");
AddFormula(sh,9,"Formula with operators, constants, names and functions","(Pi()*($C5-E4)*ROW())/F5");
AddFormula(sh,11,"The enclosed functions","If(Average(F3:F6)>6;Sum(E3:F6);0)");

if (CFormulasBooks->Checked)
  {
  sh_f = wb_f->AddSheet();
  sh_f->Title = "Formulas_Data";
  for (i=2; i<=5; i++)
    for (j=3; j<=5; j++)
      {
      r = sh_f->Ranges[i][j][i][j];
      r->FillPattern = vtefpSolid;
      r->ForegroundFillPatternColor = clSilver;
      r->Font->Style = r->Font->Style << fsBold;
      r->Value = i-j;
      }

  r = sh_f->Ranges[0][3][0][5];
  r->Value = "Data";
  r->FillPattern = vtefpSolid;
  r->ForegroundFillPatternColor = clGray;
  r->Font->Style = r->Font->Style << fsBold;
  r->VerticalAlignment = vtexlVAlignCenter;

  AddFormula(sh,13,"Link to another sheet","Sum(Formulas_Data!C4:F6)");
  }

if (CFormulasSheets->Checked)
  {
  sh_f = wb->AddSheet();
  sh_f->Title = "Formulas_Data";
  for (i=2; i<=5; i++)
    for (j=3; j<=5; j++)
      {
      r = sh_f->Ranges[i][j][i][j];
      r->FillPattern = vtefpSolid;
      r->ForegroundFillPatternColor = clSilver;
      r->Font->Style = r->Font->Style << fsBold;
      r->Value = i*j;
      }

  r = sh_f->Ranges[0][3][0][5];
  r->Value = "Data";
  r->FillPattern = vtefpSolid;
  r->ForegroundFillPatternColor = clGray;
  r->Font->Style = r->Font->Style << fsBold;
  r->VerticalAlignment = vtexlVAlignCenter;

  wsprintf(s,"Sum([%s.xls]Formulas_Data!C4:F6)",ExtractFileName(FileName));
  AddFormula(sh,15,"Link to another book",s);
  }
}
//---------------------------------------------------------------------------

void __fastcall TFormMain::CFormulasClick(TObject *Sender)
{
  CFormulasSheets->Enabled = CFormulas->Checked;
  if (CFormulas->Checked) CFormulasSheets->Checked = true;
  CFormulasBooks->Enabled = CFormulas->Checked;
  if (CFormulas->Checked) CFormulasBooks->Checked = true;
}
//---------------------------------------------------------------------------

