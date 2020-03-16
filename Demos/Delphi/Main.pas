
{******************************************}
{                                          }
{           vtk Export library             }
{            Example  program              }
{      Copyright (c) 2002 by vtkTools      }
{                                          }
{******************************************}

unit Main;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, ShellApi,TypInfo,
  BIFF8_Types, vteExcel, vteExcelTypes, vteWriters, vteExcelFormula, vteExcelFormula_iftab,
  ImgList, jpeg;

const
  LargeText =
'Welcome to vtkExport Library'#13#10+
'VtkExport library intended for export from your programs in the Excel and HTML formats.'#13#10+
'The shaping XLS of the file happens without use DDE, OLE, i.e. to receive XLS the file necessarily, that on the computer Excel was installed!.'#13#10+
'The method of export is very simple - you form object TvteXLSWorkBook, which has properties and methods similar to properties and methods OLE of the Excel server and call  him method SaveAsXLS or SaveAsHTML.'#13#10+
'(c) VtkTools, 2003';

type                 
  TFormMain = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    ButtonSave: TButton;
    COpen: TCheckBox;
    Image1: TImage;
    SaveDialog: TSaveDialog;
    Panel3: TPanel;
    CBorders: TCheckBox;
    CFonts: TCheckBox;
    CMerges: TCheckBox;
    CBackgrounds: TCheckBox;
    CRotations: TCheckBox;
    CLargeText: TCheckBox;
    CFormats: TCheckBox;
    CImages: TCheckBox;
    ImageList: TImageList;
    CFormulas: TCheckBox;
    CFormulasSheets: TCheckBox;
    CFormulasBooks: TCheckBox;
    procedure Image1Click(Sender: TObject);
    procedure ButtonSaveClick(Sender: TObject);
    procedure CFormulasClick(Sender: TObject);
  private
    { Private declarations }
    procedure Build(wb,wb_f : TvteXLSWorkbook; FileName: string);
    procedure AddBordersSheet(wb : TvteXLSWorkbook);
    procedure AddFontsSheet(wb : TvteXLSWorkbook);
    procedure AddMergesSheet(wb : TvteXLSWorkbook);
    procedure AddFillsSheet(wb : TvteXLSWorkbook);
    procedure AddRotationsSheet(wb : TvteXLSWorkbook);
    procedure AddLargeTextSheet(wb : TvteXLSWorkbook);
    procedure AddNumberFormatsSheet(wb : TvteXLSWorkbook);
    procedure AddPictures(wb : TvteXLSWorkbook);
    procedure AddFormulas(wb,wb_f : TvteXLSWorkbook; FileName: string);
  public
    { Public declarations }
  end;

var
  FormMain: TFormMain;

implementation

{$R *.DFM}
{$R windowsxp.RES}

procedure TFormMain.Image1Click(Sender: TObject);
begin
  ShellExecute(0,'open',PChar('http://www.vtktools.ru'),'','',SW_SHOWNORMAL);
end;

procedure TFormMain.ButtonSaveClick(Sender: TObject);
var
  wb, wb_f : TvteXLSWorkbook;
  FileName : string;
  Writer : TvteCustomWriter;
begin
  if SaveDialog.Execute then
  begin
    FileName := SaveDialog.FileName;
    if ExtractFileExt(SaveDialog.FileName) = '' then
      case SaveDialog.FilterIndex of
        1 : FileName := FileName + '.xls';
        2 : FileName := FileName + '.htm';
      end;
    wb := TvteXLSWorkbook.Create;
    if (CFormulas.Checked) and (CFormulasBooks.Checked) then
      wb_f := TvteXLSWorkbook.Create
    else
      wb_f := nil;
    try
      Build(wb,wb_f,FileName);
      case SaveDialog.FilterIndex of
        1 : Writer := TvteExcelWriter.Create;
        2 : Writer := TvteHTMLWriter.Create;
      else
        Raise Exception.Create('Logical error');
      end;
      try
        Writer.Save(wb,FileName);
        if (CFormulas.Checked) and (CFormulasBooks.Checked) then
          Writer.Save(wb_f,FileName+'.xls');
      finally
        Writer.Free;
      end;
    finally
      wb.Free;
      if wb_f<>nil then
        wb_f.Free;
    end;
    if COpen.Checked then
      ShellExecute(0,'open',PChar(FileName),'','',SW_SHOWNORMAL);
  end;
end;

procedure TFormMain.AddBordersSheet(wb : TvteXLSWorkbook);
var
  sh : TvteXLSWorksheet;
  i,j : integer;
  s1,s2 : string;
  Row,Col : integer;
begin
sh := wb.AddSheet;
sh.Title := 'Borders';
sh.PageSetup.PrintHeaders := true;
sh.PageSetup.Zoom := 120;
sh.PageSetup.FitToPagesTall := 3;
for j:=integer(Low(TvteXLSLineStyleType)) to integer(High(TvteXLSLineStyleType)) do
begin
  for i:=integer(Low(TvteXLSBorderType)) to integer(High(TvteXLSBorderType)) do
  begin
    Row := 1+j*2;
    Col := 3+(integer(High(TvteXLSBorderType))-i)*2;
    with sh.Ranges[Col,Row,Col,Row] do
      begin
        s1 := 'BorderType : '+Copy(GetEnumName(TypeInfo(TvteXLSBorderType),i),5,1000);
        s2 := 'LineStyleType : '+Copy(GetEnumName(TypeInfo(TvteXLSLineStyleType),j),5,1000);
        Value := s1+#10+s2;
        Borders[TvteXLSBorderType(i)].LineStyle := TvteXLSLineStyleType(j);
        Borders[TvteXLSBorderType(i)].Color := aDefaultColors[Random(MaxDefaultColors-2)+1];
        Borders[TvteXLSBorderType(i)].Weight := vtexlThick;
        WrapText := true;
      end;
  end;
  Col := 1;
  with sh.Ranges[Col,Row,Col,Row] do
    begin
      Value := 'LineStyleType : '+Copy(GetEnumName(TypeInfo(TvteXLSLineStyleType),j),5,1000);
      Borders.SetAttributes([vtexlEdgeBottom,vtexlEdgeLeft,vtexlEdgeRight,vtexlEdgeTop],aDefaultColors[Random(MaxDefaultColors-2)+1],TvteXLSLineStyleType(j),vtexlThin);
      WrapText := true;
    end;
end;
for i:=integer(Low(TvteXLSBorderType)) to integer(High(TvteXLSBorderType))+1 do
    sh.Cols[i*2+1].CentimeterWidht := 6
end;

procedure TFormMain.AddFontsSheet(wb : TvteXLSWorkbook);
const
  MaxFontsVariants = 100;
  MaxFontsNames = 6;
  MaxColors = 16;
  aFonts : array[0..MaxFontsNames-1] of string =
           ('Arial',
            'Courier',
            'Lucida console',
            'Verdana',
            'Tahoma',
            'Times New Roman');
  aColors : array[0..MaxColors-1] of TColor =
            (clWhite,clBlack,clSilver,clGray,
             clRed,clMaroon,clYellow,clOlive,
             clLime,clGreen,clAqua,clTeal,
             clBlue,clNavy,clFuchsia,clPurple);
var
  sh : TvteXLSWorksheet;
  i: integer;
  s1 : string;
begin
  sh := wb.AddSheet;
  sh.Title := 'Fonts';
  for i:=0 to MaxFontsVariants do
    with sh.Ranges[1,i,1,i] do
      begin
        Font.Name := aFonts[Random(MaxFontsNames)];
        Font.Size := 6+Random(40);
        s1 := '';
        if Random(2)=1 then
          begin
            Font.Style := Font.Style+[fsBold];
            s1 := s1+'Bold,';
          end;
        if Random(2)=1 then
          begin
            Font.Style := Font.Style+[fsItalic];
            s1 := s1+'Italic,';
          end;
        if Random(2)=1 then
          begin
            Font.Style := Font.Style+[fsUnderline];
            s1 := s1+'Underline,';
          end;
        if Random(2)=1 then
          begin
            Font.Style := Font.Style+[fsStrikeout];
            s1 := s1+'Strikeout,';
          end;
        Delete(s1,Length(s1),1);
        Font.Color := aColors[1+Random(MaxColors-1)];
        Value := Font.Name+', '+IntToStr(Font.Size)+', ['+s1+']';
      end;
end;

procedure TFormMain.AddMergesSheet(wb : TvteXLSWorkbook);
var
  sh : TvteXLSWorksheet;
  i,x1,y1,x2,y2 : integer;
begin
  sh := wb.AddSheet;
  sh.Title := 'Merge';
  for i := 0 to 30 do
    sh.Cols[i].Width := 256*15;
  for i:= 0 to 30 do
    begin
      x1 := Random(10);
      x2 := x1+Random(4);
      y1 := Random(30);
      y2 := y1+Random(5);
      with sh.Ranges[x1,y1,x2,y2] do
        begin
          VerticalAlignment := vtexlVAlignCenter;
          HorizontalAlignment := vtexlHAlignCenter;
          Value := SysUtils.Format('%d,%d,%d,%d',[x1,y1,x2,y2]);
          Borders[vtexlEdgeBottom].LineStyle := vtelsThin;
          Borders[vtexlEdgeTop].LineStyle := vtelsThin;
          Borders[vtexlEdgeLeft].LineStyle := vtelsThin;
          Borders[vtexlEdgeRight].LineStyle := vtelsThin;
        end;
    end;
end;

procedure TFormMain.AddFillsSheet(wb : TvteXLSWorkbook);
var
  sh : TvteXLSWorksheet;
  i,j : integer;
begin
  sh := wb.AddSheet;
  sh.Title := 'Fill';
  for j:=integer(Low(TvteXLSFillPattern)) to integer(High(TvteXLSFillPattern)) do
    begin
      for i:=0 to XLSMaxColorsInPalette-1 do
        begin
          with sh[j,i,j,i] do
            begin
              Value := GetEnumName(TypeInfo(TvteXLSFillPattern),j);
              FillPattern := TvteXLSFillPattern(j);
              ForegroundFillPatternColor := aDefaultColorPalette[i];
              BackgroundFillPatternColor := clWhite;
              VerticalAlignment := vtexlVAlignCenter;
            end;
          if j=0 then
            sh.Rows[i].PixelHeight := 20;
        end;
    end;
for j:=integer(Low(TvteXLSFillPattern)) to integer(High(TvteXLSFillPattern)) do
  sh.Cols[j].Width := Length(GetEnumName(TypeInfo(TvteXLSFillPattern),j))*256;
end;

procedure TFormMain.AddRotationsSheet(wb : TvteXLSWorkbook);
var
  sh : TvteXLSWorksheet;
  i : integer;
begin
  sh := wb.AddSheet;
  sh.Title := 'Rotation';
  for i:=0 to 255 do
    with sh[i mod 10,i div 10,i mod 10,i div 10] do
      begin
        Value := i;
        Rotation := i;
        VerticalAlignment := vtexlVAlignCenter;
        HorizontalAlignment := vtexlHAlignCenter;
      end;
end;

procedure TFormMain.AddLargeTextSheet(wb : TvteXLSWorkbook);
var
  sh : TvteXLSWorksheet;
begin
  sh := wb.AddSheet;
  sh.Title := 'LargeText';
  sh.Cols[0].ExcelWidth := 80;
  with sh.Ranges[0,0,2,0] do
  begin
    Value := Label1.Caption+#13#10+LargeText;
    WrapText := true;
  end;
end;

procedure TFormMain.AddNumberFormatsSheet;
var
  y : integer;
  sh : TvteXLSWorksheet;

  procedure af(number : double; const format_ : string);
  begin
    sh.Ranges[0,y,0,y].Value := number;
    with sh.Ranges[2,y,2,y] do
    begin
      Value := number;
      Format := format_;
    end;
    with sh.Ranges[1,y,1,y] do
    begin
      Value := format_;
      HorizontalAlignment := vtexlHAlignRight;
    end;
    Inc(y);
  end;

begin
  sh := wb.AddSheet;
  sh.Title := 'NumberFormats';

  with  sh.Ranges[0,0,0,0] do
  begin
    Value := 'Number';
    FillPattern := vtefpSolid;
    BackgroundFillPatternColor := clWhite;
    ForegroundFillPatternColor := $00FF00;
  end;
  sh.Cols[0].Width := 10*256;

  with sh.Ranges[1,0,1,0] do
  begin
    Value := 'Format';
    FillPattern := vtefpSolid;
    BackgroundFillPatternColor := clWhite;
    ForegroundFillPatternColor := $00FF00;
  end;
  sh.Cols[1].Width := 10*256;

  with  sh.Ranges[2,0,2,0] do
  begin
    Value := 'Formatted number';
    FillPattern := vtefpSolid;
    BackgroundFillPatternColor := clWhite;
    ForegroundFillPatternColor := $00FF00;
  end;
  sh.Cols[2].Width := 20*256;

  y := 1;
  af(10123.456,'# ##0.00');
  af(10123.456,'#,0.00');
  af(10123.456,'0');
end;

procedure TFormMain.AddPictures(wb : TvteXLSWorkbook);
var
  sh : TvteXLSWorksheet;
  Pic : TPicture;

  function AddImage(ImageIndex : integer) : TvteImage;
  begin
  ImageList.GetBitmap(ImageIndex,Pic.Bitmap);
  Result := sh.AddImage(0,0,0,0,Pic,true);
  end;

begin
  sh := wb.AddSheet;
  sh.Title := 'Pictures';
  Pic := TPicture.Create;
  Pic.Bitmap.Width := ImageList.Width;
  Pic.Bitmap.Height := ImageList.Height;

  sh.Ranges[0,0,0,0].Value := 'Image stretched into cells:';
  with AddImage(0) do
    begin
      Left := 4;
      Top := 0;
      Right := 7;
      Bottom := 2;
    end;

  sh.Ranges[0,3,0,3].Value := 'Image with real size:';
  with AddImage(1) do
    begin
      Left := 4;
      Top := 3;
      ScalePercentX := 100;
      ScalePercentY := 100;
    end;

  sh.Ranges[0,16,0,16].Value := 'Image scaled 50%:';
  with AddImage(2) do
    begin
      Left := 4;
      Top := 16;
      ScalePercentX := 50;
      ScalePercentY := 50;
    end;

  sh.Ranges[0,23,0,23].Value := 'Scaled image with border:';
  with AddImage(3) do
    begin
      Left := 4;
      Top := 23;
      ScalePercentX := 50;
      ScalePercentY := 25;
      BorderLineColor := clRed;
      BorderLineStyle := vteblsDot;
      BorderLineWeight := vteblwSingle;
    end;

  sh.Ranges[0,28,0,28].Value := 'Image not aligned to cell:';
  with AddImage(4) do
    begin
      Left := 4;
      LeftCO := 512; // offset measured in 1/1024 of col width
      Top := 28;
      TopCO := 128; // offset measured in 1/256 of row height
      Right := 8;
      RightCO := 256;
      Bottom := 33;
      BottomCO := 128;
    end;

  Pic.Free;
end;

procedure TFormMain.AddFormulas(wb,wb_f : TvteXLSWorkbook; FileName: string);
var
  sh, sh_f : TvteXLSWorksheet;
  i,j : integer;

  procedure AddFormula(ARow : integer; const Description, AFormula : string);
  begin
    with sh.Ranges[0,ARow,1,ARow] do
    begin
      sh.Rows[ARow].PixelHeight := 60;
      Value := Description;
      FillPattern := vtefpSolid;
      ForegroundFillPatternColor := clGray;
      Font.Style := [fsBold];
      VerticalAlignment := vtexlVAlignCenter;
      WrapText := True;
      Borders.SetAttributes([vtexlEdgeTop,vtexlEdgeBottom],clRed,vtelsThin,vtexlThin)
    end;

    with sh.Ranges[2,ARow,3,ARow] do
    begin
      Value := AFormula;
      FillPattern := vtefpSolid;
      ForegroundFillPatternColor := clWhite;
      Font.Style := [fsItalic];
      VerticalAlignment := vtexlVAlignCenter;
      Borders.SetAttributes([vtexlEdgeTop,vtexlEdgeBottom],clRed,vtelsThin,vtexlThin)
    end;

    with sh.Ranges[4,ARow,5,ARow] do
    begin
      Formula := AFormula;
      FillPattern := vtefpSolid;
      ForegroundFillPatternColor := clSilver;
      Font.Style := [fsBold];
      VerticalAlignment := vtexlVAlignCenter;
      Borders.SetAttributes([vtexlEdgeTop,vtexlEdgeBottom],clRed,vtelsThin,vtexlThin)
    end;
  end;

begin
  sh := wb.AddSheet;
  sh.Title := 'Formulas';
  sh.Cols[0].PixelWidth := 100;
  sh.Cols[1].PixelWidth := 100;
  sh.Cols[2].PixelWidth := 100;
  sh.Cols[3].PixelWidth := 100;
  sh.Cols[4].PixelWidth := 100;
  sh.Cols[5].PixelWidth := 100;
  with sh.Ranges[0,1,5,1] do
  begin
    Value := 'Example of using the formulas';
    HorizontalAlignment := vtexlHAlignCenter;
    Font.Color := clNavy;
    Font.Style := [fsBold];
    Font.Size := 18;
  end;

  for i := 2 to 5 do
    for j := 3 to 5 do
      with sh.Ranges[i,j,i,j] do
      begin
        FillPattern := vtefpSolid;
        ForegroundFillPatternColor := clSilver;
        Font.Style := [fsBold];
        Value := i+j;
      end;

  with sh.Ranges[0,3,0,5] do
  begin
    Value := 'Data';
    FillPattern := vtefpSolid;
    ForegroundFillPatternColor := clGray;
    Font.Style := [fsBold];
    VerticalAlignment := vtexlVAlignCenter;
  end;

  AddFormula(7, 'Simple formula with operators and constants','((2+3)*2)-1.5');
  AddFormula(9, 'Formula with operators, constants, names and functions','(Pi()*($C5-E4)*ROW())/F5');
  AddFormula(11,'The enclosed functions','If(Average(F3:F6)>6;Sum(E3:F6);0)');

  if CFormulasBooks.Checked then
  begin
    sh_f := wb_f.AddSheet;
    sh_f.Title := 'Formulas_Data';
    for i := 2 to 5 do
      for j := 3 to 5 do
        with sh_f.Ranges[i,j,i,j] do
        begin
          FillPattern := vtefpSolid;
          ForegroundFillPatternColor := clSilver;
          Font.Style := [fsBold];
          Value := i-j;
        end;

    with sh_f.Ranges[0,3,0,5] do
    begin
      Value := 'Data';
      FillPattern := vtefpSolid;
      ForegroundFillPatternColor := clGray;
      Font.Style := [fsBold];
      VerticalAlignment := vtexlVAlignCenter;
    end;
    AddFormula(13,'Link to another sheet','Sum(Formulas_Data!C4:F6)');
  end;

  if CFormulasSheets.Checked then
  begin
    sh_f := wb.AddSheet;
    sh_f.Title := 'Formulas_Data';
    for i := 2 to 5 do
      for j := 3 to 5 do
        with sh_f.Ranges[i,j,i,j] do
        begin
          FillPattern := vtefpSolid;
          ForegroundFillPatternColor := clSilver;
          Font.Style := [fsBold];
          Value := i*j;
        end;

    with sh_f.Ranges[0,3,0,5] do
    begin
      Value := 'Data';
      FillPattern := vtefpSolid;
      ForegroundFillPatternColor := clGray;
      Font.Style := [fsBold];
      VerticalAlignment := vtexlVAlignCenter;
    end;
    AddFormula(15,'Link to another book',Format('Sum([%s.xls]Formulas_Data!C4:F6)',[ExtractFileName(FileName)]));
  end;
{
  with sh_f.Ranges[10, 10, 10, 10] do
  begin
    Formula := 'Indirect("A1")';
  end;
}
end;

procedure TFormMain.Build(wb,wb_f : TvteXLSWorkbook; FileName: string);
begin
  wb.Clear;
  if CBorders.Checked then AddBordersSheet(wb);
  if CFonts.Checked then AddFontsSheet(wb);
  if CMerges.Checked then AddMergesSheet(wb);
  if CBackgrounds.Checked then AddFillsSheet(wb);
  if CRotations.Checked then AddRotationsSheet(wb);
  if CLargeText.Checked then AddLargeTextSheet(wb);
  if CFormats.Checked then AddNumberFormatsSheet(wb);
  if CImages.Checked then AddPictures(wb);
  if CFormulas.Checked then AddFormulas(wb,wb_f,FileName);
end;

procedure TFormMain.CFormulasClick(Sender: TObject);
begin
  CFormulasSheets.Enabled := CFormulas.Checked;
  if (CFormulas.Checked) then CFormulasSheets.Checked := True;
  CFormulasBooks.Enabled := CFormulas.Checked;
  if (CFormulas.Checked) then CFormulasBooks.Checked := True;
end;

end.
