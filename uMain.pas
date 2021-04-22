unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.OleCtrls, libHESKontrol,
  SHDocVw, ComObj, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit,
  cxNavigator, dxDateRanges, dxScrollbarAnnotations, Data.DB, cxDBData,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  dxRibbonSkins, dxRibbonCustomizationForm, dxBar, cxClasses, dxRibbon,
  dxSkinsForm, FireDAC.Comp.DataSet, FireDAC.Comp.Client, cxGridLevel,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, cxGridExportLink, dxSkinsCore, dxSkinOffice2019Colorful;

type
  TfMain = class(TForm)
    gridOgrencilerDBTableView1: TcxGridDBTableView;
    gridOgrencilerLevel1: TcxGridLevel;
    gridOgrenciler: TcxGrid;
    excelTable: TFDMemTable;
    dsMemTable: TDataSource;
    excelTableOGRENCI_NO: TStringField;
    excelTableAD: TStringField;
    excelTableSOYAD: TStringField;
    excelTableRANDEVU_TARIH: TStringField;
    excelTableRANDEVU_SLOT: TStringField;
    excelTableYAKINLIK_DERECESI: TStringField;
    excelTableTC_KIMLIK_NO: TStringField;
    excelTableHES_KODU: TStringField;
    excelTableSONUC: TStringField;
    dxSkinController1: TdxSkinController;
    dxBarManager1: TdxBarManager;
    dxRibbon1Tab1: TdxRibbonTab;
    dxRibbon1: TdxRibbon;
    dxBarManager1Bar1: TdxBar;
    dxBarLargeButton1: TdxBarLargeButton;
    dxBarLargeButton2: TdxBarLargeButton;
    dxBarManager1Bar2: TdxBar;
    dxBarLargeButton3: TdxBarLargeButton;
    dxBarLargeButton4: TdxBarLargeButton;
    dxBarLargeButton5: TdxBarLargeButton;
    dxBarLargeButton6: TdxBarLargeButton;
    dialogExcel: TOpenDialog;
    gridOgrencilerDBTableView1OGRENCI_NO: TcxGridDBColumn;
    gridOgrencilerDBTableView1AD: TcxGridDBColumn;
    gridOgrencilerDBTableView1SOYAD: TcxGridDBColumn;
    gridOgrencilerDBTableView1RANDEVU_TARIH: TcxGridDBColumn;
    gridOgrencilerDBTableView1RANDEVU_SLOT: TcxGridDBColumn;
    gridOgrencilerDBTableView1YAKINLIK_DERECESI: TcxGridDBColumn;
    gridOgrencilerDBTableView1TC_KIMLIK_NO: TcxGridDBColumn;
    gridOgrencilerDBTableView1HES_KODU: TcxGridDBColumn;
    gridOgrencilerDBTableView1SONUC: TcxGridDBColumn;
    dxBarManager1Bar3: TdxBar;
    dxBarLargeButton7: TdxBarLargeButton;
    dxBarLargeButton8: TdxBarLargeButton;
    dialogSave: TSaveDialog;
    WebBrowser1: TWebBrowser;
    dxRibbon1Tab2: TdxRibbonTab;
    dxBarManager1Bar4: TdxBar;
    dxBarLargeButton9: TdxBarLargeButton;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure dxBarLargeButton3Click(Sender: TObject);
    procedure dxBarLargeButton1Click(Sender: TObject);
    procedure dxBarLargeButton2Click(Sender: TObject);
    procedure dxBarLargeButton4Click(Sender: TObject);
    procedure dxBarLargeButton5Click(Sender: TObject);
    procedure dxBarLargeButton6Click(Sender: TObject);
    procedure dxBarLargeButton7Click(Sender: TObject);
    procedure dxBarLargeButton9Click(Sender: TObject);
    procedure dxBarLargeButton8Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    libHESKontrol: THESKontrol;
    Durdur: Boolean;
  end;

var
  fMain: TfMain;

implementation

{$R *.dfm}

uses uBilgi;

procedure TfMain.Button1Click(Sender: TObject);
begin
 libHESKontrol.loginEDevlet('tc_kimlik', 'edevlet-sifre');
end;

procedure TfMain.Button2Click(Sender: TObject);
begin
  libHESKontrol.goHESModule;
end;

procedure TfMain.Button3Click(Sender: TObject);
begin
 ShowMessage(libHESKontrol.HESControl('2222'));
end;

procedure TfMain.dxBarLargeButton1Click(Sender: TObject);
begin
 fBilgi.ShowModal;
end;

procedure TfMain.dxBarLargeButton2Click(Sender: TObject);
begin
 if libHESKontrol.goHESModule then
 begin
   MessageBox(handle, 'HES modülüne BAŞARIYLA gidildi.', 'BAŞARILI', MB_OK + MB_ICONINFORMATION);
 end else
 begin
   MessageBox(handle, 'Bir sorun oluştu.', 'Başarısız', MB_OK + MB_ICONERROR);
 end;
end;

procedure TfMain.dxBarLargeButton3Click(Sender: TObject);
var
 Excel: variant;
  I: Integer;
  I2: Integer;
begin
 if dialogExcel.Execute then
 begin
   if dialogExcel.FileName <> '' then
   begin
     Excel := CreateOleObject('Excel.Application');
     Excel.Workbooks.Open(dialogExcel.FileName);

     for I := 2 to Excel.WorkBooks[1].Sheets[1].UsedRange.Rows.Count - 1 do
     begin
       excelTable.Edit;
       excelTableOGRENCI_NO.Value := Excel.ActiveSheet.Cells[I, 2].Value;
       excelTableAD.Value := Excel.ActiveSheet.Cells[I, 3].Value;
       excelTableSOYAD.Value := Excel.ActiveSheet.Cells[I, 4].Value;
       excelTableRANDEVU_TARIH.Value := Excel.ActiveSheet.Cells[I, 5].Value;
       excelTableRANDEVU_SLOT.Value := Excel.ActiveSheet.Cells[I, 6].Value;
       excelTableYAKINLIK_DERECESI.Value := Excel.ActiveSheet.Cells[I, 7].Value;
       excelTableTC_KIMLIK_NO.Value := Excel.ActiveSheet.Cells[I, 8].Value;
       excelTableHES_KODU.Value := Excel.ActiveSheet.Cells[I, 9].Value;
       excelTable.Append;
     end;
   end;
 end;

end;

procedure TfMain.dxBarLargeButton4Click(Sender: TObject);
begin
 dialogExcel.DefaultExt := '*.xlsx';
 dialogExcel.Filter := 'Excel Dosyalar? (*.xlsx)|*.xlsx';
 if dialogExcel.Execute then
 begin
  if dialogExcel.FileName <> '' then
  begin
    ExportGridToXLSX(dialogExcel.FileName, gridOgrenciler);
  end;
 end;
end;

procedure TfMain.dxBarLargeButton5Click(Sender: TObject);
begin
 dialogExcel.DefaultExt := '*.csv';
 dialogExcel.Filter := 'CSV Dosyalar? (*.csv)|*.csv';

  if dialogExcel.Execute then
 begin
  if dialogExcel.FileName <> '' then
  begin
    ExportGridToCSV(dialogExcel.FileName, gridOgrenciler);
  end;
 end;
end;

procedure TfMain.dxBarLargeButton6Click(Sender: TObject);
begin
 dialogExcel.DefaultExt := '*.xml';
 dialogExcel.Filter := 'XML Dosyalar? (*.xml)|*.xml';
 if dialogExcel.Execute then
 begin
  if dialogExcel.FileName <> '' then
  begin
    ExportGridToXML(dialogExcel.FileName, gridOgrenciler);
  end;
 end;
end;

procedure TfMain.dxBarLargeButton7Click(Sender: TObject);
begin
 Durdur := False;
 excelTable.First;
 while not excelTable.Eof do
 begin
   Application.ProcessMessages;

   if Durdur = True then
   begin
     MessageBox(handle, 'Döngü durduruldu.', 'Dur', MB_OK + MB_ICONINFORMATION);
     exit;
   end;

   if libHESKontrol.goHESModule then
   begin
     excelTable.Edit;
     excelTableSONUC.Value := libHESKontrol.HESControl(excelTableHES_KODU.Value);
     excelTable.Post;
   end;
   excelTable.Next;
 end;

 ShowMessage('Tüm kayıtlar bitti!');
end;

procedure TfMain.dxBarLargeButton8Click(Sender: TObject);
begin
 Durdur := True;
end;

procedure TfMain.dxBarLargeButton9Click(Sender: TObject);
begin
 MessageBox(handle, 'Geliştirici: Halil Han Badem.' + sLineBreak
 + sLineBreak + 'Topluluğu ziyaret edebilirsiniz: https://yazilimtoplulugu.com.', 'Hakkkımda', MB_OK + MB_ICONINFORMATION);
end;

procedure TfMain.FormCreate(Sender: TObject);
begin
 libHESKontrol := THESKontrol.create(WebBrowser1);
end;

end.
