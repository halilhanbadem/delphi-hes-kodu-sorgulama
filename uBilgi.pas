unit uBilgi;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit,
   Vcl.Menus, Vcl.StdCtrls, cxButtons, cxTextEdit, cxLabel, dxSkinsCore,
  dxSkinOffice2019Colorful;

type
  TfBilgi = class(TForm)
    cxLabel1: TcxLabel;
    txtTCKimlik: TcxTextEdit;
    cxButton1: TcxButton;
    txtParola: TcxTextEdit;
    cxLabel2: TcxLabel;
    procedure cxButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fBilgi: TfBilgi;

implementation

{$R *.dfm}
     uses uMain;
procedure TfBilgi.cxButton1Click(Sender: TObject);
begin
 ShowMessage('Sisteme istek gönderildi. Ekranda giriş yapıldığından emin olunuz');
 Self.Close;
 fMain.libHESKontrol.loginEDevlet(txtTCKimlik.Text, txtParola.Text);
end;

end.
