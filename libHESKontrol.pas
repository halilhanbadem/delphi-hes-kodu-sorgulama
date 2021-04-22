unit libHESKontrol;

interface
 uses
  SysUtils,
  StrUtils,
  Classes,
  Generics.Collections,
  Vcl.StdCtrls,
  Vcl.OleCtrls,
  SHDocVw,
  Vcl.Forms,
  MSHTML,
  ActiveX,
  Windows;


type
 THESKontrol = class
    constructor create(WebBrowser: TWebBrowser);
    destructor destroy(); override;
   private
    MWB: TWebBrowser;
   public
    function loginEDevlet(TC, Sifre: String): Boolean;
    function goHESModule: Boolean;
    function HESControl(HESKodu: String): String;
 end;

resourcestring
  EDevletURL = 'https://giris.turkiye.gov.tr/Giris/gir';
  EDevletLogin = 'https://www.turkiye.gov.tr/';
  HESModuleURL = 'https://www.turkiye.gov.tr/saglik-bakanligi-hes-kodu-sorgulama';


implementation

{ THESKontrol }

constructor THESKontrol.create(WebBrowser: TWebBrowser);
begin
 MWB := WebBrowser;
end;

destructor THESKontrol.destroy;
begin
 inherited;
 MWB.Free;
end;

procedure Delay(dwMilliseconds: Longint);
var
  iStart, iStop: DWORD;
begin
  iStart := GetTickCount;
  repeat
    iStop := GetTickCount;
    Application.ProcessMessages;
    Sleep(1);
  until (iStop - iStart) >= dwMilliseconds;
end;

function GetWebBrowserHTML(const WebBrowser: TWebBrowser): String;
var
  LStream: TStringStream;
  Stream : IStream;
  LPersistStreamInit : IPersistStreamInit;
begin
  if not Assigned(WebBrowser.Document) then exit;
  LStream := TStringStream.Create('');
  try
    LPersistStreamInit := WebBrowser.Document as IPersistStreamInit;
    Stream := TStreamAdapter.Create(LStream,soReference);
    LPersistStreamInit.Save(Stream,true);
    result := LStream.DataString;
  finally
    LStream.Free();
  end;
end;

function THESKontrol.goHESModule: Boolean;
begin
 Result := False;
 MWB.Navigate(HESModuleURL);

 while MWB.ReadyState <> READYSTATE_COMPLETE do
 begin
  Application.ProcessMessages;
 end;

 Result := True;
end;

function THESKontrol.HESControl(HESKodu: String): String;
var
 sourceHTML: string;
begin
 MWB.OleObject.Document.getElementById('hes_kodu').value := HESKodu;
 MWB.OleObject.Document.forms[1].submit();

 Delay(4000); 

 while MWB.LocationURL = HESModuleURL do
 begin
   Application.ProcessMessages;
 end;

 while MWB.ReadyState <> READYSTATE_COMPLETE do
 begin
   Application.ProcessMessages;
 end;


 while MWB.LocationURL <> HESModuleURL + '?sonuc=Goster' do
 begin
   Application.ProcessMessages;
 end;

 while ContainsText(sourceHTML, 'işleminiz devam ediyor, lütfen bekleyiniz') = true do
 begin
   Application.ProcessMessages;
 end;

 sourceHTML := GetWebBrowserHTML(MWB);

 if ContainsText(sourceHTML, 'Girilen HES Kodu') then
 begin
  Result := 'Girilen HES Kodu geçersizdir.';
 end else
 begin
  Result := MWB.OleObject.document.getElementsByTagName('dd').item(2).innerText;
 end;
end;

function THESKontrol.loginEDevlet(TC, Sifre: String): Boolean;
begin
 MWB.Navigate(EDevletURL);

 while MWB.ReadyState <> READYSTATE_COMPLETE do
 begin
  Application.ProcessMessages;
 end;

 MWB.OleObject.Document.getElementById('tridField').value := TC;
 MWB.OleObject.Document.getElementById('egpField').value := Sifre;
 MWB.OleObject.Document.getElementsByClassName('submitButton').item(0).click();

 while MWB.ReadyState <> READYSTATE_COMPLETE do
 begin
  Application.ProcessMessages;
 end;

 if MWB.LocationURL = EDevletURL then
 begin
  raise Exception.Create('Kullanıcı adı veya şifre yanlış');
  Result := False;
  abort;
 end else
 begin
  Result := True;
 end;
end;

end.
