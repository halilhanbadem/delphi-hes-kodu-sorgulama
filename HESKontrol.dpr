program HESKontrol;

uses
  Vcl.Forms,
  uMain in 'uMain.pas' {fMain},
  libHESKontrol in 'libHESKontrol.pas',
  uBilgi in 'uBilgi.pas' {fBilgi};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfMain, fMain);
  Application.CreateForm(TfBilgi, fBilgi);
  Application.Run;
end.
