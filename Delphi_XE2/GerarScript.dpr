program GerarScript;

uses
  Vcl.Forms,
  uGerarScript in 'uGerarScript.pas' {frmGerarSqlQuery},
  Vcl.Themes,
  Vcl.Styles,
  uSobre in 'uSobre.pas' {frmSobre};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Metro Black');
  Application.CreateForm(TfrmGerarSqlQuery, frmGerarSqlQuery);
  Application.Run;
end.

