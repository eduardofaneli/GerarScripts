program GerarScript;

uses
  Forms,
  uGerarScript in 'uGerarScript.pas' {frmGerarSqlQuery},
  Themes,

  uSobre in 'uSobre.pas' {frmSobre};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmGerarSqlQuery, frmGerarSqlQuery);
  Application.Run;
end.

