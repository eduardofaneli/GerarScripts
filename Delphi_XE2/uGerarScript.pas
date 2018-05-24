unit uGerarScript;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.ExtCtrls, PngSpeedButton, Vcl.Menus, Vcl.ImgList, Registry, Winapi.ShellApi,
  uADStanIntf, uADStanOption, uADStanParam, uADStanError, uADDatSManager,
  uADPhysIntf, uADDAptIntf, uADStanAsync, uADDAptManager, Data.DB,
  uADCompDataSet, uADCompClient, NDQuery, Vcl.ActnList, System.StrUtils;

type
  TfrmGerarSqlQuery = class(TForm)
    pnlBotoes: TPanel;
    GroupBox1: TGroupBox;
    btnGerarSQL: TPngSpeedButton;
    btnLimpar: TPngSpeedButton;
    btnSair: TPngSpeedButton;
    btnAbrirArquivo: TPngSpeedButton;
    odSQL: TOpenDialog;
    memoSQL: TMemo;
    memoDelphi: TMemo;
    ppOpcoesMemo: TPopupMenu;
    imgPopUp: TImageList;
    Selecionartudo1: TMenuItem;
    Copiar1: TMenuItem;
    Recortar1: TMenuItem;
    Colar1: TMenuItem;
    split: TSplitter;
    gbOpcoes: TGroupBox;
    ckbTryFinally: TCheckBox;
    ckbInsertUpdate: TCheckBox;
    gbNomeQuery: TGroupBox;
    edtNomeQuery: TEdit;
    ckbNomeQuery: TCheckBox;
    ckbParamByName: TCheckBox;
    pnlPesquisa: TPanel;
    pnlExpansor: TPanel;
    btnExpansor: TPngSpeedButton;
    edtPesquisa: TEdit;
    lblPesquisa: TLabel;
    btnPesquisar: TPngSpeedButton;
    fdPesquisar: TFindDialog;
    rbDelphi: TRadioButton;
    rbSQL: TRadioButton;
    rgTipoSQL: TRadioGroup;
    btnSobre: TPngSpeedButton;
    procedure ckbNomeQueryClick(Sender: TObject);
    procedure btnGerarSQLClick(Sender: TObject);
    procedure btnLimparClick(Sender: TObject);
    procedure memoSQLKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnSairClick(Sender: TObject);
    procedure ckbTryFinallyClick(Sender: TObject);
    procedure btnAbrirArquivoClick(Sender: TObject);
    procedure memoSQLChange(Sender: TObject);
    procedure Selecionartudo1Click(Sender: TObject);
    procedure Copiar1Click(Sender: TObject);
    procedure memoDelphiMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Recortar1Click(Sender: TObject);
    procedure Colar1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ckbInsertUpdateClick(Sender: TObject);
    procedure memoSQLKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnExpansorClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fdPesquisarFind(Sender: TObject);
    procedure btnPesquisarClick(Sender: TObject);
    procedure btnSobreClick(Sender: TObject);
  private
    vNumLinhaSQL, vNumLinhaDelphi   : Integer;
    FSelPos                         : Integer;
    FMemoPesquisa                   : TMemo;
    procedure CriarRegistroWindows;
    procedure AbrirArquivoWindows;
    procedure WMDropFiles(var Msg: TMessage); message wm_DropFiles;
    procedure GerarSQL;
    function ConcatenarQuery(ASQL: String): String;
    function NumeroLinhaDelphi(ALinhaSQL: String): Integer;
    function GerarParamByName : String;
    function ConcatenarParamByName(AParametro: String): String;
    function Localizar(const StrOri, StrLoc: string; const PosInicial: Longint; DifMaieMin: Boolean = False;
                       ParaCima: Boolean = False; CoincidirPalavra: Boolean = False): Longint;
    procedure Pesquisar;
    { Private declarations }
  public
    class function getQueryTratada(): String;
    { Public declarations }
  end;

var
  frmGerarSqlQuery: TfrmGerarSqlQuery;

implementation

uses
  Clipbrd, uSobre;

{$R *.dfm}

class function TfrmGerarSqlQuery.getQueryTratada(): String;
begin

  frmGerarSqlQuery := TfrmGerarSqlQuery.Create(nil);
  try
    if frmGerarSqlQuery.ShowModal = mrOk then
      Result := frmGerarSqlQuery.memoDelphi.Lines.Text;
  finally
    FreeAndNil(frmGerarSqlQuery);
  end;

end;

function TfrmGerarSqlQuery.Localizar(const StrOri, StrLoc: string;
  const PosInicial: Integer; DifMaieMin, ParaCima,
  CoincidirPalavra: Boolean): Longint;
var
  I    : Longint;
  Achou: Boolean;

  procedure ConferePalavraInteira;
  begin

    if Achou and CoincidirPalavra then

    begin

      if ((IfThen(I = 0, '', Copy(StrOri, I - 1             , 1)) <> '') and (Copy(StrOri, I              - 1, 1)[1] in ['0'..'9','A'..'Z','a'..'z'])) or
         ((IfThen(I = 0, '', Copy(StrOri, I + Length(StrLoc), 1)) <> '') and (Copy(StrOri, I + Length(StrLoc), 1)[1] in ['0'..'9','A'..'Z','a'..'z'])) then

        Achou := False;

    end;

  end;
begin

  Result := -1;

  if ParaCima then // se for para cima ele faz o for (loop) diminuindo o valor.

  begin

    for I := PosInicial - Length(StrLoc) downto 0 do

    begin

      if DifMaieMin then // a var achou deve ser TRUE para sair do looping achando a string

        Achou := StrLoc = Copy(StrOri, I, Length(StrLoc))

      else

        Achou := AnsiUpperCase(StrLoc) = AnsiUpperCase(Copy(StrOri, I, Length(StrLoc)));

      ConferePalavraInteira;

      if Achou then

      begin

        Result := I - 1; // contém a POSICAO do bicho.

        if Result < 0 then Result := 0;

        Break;

      end;

    end;

  end
  else  // Normal, do cursor para baixo
    for I := PosInicial to (Length(StrOri) - Length(StrLoc) + 1) do
    begin

      if DifMaieMin then

        Achou := StrLoc = Copy(StrOri, I, Length(StrLoc))

      else
        Achou := AnsiUpperCase(StrLoc) = AnsiUpperCase(Copy(StrOri, I, Length(StrLoc)));

      ConferePalavraInteira;

      if Achou then
      begin

        Result := I - 1;

        if Result < 0 then Result := 0;

        Break;

      end;

    end;

end;

procedure TfrmGerarSqlQuery.btnLimparClick(Sender: TObject);
begin
  memoSQL.Clear;
  memoDelphi.Clear;
end;

procedure TfrmGerarSqlQuery.btnPesquisarClick(Sender: TObject);
begin
  Pesquisar;
end;

procedure TfrmGerarSqlQuery.btnSairClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmGerarSqlQuery.btnSobreClick(Sender: TObject);
begin
  frmSobre := TfrmSobre.Create(nil);
  try

    frmSobre.ShowModal;

  finally

    FreeAndNil(frmSobre);

  end;
end;

procedure TfrmGerarSqlQuery.ckbInsertUpdateClick(Sender: TObject);
begin
  GerarSQL;
end;

procedure TfrmGerarSqlQuery.ckbNomeQueryClick(Sender: TObject);
begin
  edtNomeQuery.Enabled  := ckbNomeQuery.Checked;
  ckbTryFinally.Enabled := ckbNomeQuery.Checked;

  edtNomeQuery.Clear;

  GerarSQL;

end;

procedure TfrmGerarSqlQuery.ckbTryFinallyClick(Sender: TObject);
begin
  GerarSQL;
end;

procedure TfrmGerarSqlQuery.Colar1Click(Sender: TObject);
begin
  if  ActiveControl is TMemo then
  begin
    TMemo(ActiveControl).PasteFromClipboard;
    GerarSQL;
  end;
end;

function TfrmGerarSqlQuery.ConcatenarParamByName(AParametro: String): String;
begin
  if (ckbNomeQuery.Checked) and (Trim(edtNomeQuery.Text) <> EmptyStr) then
    Result := ('    ' + Trim(edtNomeQuery.Text) + '.ParamByName(''' + AParametro + ''').Value := EmptyStr;')
  else
    Result := ('  ParamByName(''' + AParametro + ''').Value := EmptyStr;');
end;

function TfrmGerarSqlQuery.ConcatenarQuery(ASQL: String): String;
begin
  if (ckbNomeQuery.Checked) and (Trim(edtNomeQuery.Text) <> EmptyStr) then
  begin

    case rgTipoSQL.ItemIndex of
      0: Result := '    ' + Trim(edtNomeQuery.Text) + '.SQL.Add('' ' + StringReplace(ASQL, '''', '''''', [rfReplaceAll]) + ' '');';
      1: Result := '    ' + Trim(edtNomeQuery.Text) + '.SQLOriginal.Add('' ' + StringReplace(ASQL, '''', '''''', [rfReplaceAll]) + ' '');';
    end;

  end
  else
  begin

    case rgTipoSQL.ItemIndex of
      0: Result := '  SQL.Add('' ' + StringReplace(ASQL, '''', '''''', [rfReplaceAll]) + ' '');' ;
      1: Result := '  SQLOriginal.Add('' ' + StringReplace(ASQL, '''', '''''', [rfReplaceAll]) + ' '');' ;
    end;

  end;
end;

procedure TfrmGerarSqlQuery.Copiar1Click(Sender: TObject);
begin
  if ActiveControl is TMemo then
    TMemo(ActiveControl).CopyToClipboard
end;

procedure TfrmGerarSqlQuery.CriarRegistroWindows;
var
  VRegistro : TRegistry;
begin
  try
    VRegistro := TRegistry.Create;
    try
      VRegistro.RootKey := HKEY_CLASSES_ROOT;
      VRegistro.LazyWrite := False;

      {Define o nome interno e uma legenda para aparecer no Windows Explorer}
      VRegistro.OpenKey('\GerarScriptDelphi', True);
      VRegistro.WriteString('', 'Gerar Script - Arquivo SQL e Txt');


      VRegistro.CloseKey;
      VRegistro.OpenKey('GerarScriptDelphi\shell\open\command', True);
      VRegistro.WriteString('',ParamStr(0) + ' %1'); {NomeDoExe %1}
      VRegistro.CloseKey;

      {Define o ícone a ser usado no Windows Explorer}
      VRegistro.OpenKey('GerarScriptDelphi\DefaultIcon', True);
      VRegistro.WriteString('', ParamStr(0) + ',0');
      VRegistro.CloseKey;

      VRegistro.OpenKey('.txt', True);
      VRegistro.WriteString('', 'Arquivo de Texto');
      VRegistro.CloseKey;

      VRegistro.OpenKey('.SQL', True);
      VRegistro.WriteString('', 'Structured Query Language');
      VRegistro.CloseKey;

    finally
      VRegistro.CloseKey;
      FreeAndNil(VRegistro);
    end;

  except

  end;
end;

procedure TfrmGerarSqlQuery.fdPesquisarFind(Sender: TObject);
var
  P: Integer;
begin

  if rbDelphi.Checked then
    FMemoPesquisa := memoDelphi
  else
  if rbSQL.Checked then
    FMemoPesquisa := memoSQL;

  P:= Localizar( FMemoPesquisa.Text, fdPesquisar.FindText, FMemoPesquisa.SelStart + FMemoPesquisa.SelLength,
                 frMatchCase in fdPesquisar.Options, not (frDown in fdPesquisar.Options), frWholeWord in fdPesquisar.Options);

  if P > -1 then
  begin

    FMemoPesquisa.SelStart  := P;
    FMemoPesquisa.SelLength := Length(fdPesquisar.FindText);
    FMemoPesquisa.SetFocus;

  end
  else
    FMemoPesquisa.SelStart  := 0;

end;

procedure TfrmGerarSqlQuery.FormCreate(Sender: TObject);
begin
  CriarRegistroWindows;
  DragAcceptFiles(Handle, True);
end;

procedure TfrmGerarSqlQuery.FormDestroy(Sender: TObject);
begin
  DragAcceptFiles(Handle, False);
end;

procedure TfrmGerarSqlQuery.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

  if (Key = Vk_F3) and (pnlPesquisa.Height = pnlPesquisa.Constraints.MaxHeight) then
    if (not (ssCtrl in Shift)) then
     Pesquisar;


  if (Key = Ord('F')) and (ssCtrl in Shift) then
  begin
    btnExpansor.Click;

    Key := 0;
  end;

  if (Key = VK_F9) then
  begin
    GerarSQL;

    Key := 0;
  end;

  if (Key = VK_DELETE) and ((ssCtrl in Shift) and (ssShift in Shift)) then
  begin

    memoSQL.Clear;
    memoDelphi.Clear;

    Key := 0;

  end;

end;

procedure TfrmGerarSqlQuery.FormShow(Sender: TObject);
begin
  AbrirArquivoWindows;
  pnlPesquisa.Height := pnlPesquisa.Constraints.MinWidth;
end;

function TfrmGerarSqlQuery.GerarParamByName: String;
var
  Posicao     , PosEspaco    , PosParenteses, PosVirgula, MenorPosicao : Integer;
  Parametro   , Resto                                                  : String;
  ParamByName , ParamMemoSQL                                           : TStringList;
begin

    Result := EmptyStr;

    ParamByName  := TStringList.Create;
    ParamMemoSQL := TStringList.Create;

    try
      ParamMemoSQL.Clear;
      ParamMemoSQL.Text := memoSQL.Lines.Text;

      Posicao := Pos(':', ParamMemoSQL.Text);

      Resto   := stringreplace(copy(ParamMemoSQL.Text, Posicao + 1, Length(ParamMemoSQL.Text)), sLineBreak, ' ', [rfReplaceAll]) + ' ';

      if ParamMemoSQL.Text <> EmptyStr then
      begin
        while Posicao > 0 do
        begin

          PosEspaco     := pos(' ', Resto);
          PosParenteses := pos(')', Resto);
          PosVirgula    := pos(',', Resto);

          if (PosEspaco > PosParenteses) and (PosParenteses > 0) then
            MenorPosicao := PosParenteses
          else
            MenorPosicao := PosEspaco;

          if (PosVirgula < MenorPosicao) and (PosVirgula > 0) then
            MenorPosicao := PosVirgula;

          Parametro := UpperCase(copy(Resto, 0, MenorPosicao - 1));


          if ParamByName.IndexOf(ConcatenarParamByName(Parametro)) < 0 then
            ParamByName.Add(ConcatenarParamByName(Parametro));

          Posicao := Pos(':', Resto);
          Resto   := copy(Resto, Posicao + 1, length(resto + ' ' ));


        end;

        Result := ParamByName.Text;
      end;
    finally
      FreeAndNil(ParamByName);
      FreeAndNil(ParamMemoSQL);
    end;

end;

procedure TfrmGerarSqlQuery.GerarSQL;
var
  i : integer;
  vSQL: TStringList;
begin

  vSQL := TStringList.Create;
  try

    MemoDelphi.Clear;

    if (ckbNomeQuery.Checked) and (Trim(edtNomeQuery.Text) <> EmptyStr) and (memoSQL.Text <> EmptyStr) then
    begin

      if ckbTryFinally.Checked then
        vSQL.Add('  try ');

      vSQL.Add('');
      vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.Close;');

      case rgTipoSQL.ItemIndex of
        0: vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.SQL.Clear;');
        1: vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.SQLOriginal.Clear;');
      end;

    end
    else
    if memoSQL.Text <> EmptyStr then
    begin

      vSQL.Add('  Close;');

      case rgTipoSQL.ItemIndex of
        0: vSQL.Add('  SQL.Clear;');
        1: vSQL.Add('  SQLOriginal.Clear;');
      end;

    end;

    for I := 0 to memoSQL.Lines.Count -1 do
    begin
      vSQL.Add(ConcatenarQuery(memoSQL.Lines[i]));
    end;

    vSQL.Add(' ');

    if ckbParamByName.Checked then
    begin
      if GerarParamByName <> EmptyStr then
        vSQL.Add(GerarParamByName);
    end;

    if (ckbNomeQuery.Checked) and (Trim(edtNomeQuery.Text) <> EmptyStr) and (memoSQL.Text <> EmptyStr) then
    begin

      if ckbInsertUpdate.Checked then
        vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.ExecSQL;')
      else
        vSQL.Add('    ' + Trim(edtNomeQuery.Text) + '.Open;');

      vSQL.Add('');

      if ckbTryFinally.Checked then
      begin
        vSQL.Add('  finally ');
        vSQL.Add('');
        vSQL.Add('    FreeAndNil('+ Trim(edtNomeQuery.Text) +');');
        vSQL.Add('');
        vSQL.Add('  end; ');
      end;
    end
    else
    if memoSQL.Text <> EmptyStr then
    begin
      if ckbInsertUpdate.Checked then
        vSQL.Add('  ExecSQL;')
      else
        vSQL.Add('  Open;')
    end;

  finally
    memoDelphi.Text := vSQL.Text;
    FreeAndNil(vSQL);
  end;
end;

procedure TfrmGerarSqlQuery.memoDelphiMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Button = mbRight then
  begin

    Recortar1.Visible := not (Sender = memoDelphi);
    Colar1.Visible    := not (Sender = memoDelphi);
    TMemo(Sender).SetFocus;
  end;

end;

procedure TfrmGerarSqlQuery.memoSQLChange(Sender: TObject);
begin
//  memoDelphi.Lines[vNumLinhaDelphi] := ConcatenarQuery(memoSQL.Lines[vNumLinhaSQL]);
end;

procedure TfrmGerarSqlQuery.memoSQLKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key = Ord('A')) and (ssCtrl in Shift) then
  begin
    TMemo(Sender).SelectAll;

    Key := 0;
  end;

  vNumLinhaSQL     := memoSQL.CaretPos.Y;

  vNumLinhaDelphi  := NumeroLinhaDelphi(ConcatenarQuery(memoSQL.Lines[vNumLinhaSQL]));

end;

procedure TfrmGerarSqlQuery.memoSQLKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key = Ord('V')) and (ssCtrl in Shift) then
  begin
    GerarSQL;

    Key := 0;
  end;

  if (Key = Ord('X')) and (ssCtrl in Shift) then
  begin
    GerarSQL;

    Key := 0;
  end;

end;

function TfrmGerarSqlQuery.NumeroLinhaDelphi(ALinhaSQL: String): Integer;
var
  I: Integer;
begin
  for I := 0 to memoDelphi.Lines.Count -1 do
  begin
    if memoDelphi.Lines[I] = ALinhaSQL then
    begin
      Result := I;
      Break;
    end;
  end;

end;

procedure TfrmGerarSqlQuery.Pesquisar;
begin
  fdPesquisar.FindText := edtPesquisa.Text;
  fdPesquisar.Options  := [frDown,frShowHelp];
  fdPesquisarFind(Self);
end;

procedure TfrmGerarSqlQuery.Recortar1Click(Sender: TObject);
begin
  if TMemo(ActiveControl) is TMemo then
  begin
    TMemo(ActiveControl).CutToClipboard;
    GerarSQL;
  end;

end;

procedure TfrmGerarSqlQuery.Selecionartudo1Click(Sender: TObject);
begin
  if ActiveControl is TMemo then
    TMemo(ActiveControl).SelectAll;
end;

procedure TfrmGerarSqlQuery.WMDropFiles(var Msg: TMessage);
var
  BufferSize: word;
  Drop: HDROP;
  FileName: string;
  Pt: TPoint;
  RctMemo: TRect;
begin
  { Pega o manipulador (handle) da operação
    "arrastar e soltar" (drag-and-drop) }
  Drop := Msg.wParam;

  { Pega o retângulo do Memo }
  RctMemo := memoSQL.BoundsRect;

  if PtInRect(RctMemo, Pt) then
  begin
    { Obtém o comprimento necessário para o nome do arquivo,
      sem contar o caractere nulo do fim da string.
      O segundo parâmetro (zero) indica o primeiro arquivo da lista }
    BufferSize := DragQueryFile(Drop, 0, nil, 0);
    SetLength(FileName, BufferSize +1); { O +1 é p/ nulo do fim da string }
    if DragQueryFile(Drop, 0, PChar(FileName), BufferSize+1) = BufferSize then
    begin
      memoSQL.Lines.LoadFromFile(string(PChar(FileName)));
      GerarSQL;
    end;
  end;

  Msg.Result := 0;
end;

procedure TfrmGerarSqlQuery.AbrirArquivoWindows;
begin
  {Se o primeiro parâmetro for um nome de arquivo existente...}
  if FileExists(ParamStr(1)) then
  begin
    memoSQL.Lines.LoadFromFile(ParamStr(1));
    GerarSQL;
  end;
end;

procedure TfrmGerarSqlQuery.btnAbrirArquivoClick(Sender: TObject);
begin

  odSQL.InitialDir := GetCurrentDir;

  if odSQL.Execute(Handle) then
  begin
    memoSQL.Lines.LoadFromFile(odSQL.FileName);
    GerarSQL;
  end;

end;

procedure TfrmGerarSqlQuery.btnExpansorClick(Sender: TObject);
begin
  if pnlPesquisa.Height = pnlPesquisa.Constraints.MaxHeight then
    pnlPesquisa.Height := pnlPesquisa.Constraints.MinWidth
  else
    pnlPesquisa.Height := pnlPesquisa.Constraints.MaxHeight;

  edtPesquisa.Clear;

  if edtPesquisa.CanFocus then
    edtPesquisa.SetFocus;
end;

procedure TfrmGerarSqlQuery.btnGerarSQLClick(Sender: TObject);
begin
  GerarSQL;
end;

end.

