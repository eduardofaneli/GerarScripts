unit uSobre;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  PngBitBtn, Vcl.ExtCtrls, Winapi.ShellAPI, Vcl.Imaging.pngimage;

type
  TfrmSobre = class(TForm)
    pnlBotoes: TPanel;
    PngBitBtn1: TPngBitBtn;
    pnlSobre: TPanel;
    lblTitulo: TLabel;
    imgSobre: TImage;
    lblGitHub: TLabel;
    lblVersao: TLabel;
    lblAutoria: TLabel;
    lblEmail: TLabel;
    procedure PngBitBtn1Click(Sender: TObject);
    procedure lblGitHubClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSobre: TfrmSobre;

implementation

{$R *.dfm}

function VersaoExe: String;
type
  PFFI = ^vs_FixedFileInfo;
var
  F : PFFI;
  Handle : Dword;
  Len : Longint;
  Data : Pchar;
  Buffer : Pointer;
  Tamanho : Dword;
  Parquivo: Pchar;
  Arquivo : String;
begin

  Arquivo := Application.ExeName;
  Parquivo := StrAlloc(Length(Arquivo) + 1);
  StrPcopy(Parquivo, Arquivo);
  Len := GetFileVersionInfoSize(Parquivo, Handle);
  Result := '';

  if Len > 0 then
  begin
    Data:=StrAlloc(Len+1);
    if GetFileVersionInfo(Parquivo,Handle,Len,Data) then
    begin
      VerQueryValue(Data, '\',Buffer,Tamanho);
      F := PFFI(Buffer);
      Result := Format('%d.%d.%d.%d',
      [HiWord(F^.dwFileVersionMs),
      LoWord(F^.dwFileVersionMs),
      HiWord(F^.dwFileVersionLs),
      Loword(F^.dwFileVersionLs)]);
    end;
    StrDispose(Data);
  end;

  StrDispose(Parquivo);
end;

procedure TfrmSobre.FormShow(Sender: TObject);
begin

  lblVersao.Caption := 'Versão ' + VersaoExe;

end;

procedure TfrmSobre.lblGitHubClick(Sender: TObject);
begin
  ShellExecute(Handle, 'open', pChar(lblGitHub.Caption), nil, nil, SW_SHOW );
end;

procedure TfrmSobre.PngBitBtn1Click(Sender: TObject);
begin
  Close;
end;

end.

