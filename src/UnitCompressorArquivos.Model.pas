unit UnitCompressorArquivos.Model;

interface

uses System.ZIP, Vcl.Taskbar, Vcl.Forms;

type
  iCompressorArquivos = interface
    ['{F863A10B-E4CC-4FBC-A0FB-F89299987EBC}']
    function ComprimeArquivo(ArquivoCompacto: string; ArquivosCompactar: array of string; Sobrescrever: Boolean = True): iCompressorArquivos;
    function DescomprimeArquivo(ArquivoZip, DiretorioDestino: string): iCompressorArquivos;
    function ExbirProgresso(Value: TForm): iCompressorArquivos;
  end;

  TCompressorArquivos = class(TInterfacedObject, iCompressorArquivos)
  private
    FTaskBar: TTaskbar;
    FTitulo: string;
    FForm: TForm;
    function ObterTamanhoArquivo(Value: string): integer;
    procedure OnProgress(Sender: TObject; FileName: string; Header: TZipHeader; Position: Int64);
  public
    constructor Create;
    destructor Destroy; override;
    class function New: iCompressorArquivos;
    function ComprimeArquivo(ArquivoCompacto: string; ArquivosCompactar: array of string; Sobrescrever: Boolean = True): iCompressorArquivos;
    function DescomprimeArquivo(ArquivoZip, DiretorioDestino: string): iCompressorArquivos;
    function ExbirProgresso(Value: TForm): iCompressorArquivos;
  end;

implementation

uses
  System.SysUtils, System.Classes, System.Win.TaskbarCore;

{ TCompressorArquivos }

constructor TCompressorArquivos.Create;
begin

end;

destructor TCompressorArquivos.Destroy;
begin
  if Assigned(FTaskBar) then
  begin
    FTaskBar.ProgressState := TTaskBarProgressState.None;
    FForm.Caption          := FTitulo;
    FreeAndNil(FTaskBar);
  end;
  inherited;
end;

function TCompressorArquivos.ExbirProgresso(Value: TForm): iCompressorArquivos;
begin
  Result                 := Self;
  FForm                  := Value;
  FTaskBar               := TTaskbar.Create(FForm);
  FTitulo                := FForm.Caption;
  FForm.Caption          := 'Compactando Arquivos';
  FTaskBar.ProgressState := TTaskBarProgressState.Indeterminate
end;

class function TCompressorArquivos.New: iCompressorArquivos;
begin
  Result := Self.Create;
end;

function TCompressorArquivos.ObterTamanhoArquivo(Value: string): integer;
var
  StreamArquivo: TFileStream;
begin
  StreamArquivo := TFileStream.Create(Value, fmOpenRead);
  try
    Result := StreamArquivo.Size;
  finally
    StreamArquivo.Free;
  end;
end;

procedure TCompressorArquivos.OnProgress(Sender: TObject; FileName: string; Header: TZipHeader; Position: Int64);
begin
  FTaskBar.ProgressValue := Position;
end;

function TCompressorArquivos.ComprimeArquivo(ArquivoCompacto: string; ArquivosCompactar: array of string; Sobrescrever: Boolean = True): iCompressorArquivos;
var
  ZIP: TZipFile;
  i: integer;
  ArquivoDeletado: Boolean;
begin
  Result := Self;
  ZIP    := TZipFile.Create;
  try
    // associa ao evento
    if Assigned(FTaskBar) then
      ZIP.OnProgress := OnProgress;
    if Sobrescrever then
    begin
      DeleteFile(ArquivoCompacto);
      Sleep(1000);
    end;
    // verifico a existencia do arquivo .zip
    if FileExists(ArquivoCompacto) then
    begin
      ZIP.Open(ArquivoCompacto, zmReadWrite);
    end
    else
    begin
      ZIP.Open(ArquivoCompacto, zmWrite);
    end;
    // adiciona arquivos para serem compactados
    for i := 0 to Length(ArquivosCompactar) - 1 do
    begin
      // exibir progresso na barra de tarefas
      if Assigned(FTaskBar) then
        FTaskBar.ProgressMaxValue := ObterTamanhoArquivo(ArquivosCompactar[i]);
      ZIP.Add(ArquivosCompactar[i]);
    end;
    ZIP.Close;
  finally
    FreeAndNil(ZIP);
  end;
end;

function TCompressorArquivos.DescomprimeArquivo(ArquivoZip, DiretorioDestino: string): iCompressorArquivos;
var
  ZIP: TZipFile;
begin
  Result := Self;
  try
    ZIP := TZipFile.Create;
    // associa ao evento
    if Assigned(FTaskBar) then
      ZIP.OnProgress := OnProgress;
    if FileExists(ArquivoZip) then
    begin
      try
        ZIP.Open(ArquivoZip, zmReadWrite);
      except
        on E: Exception do
          raise Exception.Create('Erro abrir aquivo .zip' + sLineBreak + E.Message);
      end;
    end
    else
      raise Exception.Create('Arquivo .zip não encontrado.' + sLineBreak + ArquivoZip);
    // exibir progresso na barra de tarefas
    if Assigned(FTaskBar) then
      FTaskBar.ProgressMaxValue := ObterTamanhoArquivo(ArquivoZip);
    ZIP.ExtractAll(DiretorioDestino);
    ZIP.Close;
  finally
    FreeAndNil(ZIP);
  end;
end;

end.
