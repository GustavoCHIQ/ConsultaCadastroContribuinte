unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ACBrBase,
  ACBrDFe, ACBrNFe, Vcl.Buttons, Vcl.StdCtrls, ACBrUtil, ACBrDFESSL,
  System.TypInfo, System.IniFiles, System.ImageList, Vcl.ImgList, Blcksock,
  Xml.xmldom, Xml.XMLIntf, Xml.XMLDoc, ACBrDFeConfiguracoes, Winapi.ShellAPI,
  Vcl.ExtCtrls, Vcl.Menus, Vcl.WinXCtrls;

type
  TForm1 = class(TForm)
    ACBrNFe1: TACBrNFe;
    cbbSSLLib: TComboBox;
    lblSSLLib: TLabel;
    lblCryptLib: TLabel;
    cbbCryptLib: TComboBox;
    cbbHttpLib: TComboBox;
    lblHttpLib: TLabel;
    lblXmlSign: TLabel;
    cbbXmlSignLib: TComboBox;
    btnGetCert: TSpeedButton;
    edtNumSerie: TEdit;
    GroupBox1: TGroupBox;
    lblInscricaoEstadual: TLabel;
    btnConsultar: TButton;
    ImageList1: TImageList;
    btnStatus: TButton;
    cbbSSLType: TComboBox;
    lblSSLLib1: TLabel;
    edtInscricao: TEdit;
    btnSalvarConfig: TButton;
    XMLDocument1: TXMLDocument;
    edtCNPJ: TEdit;
    edtUF: TEdit;
    lblCNPJ: TLabel;
    lblUF: TLabel;
    edtRazao: TEdit;
    lblRazao: TLabel;
    edtFantasia: TEdit;
    lblFantasia: TLabel;
    edtRegime: TEdit;
    lblRegime: TLabel;
    edtRua: TEdit;
    edtBairro: TEdit;
    edtCodMunicipio: TEdit;
    edtMunicipio: TEdit;
    edtCEP: TEdit;
    lblRua: TLabel;
    lblCodMunicipio: TLabel;
    lblBairro: TLabel;
    lblCidade: TLabel;
    lblCEP: TLabel;
    lblResposta: TLabel;
    shpStatus: TShape;
    procedure btnSalvarConfigClick(Sender: TObject);
    procedure btnGetCertClick(Sender: TObject);
    procedure cbbSSLLibChange(Sender: TObject);
    procedure cbbCryptLibChange(Sender: TObject);
    procedure cbbHttpLibChange(Sender: TObject);
    procedure cbbXmlSignLibChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnConsultarClick(Sender: TObject);
    procedure btnStatusClick(Sender: TObject);
    procedure cbbSSLTypeChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    procedure AtualizarSSLLibsCombo;
    procedure gravarConfiguracao;
    procedure lerConfiguracao;
    procedure ConfigurarComponente;
    procedure ApagaDir(hHandle: THandle; const sPath: string; Confirm: boolean);
    function ListFiles(const Dir, Wildcard: string; const List: TStrings): Boolean;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

{ TForm1 }

procedure TForm1.btnConsultarClick(Sender: TObject);
var
  UF, Documento: string;
  NodeCount: IXMLNode;
  i: integer;
  List: TStringList;
  V, N: Integer;
  MostRecentFile: string;
  MostRecentDate: TDateTime;
begin
  UF := '';
  if not (InputQuery('WebServices Consulta Cadastro ', 'UF do Documento a ser Consultado:', UF)) then
    exit;

  Documento := '';
  if not (InputQuery('WebServices Consulta Cadastro ', 'Documento(CPF/CNPJ)', Documento)) then
    exit;

  Documento := Trim(OnlyNumber(Documento));

  ACBrNFe1.WebServices.ConsultaCadastro.UF := UF;

  if Length(Documento) > 11 then
    ACBrNFe1.WebServices.ConsultaCadastro.CNPJ := Documento
  else
    ACBrNFe1.WebServices.ConsultaCadastro.CPF := Documento;

  ACBrNFe1.WebServices.ConsultaCadastro.Executar;

  List := TStringList.Create;
  if ListFiles(ExtractFilePath(ParamStr(0)) + 'Docs', '*-cad.xml', List) then
  begin
    N := List.Count - 1;
    for V := 0 to N do
    begin
      if FileDateToDateTime(FileAge(List.Strings[V])) > MostRecentDate then
      begin
        MostRecentDate := FileDateToDateTime(FileAge(List.Strings[V]));
        MostRecentFile := List.Strings[V];
      end;
    end;
    XMLDocument1.FileName := MostRecentFile;
  end;

  XMLDocument1.Active := True;
  NodeCount := XMLDocument1.ChildNodes.FindNode('infCad');
  begin
    for i := 0 to XMLDocument1.DocumentElement.ChildNodes.Count - 1 do
      with XMLDocument1.DocumentElement.ChildNodes[i] do
      begin
        edtInscricao.Text := ChildNodes['infCad'].ChildNodes['IE'].text;
        edtCNPJ.Text := ChildNodes['infCad'].ChildNodes['CNPJ'].Text;
        edtUF.Text := ChildNodes['infCad'].ChildNodes['UF'].Text;
        edtRazao.Text := ChildNodes['infCad'].ChildNodes['xNome'].Text;
        edtFantasia.Text := ChildNodes['infCad'].ChildNodes['xFant'].Text;
        edtRegime.Text := ChildNodes['infCad'].ChildNodes['xRegApur'].Text;

        edtRua.Text := ChildNodes['infCad'].ChildNodes['ender'].ChildNodes['xLgr'].Text;
        edtBairro.Text := ChildNodes['infCad'].ChildNodes['ender'].ChildNodes['xBairro'].Text;
        edtCodMunicipio.Text := ChildNodes['infCad'].ChildNodes['ender'].ChildNodes['cMun'].Text;
        edtMunicipio.Text := ChildNodes['infCad'].ChildNodes['ender'].ChildNodes['xMun'].Text;
        edtCEP.Text := ChildNodes['infCad'].ChildNodes['ender'].ChildNodes['CEP'].Text;
      end;
  end;
end;

procedure TForm1.btnGetCertClick(Sender: TObject);
begin
  edtNumSerie.Text := ACBrNFe1.SSL.SelecionarCertificado;
end;

procedure TForm1.btnSalvarConfigClick(Sender: TObject);
begin
  gravarConfiguracao;
end;

procedure TForm1.btnStatusClick(Sender: TObject);
var
  status: Integer;
begin
  ACBrNFe1.Configuracoes.WebServices.Visualizar := false;
  ACBrNFe1.WebServices.StatusServico.Executar;
  shpStatus.Visible := true;
  lblResposta.Visible := True;
  status := ACBrNFe1.WebServices.StatusServico.cStat;

  if status = 107 then
  begin
    shpStatus.Brush.Color := clGreen;
    lblResposta.Caption := 'Serviço em Operação';
  end
  else
  begin
    shpStatus.Brush.Color := clRed;
    lblResposta.Caption := 'Serviço nao esta em Operação';
  end;

end;

procedure TForm1.cbbCryptLibChange(Sender: TObject);
begin
  try
    if cbbCryptLib.ItemIndex <> -1 then
      ACBrNFe1.Configuracoes.Geral.SSLCryptLib := TSSLCryptLib(cbbCryptLib.ItemIndex);
  finally
    AtualizarSSLLibsCombo;
  end;
end;

procedure TForm1.cbbHttpLibChange(Sender: TObject);
begin
  try
    if cbbHttpLib.ItemIndex <> -1 then
      ACBrNFe1.Configuracoes.Geral.SSLHttpLib := TSSLHttpLib(cbbHttpLib.ItemIndex);
  finally
    AtualizarSSLLibsCombo;
  end;
end;

procedure TForm1.cbbSSLLibChange(Sender: TObject);
begin
  try
    if cbbSSLLib.ItemIndex <> -1 then
      ACBrNFe1.Configuracoes.Geral.SSLLib := TSSLLib(cbbSSLLib.ItemIndex);
  finally
    AtualizarSSLLibsCombo;
  end;
end;

procedure TForm1.cbbSSLTypeChange(Sender: TObject);
begin
  if cbbSSLType.ItemIndex <> -1 then
    ACBrNFe1.SSL.SSLType := TSSLType(cbbSSLType.ItemIndex);
end;

procedure TForm1.cbbXmlSignLibChange(Sender: TObject);
begin
  try
    if cbbXmlSignLib.ItemIndex <> -1 then
      ACBrNFe1.Configuracoes.Geral.SSLXmlSignLib := TSSLXmlSignLib(cbbXmlSignLib.ItemIndex);
  finally
    AtualizarSSLLibsCombo;
  end;
end;

procedure TForm1.ConfigurarComponente;
begin
  ACBrNFe1.Configuracoes.Certificados.NumeroSerie := edtNumSerie.Text;

  ACBrNFe1.SSL.DescarregarCertificado;

  with ACBrNFe1.Configuracoes.Geral do
  begin
    SSLLib := TSSLLib(cbbSSLLib.ItemIndex);
    SSLCryptLib := TSSLCryptLib(cbbCryptLib.ItemIndex);
    SSLHttpLib := TSSLHttpLib(cbbHttpLib.ItemIndex);
    SSLXmlSignLib := TSSLXmlSignLib(cbbXmlSignLib.ItemIndex);

    AtualizarSSLLibsCombo;
  end;

  ACBrNFe1.SSL.SSLType := TSSLType(cbbSSLType.ItemIndex);
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if DirectoryExists(ExtractFilePath(GetCurrentDir)) then
    ApagaDir(Self.Handle, 'Docs', False);
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  T: TSSLLib;
  U: TSSLCryptLib;
  V: TSSLHttpLib;
  X: TSSLXmlSignLib;
  Y: TSSLType;
begin
  cbbSSLLib.Items.Clear;
  for T := Low(TSSLLib) to High(TSSLLib) do
    cbbSSLLib.Items.Add(GetEnumName(TypeInfo(TSSLLib), integer(T)));
  cbbSSLLib.ItemIndex := 0;

  cbbCryptLib.Items.Clear;
  for U := Low(TSSLCryptLib) to High(TSSLCryptLib) do
    cbbCryptLib.Items.Add(GetEnumName(TypeInfo(TSSLCryptLib), integer(U)));
  cbbCryptLib.ItemIndex := 0;

  cbbHttpLib.Items.Clear;
  for V := Low(TSSLHttpLib) to High(TSSLHttpLib) do
    cbbHttpLib.Items.Add(GetEnumName(TypeInfo(TSSLHttpLib), integer(V)));
  cbbHttpLib.ItemIndex := 0;

  cbbXmlSignLib.Items.Clear;
  for X := Low(TSSLXmlSignLib) to High(TSSLXmlSignLib) do
    cbbXmlSignLib.Items.Add(GetEnumName(TypeInfo(TSSLXmlSignLib), integer(X)));
  cbbXmlSignLib.ItemIndex := 0;

  cbbSSLType.Items.Clear;
  for Y := Low(TSSLType) to High(TSSLType) do
    cbbSSLType.Items.Add(GetEnumName(TypeInfo(TSSLType), integer(Y)));
  cbbSSLType.ItemIndex := 0;

  lerConfiguracao;
  shpStatus.Visible := false;
  lblResposta.Visible := false;
end;

procedure TForm1.gravarConfiguracao;
var
  IniFile: string;
  Ini: TIniFile;
  StreamMemo: TMemoryStream;
begin
  IniFile := ChangeFileExt(ParamStr(0), '.ini');

  Ini := TIniFile.Create(IniFile);
  try
    Ini.WriteInteger('Certificado', 'SSLLib', cbbSSLLib.ItemIndex);
    Ini.WriteInteger('Certificado', 'CryptLib', cbbCryptLib.ItemIndex);
    Ini.WriteInteger('Certificado', 'HttpLib', cbbHttpLib.ItemIndex);
    Ini.WriteInteger('Certificado', 'XmlSignLib', cbbXmlSignLib.ItemIndex);
    Ini.WriteString('Certificado', 'NumSerie', edtNumSerie.Text);

    Ini.WriteInteger('WebService', 'SSLType', cbbSSLType.ItemIndex);

    StreamMemo := TMemoryStream.Create;
    StreamMemo.Seek(0, soBeginning);
    StreamMemo.Free;
  finally
    Ini.Free;
  end;
end;

procedure TForm1.lerConfiguracao;
var
  IniFile: string;
  Ini: TIniFile;
begin
  IniFile := ChangeFileExt(ParamStr(0), '.ini');

  Ini := TIniFile.Create(IniFile);
  try
    cbbSSLLib.ItemIndex := Ini.ReadInteger('Certificado', 'SSLLib', 0);
    cbbCryptLib.ItemIndex := Ini.ReadInteger('Certificado', 'CryptLib', 0);
    cbbHttpLib.ItemIndex := Ini.ReadInteger('Certificado', 'HttpLib', 0);
    cbbXmlSignLib.ItemIndex := Ini.ReadInteger('Certificado', 'XmlSignLib', 0);
    edtNumSerie.Text := Ini.ReadString('Certificado', 'NumSerie', '');

    cbbSSLType.ItemIndex := Ini.ReadInteger('WebService', 'SSLType', 0);

    ConfigurarComponente;
  finally
    Ini.Free;
  end;
end;

function TForm1.ListFiles(const Dir, Wildcard: string; const List: TStrings): Boolean;
var
  FileSpec: string;
  SR: TSearchRec;
  Success: Integer;
begin
  Assert(Assigned(List));
  Result := DirectoryExists(Dir);
  if not Result then
    Exit;
  FileSpec := Dir + '\';
  if Wildcard = '' then
    FileSpec := FileSpec + '*.*'
  else
    FileSpec := FileSpec + Wildcard;
  Success := FindFirst(FileSpec, faAnyFile, SR);
  try
    while Success = 0 do
    begin
      if (SR.Name <> '.') and (SR.Name <> '..') and (SR.Attr and faVolumeId = 0) then
        List.Add(Dir + '\' + SR.Name);
      Success := FindNext(SR);
    end;
  finally
    FindClose(SR);
  end;
end;

procedure TForm1.ApagaDir(hHandle: THandle; const sPath: string; Confirm: boolean);
var
  OpStruc: TSHFileOpStruct;
  FromBuffer, ToBuffer: array[0..128] of Char;
begin
  fillChar(OpStruc, Sizeof(OpStruc), 0);
  FillChar(FromBuffer, Sizeof(FromBuffer), 0);
  FillChar(ToBuffer, Sizeof(ToBuffer), 0);
  StrPCopy(FromBuffer, sPath);
  with OpStruc do
  begin
    Wnd := hHandle;
    wFunc := FO_DELETE;
    pFrom := @FromBuffer;
    pTo := @ToBuffer;
    if not Confirm then
    begin
      fFlags := FOF_NOCONFIRMATION;
    end;
    fAnyOperationsAborted := False;
    hNameMappings := nil;
  end;
  ShFileOperation(OpStruc);
end;

procedure TForm1.AtualizarSSLLibsCombo;
begin
  cbbSSLLib.ItemIndex := Integer(ACBrNFe1.Configuracoes.Geral.SSLLib);
  cbbCryptLib.ItemIndex := Integer(ACBrNFe1.Configuracoes.Geral.SSLCryptLib);
  cbbHttpLib.ItemIndex := Integer(ACBrNFe1.Configuracoes.Geral.SSLHttpLib);
  cbbXmlSignLib.ItemIndex := Integer(ACBrNFe1.Configuracoes.Geral.SSLXmlSignLib);

  cbbSSLType.Enabled := (ACBrNFe1.Configuracoes.Geral.SSLHttpLib in [httpWinHttp, httpOpenSSL]);
end;

end.

