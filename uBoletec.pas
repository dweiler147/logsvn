{
 Data: 25/10/2016
 Autor: Diego Weiler
 Função: Gerar o arquivo por período do Boletec
 ----------------------------------------------
 Versão 1.0.0: Gera o arquivo XML com validações básicas
 Versão 1.1.0: Gera o arquivo CSV sem formatação
 Versão 1.2.0: Opção para excluir arquivos antigos CSV e XML
 Versão 1.3.0: Configuração de autenticação de usuário SVN -> Arquivo bat
 Versão 1.3.1: Pequenos ajustes em mensagens de erro e campos no arquivo CSV pedido pela Rosilei
 Versão 1.3.2: Ajustado ao novo layout proposto na reunião, para remoção do Chr(10) e Chr(13) da mensagem
 Versão 1.3.3: Implementado para abrir o arquivo CSV gerado
 Versão 1.4.0: Implementar leitura de um arquivo de configuração(config.ini)
 Versão 1.4.1: Inclusão do autor do fonte no arquivo CSV
 Versão 1.4.2: Opção para remoção da tag padrão no arquivo CSV
}
unit uBoletec;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, StdCtrls, Buttons, ComCtrls, ShellAPI, xmldom, XMLIntf,
  msxmldom, XMLDoc, IniFiles;

type
  TFrPrincipal = class(TForm)
    dtIni: TDateTimePicker;
    dtFim: TDateTimePicker;
    lbDtIni: TLabel;
    lbDtFim: TLabel;
    cbVersao: TComboBox;
    lbVersao: TLabel;
    OpenDialog1: TOpenDialog;
    edCaminho: TEdit;
    lbCaminho: TLabel;
    btAbrir: TBitBtn;
    btGerarXML: TButton;
    btGerarCSV: TButton;
    XMLDocument1: TXMLDocument;
    btExcluir: TButton;
    ckRemovePadrao: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure btAbrirClick(Sender: TObject);
    procedure cbVersaoChange(Sender: TObject);
    procedure btGerarXMLClick(Sender: TObject);
    procedure btGerarCSVClick(Sender: TObject);
    procedure btExcluirClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    function fRetornaDataArquivo(data: TDateTime): String;
    function fValidarGeracao: boolean;
    function fValidaArqGerado(tipo: Integer): Boolean;
    function fFormaUsrSvn: String;
    procedure pLerArquivoIni;
    procedure pCriarArquivoIniPadrao;
    procedure pGerarCSV;
    procedure pGerarXML;
    procedure pDeletarCSV;
    procedure pDeletarXML;
  public
    { Public declarations }
  end;

const
  extXML = '.xml';
  extCSV = '.csv';

var
  FrPrincipal: TFrPrincipal;
  nomeArq: String;
  svn, usr, pwd : String;

implementation

uses DateUtils;

{$R *.dfm}

procedure TFrPrincipal.FormCreate(Sender: TObject);
var
  sl : TStringList;
  comando : String;
  existe : Boolean;
begin
  try
    //Carrega arquivo config.ini
    pLerArquivoIni();

    //Carrega data atual nos campos DtIni = data atual-7 e DtFim = data atual
    dtFim.DateTime := Now;
    dtIni.DateTime := IncDay(Now, -7);

    //Preenche o diretório padrão com o de execução do aplicativo
    edCaminho.Text := GetCurrentDir;

    //Executa comando para buscar versões do SVN
    comando := '/c svn ' + fFormaUsrSvn() + ' ls ' + svn + 'Releases > ' + getcurrentdir + '\lista.txt';
    ShellExecute(0, nil, 'cmd.exe', pChar(comando), nil, SW_HIDE);

    existe := false;
    //Preenche as versões no combobox
    while not(existe) do
      begin
        if (FileExists(getcurrentdir + '\lista.txt')) then
           existe := True
        else
           Sleep(500);
      end;

    Sleep(1000);  
    sl := TStringList.Create;
    sl.LoadFromFile(getCurrentDir + '\lista.txt');
    sl.Text := StringReplace(sl.Text, '/', '', [rfReplaceAll]);
    cbVersao.Items.Clear;
    cbVersao.Items.Add('20100');
    cbVersao.Items.AddStrings(sl);
    sl.Free;
  except
    Application.MessageBox('Não foi possível criar arquivo de versões. Reinicie a aplicação!','Boletec', mb_OK + mb_IconWarning);
  end;
end;

procedure TFrPrincipal.btAbrirClick(Sender: TObject);
var
  data : TDateTime;
begin
  if fValidarGeracao then
    begin
      data := Now;
      OpenDialog1.Title := 'Selecione um arquivo para exportar';
      OpenDialog1.InitialDir := edCaminho.Text;
      OpenDialog1.FileName := cbVersao.Text + '_' + fRetornaDataArquivo(data) + extXML;
      if OpenDialog1.Execute then
        begin
          edCaminho.Text := OpenDialog1.FileName;
        end;
    end;
end;

procedure TFrPrincipal.cbVersaoChange(Sender: TObject);
var
  data : TdateTime;
begin
  data := Now;
  edCaminho.Text := GetCurrentDir + '\' + cbVersao.Text + '_' + fRetornaDataArquivo(data) + extXML;
end;

function TFrPrincipal.fRetornaDataArquivo(data: TDateTime): String;
begin
  result := StringReplace(formatdatetime('yyyy/mm/dd', data), '/', '-', [rfReplaceAll]);
end;

procedure TFrPrincipal.btGerarXMLClick(Sender: TObject);
begin
  //Valida campos
  if fValidarGeracao then
    begin
      if fValidaArqGerado(1) then
        Application.MessageBox('Arquivo XML gerado com sucesso!','Boletec',mb_OK);
    end;
end;

procedure TFrPrincipal.btGerarCSVClick(Sender: TObject);
begin
  if fValidarGeracao then
    begin
      if fValidaArqGerado(1) then
      begin
        if fValidaArqGerado(2) then
        begin
          if Application.MessageBox('Arquivo CSV gerado com sucesso!' + Chr(10) + chr(13) +
                                    'Deseja abrir o arquivo recém criado?','Boletec',MB_YESNO + mb_iconquestion) = id_YES then
            ShellExecute(handle,'open',PChar(nomeArq), '','',SW_SHOWNORMAL);
        end;
      end;
    end;
end;

procedure TFrPrincipal.pGerarCSV;
var
  arq: TextFile;
  data : TDateTime;
  i, j : integer;
  lista: TStringList;
  NodeRec,NodeMsg:IXmlNode;
  revisao, dataArq, mensagem, autor : String;
  removePadrao: Boolean;
begin
  data := Now;
  removePadrao := ckRemovePadrao.Checked;
  //Nome padrão: boletec_2016-10-06.csv
  nomeArq := '';
  nomeArq := GetCurrentDir + '\boletec_' + cbVersao.Text + '_' + fRetornaDataArquivo(data) + extCSV;
  //Cria o arquivo no diretório
  AssignFile(arq, nomeArq);
  Rewrite(arq);
  writeln(arq, 'Revisão;Data;Autor;Ticket Cliente;Ticket Interno;Tema;Solução;Tela;Módulo;Versão;');
  lista := TStringList.Create;

  XMLDocument1.LoadFromFile(edCaminho.Text);
  NodeRec := XMLDocument1.DocumentElement;

  for i := 0 to NodeRec.ChildNodes.Count-1 do
    begin
      //inicialização das variáveis
      mensagem := '';
      dataArq := '';
      revisao := '';
      autor := '';

      NodeMsg := NodeRec.ChildNodes[i];

      //Busca o número pelo atributo do XML
      revisao := NodeMsg.AttributeNodes.Nodes['revision'].Text;

      for j := 0 to NodeMsg.ChildNodes.Count -1 do
      begin
        if (NodeMsg.ChildNodes[j].NodeName='author') then
        begin
          autor := String(NodeMsg.ChildNodes[j].Text);
        end;
        if (NodeMsg.ChildNodes[j].NodeName='msg') then
        begin
          //Tratamento para caracteres especiais = { ; # Chr(10) Chr(13) }
          mensagem := StringReplace(String(NodeMsg.ChildNodes[j].Text),';','.',[rfReplaceAll]);
          mensagem := StringReplace(mensagem,'#',';',[rfReplaceAll]);
          mensagem := StringReplace(mensagem,Chr(10),'',[rfReplaceAll]);
          mensagem := StringReplace(mensagem,Chr(13),'',[rfReplaceAll]);
          if removePadrao then
          begin
            mensagem := StringReplace(mensagem,'TicketCliente:','',[rfReplaceAll]);
            mensagem := StringReplace(mensagem,'TicketInterno:','',[rfReplaceAll]);
            mensagem := StringReplace(mensagem,'Tema:','',[rfReplaceAll]);
            mensagem := StringReplace(mensagem,'Solução:','',[rfReplaceAll]);
            mensagem := StringReplace(mensagem,'Tela:','',[rfReplaceAll]);
            mensagem := StringReplace(mensagem,'Módulos:','',[rfReplaceAll]);
            mensagem := StringReplace(mensagem,'versões:','',[rfReplaceAll]);
          end;
        end;
        if (NodeMsg.ChildNodes[j].NodeName='date') then
        begin
          dataArq := Copy(StringReplace(String(NodeMsg.ChildNodes[j].Text),'#',';',[rfReplaceAll]),1,10);
        end;
      end;
      dataArq := Copy(dataArq,9,2) + '/' + Copy(dataArq,6,2) + '/' + Copy(dataArq,1,4);

      //Não inclui ; quando estiver no padrão novo
      if (mensagem[1] = ';') then
        lista.Add(revisao + ';' + dataArq + ';' + autor + mensagem)
      else
        lista.Add(revisao + ';' + dataArq + ';' + autor + ';' + mensagem);
    end;

  for i := 0 to lista.count-1 do
    begin
      Writeln(arq, lista[i]);
    end;

  lista.Free;
  CloseFile(arq);
end;

function TFrPrincipal.fValidarGeracao:boolean;
begin
  //Validações:
  //Data Final não pode ser menor que a inicial
  if (dtFim.Date < dtIni.Date) then
    begin
      Application.MessageBox('Data final não pode ser menor que a data inicial.','Atenção', mb_OK + mb_IconWarning);
      dtFim.SetFocus;
      dtFim.Perform(WM_KEYDOWN, VK_F4, 0);
      result := false;
    end
  else
    begin
      //Versão deve estar selecionada
      if (cbVersao.Text = '') then
        begin
          Application.MessageBox('Selecione uma versão a gerar','Atenção', mb_ok + mb_IconWarning);
          cbVersao.SetFocus;
          cbVersao.Perform(WM_KEYDOWN, VK_F4, 0);
          result := false;
        end
    else
      result := true;
    end;
end;

procedure TFrPrincipal.pGerarXML;
var
  comando : String;
  i : integer;
  existe : Boolean;
begin
  //Comandos para gerar o arquivo XML
  comando := '/c svn log ' + svn;
  i := cbVersao.itemindex;

  //Avaliar se é maior que a versão 2.01.00
  if (i > 0) then
    comando := comando + 'Releases/' + cbVersao.Text;

  //Adiciona restante do caminho
  comando := comando + ' -r {' + fRetornaDataArquivo(dtIni.DateTime) + '}:{' +
             fRetornaDataArquivo(dtFim.DateTime) + '} -v --xml > ' + edCaminho.Text + fFormaUsrSvn();

  //Executa o comando via shell (DOS)
  ShellExecute(0, nil, 'cmd.exe', pChar(comando), nil, SW_HIDE);
  existe := false;

  //Tratamento para criação do arquivo, ocorre erro quando não há o arquivo e carrega o combobox de versão
  while not(existe) do
    begin
      if (FileExists(edCaminho.Text)) then
        existe := True
      else
        Sleep(500);
    end;

  Sleep(1000);
end;

procedure TFrPrincipal.btExcluirClick(Sender: TObject);
begin
  pDeletarCSV;
  pDeletarXML;
  Application.MessageBox('Exclusão realizada com sucesso!','Boletec',mb_OK);
end;

procedure TFrPrincipal.pDeletarCSV;
var
  comando : String;
begin
  //Exclusão de arquivos CSV
  comando := '/c del *.csv /q /f';
  ShellExecute(0, nil, 'cmd.exe', pChar(comando), nil, SW_HIDE);
end;

procedure TFrPrincipal.pDeletarXML;
var
  comando : String;
begin
  //Exclusão de arquivos XML
  comando := '/c del *.xml /q /f';
  ShellExecute(0, nil, 'cmd.exe', pChar(comando), nil, SW_HIDE);
end;

procedure TFrPrincipal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  try
    //Exclui arquivo lista.txt ao finalizar a aplicação
    DeleteFile(GetCurrentDir + '\lista.txt');
  except
  end;
end;

function TFrPrincipal.fValidaArqGerado(tipo: Integer): Boolean;
begin
  case tipo of
    //Gerar XML
    1:
    begin
      try
        pGerarXML;
        result := True;
      except
        Application.MessageBox('Não foi possível gerar o arquivo XML.','Atenção', mb_OK + mb_IconWarning);
        result := False;
      end;
    end;
    //Gerar CSV
    2:
    begin
      try
        pGerarCSV;
        Result := True;
      except
        Application.MessageBox('Não foi possível gerar o arquivo CSV. ' + Chr(10) + Chr(13) +
                           'Verifique se o arquivo não se encontra aberto!','Atenção', mb_OK + mb_IconWarning);
        result := false;
      end;
    end;
  end;
end;

function TFrPrincipal.fFormaUsrSvn: String;
begin
  result := ' --username ' + usr + ' --password ' + pwd;
end;

procedure TFrPrincipal.pLerArquivoIni;
var
  arquivoIni : TIniFile;
begin
  if FileExists(GetCurrentDir + '\config.ini') then
    begin
      arquivoIni := TIniFile.Create(GetCurrentDir + '\config.ini');
      svn := arquivoIni.ReadString('Config','localsvn',svn);
      usr := arquivoIni.ReadString('Config','usr',usr);
      pwd := arquivoIni.ReadString('Config','pwd',pwd);
      arquivoIni.Free;
    end
  else
    begin
      Application.MessageBox('Não foi possível criar o arquivo de configuração' + chr(10) + Chr(13) +
                           'Será gerado um arquivo padrão. Favor verificar!','Atenção',MB_OK + MB_ICONWARNING);
      pCriarArquivoIniPadrao;
    end;
end;

procedure TFrPrincipal.pCriarArquivoIniPadrao;
var
  arquivoIni : TIniFile;
begin
  //Rotina que cria um arquivo INI padrão
  try
    arquivoIni := TIniFile.Create(GetCurrentDir + '\config.ini');
    try
      arquivoIni.WriteString('Config','localsvn','https://_server_mtm/svn/MultM3/');
      arquivoIni.WriteString('Config','usr','diego.weiler@mult.com.br');
      arquivoIni.WriteString('Config','pwd','mult@2016');
    except
      Application.MessageBox('Não foi possível criar o arquivo de configuração' + chr(10) + Chr(13) +
                           'Verifique as permissões de gravação do diretório. Aplicação encerrada!','Erro',MB_OK + MB_ICONERROR);
      close();
    end;
  finally
    arquivoIni.Free;
  end;
end;

end.
