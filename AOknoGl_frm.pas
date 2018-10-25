unit AOknoGl_frm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  IdIMAP4, IdSSLOpenSSL, IdText, IdMessage, IdExplicitTLSClientServerBase,
  Vcl.StdCtrls, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdMessageClient, IdAttachment, Vcl.ComCtrls, Vcl.ExtCtrls, Data.DB, Data.Win.ADODB;

type
  TAOknoGl = class(TForm)
    btn_get_messages: TButton;
    ListBox1: TListBox;
    ProgressBar: TProgressBar;
    btn_stop: TButton;
    ProgressBarAutoStart: TProgressBar;
    Start: TTimer;
    AutoRun: TTimer;

    procedure GetMessages(const UserName, Password: string; Logi: TStrings);
    procedure Load_file(path_to_file: String);
    procedure Add_to_log(text: String; insert_date_time: Boolean);

    procedure btn_get_messagesClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DeleteAllFiles(dir : string);
    procedure StartTimer(Sender: TObject);
    procedure btn_stopClick(Sender: TObject);
    procedure AutoRunTimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const
 wersja = '1.0.0';
 data_kompilacji = '2018-10-25';

var
  AOknoGl: TAOknoGl;
  lista_odebranych : TStringList;
  lista_odebranych_plik : String;
  folder_tmp : String;
  plik_logu : String;
  logi : TStringList;
  ostatniSQL : String;

implementation

{$R *.dfm}

procedure TAOknoGl.Add_to_log(text: String; insert_date_time: Boolean);
Begin
 if insert_date_time then logi.Add('['+DateTimeToStr(Now)+'] - '+text)
 else logi.Add(text);
End;

procedure TAOknoGl.DeleteAllFiles(dir : string);
  var sl   : tstringlist;
      sr   : tsearchrec;
      i,mx : integer;
Begin
  sl := tstringlist.create;
  try
    sl.clear;
    if findfirst(dir+'*.*',faanyfile,sr) = 0
    then begin
         repeat
           sl.add(sr.name);
         until findnext(sr) <> 0;
         findclose(sr);
         end;
       Begin
         mx := sl.count - 1;
         for i := 0 to mx
         do deletefile(dir+sl[i]);
         end;
  finally
    sl.free;
  end;
end;

procedure TAOknoGl.StartTimer(Sender: TObject);
begin
 ProgressBarAutoStart.Position:=ProgressBarAutoStart.Position-1;
 if ProgressBarAutoStart.Position=0 then
  Begin
   Start.Enabled:=False;
   btn_stop.Visible:=False;
   Application.ProcessMessages;
   btn_get_messagesClick(Self);
  End;
end;

procedure TAOknoGl.btn_get_messagesClick(Sender: TObject);
begin
 btn_stopClick(Self);
 btn_get_messages.Enabled:=False;
 Application.ProcessMessages;
 GetMessages('some e-mail address','some password', ListBox1.Items);
 Close;
end;

procedure TAOknoGl.btn_stopClick(Sender: TObject);
begin
 Start.Enabled:=False;
 ProgressBarAutoStart.Visible:=False;
 btn_stop.Visible:=False;
end;

procedure TAOknoGl.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 lista_odebranych.SaveToFile(lista_odebranych_plik);
 Add_to_log('I''m closing the application',True);
 Add_to_log('===============================================================================================',False);
 Add_to_log('',False);
 logi.SaveToFile(plik_logu);
end;

procedure TAOknoGl.AutoRunTimer(Sender: TObject);
Var
  struktura : String;
begin
 AutoRun.Enabled:=False;

 struktura:=ExtractFilePath(Application.ExeName)+'Dane';
 if DirectoryExists(struktura)=False then CreateDir(struktura);

 if DirectoryExists(struktura+'\Kopie')=False then CreateDir(struktura+'\Kopie');

 plik_logu:=ExtractFilePath(Application.ExeName)+'Logi';
 if DirectoryExists(plik_logu)=False then CreateDir(plik_logu);
 plik_logu:=plik_logu+'\log_'+DateToStr(Now)+'.txt';

 logi := TStringList.Create;
 if FileExists(plik_logu) then logi.LoadFromFile(plik_logu);

 folder_tmp:=struktura+'\temp\';
 if DirectoryExists(folder_tmp)=False then CreateDir(folder_tmp);

 DeleteAllFiles(folder_tmp);

 lista_odebranych := TStringList.Create;
 lista_odebranych.Sorted := True;
 lista_odebranych_plik:=struktura+'\lista_odebranych.dat';

 if FileExists(lista_odebranych_plik) then
  Begin
   struktura:=struktura+'\Kopie\lista_odebranych_'+DateToStr(Now)+'.dat';
   CopyFile(PWideChar(lista_odebranych_plik),PWideChar(struktura),False);
   lista_odebranych.LoadFromFile(lista_odebranych_plik);
  End;

 ProgressBarAutoStart.Position:=ProgressBarAutoStart.Max;
 Start.Enabled:=True;
 btn_get_messages.Enabled:=True;
 btn_stop.Enabled:=True;

 Add_to_log('===============================================================================================',False);
 Add_to_log('Launch the receiving application from e-mail, version: '+wersja,True);
 Add_to_log('',False);
end;


procedure TAOknoGl.FormCreate(Sender: TObject);
begin
 Caption:='Mailer Version '+wersja;
 btn_get_messages.Enabled:=False;
 btn_stop.Enabled:=False;
 ProgressBarAutoStart.Position:=0;
end;

procedure TAOknoGl.FormShow(Sender: TObject);
begin
 AutoRun.Enabled:=True;
end;

procedure TAOknoGl.GetMessages(const UserName, Password: string; Logi: TStrings);
var
  MsgIndex: Integer;
  MsgObject: TIdMessage;
  PartIndex: Integer;
  IMAPClient: TIdIMAP4;
  OpenSSLHandler: TIdSSLIOHandlerSocketOpenSSL;
  nazwa_pliku: string;
  identyfikator_wiadomosci: string;
  data_wiadomosci: string;
  adres_wysylajacego: string;
  temat_wiadomosci: string;
  nowy_plik: string;
begin
  Logi.Clear;
  IMAPClient := TIdIMAP4.Create(nil);
  try
    OpenSSLHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
    try
      IMAPClient.IOHandler  := OpenSSLHandler;
      IMAPClient.Host       := 'host address';
      IMAPClient.Port       := 993; //port number for host
      IMAPClient.UseTLS     := utUseImplicitTLS;
      IMAPClient.Username   := UserName;
      IMAPClient.Password   := Password;
      IMAPClient.Connect;
      try
        if IMAPClient.SelectMailBox('INBOX') then
        begin
          Logi.BeginUpdate;
          try
           ProgressBar.Position:=0;
           ProgressBar.Max:=IMAPClient.MailBox.TotalMsgs;
            for MsgIndex := 1 to IMAPClient.MailBox.TotalMsgs do
            begin
              MsgObject := TIdMessage.Create(nil);
              try
                IMAPClient.Retrieve(MsgIndex, MsgObject);
                MsgObject.MessageParts.CountParts;

                identyfikator_wiadomosci:=MsgObject.MsgId;

                if lista_odebranych.IndexOf(identyfikator_wiadomosci) = -1 then
                 Begin
                  data_wiadomosci := DateTimeToStr(MsgObject.Date);
                  adres_wysylajacego := MsgObject.From.Address;
                  temat_wiadomosci := MsgObject.Subject;

                  // if adres_wysylajacego='some excepted e-mail adress' then
                  Begin

                    Logi.Add(identyfikator_wiadomosci + ' - ' + data_wiadomosci + ' - ' + adres_wysylajacego + ' - ' + temat_wiadomosci);
                    Application.ProcessMessages;
                    if MsgObject.MessageParts.AttachmentCount > 0 then
                    Begin
                      Logi.Add(' >> The message contains attachments: ' + IntToStr(MsgObject.MessageParts.AttachmentCount));
                      for PartIndex := 0 to MsgObject.MessageParts.Count - 1 do
                      Begin
                        if MsgObject.MessageParts[PartIndex] is TIdAttachment then
                        Begin
                          nazwa_pliku := TIdAttachment(MsgObject.MessageParts[PartIndex]).FileName;
                          Logi.Add(' -> ' + nazwa_pliku);
                          TIdAttachment(MsgObject.MessageParts[PartIndex]).SaveToFile(folder_tmp + nazwa_pliku);
                          nowy_plik:=folder_tmp+nazwa_pliku;

                          if AnsiLowerCase(ExtractFileExt(nowy_plik))='.csv' then
                          Load_file(nowy_plik);

                        End;
                      End;
                    End;

                  End;

                  lista_odebranych.Add(identyfikator_wiadomosci);
                 End;

              finally
                MsgObject.Free;
              end;
             ProgressBar.Position:=ProgressBar.Position+1;
             Application.ProcessMessages;
            end;
          finally
            Logi.EndUpdate;
          end;
        end;
      finally
        IMAPClient.Disconnect;
      end;
    finally
      OpenSSLHandler.Free;
    end;
  finally
    IMAPClient.Free;
  end;
end;

procedure TAOknoGl.Load_file(path_to_file: String);
Begin
 { TODO : In this place you can write import from file }
End;


end.
