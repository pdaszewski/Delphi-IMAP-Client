unit AOknoGl_frm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  IdIMAP4, IdSSLOpenSSL, IdText, IdMessage, IdExplicitTLSClientServerBase,
  Vcl.StdCtrls, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdMessageClient, IdAttachment, Vcl.ComCtrls, Vcl.ExtCtrls, Data.DB, Data.Win.ADODB;

type
  TMainForm = class(TForm)
    btn_get_messages: TButton;
    ListBox1: TListBox;
    ProgressBar: TProgressBar;
    btn_stop: TButton;
    ProgressBarAutoStart: TProgressBar;
    Start: TTimer;
    AutoRun: TTimer;

    ///<summary>Connection procedure to INBOX for given user and password</summary>
    procedure GetMessages(const UserName, Password: string; Logi: TStrings);

    ///<summary>Procedure to load new files from attachements</summary>
    procedure Load_file(path_to_file: String);

    ///<summary>Procedure to insert text to logs list</summary>
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
 version = '1.0.0';
 compilation_date = '2018-10-26';

var
  MainForm: TMainForm;
  inbox_list : TStringList;
  file_with_inbox_list : String;
  tmp_directory : String;
  log_file : String;
  logs : TStringList;

implementation

{$R *.dfm}

procedure TMainForm.Add_to_log(text: String; insert_date_time: Boolean);
Begin
 if insert_date_time then logs.Add('['+DateTimeToStr(Now)+'] - '+text)
 else logs.Add(text);
End;

procedure TMainForm.DeleteAllFiles(dir : string);
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

procedure TMainForm.StartTimer(Sender: TObject);
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

procedure TMainForm.btn_get_messagesClick(Sender: TObject);
begin
 btn_stopClick(Self);
 btn_get_messages.Enabled:=False;
 Application.ProcessMessages;
 GetMessages('some e-mail address','some password', ListBox1.Items);
 Close;
end;

procedure TMainForm.btn_stopClick(Sender: TObject);
begin
 Start.Enabled:=False;
 ProgressBarAutoStart.Visible:=False;
 btn_stop.Visible:=False;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 inbox_list.SaveToFile(file_with_inbox_list);
 Add_to_log('I''m closing the application',True);
 Add_to_log('===============================================================================================',False);
 Add_to_log('',False);
 logs.SaveToFile(log_file);
end;

procedure TMainForm.AutoRunTimer(Sender: TObject);
Var
  path_name : String;
begin
 AutoRun.Enabled:=False;

 path_name:=ExtractFilePath(Application.ExeName)+'Dane';
 if DirectoryExists(path_name)=False then CreateDir(path_name);

 if DirectoryExists(path_name+'\Kopie')=False then CreateDir(path_name+'\Kopie');

 log_file:=ExtractFilePath(Application.ExeName)+'Logi';
 if DirectoryExists(log_file)=False then CreateDir(log_file);
 log_file:=log_file+'\log_'+DateToStr(Now)+'.txt';

 logs := TStringList.Create;
 if FileExists(log_file) then logs.LoadFromFile(log_file);

 tmp_directory:=path_name+'\temp\';
 if DirectoryExists(tmp_directory)=False then CreateDir(tmp_directory);

 DeleteAllFiles(tmp_directory);

 inbox_list := TStringList.Create;
 inbox_list.Sorted := True;
 file_with_inbox_list:=path_name+'\lista_odebranych.dat';

 if FileExists(file_with_inbox_list) then
  Begin
   path_name:=path_name+'\Kopie\lista_odebranych_'+DateToStr(Now)+'.dat';
   CopyFile(PWideChar(file_with_inbox_list),PWideChar(path_name),False);
   inbox_list.LoadFromFile(file_with_inbox_list);
  End;

 ProgressBarAutoStart.Position:=ProgressBarAutoStart.Max;
 Start.Enabled:=True;
 btn_get_messages.Enabled:=True;
 btn_stop.Enabled:=True;

 Add_to_log('===============================================================================================',False);
 Add_to_log('Launch the receiving application from e-mail, version: '+version,True);
 Add_to_log('',False);
end;


procedure TMainForm.FormCreate(Sender: TObject);
begin
 Caption:='Mailer Version '+version;
 btn_get_messages.Enabled:=False;
 btn_stop.Enabled:=False;
 ProgressBarAutoStart.Position:=0;
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
 AutoRun.Enabled:=True;
end;

procedure TMainForm.GetMessages(const UserName, Password: string; Logi: TStrings);
var
  MsgIndex: Integer;
  MsgObject: TIdMessage;
  PartIndex: Integer;
  IMAPClient: TIdIMAP4;
  OpenSSLHandler: TIdSSLIOHandlerSocketOpenSSL;
  file_name: string;
  message_identifier: string;
  message_date: string;
  sending_address: string;
  message_topic: string;
  new_file: string;
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

                message_identifier:=MsgObject.MsgId;

                if inbox_list.IndexOf(message_identifier) = -1 then
                 Begin
                  message_date := DateTimeToStr(MsgObject.Date);
                  sending_address := MsgObject.From.Address;
                  message_topic := MsgObject.Subject;

                  // if sending_address='some excepted e-mail adress' then
                  Begin

                    Logi.Add(message_identifier + ' - ' + message_date + ' - ' + sending_address + ' - ' + message_topic);
                    Application.ProcessMessages;
                    if MsgObject.MessageParts.AttachmentCount > 0 then
                    Begin
                      Logi.Add(' >> The message contains attachments: ' + IntToStr(MsgObject.MessageParts.AttachmentCount));
                      for PartIndex := 0 to MsgObject.MessageParts.Count - 1 do
                      Begin
                        if MsgObject.MessageParts[PartIndex] is TIdAttachment then
                        Begin
                          file_name := TIdAttachment(MsgObject.MessageParts[PartIndex]).FileName;
                          Logi.Add(' -> ' + file_name);
                          TIdAttachment(MsgObject.MessageParts[PartIndex]).SaveToFile(tmp_directory + file_name);
                          new_file:=tmp_directory+file_name;

                          if AnsiLowerCase(ExtractFileExt(new_file))='.csv' then
                          Load_file(new_file);

                        End;
                      End;
                    End;

                  End;

                  inbox_list.Add(message_identifier);
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

procedure TMainForm.Load_file(path_to_file: String);
Begin
 { TODO : In this place you can write import from file }
End;


end.
