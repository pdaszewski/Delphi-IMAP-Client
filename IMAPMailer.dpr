program IMAPMailer;

uses
  Vcl.Forms,
  AOknoGl_frm in 'AOknoGl_frm.pas' {MainForm},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Mailer';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
