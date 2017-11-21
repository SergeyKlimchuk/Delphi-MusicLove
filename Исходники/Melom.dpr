program Melom;

uses
  Forms,
  ConnectForm in 'Units\ConnectForm.pas' {Form1},
  MessagesUnit in 'Libs\MessagesUnit.pas',
  RepositoryUnit in 'Libs\RepositoryUnit.pas',
  MainFormUnit in 'Units\MainFormUnit.pas' {MainForm};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Melom (The sounds info collector)';
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
