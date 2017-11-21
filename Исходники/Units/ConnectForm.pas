unit ConnectForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, RepositoryUnit, MessagesUnit, jpeg,
  ComCtrls;

type
  TForm1 = class(TForm)
    imgLoading: TImage;
    lbl1: TLabel;
    lbl2: TLabel;
    img1: TImage;
    pnlInfo: TPanel;
    aniConnect: TAnimate;
    pnlMain: TPanel;
    lbledtServerAddress: TLabeledEdit;
    lbledtPathToDB: TLabeledEdit;
    lbledtPassword: TLabeledEdit;
    lbledtLogin: TLabeledEdit;
    lbl4: TLabel;
    lbl3: TLabel;
    chkServerConnection: TCheckBox;
    bvl3: TBevel;
    bvl2: TBevel;
    bvl1: TBevel;
    pnl1: TPanel;
    btnTestConnection: TBitBtn;
    btnEnter: TBitBtn;
    btnExit: TBitBtn;
    procedure btnEnterClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnTestConnectionClick(Sender: TObject);
    procedure chkServerConnectionClick(Sender: TObject);
    procedure ClearOnKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    function CheckEditValid(Lbled: TLabeledEdit): Boolean;
    function CheckEditValidColor(Lbled: TLabeledEdit): Boolean;
    function CheckFieldsValid(): Boolean;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  
implementation

uses MainFormUnit;

{$R *.dfm}

// Проверяем на пустоту поле для ввода
function TForm1.CheckEditValid(Lbled: TLabeledEdit): Boolean;
begin
  Result := not (Lbled.Text = '');
end;

// Проверка на заполненность компонент
function TForm1.CheckEditValidColor(Lbled: TLabeledEdit):Boolean;
begin
  Result := CheckEditValid(Lbled);
  if (Result) then
    Lbled.Color := TColor($88FF88)
  else
    Lbled.Color := TColor($8888FF);
end;


function TForm1.CheckFieldsValid(): Boolean;
begin
  // Если путь не введен
  if (not CheckEditValidColor(lbledtPathToDB)) then
  begin
    Result := False;
    ShowError('Вы не ввели путь к базе!');
    pnlInfo.Caption := 'Ошибка: Вы не ввели путь к базе!';
    pnlInfo.Color := TColor($8888FF);
    Exit;
  end;

  if (lbledtServerAddress.Enabled and (not CheckEditValid(lbledtServerAddress))) then
  begin
    Result := False;
    ShowError('Вы не ввели адрес для подключения к базе!');
    pnlInfo.Caption := 'Ошибка: Вы не ввели адрес для подключения к базе!';
    pnlInfo.Color := TColor($8888FF);
    Exit;
  end;

  // Если введен логин или пароль
  if (CheckEditValid(lbledtLogin) or CheckEditValid(lbledtPassword)) then
  begin
    // Если введен логин
    if (CheckEditValidColor(lbledtLogin)) then
    begin
      lbledtPassword.Color := TColor($88FF88);
    end
    else
    begin
      lbledtPassword.Color := TColor($8888FF);
      Result := False;
      ShowError('Логин не задан!');
      pnlInfo.Caption := 'Ошибка: Логин не задан!';
      pnlInfo.Color := TColor($8888FF);
      Exit;
    end;
  end;

  if (chkServerConnection.Checked and not CheckEditValidColor(lbledtServerAddress)) then
  begin
    Result := False;
    ShowError('Адресс сервера не указан! ' +
      'Если вы хотите подключиться локально, то снемите галочку!');
    pnlInfo.Caption := 'Ошибка: Адресс сервера не указан! ' +
      'Если вы хотите подключиться локально, то снемите галочку!';
    Exit;
  end;

  Result := True;
end;


// Вход в глвное меню
procedure TForm1.btnEnterClick(Sender: TObject);
var
  repo: TRepository;
begin
  // Проверка полей на заполненность
  if (not CheckFieldsValid()) then
    Exit;

  // Включаем анимацию
  aniConnect.Visible := True;
  aniConnect.Active := True;

  // Подключение к базе данных
  repo := TRepository.Create(MainForm);

  if (chkServerConnection.Enabled) then
  begin
    repo.Path :=
      lbledtServerAddress.Text + ':' + lbledtPathToDB.Text;
  end
  else
  begin
    repo.Path := lbledtPathToDB.Text;
  end;
  repo.Login := lbledtLogin.Text;
  repo.Password := lbledtPassword.Text;

  // Проверяем подключение
  if (not repo.TestConnection()) then
  begin
    ShowError(FAILED_CONNECT_TO_DATABASE);
    pnlInfo.Caption := 'Ошибка: ' + FAILED_CONNECT_TO_DATABASE;
    pnlInfo.Color := TColor($8888FF);
    // Выключаем анимацию
    aniConnect.Visible := False;
    aniConnect.Active := False;
    Exit;
  end;

  // Отключаем анимацию
  aniConnect.Visible := False;
  aniConnect.Active := False;

  pnlInfo.Caption := 'Подключение прошло успешно! Формируем главное меню.';
  pnlInfo.Color := TColor($FF8888);

  // Переходим в главное меню
  MainForm.ShowWithRepo(Self, repo);
end;

procedure TForm1.btnExitClick(Sender: TObject);
begin
  // Подтверждение выхода
  if (ShowConfirm(CONFIRM_EXIT)) then
    ExitProcess(0)
  else
    Abort;
end;

// Подключение к базе
procedure TForm1.btnTestConnectionClick(Sender: TObject);
var
  repo : TRepository;
begin
  // Проверка полей на заполненность
  if (not CheckFieldsValid()) then
    Exit;

  // Подключение к базе данных
  repo := TRepository.Create(MainForm);
  if (chkServerConnection.Enabled) then
  begin
    repo.Path :=
      lbledtServerAddress.Text + ':' + lbledtPathToDB.Text;
  end
  else
  begin
    repo.Path := lbledtPathToDB.Text;
  end;
  repo.Login := lbledtLogin.Text;
  repo.Password := lbledtPassword.Text;
  // Проверяем подключение
  if (repo.TestConnection()) then
  begin
    ShowInfo(SUCCESSIVELY_TEST_CONNECT);
    pnlInfo.Caption := SUCCESSIVELY_TEST_CONNECT;
    pnlInfo.Color := TColor($88FF88);
  end
  else
  begin
    ShowError(FAILED_CONNECT_TO_DATABASE);
    pnlInfo.Caption := 'Ошибка: ' + FAILED_CONNECT_TO_DATABASE;
    pnlInfo.Color := TColor($8888FF);
    Exit;
  end;
end;

procedure TForm1.chkServerConnectionClick(Sender: TObject);
begin
  lbledtServerAddress.Enabled := chkServerConnection.Checked;
end;

procedure TForm1.ClearOnKeyPress(Sender: TObject;
  var Key: Char);
begin
  (Sender as TLabeledEdit).Color := clWindow;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // Подтверждение выхода
  if (ShowConfirm(CONFIRM_EXIT)) then
    ExitProcess(0)
  else
    Abort;
end;

end.
