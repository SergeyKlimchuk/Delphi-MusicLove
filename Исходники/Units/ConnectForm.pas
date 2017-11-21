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

// ��������� �� ������� ���� ��� �����
function TForm1.CheckEditValid(Lbled: TLabeledEdit): Boolean;
begin
  Result := not (Lbled.Text = '');
end;

// �������� �� ������������� ���������
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
  // ���� ���� �� ������
  if (not CheckEditValidColor(lbledtPathToDB)) then
  begin
    Result := False;
    ShowError('�� �� ����� ���� � ����!');
    pnlInfo.Caption := '������: �� �� ����� ���� � ����!';
    pnlInfo.Color := TColor($8888FF);
    Exit;
  end;

  if (lbledtServerAddress.Enabled and (not CheckEditValid(lbledtServerAddress))) then
  begin
    Result := False;
    ShowError('�� �� ����� ����� ��� ����������� � ����!');
    pnlInfo.Caption := '������: �� �� ����� ����� ��� ����������� � ����!';
    pnlInfo.Color := TColor($8888FF);
    Exit;
  end;

  // ���� ������ ����� ��� ������
  if (CheckEditValid(lbledtLogin) or CheckEditValid(lbledtPassword)) then
  begin
    // ���� ������ �����
    if (CheckEditValidColor(lbledtLogin)) then
    begin
      lbledtPassword.Color := TColor($88FF88);
    end
    else
    begin
      lbledtPassword.Color := TColor($8888FF);
      Result := False;
      ShowError('����� �� �����!');
      pnlInfo.Caption := '������: ����� �� �����!';
      pnlInfo.Color := TColor($8888FF);
      Exit;
    end;
  end;

  if (chkServerConnection.Checked and not CheckEditValidColor(lbledtServerAddress)) then
  begin
    Result := False;
    ShowError('������ ������� �� ������! ' +
      '���� �� ������ ������������ ��������, �� ������� �������!');
    pnlInfo.Caption := '������: ������ ������� �� ������! ' +
      '���� �� ������ ������������ ��������, �� ������� �������!';
    Exit;
  end;

  Result := True;
end;


// ���� � ������ ����
procedure TForm1.btnEnterClick(Sender: TObject);
var
  repo: TRepository;
begin
  // �������� ����� �� �������������
  if (not CheckFieldsValid()) then
    Exit;

  // �������� ��������
  aniConnect.Visible := True;
  aniConnect.Active := True;

  // ����������� � ���� ������
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

  // ��������� �����������
  if (not repo.TestConnection()) then
  begin
    ShowError(FAILED_CONNECT_TO_DATABASE);
    pnlInfo.Caption := '������: ' + FAILED_CONNECT_TO_DATABASE;
    pnlInfo.Color := TColor($8888FF);
    // ��������� ��������
    aniConnect.Visible := False;
    aniConnect.Active := False;
    Exit;
  end;

  // ��������� ��������
  aniConnect.Visible := False;
  aniConnect.Active := False;

  pnlInfo.Caption := '����������� ������ �������! ��������� ������� ����.';
  pnlInfo.Color := TColor($FF8888);

  // ��������� � ������� ����
  MainForm.ShowWithRepo(Self, repo);
end;

procedure TForm1.btnExitClick(Sender: TObject);
begin
  // ������������� ������
  if (ShowConfirm(CONFIRM_EXIT)) then
    ExitProcess(0)
  else
    Abort;
end;

// ����������� � ����
procedure TForm1.btnTestConnectionClick(Sender: TObject);
var
  repo : TRepository;
begin
  // �������� ����� �� �������������
  if (not CheckFieldsValid()) then
    Exit;

  // ����������� � ���� ������
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
  // ��������� �����������
  if (repo.TestConnection()) then
  begin
    ShowInfo(SUCCESSIVELY_TEST_CONNECT);
    pnlInfo.Caption := SUCCESSIVELY_TEST_CONNECT;
    pnlInfo.Color := TColor($88FF88);
  end
  else
  begin
    ShowError(FAILED_CONNECT_TO_DATABASE);
    pnlInfo.Caption := '������: ' + FAILED_CONNECT_TO_DATABASE;
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
  // ������������� ������
  if (ShowConfirm(CONFIRM_EXIT)) then
    ExitProcess(0)
  else
    Abort;
end;

end.
