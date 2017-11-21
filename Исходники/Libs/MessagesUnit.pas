unit MessagesUnit;

interface

uses Forms, StdCtrls, Classes, Dialogs, SysUtils, Windows;

const
  FAILED_CONNECT_TO_DATABASE  = '�� ������� ������������ � ���� ������!';
  FAILED_CREATE_CONTEXT       = '�� ������� ������� ��������!';
  FAILED_REQUEST_TABLE        = '�� ������� �������� �������!';
  FAILED_REQUEST_COLUMN_NAMES = '�� ������� �������� ��������� ��������!';
  FAILED_SQL_REQUEST          = '�� ������� ��������� ������!';
  FAILED_INSERT_RECORD        = '�� ������� �������� ������!';

  SUCCESSIVELY_CONNECT        = '����������� ������ �������!';
  SUCCESSIVELY_TEST_CONNECT   = '�������� ����������� � ���� ������ ������ �������!';

  SQL_SELECT_ALL              = 'SELECT * FROM ';
  SQL_DELETE_ALL              = 'DELETE * FROM ';

  CONFIRM_EXIT                = '�� ������� ��� ������ �����?';
  CONFIRM_DELETE              = '�� ������� ��� ������ ������� ������?';

  WARNING_FIELDS_IS_EMPTY     = '��������! ���� ��� ��������� ������������ ����� �� ���� ���������!';

  ERROR_RECORD_NOT_SELECTED   = '������! �� ���� ������ �� ���� ��������!';

  BUTTON_CANCEL = '������';
  BUTTON_YES    = '��';
  BUTTON_NO     = '���';
  BUTTON_OK     = '��';

  procedure ShowError(msg: String);
  procedure ShowWarning(msg: String);
  procedure ShowInfo(msg: String);
  function ShowConfirm(msg: String): Boolean;

implementation

function Show(CONST Msg: string; DlgTypt: TmsgDlgType; button: TMsgDlgButtons;
  Captions: ARRAY OF string; dlgcaption: string): Integer;
var
  aMsgdlg: TForm;
  i: Integer;
  Dlgbutton: Tbutton;
  Captionindex: Integer;
begin
  aMsgdlg := createMessageDialog(Msg, DlgTypt, button);
  aMsgdlg.Caption := dlgcaption;
  aMsgdlg.BiDiMode := bdRightToLeft;
  Captionindex := 0;
  for i := 0 to aMsgdlg.componentcount - 1 Do
  begin
    if (aMsgdlg.components[i] is Tbutton) then
    Begin
      Dlgbutton := Tbutton(aMsgdlg.components[i]);
      Dlgbutton.Caption := Captions[Captionindex];
      
      if (Dlgbutton.Caption = BUTTON_CANCEL) then
        Dlgbutton.TabOrder := 0;

      inc(Captionindex);
    end;
  end;

  Result := aMsgdlg.Showmodal;
end;


// ����� ������
procedure ShowError(msg: String);
begin
  MessageBeep(MB_ICONHAND);
  Show(msg, mtError, [mbOK], [BUTTON_OK], '������!');
end;

// ����� ��������������
procedure ShowWarning(msg: String);
begin
  Show(msg, mtWarning, [mbOK], [BUTTON_OK], '��������!');
end;

// ����� ����������
procedure ShowInfo(msg: String);
begin
  Show(msg, mtInformation, [mbOK], [BUTTON_OK], '����������');
end;

// ����� �������������
function ShowConfirm(msg: String): Boolean;
begin
  MessageBeep(MB_ICONEXCLAMATION);
  Result := Show(msg, mtConfirmation, [mbYes, mbNo], [BUTTON_YES, BUTTON_CANCEL], '�������������') = 6;
end;

end.
