unit MessagesUnit;

interface

uses Forms, StdCtrls, Classes, Dialogs, SysUtils, Windows;

const
  FAILED_CONNECT_TO_DATABASE  = 'Не удалось подключиться к базе данных!';
  FAILED_CREATE_CONTEXT       = 'Не удалось создать контекст!';
  FAILED_REQUEST_TABLE        = 'Не удалось получить таблицу!';
  FAILED_REQUEST_COLUMN_NAMES = 'Не удалось получить заголовки столбцов!';
  FAILED_SQL_REQUEST          = 'Не удалось выполнить запрос!';
  FAILED_INSERT_RECORD        = 'Не удалось добавить запись!';

  SUCCESSIVELY_CONNECT        = 'Подключение прошло успешно!';
  SUCCESSIVELY_TEST_CONNECT   = 'Проверка подключения к базе данных прошла успешно!';

  SQL_SELECT_ALL              = 'SELECT * FROM ';
  SQL_DELETE_ALL              = 'DELETE * FROM ';

  CONFIRM_EXIT                = 'Вы уверены что хотите выйти?';
  CONFIRM_DELETE              = 'Вы уверены что хотите удалить запись?';

  WARNING_FIELDS_IS_EMPTY     = 'Внимение! Одно или несколько обязательных полей не были заполнены!';

  ERROR_RECORD_NOT_SELECTED   = 'Ошибка! Ни одна запись не была выделена!';

  BUTTON_CANCEL = 'Отмена';
  BUTTON_YES    = 'Да';
  BUTTON_NO     = 'Нет';
  BUTTON_OK     = 'Ок';

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


// Вывод ошибки
procedure ShowError(msg: String);
begin
  MessageBeep(MB_ICONHAND);
  Show(msg, mtError, [mbOK], [BUTTON_OK], 'Ошибка!');
end;

// Вывод предупреждения
procedure ShowWarning(msg: String);
begin
  Show(msg, mtWarning, [mbOK], [BUTTON_OK], 'Внимание!');
end;

// Вывод информации
procedure ShowInfo(msg: String);
begin
  Show(msg, mtInformation, [mbOK], [BUTTON_OK], 'Информация');
end;

// Вывод подтверждения
function ShowConfirm(msg: String): Boolean;
begin
  MessageBeep(MB_ICONEXCLAMATION);
  Result := Show(msg, mtConfirmation, [mbYes, mbNo], [BUTTON_YES, BUTTON_CANCEL], 'Подтверждение') = 6;
end;

end.
