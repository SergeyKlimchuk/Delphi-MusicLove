unit MainFormUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, RepositoryUnit, IBQuery, DB, IBDatabase, ComCtrls,
  jpeg, StdCtrls, MessagesUnit, Grids, DBGrids, Clipbrd, ComObj;

type
  TMainForm = class(TForm)
    imgs: TImage;
    pgcTables: TPageControl;
    tsSounds: TTabSheet;
    stat1: TStatusBar;
    tsAlbums: TTabSheet;
    tsGroups: TTabSheet;
    tsMusicians: TTabSheet;
    lbl2: TLabel;
    img1: TImage;
    lbl1: TLabel;
    dsMusicians: TDataSource;
    dbgrdSounds: TDBGrid;
    pnlBackTools: TPanel;
    pnlTools: TPanel;
    btnDelete: TButton;
    btnEdit: TButton;
    btnAdd: TButton;
    dsAlbums: TDataSource;
    dsGroups: TDataSource;
    dsSounds: TDataSource;
    pnlSoundsSql: TPanel;
    lbledtSoundsName: TLabeledEdit;
    lbl3: TLabel;
    cbbSoundsAlbum: TComboBox;
    dtpSoundTime: TDateTimePicker;
    lbl5: TLabel;
    btnSoundsCancel: TButton;
    btnSoundsEnter: TButton;
    bvl1: TBevel;
    lblSoundsCaption: TLabel;
    dbgrdAlbums: TDBGrid;
    dbgrdGroups: TDBGrid;
    dbgrdMusicians: TDBGrid;
    pnlAlbumsSql: TPanel;
    lbl4: TLabel;
    bvl2: TBevel;
    lblAlbumsCaption: TLabel;
    lbledtAlbumsName: TLabeledEdit;
    cbbAlbumsGroup: TComboBox;
    btnAlbumsCancel: TButton;
    btnAlbumsEnter: TButton;
    lbl6: TLabel;
    mmoAlbumsInfo: TMemo;
    pnlGroupsSql: TPanel;
    bvl3: TBevel;
    lblGroupsCaption: TLabel;
    lbl9: TLabel;
    lbledtGroupsName: TLabeledEdit;
    btnGroupsCancel: TButton;
    btnGroupsEnter: TButton;
    mmoGroupsInfo: TMemo;
    pnlMusiciansSql: TPanel;
    bvl4: TBevel;
    lblMusiciansCaption: TLabel;
    lbledtMusiciansName: TLabeledEdit;
    btnMusiciansCancel: TButton;
    btnMusiciansEnter: TButton;
    cbbMusiciansSex: TComboBox;
    lbl8: TLabel;
    lbl11: TLabel;
    dtpMusiciansDate: TDateTimePicker;
    lbl7: TLabel;
    cbbMusiciansGroup: TComboBox;
    lbl10: TLabel;
    btn1: TButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure pgcTablesChange(Sender: TObject);
    // Общие
    function CheckIdOnUnique(TableName: string; IdentityColumn: string; Indx: Integer): Boolean;
    function CheckFieldOnUnique(TableName: string; IdentityColumn: string; val: String): Boolean;
    function GetFieldsById(TableName: string; IdentityColumn: string; Id: Variant): TStrings;
    function GetUniqueId(TableName: string; IdentityColumn: string): Integer;
    function GetCurrentRecordID(): Integer;
    procedure btnDeleteClick(Sender: TObject);
    procedure FillClearColor(Sender: TObject);
    procedure LockCurrentPage(Value: Boolean = False);
    procedure UnLockCurrentPage();
    procedure FormResize(Sender: TObject);
    procedure pgcTablesChanging(Sender: TObject; var AllowChange: Boolean);
    procedure CreateReportForCurrentTable(Sender: TObject);
    // Музыка
    procedure ShowSoundAddPanel(Sender: TObject);
    procedure ShowSoundEditPanel(Sender: TObject);
    procedure btnSoundsPanelCancelClick(Sender: TObject);
    procedure btnSoundsPanelEnterClick(Sender: TObject);
    // Альбомы
    procedure ShowAlbumsAddPanel(Sender: TObject);
    procedure ShowAlbumsEditPanel(Sender: TObject);
    procedure btnAlbumsPanelCancelClick(Sender: TObject);
    procedure btnAlbumsPanelEnterClick(Sender: TObject);
    // Группы
    procedure ShowGroupsAddPanel(Sender: TObject);
    procedure ShowGroupsEditPanel(Sender: TObject);
    procedure btnGroupsCancelClick(Sender: TObject);
    procedure btnGroupsEnterClick(Sender: TObject);
    // Музыканты
    procedure ShowMusiciansAddPanel(Sender: TObject);
    procedure ShowMusiciansEditPanel(Sender: TObject);
    procedure btnMusiciansCancelClick(Sender: TObject);
    procedure btnMusiciansEnterClick(Sender: TObject);
  private
    Repository: TRepository;
  public
    procedure ShowWithRepo(Sender: TObject; Repository: TRepository);
  end;

var
  MainForm: TMainForm;

implementation

uses DateUtils;

{$R *.dfm}


// ОБЩЕЕ -----------------------------------------------------------------------

// Форматируем запрос на вывод всего содержимого
procedure SetSelectAllCommand(Query: TIBQuery; TableName: string);
begin
  Query.SQL.Clear;
  
  if (TableName = 'SOUNDS') then
  begin
    Query.SQL.Add(
      'SELECT SOUNDS.ID AS ID, SOUNDS.NAME AS SOUNDNAME, ' +
      'ALBUMS.NAME AS ALBUMNAME, GROUPS.NAME AS GROUPNAME, SOUNDS.PLAYTIME ' +
      'FROM SOUNDS ' +
      'LEFT JOIN ALBUMS ON SOUNDS.ALBUMID = ALBUMS.ID ' +
      'LEFT JOIN GROUPS ON ALBUMS.GROUPID = GROUPS.ID');
  end
  else if (TableName = 'ALBUMS') then
  begin
    Query.SQL.Add('SELECT ALBUMS.ID AS ID, ALBUMS.NAME AS ALBUM_NAME,' +
      'GROUPS.NAME AS GROUP_NAME, ALBUMS.INFO AS ALBUM_INFO ' +
      'FROM ALBUMS ' +
      'LEFT JOIN GROUPS ON ALBUMS.GROUPID = GROUPS.ID');
  end
  else if (TableName = 'GROUPS') then
  begin
    Query.SQL.Add('SELECT ID AS ID,NAME AS GROUP_NAME, INFO AS GROUP_INFO ' +
      'FROM GROUPS');
  end
  else if (TableName = 'MUSICIANS') then
  begin
    Query.SQL.Add('SELECT MUSICIANS.ID AS ID, MUSICIANS.NAME AS MUSICIAN_NAME, ' +
      'GROUPS.NAME AS GROUP_NAME, AGE AS MUSICIAN_AGE, SEX AS MUSICIAN_SEX ' +
      'FROM MUSICIANS ' +
      'LEFT JOIN GROUPS ON GROUPS.ID = MUSICIANS.GROUPID')
  end
  else
  // Временная замена
  Query.SQL.Add('SELECT DISTINCT * FROM ' + TableName);
end;

// Форматируем запрос на добавление
procedure SetInsertCommand(Query: TIBQuery; TableName: string;
  ParamsNames: array of String; ParamsValues: Array of Variant);
var
  i: Integer;
begin
  With Query do
  begin
    SQL.Clear;
    SQL.Add('INSERT INTO ' + TableName);
    SQL.Add('('); // Навзания полей
    SQL.Add('VALUES('); // Значения полей
    i := 0;
    Params.Clear;
    // Вносим параметры
    while (i < Length(ParamsNames)) do
    begin
      // Разделитель между данными
      if (i > 0) then
      begin
        SQL[1] := SQL[1] + ', ';
        SQL[2] := SQL[2] + ', ';
      end;

      SQL[1] := SQL[1] + ParamsNames[i];
      SQL[2] := SQL[2] + ':' + ParamsNames[i];
      Params.ParamByName(ParamsNames[i]).Value := ParamsValues[i];

      Inc(i);
    end;
    SQL[1] := SQL[1] + ')';
    SQL[2] := SQL[2] + ')';
  end;
end;

// Форматируем запрос на обнавление записи
procedure SetUpdateCommand(Query: TIBQuery; TableName: string;
  ParamsNames: array of String; ParamsValues: Array of Variant; ID: Integer);
var
  i: Integer;
begin
  With Query do
  begin
    SQL.Clear;
    SQL.Add('UPDATE ' + TableName);
    SQL.Add('SET ');
    i := 0;
    Params.Clear;
    // Вносим параметры
    while (i < Length(ParamsNames)) do
    begin
      // Разделитель между данными
      if (i > 0) then
        SQL[1] := SQL[1] + ', ';

      SQL[1] := SQL[1] + ParamsNames[i] + '=:' + ParamsNames[i];
      Params.ParamByName(ParamsNames[i]).Value := ParamsValues[i];
      Inc(i);
    end;
    SQL.Add('WHERE ID=:ID');
    ParamByName('ID').AsInteger := ID;
  end;
end;

// Форматируем запрос на добавление
procedure SetDeleteCommand(Query: TIBQuery; TableName: String; ID: Integer);
begin
  with Query do
  begin
    SQL.Clear;
    SQL.Add('DELETE FROM ' + TableName + ' WHERE ID=:ID');
    ParamByName('ID').AsInteger := ID;
  end;
end;

// Функция проверки компонента на пустоту
function InputIsEmpty(Obj: TComponent): Boolean;
var
  BadColor: TColor;
  GoodColor: TColor;
begin
  BadColor := TColor($8888FF);
  GoodColor := TColor($88FF88);
  if (Obj is TLabeledEdit) then
  begin
    Result := (Obj as TLabeledEdit).Text = '';
    if (Result) then
      (Obj as TLabeledEdit).Color := BadColor
    else
      (Obj as TLabeledEdit).Color := GoodColor;
    Exit;
  end;

  if (Obj is TDateTimePicker) then
  begin
    Result := TimeToStr((Obj as TDateTimePicker).Time) = '0:00:00';
    if (Result) then
      (Obj as TDateTimePicker).Color := BadColor
    else
      (Obj as TDateTimePicker).Color := GoodColor;
    Exit;
  end;

  if (Obj is TMemo) then
  begin
    Result := ((Obj as TMemo).Lines.Count = 0) or ((Obj as TMemo).Lines[0] = '');
    if (Result) then
      (Obj as TMemo).Color := BadColor
    else
      (Obj as TMemo).Color := GoodColor;
    Exit;
  end;

  Result := False;
  ShowError('Функция "InputIsEmpty" приняла не верный формат поля!');
  Abort;
end;

// Получаем идентификатор текущей записи
function TMainForm.GetCurrentRecordID(): Integer;
var
  DataSource: TDataSource;
begin
  // Выбираем распределитель
  case pgcTables.ActivePageIndex of
    0:DataSource := dsAlbums;
    1:DataSource := dsGroups;
    2:DataSource := dsMusicians;
    3:DataSource := dsSounds;
    else
    begin
      ShowError('Индекс таблицы вишел за границы!');
      Result := -1;
      Exit;
    end;
  end;
  // проверяем привязку источника данных
  if (DataSource.DataSet <> nil) then
    Result := DataSource.DataSet.FieldByName('ID').AsInteger
  else
    Result := -1;
end;

// Проверяет иденификатор в таблице на уникальность
function TMainForm.CheckIdOnUnique(TableName: string;
  IdentityColumn: string; Indx: Integer): Boolean;
var
  Query: TIBQuery;
begin
  // осуществляем поиск по идентификатору
  Query := Self.Repository.GetQuery(Self);
  Query.SQL.Add('SELECT ' + IdentityColumn + ' FROM ' + TableName +
    ' WHERE ' + IdentityColumn + '=:IDD');
  Query.Params[0].AsInteger := Indx;
  Query.Active;

  // Проверяем на кол-во записей по запросу
  Result := (Query.RecordCount = 0);
  Query.Free;
end;

function TMainForm.CheckFieldOnUnique(TableName: string;
  IdentityColumn: string; val: String): Boolean;
var
  Query: TIBQuery;
begin
  Query := Repository.GetQuery(Self);
  Query.SQL.Add(
    'SELECT * FROM ' + TableName + ' WHERE ' + IdentityColumn + '=''' + val + ''''
  );
  Query.Active := True;
  // Проверяем на кол-во записей по запросу
  Result := (Query.RecordCount = 0);
end;

// Делает не активными ключевые компоненты текущего меню
procedure TMainForm.LockCurrentPage(Value: Boolean = False);
var
  DBGrid: TDBGrid;
begin

  case pgcTables.TabIndex of
    0:DBGrid := dbgrdAlbums;
    1:DBGrid := dbgrdGroups;
    2:DBGrid := dbgrdMusicians;
    3:DBGrid := dbgrdSounds;
    else
    begin
      DBGrid := TDBGrid.Create(Self);
      ShowError('Значение индекса вкладки вышло за ожидаемые границы!');
      Abort;
    end;
  end;

  DBGrid.Enabled := Value;
  DBGrid.Refresh;
  btnDelete.Enabled := Value;
end;

// Делает активными ключевые компоненты текущего меню
procedure TMainForm.UnLockCurrentPage();
begin
  LockCurrentPage(True);
end;

// Возвращает поля по идентификатору записи
function TMainForm.GetFieldsById(TableName: string;
  IdentityColumn: string; Id: Variant): TStrings;
var
  Query: TIBQuery;
  ColIndx: Integer;
begin
  // Строим запрос для получения полей
  Query := Self.Repository.GetQuery(Self);
  SetSelectAllCommand(Query, TableName);
  Query.SQL[0] := Query.SQL[0] + ' WHERE ' + IdentityColumn + '=:IDD';
  Query.ParamByName('IDD').AsString := VarTostr(Id);
  Query.Active := True;

  // Если есть записи то взвращаем набор полей
  if (Query.RecordCount > 0) then
  begin
    Result := TStringList.Create;
    for ColIndx := 0 to Query.FieldCount - 1 do
      Result.Add(Query.Fields[ColIndx].AsString);
  end
  else
    Result := nil;

  Query.Free;
end;

// Получаем уникальный идентификатор из таблиицы
function TMainForm.GetUniqueId(TableName: string;
  IdentityColumn: string): Integer;
var
  Query: TIBQuery;
begin
  Query := Self.Repository.GetQuery(Self);
  Randomize;
  Result := Random(1000000);

  // Генерируем новое значение каждый раз когда оно уже находится в базе
  while (not CheckIdOnUnique(TableName, IdentityColumn, Result)) do
    Result := Random(1000000);

  Query.Free;
end;

// Заполняет список данными
procedure FillInputList(Query: TIBQuery; ComboBox: TComboBox;
  TableName: string; ColumnName: String);
begin
  // Берем все записи
  SetSelectAllCommand(Query, TableName);
  Query.Active := True;
  // Очищаем прошлые данные
  ComboBox.Clear;
  // Проверяем результат на пустоту
  ComboBox.Enabled := (Query.RecordCount > 0);
  // Поочередно запихиваем их в инпут
  While (not Query.Eof) do
  begin
    ComboBox.Items.Add(Query.FieldByName(ColumnName).AsString);
    Query.Next;
  end;
  // Если есть записи - ставим первую в списке
  if (ComboBox.Enabled) then
    ComboBox.ItemIndex := 0;
end;

// Заполняет список данными и устанавливаем фокус на нужном элементе
procedure FillInputListAndSelect(Query: TIBQuery; ComboBox: TComboBox;
  Tablename: string; FieldName: String; SelectedValue: String);
begin
  // Берем все записи
  SetSelectAllCommand(Query, Tablename);
  Query.Active := True;
  // Очищаем прошлые данные
  ComboBox.Clear;
  // Проверяем результат на пустоту
  ComboBox.Enabled := (Query.RecordCount > 0);
  // Поочередно запихиваем их в инпут
  While (not Query.Eof) do
  begin
    // Добавили объект в список
    ComboBox.Items.Add(Query.FieldByName(FieldName).AsString);
    // Если совпал в выбранным то выставляем на него фокус
    if (Query.FieldByName(FieldName).AsString = SelectedValue) then
      ComboBox.ItemIndex := ComboBox.Items.Count -1;
    Query.Next;
  end;
  // Если есть записи - ставим первую в списке
  if (ComboBox.Enabled) then
    ComboBox.ItemIndex := 0;
end;

// Настройка действий кнопок
procedure TMainForm.pgcTablesChange(Sender: TObject);
var
  Query: TIBQuery;
  DataSource: TDataSource;
begin
  case pgcTables.TabIndex of
    0:
    begin
      DataSource := dsAlbums;
      // Устанавливаем ивенты
      btnAdd.OnClick := ShowAlbumsAddPanel;
      btnEdit.OnClick := ShowAlbumsEditPanel;
      btnDelete.Enabled := (not pnlAlbumsSql.Visible);
    end;
    1:
    begin
      DataSource := dsGroups;
      // Устанавливаем ивенты
      btnAdd.OnClick := ShowGroupsAddPanel;
      btnEdit.OnClick := ShowGroupsEditPanel;
      btnDelete.Enabled := (not pnlGroupsSql.Visible);
    end;
    2:
    begin
      DataSource := dsMusicians;
      // Устанавливаем ивенты
      btnAdd.OnClick := ShowMusiciansAddPanel;
      btnEdit.OnClick := ShowMusiciansEditPanel;
      btnDelete.Enabled := (not pnlMusiciansSql.Visible);
    end;
    3:
    begin
      DataSource := dsSounds;
      // Устанавливаем ивенты
      btnAdd.OnClick := ShowSoundAddPanel;
      btnEdit.OnClick := ShowSoundEditPanel;
      btnDelete.Enabled := (not pnlSoundsSql.Visible);
    end;
    else
    begin
      DataSource := TDataSource.Create(Self);
      ShowError('Значение индекса вкладки вышло за ожидаемые границы!');
      Abort;
    end;
  end;
  // Если таблица не была заполнена
  if (DataSource.DataSet = nil) then
  begin
    // Получаем запросник для новой странички
    Query := Repository.GetQuery(Self);
    // Вставляем команду выборки всех элементов
    SetSelectAllCommand(Query,
      Repository.GetTablesNames()[pgcTables.ActivePageIndex]);
    // Связываем таблицу и элемент отображения
    DataSource.DataSet := Query;
  end;

  DataSource.DataSet.Active := False;
  DataSource.DataSet.Active := True;
end;

// Отображение с настройками
procedure TMainForm.ShowWithRepo(Sender: TObject; Repository: TRepository);
begin
  // Переносим репозиторий данных
  Self.Repository := Repository;
  Self.Repository.Connect;
  // Формируем первую страницу
  pgcTablesChange(Self);
  // Скрываем старую форму
  (Sender as TForm).Hide;
  // Отображаем новоую форму полностью
  Show();
end;

// Закртыие формы
procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    // Подтверждение выхода
  if (ShowConfirm(CONFIRM_EXIT)) then
    ExitProcess(0)
  else
    Abort;
end;

// Инициализация формы
procedure TMainForm.FormCreate(Sender: TObject);
begin
  Self.DoubleBuffered := True;
end;

// Кнопка на панели инструментов Удаления записи
// Для всех таблиц одна -> хороший алгоритм :)
procedure TMainForm.btnDeleteClick(Sender: TObject);
var
  Query: TIBQuery;
  TabIndex: Integer;
  TableName: String;
  DataSource: TDataSource;
begin
  // Получаем номер таблицы
  TabIndex := pgcTables.TabIndex;
  // Определяем откдуа будем удалять
  case TabIndex of
    0:DataSource := dsAlbums;
    1:DataSource := dsGroups;
    2:DataSource := dsMusicians;
    3:DataSource := dsSounds;
    else
    begin
      DataSource := TDataSource.Create(Self);
      ShowError('Значение индекса вкладки вышло за ожидаемые границы!');
      Abort;
    end;
  end;
  // Если в таблице ничего нет, то оповестим пользователя
  if (DataSource.DataSet.RecordCount = 0) then
  begin
    ShowWarning('Отсутствуют записи для удаления!');
    Abort;
  end;

  // Просим подтверждение удаления
  if (not ShowConfirm(CONFIRM_DELETE)) then
    Abort;
  // Получаем название таблицы
  TableName := Repository.GetTablesNames()[TabIndex];
  Query := Repository.GetQuery(Self);
  // Устанавливаем и выполняем команду удаления
  SetDeleteCommand(Query,
    TableName,
    DataSource.DataSet.FieldByName('ID').AsInteger);
  Query.ExecSQL;
  Repository.SaveChanges;

  // Обновляем текущую таблицу
  DataSource.DataSet.Active := True;
end;

// Заполняет текущщий компонент белым цветом
procedure TMainForm.FillClearColor(Sender: TObject);
begin
  if (Sender is TLabeledEdit) then
    (Sender as TLabeledEdit).Color := clWindow;

  if (Sender is TDateTimePicker) then
    (Sender as TDateTimePicker).Color := clWindow;
end;


// МУЗЫКА ----------------------------------------------------------------------

// Кнопка на панели инструментов Добавления записи
procedure TMainForm.ShowSoundAddPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // Очизаем прошлые данные
  btnSoundsPanelCancelClick(Self);
  // Заполняем выпадающий список с альбомами
  Query := Repository.GetQuery(Self);
  FillInputList(Query, cbbSoundsAlbum, 'ALBUMS', 'ALBUM_NAME');
  Query.Free;
  // Скрываем кнопку
  // Только если отсутствуют данные в выпадающих списках
  btnSoundsEnter.Enabled := cbbSoundsAlbum.Enabled;
  // Устанавливаем режим добавления
  pnlSoundsSql.Tag := 0;
  lblSoundsCaption.Caption := 'Добавление новой музыки';
  btnSoundsEnter.Caption := 'Добавить';
  // отображение панели
  pnlSoundsSql.Show;
  LockCurrentPage();
end;

// Кнопка на панели инструментов Редактирования записи
procedure TMainForm.ShowSoundEditPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // Если в таблице ничего нет, то оповестим пользователя
  if (dsSounds.DataSet.RecordCount = 0) then
  begin
    ShowWarning('Отсутствуют записи для редактирования!');
    Abort;
  end;
  // Заполняем выпадающий список с альбомами
  // и выбираем там выбранный
  Query := Repository.GetQuery(Self);
  FillInputListAndSelect(Query, cbbSoundsAlbum,
    'ALBUMS',
    'ALBUM_NAME',
    dbgrdSounds.Columns[2].Field.AsString);
  Query.Free;
  // Устанавливаем название музыки
  lbledtSoundsName.Text := dbgrdSounds.Columns[0].Field.AsString;
  // Устанавливаем время музыки
  dtpSoundTime.Time := dbgrdSounds.Columns[3].Field.AsVariant;
  // Скрываем кнопку
  // Только если отсутствуют данные в выпадающих списках
  btnSoundsEnter.Enabled := cbbSoundsAlbum.Enabled;
  // Устанавливаем режим добавления
  pnlSoundsSql.Tag := 1;
  // Меняем текста на панели
  lblSoundsCaption.Caption := 'Редактирование музыки';
  btnSoundsEnter.Caption := 'Применить';
  // отображение панели
  pnlSoundsSql.Show;
  LockCurrentPage();
end;

// Закрытие формы (добавления / редактирования)
procedure TMainForm.btnSoundsPanelCancelClick(Sender: TObject);
begin
  // Очищаем название музыки
  lbledtSoundsName.Text := '';
  // Скрываем панель
  pnlSoundsSql.Hide;
  // Выставляем минимальное время музыки
  dtpSoundTime.Time := StrToTime('00:00:01');
  // Очищаем цвета у полей
  lbledtSoundsName.Color := clWindow;
  dtpSoundTime.Color := clWindow;
  UnLockCurrentPage();
end;
// Применение действий формы (добавления / редактирования)
procedure TMainForm.btnSoundsPanelEnterClick(Sender: TObject);
var
  Query: TIBQuery;
  AlbumId: Integer;
begin
  Query := Repository.GetQuery(Self);
  // Проверка полей
  if (InputIsEmpty(lbledtSoundsName)) then
  begin
    ShowError('Поле с названием песни не может быть пустым!');
    Abort;
  end;
  if (InputIsEmpty(dtpSoundTime)) then
  begin
    ShowError('Поле с временем песни не может быть равным нулю!');
    Abort;
  end;
  // Если стоит режим добавления
  if (pnlSoundsSql.Tag = 0) then
  begin
    // Проверяем на уникальность запись
    if (not CheckFieldOnUnique('SOUNDS', 'NAME', lbledtSoundsName.Text)) then
    begin
      lbledtSoundsName.Color := TColor($88FFFF);
      ShowError('Название музыки не уникально!');
      Abort;
    end;
    // Получаем идентификатор альбома
    AlbumId := StrToInt(GetFieldsById('ALBUMS', 'ALBUMS.NAME', cbbSoundsAlbum.Text)[0]);
    // Формируем команду на вставку
    SetInsertCommand(Query,
      'SOUNDS',
      ['ID', 'ALBUMID', 'NAME', 'PLAYTIME'],
      [GetUniqueId('SOUNDS', 'ID'),
        AlbumId,
        lbledtSoundsName.Text,
        TimeToStr(dtpSoundTime.Time)]
      );
    // Выполняем сформированный запрос
    Query.ExecSQL;
    Repository.SaveChanges;
    SetSelectAllCommand(Query, 'SOUNDS');
    Query.Active := True;
  end
  else
  // Если стоит режим обновления
  begin
    // Получаем идентификатор альбома
    AlbumId := StrToInt(GetFieldsById('ALBUMS', 'ALBUMS.NAME', cbbSoundsAlbum.Text)[0]);
    SetUpdateCommand(Query,
      'SOUNDS',
      ['ALBUMID', 'NAME', 'PLAYTIME'],
      [AlbumId,lbledtSoundsName.Text, TimeToStr(dtpSoundTime.Time)],
      GetCurrentRecordID());

    // Выполняем сформированный запрос
    Query.ExecSQL;
    Repository.SaveChanges;
    Query.Free;
  end;
  // Перезагружаем данные из таблицы
  dsSounds.DataSet.Active := True;
  // Закрытие панели
  btnSoundsCancel.Click;
end;

// АЛЬБОМЫ ---------------------------------------------------------------------

// Кнопка на панели инструментов Добавления записи
procedure TMainForm.ShowAlbumsAddPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // Заполняем выпадающий список с группами
  Query := Repository.GetQuery(Self);
  FillInputList(Query, cbbAlbumsGroup, 'GROUPS', 'GROUP_NAME');
  Query.Free;
  // Скрываем кнопку
  // Только если отсутствуют данные в выпадающих списках
  btnAlbumsEnter.Enabled := cbbAlbumsGroup.Enabled;
  // Устанавливаем режим добавления
  pnlAlbumsSql.Tag := 0;
  lblAlbumsCaption.Caption := 'Добавление нового альбома';
  btnAlbumsEnter.Caption := 'Добавить';
  // отображение панели
  pnlAlbumsSql.Show;
  LockCurrentPage();
end;

// Кнопка на панели инструментов Режактировать запись
procedure TMainForm.ShowAlbumsEditPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // Если в таблице ничего нет, то оповестим пользователя
  if (dsAlbums.DataSet.RecordCount = 0) then
  begin
    ShowWarning('Отсутствуют записи для редактирования!');
    Abort;
  end;
  // Заполняем выпадающий список с альбомами
  // и выбираем там выбранный
  Query := Repository.GetQuery(Self);
  FillInputListAndSelect(Query, cbbAlbumsGroup,
    'GROUPS',
    'GROUP_NAME',
    dbgrdAlbums.Columns[1].Field.AsString);
  Query.Free;
  // Устанавливаем название альбома
  lbledtAlbumsName.Text := dbgrdAlbums.Columns[0].Field.AsString;
  // Устанавливаем описание альбома
  mmoAlbumsInfo.Lines.Add(dbgrdAlbums.Columns[2].Field.AsString);
  // Делаем не активной кнопку
  // Только если отсутствуют данные в выпадающих списках
  btnAlbumsEnter.Enabled := cbbAlbumsGroup.Enabled;
  // Устанавливаем режим добавления
  pnlAlbumsSql.Tag := 1;
  // Меняем текста на панели
  lblAlbumsCaption.Caption := 'Редактирование альбома';
  btnAlbumsEnter.Caption := 'Применить';
  // отображение панели
  pnlAlbumsSql.Show;
  LockCurrentPage();
end;

// Закрытие панели
procedure TMainForm.btnAlbumsPanelCancelClick(Sender: TObject);
begin
  // Сбрасываем текста
  lbledtAlbumsName.Text := '';
  cbbAlbumsGroup.Items.Clear;
  mmoAlbumsInfo.Clear;
  // Сбрасываем цвета
  lbledtAlbumsName.Color := clWindow;
  cbbAlbumsGroup.Color := clWindow;
  mmoAlbumsInfo.Color := clWindow;
  // Скрываем панель
  pnlAlbumsSql.Hide;
  UnLockCurrentPage();
end;

// Действие на принятие в панели
procedure TMainForm.btnAlbumsPanelEnterClick(Sender: TObject);
var
  Query: TIBQuery;
  GroupId: Integer;
begin
  Query := Repository.GetQuery(Self);
  // Проверка полей
  if (InputIsEmpty(lbledtAlbumsName)) then
  begin
    ShowError('Поле с названием альбома не может быть пустым!');
    Abort;
  end;
  if (InputIsEmpty(mmoAlbumsInfo)) then
  begin
    ShowError('Поле с информацией о альбоме не может быть пустым!!');
    Abort;
  end;
  // Если стоит режим добавления
  if (pnlAlbumsSql.Tag = 0) then
  begin
    // Проверяем на уникальность запись
    if (not CheckFieldOnUnique('ALBUMS', 'NAME', lbledtAlbumsName.Text)) then
    begin
      lbledtAlbumsName.Color := TColor($88FFFF);
      ShowError('Название альбома не уникально!');
      Abort;
    end;
    // Получаем идентификатор группы
    GroupId := StrToInt(GetFieldsById('GROUPS', 'GROUPS.NAME', cbbAlbumsGroup.Text)[0]);
    // Формируем команду на вставку
    SetInsertCommand(Query,
      'ALBUMS',
      ['ID',                        'GROUPID', 'NAME',               'INFO'],
      [GetUniqueId('ALBUMS', 'ID'), GroupId,   lbledtAlbumsName.Text, mmoAlbumsInfo.Lines.Text]
      );
    // Выполняем сформированный запрос и сохраняем
    Query.ExecSQL;
    Repository.SaveChanges;
  end
  else
  // Если стоит режим обновления
  begin
    // Получаем идентификатор группы
    GroupId := StrToInt(GetFieldsById('GROUPS', 'GROUPS.NAME', cbbAlbumsGroup.Text)[0]);
    // Формируем команду на обнавление
    SetUpdateCommand(Query,
      'ALBUMS',
      ['GROUPID', 'NAME',                'INFO'],
      [GroupId,   lbledtAlbumsName.Text, mmoAlbumsInfo.Lines.Text],
      GetCurrentRecordID());
    // Выполняем сформированный запрос и сохраняем
    Query.ExecSQL;
    Repository.SaveChanges;
  end;
  // Очищаем память
  Query.Free;
  // Перезагружаем данные из таблицы
  dsAlbums.DataSet.Active := True;
  // Очищаем форму
  btnAlbumsPanelCancelClick(Sender);
end;

// ГРУППЫ ----------------------------------------------------------------------

// Кнопка на панели инструментов Добавления записи
procedure TMainForm.ShowGroupsAddPanel(Sender: TObject);
begin
  // Устанавливаем режим добавления
  pnlGroupsSql.Tag := 0;
  lblGroupsCaption.Caption := 'Добавление новоую группу';
  btnGroupsEnter.Caption := 'Добавить';
  // отображение панели
  pnlGroupsSql.Show;
  LockCurrentPage();
end;

// Кнопка на панели инструментов Режактировать запись
procedure TMainForm.ShowGroupsEditPanel(Sender: TObject);
begin
  // Если в таблице ничего нет, то оповестим пользователя
  if (dsGroups.DataSet.RecordCount = 0) then
  begin
    ShowWarning('Отсутствуют записи для редактирования!');
    Abort;
  end;
  // Устанавливаем название альбома
  lbledtGroupsName.Text := dbgrdGroups.Columns[0].Field.AsString;
  // Устанавливаем описание альбома
  mmoGroupsInfo.Lines.Add(dbgrdGroups.Columns[1].Field.AsString);
  // Устанавливаем режим добавления
  pnlGroupsSql.Tag := 1;
  // Меняем текста на панели
  lblGroupsCaption.Caption := 'Редактирование группы';
  btnGroupsEnter.Caption := 'Применить';
  // отображение панели
  pnlGroupsSql.Show;
  LockCurrentPage();
end;


// Закрытие панели
procedure TMainForm.btnGroupsCancelClick(Sender: TObject);
begin
  // Сбрасываем текста
  lbledtGroupsName.Text := '';
  mmoGroupsInfo.Lines.Clear;
  // Сбрасываем цвета
  lbledtGroupsName.Color := clWindow;
  mmoGroupsInfo.Color := clWindow;
  // Скрываем панель
  pnlGroupsSql.Hide;
  UnLockCurrentPage();
end;

// Действие на принятие в панели
procedure TMainForm.btnGroupsEnterClick(Sender: TObject);
var
  Query: TIBQuery;
begin
  Query := Repository.GetQuery(Self);
  // Проверка полей
  if (InputIsEmpty(lbledtGroupsName)) then
  begin
    ShowError('Поле с названием группы не может быть пустым!');
    Abort;
  end;
  if (InputIsEmpty(mmoGroupsInfo)) then
  begin
    ShowError('Поле с информацией о группе не может быть пустым!!');
    Abort;
  end;
  // Если стоит режим добавления
  if (pnlGroupsSql.Tag = 0) then
  begin
    // Проверяем на уникальность запись
    if (not CheckFieldOnUnique('GROUPS', 'NAME', lbledtGroupsName.Text)) then
    begin
      lbledtGroupsName.Color := TColor($88FFFF);
      ShowError('Название группы не уникально!');
      Abort;
    end;
    // Формируем команду на вставку
    SetInsertCommand(Query,
      'GROUPS',
      ['ID',                        'NAME',                'INFO'],
      [GetUniqueId('GROUPS', 'ID'), lbledtGroupsName.Text, mmoGroupsInfo.Lines.Text]
      );
    // Выполняем сформированный запрос и сохраняем
    Query.ExecSQL;
    Repository.SaveChanges;
  end
  else
  // Если стоит режим обновления
  begin
    // Формируем команду на обнавление
    SetUpdateCommand(Query,
      'GROUPS',
      ['NAME',                'INFO'],
      [lbledtGroupsName.Text, mmoGroupsInfo.Lines.Text],
      GetCurrentRecordID());
    // Выполняем сформированный запрос и сохраняем
    Query.ExecSQL;
    Repository.SaveChanges;
  end;
  // Очищаем память
  Query.Free;
  // Перезагружаем данные из таблицы
  dsGroups.DataSet.Active := True;
  // Очищаем форму
  btnGroupsCancelClick(Sender);
end;

// МУЗЫКАНТЫ -------------------------------------------------------------------

// Закрытие панели
procedure TMainForm.btnMusiciansCancelClick(Sender: TObject);
begin
  // обнуляем поля
  lbledtMusiciansName.Text := '';
  dtpMusiciansDate.Date := Now;
  cbbMusiciansSex.ItemIndex := 0;
  cbbMusiciansGroup.Items.Clear;
  pnlMusiciansSql.Hide;
  UnLockCurrentPage();
end;

// Кнопка на панели инструментов Добавления записи
procedure TMainForm.ShowMusiciansAddPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // Заполняем выпадающий список с группами
  Query := Repository.GetQuery(Self);
  FillInputList(Query, cbbMusiciansGroup, 'GROUPS', 'GROUP_NAME');
  Query.Free;
  // Устанавливаем режим добавления
  pnlMusiciansSql.Tag := 0;
  lblMusiciansCaption.Caption := 'Добавление нового музыканта';
  btnMusiciansEnter.Caption := 'Добавить';
  // отображение панели
  pnlMusiciansSql.Show;
  LockCurrentPage();
end;

// Кнопка на панели инструментов Режактировать запись
procedure TMainForm.ShowMusiciansEditPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // Если в таблице ничего нет, то оповестим пользователя
  if (dsMusicians.DataSet.RecordCount = 0) then
  begin
    ShowWarning('Отсутствуют записи для редактирования!');
    Abort;
  end;
  // Устанавливаем Ф.И.О. музыканта
  lbledtMusiciansName.Text := dbgrdMusicians.Columns[0].Field.AsString;
  // Устанавливаем дату рождения
  dtpMusiciansDate.Date := dbgrdMusicians.Columns[1].Field.AsDateTime;
  // Устанавливаем пол
  cbbMusiciansSex.ItemIndex := dbgrdMusicians.Columns[2].Field.AsInteger;
   // Заполняем список группами
  Query := Repository.GetQuery(Self);
  FillInputListAndSelect(Query, cbbMusiciansGroup,
    'GROUPS',
    'GROUP_NAME',
    dbgrdMusicians.Columns[3].Field.AsString);
  Query.Free;
  // Устанавливаем режим добавления
  pnlMusiciansSql.Tag := 1;
  // Меняем текста на панели
  lblMusiciansCaption.Caption := 'Редактирование музыканта';
  btnMusiciansEnter.Caption := 'Применить';
  // отображение панели
  pnlMusiciansSql.Show;
  LockCurrentPage();
end;

// Действия при Добавлении / Редактировании музыканта
procedure TMainForm.btnMusiciansEnterClick(Sender: TObject);
var
  Query: TIBQuery;
  GroupId: Integer;
begin
  Query := Repository.GetQuery(Self);
  // Проверка полей
  if (InputIsEmpty(lbledtMusiciansName)) then
  begin
    ShowError('Поле с именем музыканта не может быть пустым!');
    Abort;
  end;
  // Если стоит режим добавления
  if (pnlMusiciansSql.Tag = 0) then
  begin
    // Проверяем на уникальность запись
    if (not CheckFieldOnUnique('MUSICIANS', 'NAME', lbledtMusiciansName.Text)) then
    begin
      lbledtMusiciansName.Color := TColor($88FFFF);
      ShowError('Ф.И.О. музыканта не уникально!');
      Abort;
    end;
    GroupId := StrToInt(GetFieldsById('GROUPS', 'GROUPS.NAME', cbbMusiciansGroup.Text)[0]);
    // Формируем команду на вставку
    SetInsertCommand(Query,
      'MUSICIANS',
      ['ID', 'GROUPID', 'NAME', 'AGE', 'SEX'],
      [GetUniqueId('MUSICIANS', 'ID'), GroupId, lbledtMusiciansName.Text,
      DateToStr(dtpMusiciansDate.Date), cbbMusiciansSex.ItemIndex]
      );
    // Выполняем сформированный запрос и сохраняем
    Query.ExecSQL;
    Repository.SaveChanges;
  end
  else
  // Если стоит режим обновления
  begin
    GroupId := StrToInt(GetFieldsById('GROUPS', 'GROUPS.NAME', cbbMusiciansGroup.Text)[0]);
    // Формируем команду на обнавление
    SetUpdateCommand(Query,
      'MUSICIANS',
      ['GROUPID', 'NAME', 'AGE', 'SEX'],
      [GroupId, lbledtMusiciansName.Text,
      DateToStr(dtpMusiciansDate.Date), cbbMusiciansSex.ItemIndex],
      GetCurrentRecordID());
    // Выполняем сформированный запрос и сохраняем
    Query.ExecSQL;
    Repository.SaveChanges;
  end;
  // Очищаем память
  Query.Free;
  // Перезагружаем данные из таблицы
  dsMusicians.DataSet.Active := True;
  // Очищаем форму
  btnMusiciansCancelClick(Sender);
end;

// Действия при изменении азмеров формы
procedure TMainForm.FormResize(Sender: TObject);
begin
  // Выбираем распределитель
  case pgcTables.ActivePageIndex of
    0:
    begin
      if (pnlAlbumsSql.Visible) then
      begin
        pnlAlbumsSql.Left := (tsAlbums.Width div 2) - 180;
        pnlAlbumsSql.Top := (tsAlbums.Height div 2) - 165;
      end;
    end;
    1:
    begin
      if (pnlGroupsSql.Visible) then
      begin
        pnlGroupsSql.Left := (tsGroups.Width div 2) - 180;
        pnlGroupsSql.Top := (tsGroups.Height div 2) - 140;
      end;
    end;
    2:
    begin
      if (pnlMusiciansSql.Visible) then
      begin
        pnlMusiciansSql.Left := (tsMusicians.Width div 2) - 180;
        pnlMusiciansSql.Top := (tsMusicians.Height div 2) - 115;
      end;
    end;
    3:
    begin
      if (pnlSoundsSql.Visible) then
      begin
        pnlSoundsSql.Left := (tsSounds.Width div 2) - 180;
        pnlSoundsSql.Top := (tsSounds.Height div 2) - 96;
      end;
    end;
    else
    begin
      ShowError('Индекс таблицы вишел за границы!');
      Exit;
    end;
  end;

  pnlTools.Left := (pnlBackTools.Width div 2) - 150;
end;

// Предустановки блокирующие не линейное изменение таблицы
procedure TMainForm.pgcTablesChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
  if (pnlAlbumsSql.Visible) then
    btnAlbumsPanelCancelClick(btnAlbumsCancel);

  if (pnlGroupsSql.Visible) then
    btnGroupsCancelClick(btnGroupsCancel);

  if (pnlMusiciansSql.Visible) then
    btnMusiciansCancelClick(btnMusiciansCancel);

  if (pnlSoundsSql.Visible) then
    btnSoundsPanelCancelClick(btnSoundsCancel);
end;

// Создаем отчет по таблице альбомов
procedure TMainForm.CreateReportForCurrentTable(Sender: TObject);
var
  ExcelApp, WorkBook, Cell1, Cell2, Range, ArrayData: Variant;
  X, Y, RowCount, ColCount: Integer;
  Query: TIBQuery;
  DBGrrid: TDBGrid;
begin
  ExcelApp := CreateOleObject('Excel.Application');
  try
    ExcelApp.Application.EnableEvents := False;
    Workbook := ExcelApp.WorkBooks.Add;

    WorkBook.WorkSheets[1].Name := 'Отчет';
    WorkBook.WorkSheets[1].Rows[2].Font.Bold := True;
    WorkBook.WorkSheets[1].Rows[2].Font.Color := clRed;
    WorkBook.WorkSheets[1].Rows[2].Font.Size := 12;
    WorkBook.WorkSheets[1].Rows[2].HorizontalAlignment := 3;

    // Формируем запрос
    Query := Repository.GetQuery(Self);
    SetSelectAllCommand(Query, Repository.GetTablesNames[pgcTables.TabIndex]);
    Query.Active := True;
    Query.Last;

    // Определяем что нужно вносить в отчет
    case pgcTables.TabIndex of
      0:DBGrrid := dbgrdAlbums;
      1:DBGrrid := dbgrdGroups;
      2:DBGrrid := dbgrdMusicians;
      3:DBGrrid := dbgrdSounds;
      else DBGrrid := TDBGrid.Create(nil);
    end;

    // Получаем кол-во столбцов и элементов
    ColCount := DBGrrid.Columns.Count;
    RowCount := Query.RecordCount;

    // Выстраиваем сетку
    ArrayData := VarArrayCreate([1, RowCount, 1, ColCount], varVariant);
    Cell1 := WorkBook.WorkSheets[1].Cells[2, 2];
    Cell2 := WorkBook.WorkSheets[1].Cells[RowCount + 2, ColCount + 1];
    Range := WorkBook.WorkSheets[1].Range[Cell1, Cell2];
    Range.Value := ArrayData;
    Range.HorizontalAlignment:=2;
    Range.Borders.LineStyle:=1;

    // Отображаем все заголовки столбцы
    for X := 0 to DBGrrid.Columns.Count - 1 do
    begin
      // Выводим все заголовки
      ExcelApp.cells[2, X + 2] := DBGrrid.Columns[X].Title.Caption;
      // отображаем все поля по данному столбцу
      DBGrrid.DataSource.DataSet.First;
      Y := 3;
      while not DBGrrid.DataSource.DataSet.Eof do
      begin
        ExcelApp.cells[Y, X + 2] :=
          DBGrrid.DataSource.DataSet.FieldByName(DBGrrid.Columns[X].FieldName).AsString;
        DBGrrid.DataSource.DataSet.Next;
        Inc(Y);
      end;
    end;

    ExcelApp.Columns.AutoFit;
    ExcelApp.Visible:=True;
    Exit;
  except
    ShowError('Не удалось сформировать отчет!');
    Abort;
  end;
end;

end.

