unit RepositoryUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, IBDatabase, IBCustomDataSet, IBTable,
  IBQuery, StdCtrls, MessagesUnit;

type

  // Репозиторий данных
  TRepository = class(TObject)
  private
    // База данных
    _database: TIBDataBase;
    // Транзактор
    _transactor: TIBTransaction;
    // Работа с путем
    procedure SetPath(Path: string);
    // Работа с адресом
    procedure SetLogin(Login: string);
    // Работа с адресом
    procedure SetPassword(Password: string);
  public
    // Путь к базе
    property Path: String write SetPath;
    // Логин для авторизации
    property Login: String write SetLogin;
    // Пароль для авторизации
    property Password: String write SetPassword;
    // Конструктор
    constructor Create(Sender: TComponent);
    // Деструктор
    destructor Destroy(); override;
    // Подключение к базе
    procedure Connect;
    // Сохраняет изменения в базе
    procedure SaveChanges();
    // Возвращает True если подключилось удачно
    function TestConnection: Boolean;
    // Возвращает Query
    function GetQuery(Sender: TComponent): TIBQuery;
    // Возвращает имена таблиц
    function GetTablesNames: TStringList;
    // Возвращает заголовки столбцов по имени таблицы
    function GetFieldNames(TableName: String): TStringList;
  end;

implementation


// Установка пути
procedure TRepository.SetPath(Path: string);
begin
  Self._database.DatabaseName := Path;
end;


// Установка Логина
procedure TRepository.SetLogin(Login: string);
var
  _paramIndex: Integer;
  _loginRowIndex: Integer;
  _leave: Boolean;
begin
  _loginRowIndex := -1;
  _paramIndex := 0;
  _leave := False;
  while ( (_paramIndex < Self._database.Params.Count) and (not _leave) ) do
  begin
    if (Pos('user_name=', Self._database.Params[_paramIndex]) = 1) then
    begin
      _loginRowIndex := _paramIndex;
      _leave := True;
    end
    else
      Inc(_paramIndex);
  end;
  
  if (_loginRowIndex = -1) then
    Self._database.Params.Add('user_name=' + Login)
  else
    Self._database.Params[_paramIndex] := 'user_name=' + Login;
end;


// Установка пароля
procedure TRepository.SetPassword(Password: string);
var
  _paramIndex: Integer;
  _passwordRowIndex: Integer;
  _leave: Boolean;
begin
  _passwordRowIndex := -1;
  _paramIndex := 0;
  _leave := False;
  while ( (_paramIndex < Self._database.Params.Count) and (not _leave)) do
  begin
    if (Pos('password=', Self._database.Params[_paramIndex]) = 1) then
    begin
      _passwordRowIndex := _paramIndex;
      _leave := True;
    end
    else
      Inc(_paramIndex);
  end;
  
  if (_passwordRowIndex = -1) then
    Self._database.Params.Add('password=' + Password)
  else
    Self._database.Params[_paramIndex] := 'password=' + Password;
end;


// Конструктор контекста 
constructor TRepository.Create(Sender: TComponent);
begin
  try
    self._database := TIBDatabase.Create(Sender);
    self._database.Params.Add('lc_ctype=WIN1251');
    self._transactor := TIBTransaction.Create(Sender);
    self._transactor.DefaultDatabase := self._database;
    Self._database.LoginPrompt := false;
  except
    ShowError(FAILED_CREATE_CONTEXT);
    Abort;
  end;
end;


// Деструткор контекста
destructor TRepository.Destroy;
begin
  _transactor.Active := False;
  _transactor.Free;
  _database.Connected := False;
  _database.Free;
  inherited;
end;


// Проверка подключения
function TRepository.TestConnection(): Boolean;
begin
  try
    Self._database.Connected := True;
    Result := True;
    Self._database.Connected := False;
  except
    Result := False;
  end;
end;


// Подключение к базе
procedure TRepository.Connect();
begin
  Self._database.Connected := True;
  Self._transactor.Active := True;
end;


// Возвращает имена таблиц
function TRepository.GetTablesNames(): TStringList;
var
  tables: TStringList;
begin
  tables := TStringList.Create();
  try
    self._database.GetTableNames(tables, false);
  except
    ShowError(FAILED_REQUEST_TABLE);
  end;
  Result := tables;
end;


// Возвращает заголовки столбцов по названию таблицы
function TRepository.GetFieldNames(TableName: String): TStringList;
var
  Columns: TStringList;
begin
  Columns := TStringList.Create();
  try
    self._database.GetFieldNames(Tablename, Columns);
  except
    ShowError(FAILED_REQUEST_COLUMN_NAMES);
  end;
  Result := Columns;
end;


// Возвращает новый Query через который можно получить доступ к базе
function TRepository.GetQuery(Sender: TComponent): TIBQuery;
var
  query: TIBQuery;
begin
  query := TIBQuery.Create(Sender);
  query.Transaction := _transactor;
  Result := query;
end;

procedure TRepository.SaveChanges();
begin
  _transactor.commit;
end;

end.
