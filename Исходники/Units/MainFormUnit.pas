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
    // �����
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
    // ������
    procedure ShowSoundAddPanel(Sender: TObject);
    procedure ShowSoundEditPanel(Sender: TObject);
    procedure btnSoundsPanelCancelClick(Sender: TObject);
    procedure btnSoundsPanelEnterClick(Sender: TObject);
    // �������
    procedure ShowAlbumsAddPanel(Sender: TObject);
    procedure ShowAlbumsEditPanel(Sender: TObject);
    procedure btnAlbumsPanelCancelClick(Sender: TObject);
    procedure btnAlbumsPanelEnterClick(Sender: TObject);
    // ������
    procedure ShowGroupsAddPanel(Sender: TObject);
    procedure ShowGroupsEditPanel(Sender: TObject);
    procedure btnGroupsCancelClick(Sender: TObject);
    procedure btnGroupsEnterClick(Sender: TObject);
    // ���������
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


// ����� -----------------------------------------------------------------------

// ����������� ������ �� ����� ����� �����������
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
  // ��������� ������
  Query.SQL.Add('SELECT DISTINCT * FROM ' + TableName);
end;

// ����������� ������ �� ����������
procedure SetInsertCommand(Query: TIBQuery; TableName: string;
  ParamsNames: array of String; ParamsValues: Array of Variant);
var
  i: Integer;
begin
  With Query do
  begin
    SQL.Clear;
    SQL.Add('INSERT INTO ' + TableName);
    SQL.Add('('); // �������� �����
    SQL.Add('VALUES('); // �������� �����
    i := 0;
    Params.Clear;
    // ������ ���������
    while (i < Length(ParamsNames)) do
    begin
      // ����������� ����� �������
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

// ����������� ������ �� ���������� ������
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
    // ������ ���������
    while (i < Length(ParamsNames)) do
    begin
      // ����������� ����� �������
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

// ����������� ������ �� ����������
procedure SetDeleteCommand(Query: TIBQuery; TableName: String; ID: Integer);
begin
  with Query do
  begin
    SQL.Clear;
    SQL.Add('DELETE FROM ' + TableName + ' WHERE ID=:ID');
    ParamByName('ID').AsInteger := ID;
  end;
end;

// ������� �������� ���������� �� �������
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
  ShowError('������� "InputIsEmpty" ������� �� ������ ������ ����!');
  Abort;
end;

// �������� ������������� ������� ������
function TMainForm.GetCurrentRecordID(): Integer;
var
  DataSource: TDataSource;
begin
  // �������� ��������������
  case pgcTables.ActivePageIndex of
    0:DataSource := dsAlbums;
    1:DataSource := dsGroups;
    2:DataSource := dsMusicians;
    3:DataSource := dsSounds;
    else
    begin
      ShowError('������ ������� ����� �� �������!');
      Result := -1;
      Exit;
    end;
  end;
  // ��������� �������� ��������� ������
  if (DataSource.DataSet <> nil) then
    Result := DataSource.DataSet.FieldByName('ID').AsInteger
  else
    Result := -1;
end;

// ��������� ������������ � ������� �� ������������
function TMainForm.CheckIdOnUnique(TableName: string;
  IdentityColumn: string; Indx: Integer): Boolean;
var
  Query: TIBQuery;
begin
  // ������������ ����� �� ��������������
  Query := Self.Repository.GetQuery(Self);
  Query.SQL.Add('SELECT ' + IdentityColumn + ' FROM ' + TableName +
    ' WHERE ' + IdentityColumn + '=:IDD');
  Query.Params[0].AsInteger := Indx;
  Query.Active;

  // ��������� �� ���-�� ������� �� �������
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
  // ��������� �� ���-�� ������� �� �������
  Result := (Query.RecordCount = 0);
end;

// ������ �� ��������� �������� ���������� �������� ����
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
      ShowError('�������� ������� ������� ����� �� ��������� �������!');
      Abort;
    end;
  end;

  DBGrid.Enabled := Value;
  DBGrid.Refresh;
  btnDelete.Enabled := Value;
end;

// ������ ��������� �������� ���������� �������� ����
procedure TMainForm.UnLockCurrentPage();
begin
  LockCurrentPage(True);
end;

// ���������� ���� �� �������������� ������
function TMainForm.GetFieldsById(TableName: string;
  IdentityColumn: string; Id: Variant): TStrings;
var
  Query: TIBQuery;
  ColIndx: Integer;
begin
  // ������ ������ ��� ��������� �����
  Query := Self.Repository.GetQuery(Self);
  SetSelectAllCommand(Query, TableName);
  Query.SQL[0] := Query.SQL[0] + ' WHERE ' + IdentityColumn + '=:IDD';
  Query.ParamByName('IDD').AsString := VarTostr(Id);
  Query.Active := True;

  // ���� ���� ������ �� ��������� ����� �����
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

// �������� ���������� ������������� �� ��������
function TMainForm.GetUniqueId(TableName: string;
  IdentityColumn: string): Integer;
var
  Query: TIBQuery;
begin
  Query := Self.Repository.GetQuery(Self);
  Randomize;
  Result := Random(1000000);

  // ���������� ����� �������� ������ ��� ����� ��� ��� ��������� � ����
  while (not CheckIdOnUnique(TableName, IdentityColumn, Result)) do
    Result := Random(1000000);

  Query.Free;
end;

// ��������� ������ �������
procedure FillInputList(Query: TIBQuery; ComboBox: TComboBox;
  TableName: string; ColumnName: String);
begin
  // ����� ��� ������
  SetSelectAllCommand(Query, TableName);
  Query.Active := True;
  // ������� ������� ������
  ComboBox.Clear;
  // ��������� ��������� �� �������
  ComboBox.Enabled := (Query.RecordCount > 0);
  // ���������� ���������� �� � �����
  While (not Query.Eof) do
  begin
    ComboBox.Items.Add(Query.FieldByName(ColumnName).AsString);
    Query.Next;
  end;
  // ���� ���� ������ - ������ ������ � ������
  if (ComboBox.Enabled) then
    ComboBox.ItemIndex := 0;
end;

// ��������� ������ ������� � ������������� ����� �� ������ ��������
procedure FillInputListAndSelect(Query: TIBQuery; ComboBox: TComboBox;
  Tablename: string; FieldName: String; SelectedValue: String);
begin
  // ����� ��� ������
  SetSelectAllCommand(Query, Tablename);
  Query.Active := True;
  // ������� ������� ������
  ComboBox.Clear;
  // ��������� ��������� �� �������
  ComboBox.Enabled := (Query.RecordCount > 0);
  // ���������� ���������� �� � �����
  While (not Query.Eof) do
  begin
    // �������� ������ � ������
    ComboBox.Items.Add(Query.FieldByName(FieldName).AsString);
    // ���� ������ � ��������� �� ���������� �� ���� �����
    if (Query.FieldByName(FieldName).AsString = SelectedValue) then
      ComboBox.ItemIndex := ComboBox.Items.Count -1;
    Query.Next;
  end;
  // ���� ���� ������ - ������ ������ � ������
  if (ComboBox.Enabled) then
    ComboBox.ItemIndex := 0;
end;

// ��������� �������� ������
procedure TMainForm.pgcTablesChange(Sender: TObject);
var
  Query: TIBQuery;
  DataSource: TDataSource;
begin
  case pgcTables.TabIndex of
    0:
    begin
      DataSource := dsAlbums;
      // ������������� ������
      btnAdd.OnClick := ShowAlbumsAddPanel;
      btnEdit.OnClick := ShowAlbumsEditPanel;
      btnDelete.Enabled := (not pnlAlbumsSql.Visible);
    end;
    1:
    begin
      DataSource := dsGroups;
      // ������������� ������
      btnAdd.OnClick := ShowGroupsAddPanel;
      btnEdit.OnClick := ShowGroupsEditPanel;
      btnDelete.Enabled := (not pnlGroupsSql.Visible);
    end;
    2:
    begin
      DataSource := dsMusicians;
      // ������������� ������
      btnAdd.OnClick := ShowMusiciansAddPanel;
      btnEdit.OnClick := ShowMusiciansEditPanel;
      btnDelete.Enabled := (not pnlMusiciansSql.Visible);
    end;
    3:
    begin
      DataSource := dsSounds;
      // ������������� ������
      btnAdd.OnClick := ShowSoundAddPanel;
      btnEdit.OnClick := ShowSoundEditPanel;
      btnDelete.Enabled := (not pnlSoundsSql.Visible);
    end;
    else
    begin
      DataSource := TDataSource.Create(Self);
      ShowError('�������� ������� ������� ����� �� ��������� �������!');
      Abort;
    end;
  end;
  // ���� ������� �� ���� ���������
  if (DataSource.DataSet = nil) then
  begin
    // �������� ��������� ��� ����� ���������
    Query := Repository.GetQuery(Self);
    // ��������� ������� ������� ���� ���������
    SetSelectAllCommand(Query,
      Repository.GetTablesNames()[pgcTables.ActivePageIndex]);
    // ��������� ������� � ������� �����������
    DataSource.DataSet := Query;
  end;

  DataSource.DataSet.Active := False;
  DataSource.DataSet.Active := True;
end;

// ����������� � �����������
procedure TMainForm.ShowWithRepo(Sender: TObject; Repository: TRepository);
begin
  // ��������� ����������� ������
  Self.Repository := Repository;
  Self.Repository.Connect;
  // ��������� ������ ��������
  pgcTablesChange(Self);
  // �������� ������ �����
  (Sender as TForm).Hide;
  // ���������� ������ ����� ���������
  Show();
end;

// �������� �����
procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    // ������������� ������
  if (ShowConfirm(CONFIRM_EXIT)) then
    ExitProcess(0)
  else
    Abort;
end;

// ������������� �����
procedure TMainForm.FormCreate(Sender: TObject);
begin
  Self.DoubleBuffered := True;
end;

// ������ �� ������ ������������ �������� ������
// ��� ���� ������ ���� -> ������� �������� :)
procedure TMainForm.btnDeleteClick(Sender: TObject);
var
  Query: TIBQuery;
  TabIndex: Integer;
  TableName: String;
  DataSource: TDataSource;
begin
  // �������� ����� �������
  TabIndex := pgcTables.TabIndex;
  // ���������� ������ ����� �������
  case TabIndex of
    0:DataSource := dsAlbums;
    1:DataSource := dsGroups;
    2:DataSource := dsMusicians;
    3:DataSource := dsSounds;
    else
    begin
      DataSource := TDataSource.Create(Self);
      ShowError('�������� ������� ������� ����� �� ��������� �������!');
      Abort;
    end;
  end;
  // ���� � ������� ������ ���, �� ��������� ������������
  if (DataSource.DataSet.RecordCount = 0) then
  begin
    ShowWarning('����������� ������ ��� ��������!');
    Abort;
  end;

  // ������ ������������� ��������
  if (not ShowConfirm(CONFIRM_DELETE)) then
    Abort;
  // �������� �������� �������
  TableName := Repository.GetTablesNames()[TabIndex];
  Query := Repository.GetQuery(Self);
  // ������������� � ��������� ������� ��������
  SetDeleteCommand(Query,
    TableName,
    DataSource.DataSet.FieldByName('ID').AsInteger);
  Query.ExecSQL;
  Repository.SaveChanges;

  // ��������� ������� �������
  DataSource.DataSet.Active := True;
end;

// ��������� �������� ��������� ����� ������
procedure TMainForm.FillClearColor(Sender: TObject);
begin
  if (Sender is TLabeledEdit) then
    (Sender as TLabeledEdit).Color := clWindow;

  if (Sender is TDateTimePicker) then
    (Sender as TDateTimePicker).Color := clWindow;
end;


// ������ ----------------------------------------------------------------------

// ������ �� ������ ������������ ���������� ������
procedure TMainForm.ShowSoundAddPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // ������� ������� ������
  btnSoundsPanelCancelClick(Self);
  // ��������� ���������� ������ � ���������
  Query := Repository.GetQuery(Self);
  FillInputList(Query, cbbSoundsAlbum, 'ALBUMS', 'ALBUM_NAME');
  Query.Free;
  // �������� ������
  // ������ ���� ����������� ������ � ���������� �������
  btnSoundsEnter.Enabled := cbbSoundsAlbum.Enabled;
  // ������������� ����� ����������
  pnlSoundsSql.Tag := 0;
  lblSoundsCaption.Caption := '���������� ����� ������';
  btnSoundsEnter.Caption := '��������';
  // ����������� ������
  pnlSoundsSql.Show;
  LockCurrentPage();
end;

// ������ �� ������ ������������ �������������� ������
procedure TMainForm.ShowSoundEditPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // ���� � ������� ������ ���, �� ��������� ������������
  if (dsSounds.DataSet.RecordCount = 0) then
  begin
    ShowWarning('����������� ������ ��� ��������������!');
    Abort;
  end;
  // ��������� ���������� ������ � ���������
  // � �������� ��� ���������
  Query := Repository.GetQuery(Self);
  FillInputListAndSelect(Query, cbbSoundsAlbum,
    'ALBUMS',
    'ALBUM_NAME',
    dbgrdSounds.Columns[2].Field.AsString);
  Query.Free;
  // ������������� �������� ������
  lbledtSoundsName.Text := dbgrdSounds.Columns[0].Field.AsString;
  // ������������� ����� ������
  dtpSoundTime.Time := dbgrdSounds.Columns[3].Field.AsVariant;
  // �������� ������
  // ������ ���� ����������� ������ � ���������� �������
  btnSoundsEnter.Enabled := cbbSoundsAlbum.Enabled;
  // ������������� ����� ����������
  pnlSoundsSql.Tag := 1;
  // ������ ������ �� ������
  lblSoundsCaption.Caption := '�������������� ������';
  btnSoundsEnter.Caption := '���������';
  // ����������� ������
  pnlSoundsSql.Show;
  LockCurrentPage();
end;

// �������� ����� (���������� / ��������������)
procedure TMainForm.btnSoundsPanelCancelClick(Sender: TObject);
begin
  // ������� �������� ������
  lbledtSoundsName.Text := '';
  // �������� ������
  pnlSoundsSql.Hide;
  // ���������� ����������� ����� ������
  dtpSoundTime.Time := StrToTime('00:00:01');
  // ������� ����� � �����
  lbledtSoundsName.Color := clWindow;
  dtpSoundTime.Color := clWindow;
  UnLockCurrentPage();
end;
// ���������� �������� ����� (���������� / ��������������)
procedure TMainForm.btnSoundsPanelEnterClick(Sender: TObject);
var
  Query: TIBQuery;
  AlbumId: Integer;
begin
  Query := Repository.GetQuery(Self);
  // �������� �����
  if (InputIsEmpty(lbledtSoundsName)) then
  begin
    ShowError('���� � ��������� ����� �� ����� ���� ������!');
    Abort;
  end;
  if (InputIsEmpty(dtpSoundTime)) then
  begin
    ShowError('���� � �������� ����� �� ����� ���� ������ ����!');
    Abort;
  end;
  // ���� ����� ����� ����������
  if (pnlSoundsSql.Tag = 0) then
  begin
    // ��������� �� ������������ ������
    if (not CheckFieldOnUnique('SOUNDS', 'NAME', lbledtSoundsName.Text)) then
    begin
      lbledtSoundsName.Color := TColor($88FFFF);
      ShowError('�������� ������ �� ���������!');
      Abort;
    end;
    // �������� ������������� �������
    AlbumId := StrToInt(GetFieldsById('ALBUMS', 'ALBUMS.NAME', cbbSoundsAlbum.Text)[0]);
    // ��������� ������� �� �������
    SetInsertCommand(Query,
      'SOUNDS',
      ['ID', 'ALBUMID', 'NAME', 'PLAYTIME'],
      [GetUniqueId('SOUNDS', 'ID'),
        AlbumId,
        lbledtSoundsName.Text,
        TimeToStr(dtpSoundTime.Time)]
      );
    // ��������� �������������� ������
    Query.ExecSQL;
    Repository.SaveChanges;
    SetSelectAllCommand(Query, 'SOUNDS');
    Query.Active := True;
  end
  else
  // ���� ����� ����� ����������
  begin
    // �������� ������������� �������
    AlbumId := StrToInt(GetFieldsById('ALBUMS', 'ALBUMS.NAME', cbbSoundsAlbum.Text)[0]);
    SetUpdateCommand(Query,
      'SOUNDS',
      ['ALBUMID', 'NAME', 'PLAYTIME'],
      [AlbumId,lbledtSoundsName.Text, TimeToStr(dtpSoundTime.Time)],
      GetCurrentRecordID());

    // ��������� �������������� ������
    Query.ExecSQL;
    Repository.SaveChanges;
    Query.Free;
  end;
  // ������������� ������ �� �������
  dsSounds.DataSet.Active := True;
  // �������� ������
  btnSoundsCancel.Click;
end;

// ������� ---------------------------------------------------------------------

// ������ �� ������ ������������ ���������� ������
procedure TMainForm.ShowAlbumsAddPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // ��������� ���������� ������ � ��������
  Query := Repository.GetQuery(Self);
  FillInputList(Query, cbbAlbumsGroup, 'GROUPS', 'GROUP_NAME');
  Query.Free;
  // �������� ������
  // ������ ���� ����������� ������ � ���������� �������
  btnAlbumsEnter.Enabled := cbbAlbumsGroup.Enabled;
  // ������������� ����� ����������
  pnlAlbumsSql.Tag := 0;
  lblAlbumsCaption.Caption := '���������� ������ �������';
  btnAlbumsEnter.Caption := '��������';
  // ����������� ������
  pnlAlbumsSql.Show;
  LockCurrentPage();
end;

// ������ �� ������ ������������ ������������� ������
procedure TMainForm.ShowAlbumsEditPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // ���� � ������� ������ ���, �� ��������� ������������
  if (dsAlbums.DataSet.RecordCount = 0) then
  begin
    ShowWarning('����������� ������ ��� ��������������!');
    Abort;
  end;
  // ��������� ���������� ������ � ���������
  // � �������� ��� ���������
  Query := Repository.GetQuery(Self);
  FillInputListAndSelect(Query, cbbAlbumsGroup,
    'GROUPS',
    'GROUP_NAME',
    dbgrdAlbums.Columns[1].Field.AsString);
  Query.Free;
  // ������������� �������� �������
  lbledtAlbumsName.Text := dbgrdAlbums.Columns[0].Field.AsString;
  // ������������� �������� �������
  mmoAlbumsInfo.Lines.Add(dbgrdAlbums.Columns[2].Field.AsString);
  // ������ �� �������� ������
  // ������ ���� ����������� ������ � ���������� �������
  btnAlbumsEnter.Enabled := cbbAlbumsGroup.Enabled;
  // ������������� ����� ����������
  pnlAlbumsSql.Tag := 1;
  // ������ ������ �� ������
  lblAlbumsCaption.Caption := '�������������� �������';
  btnAlbumsEnter.Caption := '���������';
  // ����������� ������
  pnlAlbumsSql.Show;
  LockCurrentPage();
end;

// �������� ������
procedure TMainForm.btnAlbumsPanelCancelClick(Sender: TObject);
begin
  // ���������� ������
  lbledtAlbumsName.Text := '';
  cbbAlbumsGroup.Items.Clear;
  mmoAlbumsInfo.Clear;
  // ���������� �����
  lbledtAlbumsName.Color := clWindow;
  cbbAlbumsGroup.Color := clWindow;
  mmoAlbumsInfo.Color := clWindow;
  // �������� ������
  pnlAlbumsSql.Hide;
  UnLockCurrentPage();
end;

// �������� �� �������� � ������
procedure TMainForm.btnAlbumsPanelEnterClick(Sender: TObject);
var
  Query: TIBQuery;
  GroupId: Integer;
begin
  Query := Repository.GetQuery(Self);
  // �������� �����
  if (InputIsEmpty(lbledtAlbumsName)) then
  begin
    ShowError('���� � ��������� ������� �� ����� ���� ������!');
    Abort;
  end;
  if (InputIsEmpty(mmoAlbumsInfo)) then
  begin
    ShowError('���� � ����������� � ������� �� ����� ���� ������!!');
    Abort;
  end;
  // ���� ����� ����� ����������
  if (pnlAlbumsSql.Tag = 0) then
  begin
    // ��������� �� ������������ ������
    if (not CheckFieldOnUnique('ALBUMS', 'NAME', lbledtAlbumsName.Text)) then
    begin
      lbledtAlbumsName.Color := TColor($88FFFF);
      ShowError('�������� ������� �� ���������!');
      Abort;
    end;
    // �������� ������������� ������
    GroupId := StrToInt(GetFieldsById('GROUPS', 'GROUPS.NAME', cbbAlbumsGroup.Text)[0]);
    // ��������� ������� �� �������
    SetInsertCommand(Query,
      'ALBUMS',
      ['ID',                        'GROUPID', 'NAME',               'INFO'],
      [GetUniqueId('ALBUMS', 'ID'), GroupId,   lbledtAlbumsName.Text, mmoAlbumsInfo.Lines.Text]
      );
    // ��������� �������������� ������ � ���������
    Query.ExecSQL;
    Repository.SaveChanges;
  end
  else
  // ���� ����� ����� ����������
  begin
    // �������� ������������� ������
    GroupId := StrToInt(GetFieldsById('GROUPS', 'GROUPS.NAME', cbbAlbumsGroup.Text)[0]);
    // ��������� ������� �� ����������
    SetUpdateCommand(Query,
      'ALBUMS',
      ['GROUPID', 'NAME',                'INFO'],
      [GroupId,   lbledtAlbumsName.Text, mmoAlbumsInfo.Lines.Text],
      GetCurrentRecordID());
    // ��������� �������������� ������ � ���������
    Query.ExecSQL;
    Repository.SaveChanges;
  end;
  // ������� ������
  Query.Free;
  // ������������� ������ �� �������
  dsAlbums.DataSet.Active := True;
  // ������� �����
  btnAlbumsPanelCancelClick(Sender);
end;

// ������ ----------------------------------------------------------------------

// ������ �� ������ ������������ ���������� ������
procedure TMainForm.ShowGroupsAddPanel(Sender: TObject);
begin
  // ������������� ����� ����������
  pnlGroupsSql.Tag := 0;
  lblGroupsCaption.Caption := '���������� ������ ������';
  btnGroupsEnter.Caption := '��������';
  // ����������� ������
  pnlGroupsSql.Show;
  LockCurrentPage();
end;

// ������ �� ������ ������������ ������������� ������
procedure TMainForm.ShowGroupsEditPanel(Sender: TObject);
begin
  // ���� � ������� ������ ���, �� ��������� ������������
  if (dsGroups.DataSet.RecordCount = 0) then
  begin
    ShowWarning('����������� ������ ��� ��������������!');
    Abort;
  end;
  // ������������� �������� �������
  lbledtGroupsName.Text := dbgrdGroups.Columns[0].Field.AsString;
  // ������������� �������� �������
  mmoGroupsInfo.Lines.Add(dbgrdGroups.Columns[1].Field.AsString);
  // ������������� ����� ����������
  pnlGroupsSql.Tag := 1;
  // ������ ������ �� ������
  lblGroupsCaption.Caption := '�������������� ������';
  btnGroupsEnter.Caption := '���������';
  // ����������� ������
  pnlGroupsSql.Show;
  LockCurrentPage();
end;


// �������� ������
procedure TMainForm.btnGroupsCancelClick(Sender: TObject);
begin
  // ���������� ������
  lbledtGroupsName.Text := '';
  mmoGroupsInfo.Lines.Clear;
  // ���������� �����
  lbledtGroupsName.Color := clWindow;
  mmoGroupsInfo.Color := clWindow;
  // �������� ������
  pnlGroupsSql.Hide;
  UnLockCurrentPage();
end;

// �������� �� �������� � ������
procedure TMainForm.btnGroupsEnterClick(Sender: TObject);
var
  Query: TIBQuery;
begin
  Query := Repository.GetQuery(Self);
  // �������� �����
  if (InputIsEmpty(lbledtGroupsName)) then
  begin
    ShowError('���� � ��������� ������ �� ����� ���� ������!');
    Abort;
  end;
  if (InputIsEmpty(mmoGroupsInfo)) then
  begin
    ShowError('���� � ����������� � ������ �� ����� ���� ������!!');
    Abort;
  end;
  // ���� ����� ����� ����������
  if (pnlGroupsSql.Tag = 0) then
  begin
    // ��������� �� ������������ ������
    if (not CheckFieldOnUnique('GROUPS', 'NAME', lbledtGroupsName.Text)) then
    begin
      lbledtGroupsName.Color := TColor($88FFFF);
      ShowError('�������� ������ �� ���������!');
      Abort;
    end;
    // ��������� ������� �� �������
    SetInsertCommand(Query,
      'GROUPS',
      ['ID',                        'NAME',                'INFO'],
      [GetUniqueId('GROUPS', 'ID'), lbledtGroupsName.Text, mmoGroupsInfo.Lines.Text]
      );
    // ��������� �������������� ������ � ���������
    Query.ExecSQL;
    Repository.SaveChanges;
  end
  else
  // ���� ����� ����� ����������
  begin
    // ��������� ������� �� ����������
    SetUpdateCommand(Query,
      'GROUPS',
      ['NAME',                'INFO'],
      [lbledtGroupsName.Text, mmoGroupsInfo.Lines.Text],
      GetCurrentRecordID());
    // ��������� �������������� ������ � ���������
    Query.ExecSQL;
    Repository.SaveChanges;
  end;
  // ������� ������
  Query.Free;
  // ������������� ������ �� �������
  dsGroups.DataSet.Active := True;
  // ������� �����
  btnGroupsCancelClick(Sender);
end;

// ��������� -------------------------------------------------------------------

// �������� ������
procedure TMainForm.btnMusiciansCancelClick(Sender: TObject);
begin
  // �������� ����
  lbledtMusiciansName.Text := '';
  dtpMusiciansDate.Date := Now;
  cbbMusiciansSex.ItemIndex := 0;
  cbbMusiciansGroup.Items.Clear;
  pnlMusiciansSql.Hide;
  UnLockCurrentPage();
end;

// ������ �� ������ ������������ ���������� ������
procedure TMainForm.ShowMusiciansAddPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // ��������� ���������� ������ � ��������
  Query := Repository.GetQuery(Self);
  FillInputList(Query, cbbMusiciansGroup, 'GROUPS', 'GROUP_NAME');
  Query.Free;
  // ������������� ����� ����������
  pnlMusiciansSql.Tag := 0;
  lblMusiciansCaption.Caption := '���������� ������ ���������';
  btnMusiciansEnter.Caption := '��������';
  // ����������� ������
  pnlMusiciansSql.Show;
  LockCurrentPage();
end;

// ������ �� ������ ������������ ������������� ������
procedure TMainForm.ShowMusiciansEditPanel(Sender: TObject);
var
  Query: TIBQuery;
begin
  // ���� � ������� ������ ���, �� ��������� ������������
  if (dsMusicians.DataSet.RecordCount = 0) then
  begin
    ShowWarning('����������� ������ ��� ��������������!');
    Abort;
  end;
  // ������������� �.�.�. ���������
  lbledtMusiciansName.Text := dbgrdMusicians.Columns[0].Field.AsString;
  // ������������� ���� ��������
  dtpMusiciansDate.Date := dbgrdMusicians.Columns[1].Field.AsDateTime;
  // ������������� ���
  cbbMusiciansSex.ItemIndex := dbgrdMusicians.Columns[2].Field.AsInteger;
   // ��������� ������ ��������
  Query := Repository.GetQuery(Self);
  FillInputListAndSelect(Query, cbbMusiciansGroup,
    'GROUPS',
    'GROUP_NAME',
    dbgrdMusicians.Columns[3].Field.AsString);
  Query.Free;
  // ������������� ����� ����������
  pnlMusiciansSql.Tag := 1;
  // ������ ������ �� ������
  lblMusiciansCaption.Caption := '�������������� ���������';
  btnMusiciansEnter.Caption := '���������';
  // ����������� ������
  pnlMusiciansSql.Show;
  LockCurrentPage();
end;

// �������� ��� ���������� / �������������� ���������
procedure TMainForm.btnMusiciansEnterClick(Sender: TObject);
var
  Query: TIBQuery;
  GroupId: Integer;
begin
  Query := Repository.GetQuery(Self);
  // �������� �����
  if (InputIsEmpty(lbledtMusiciansName)) then
  begin
    ShowError('���� � ������ ��������� �� ����� ���� ������!');
    Abort;
  end;
  // ���� ����� ����� ����������
  if (pnlMusiciansSql.Tag = 0) then
  begin
    // ��������� �� ������������ ������
    if (not CheckFieldOnUnique('MUSICIANS', 'NAME', lbledtMusiciansName.Text)) then
    begin
      lbledtMusiciansName.Color := TColor($88FFFF);
      ShowError('�.�.�. ��������� �� ���������!');
      Abort;
    end;
    GroupId := StrToInt(GetFieldsById('GROUPS', 'GROUPS.NAME', cbbMusiciansGroup.Text)[0]);
    // ��������� ������� �� �������
    SetInsertCommand(Query,
      'MUSICIANS',
      ['ID', 'GROUPID', 'NAME', 'AGE', 'SEX'],
      [GetUniqueId('MUSICIANS', 'ID'), GroupId, lbledtMusiciansName.Text,
      DateToStr(dtpMusiciansDate.Date), cbbMusiciansSex.ItemIndex]
      );
    // ��������� �������������� ������ � ���������
    Query.ExecSQL;
    Repository.SaveChanges;
  end
  else
  // ���� ����� ����� ����������
  begin
    GroupId := StrToInt(GetFieldsById('GROUPS', 'GROUPS.NAME', cbbMusiciansGroup.Text)[0]);
    // ��������� ������� �� ����������
    SetUpdateCommand(Query,
      'MUSICIANS',
      ['GROUPID', 'NAME', 'AGE', 'SEX'],
      [GroupId, lbledtMusiciansName.Text,
      DateToStr(dtpMusiciansDate.Date), cbbMusiciansSex.ItemIndex],
      GetCurrentRecordID());
    // ��������� �������������� ������ � ���������
    Query.ExecSQL;
    Repository.SaveChanges;
  end;
  // ������� ������
  Query.Free;
  // ������������� ������ �� �������
  dsMusicians.DataSet.Active := True;
  // ������� �����
  btnMusiciansCancelClick(Sender);
end;

// �������� ��� ��������� ������� �����
procedure TMainForm.FormResize(Sender: TObject);
begin
  // �������� ��������������
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
      ShowError('������ ������� ����� �� �������!');
      Exit;
    end;
  end;

  pnlTools.Left := (pnlBackTools.Width div 2) - 150;
end;

// ������������� ����������� �� �������� ��������� �������
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

// ������� ����� �� ������� ��������
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

    WorkBook.WorkSheets[1].Name := '�����';
    WorkBook.WorkSheets[1].Rows[2].Font.Bold := True;
    WorkBook.WorkSheets[1].Rows[2].Font.Color := clRed;
    WorkBook.WorkSheets[1].Rows[2].Font.Size := 12;
    WorkBook.WorkSheets[1].Rows[2].HorizontalAlignment := 3;

    // ��������� ������
    Query := Repository.GetQuery(Self);
    SetSelectAllCommand(Query, Repository.GetTablesNames[pgcTables.TabIndex]);
    Query.Active := True;
    Query.Last;

    // ���������� ��� ����� ������� � �����
    case pgcTables.TabIndex of
      0:DBGrrid := dbgrdAlbums;
      1:DBGrrid := dbgrdGroups;
      2:DBGrrid := dbgrdMusicians;
      3:DBGrrid := dbgrdSounds;
      else DBGrrid := TDBGrid.Create(nil);
    end;

    // �������� ���-�� �������� � ���������
    ColCount := DBGrrid.Columns.Count;
    RowCount := Query.RecordCount;

    // ����������� �����
    ArrayData := VarArrayCreate([1, RowCount, 1, ColCount], varVariant);
    Cell1 := WorkBook.WorkSheets[1].Cells[2, 2];
    Cell2 := WorkBook.WorkSheets[1].Cells[RowCount + 2, ColCount + 1];
    Range := WorkBook.WorkSheets[1].Range[Cell1, Cell2];
    Range.Value := ArrayData;
    Range.HorizontalAlignment:=2;
    Range.Borders.LineStyle:=1;

    // ���������� ��� ��������� �������
    for X := 0 to DBGrrid.Columns.Count - 1 do
    begin
      // ������� ��� ���������
      ExcelApp.cells[2, X + 2] := DBGrrid.Columns[X].Title.Caption;
      // ���������� ��� ���� �� ������� �������
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
    ShowError('�� ������� ������������ �����!');
    Abort;
  end;
end;

end.

