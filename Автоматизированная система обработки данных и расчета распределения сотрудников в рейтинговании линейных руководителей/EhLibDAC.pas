{
    EhLib
    Devart DAC Features
    (c) Dorin Marcoci
}

unit EhLibDAC;

interface

implementation

uses
 DBUtilsEh, DBGridEh, ToolCtrlsEh, DB, SysUtils, DBAccess;

type
  TDACDatasetFeaturesEh = class(TDatasetFeaturesEh)
  public
    procedure ApplySorting(Sender: TObject; DataSet: TDataSet; IsReopen: Boolean); override;
    procedure ApplyFilter(Sender: TObject; DataSet: TDataSet; IsReopen: Boolean); override;
  end;

function DateValueToIBSQLStringProc(DataSet: TDataSet; Value: Variant): string;
begin
  Result := '''' + FormatDateTime('YYYY-MM-DD', Value) + '''';
end;

procedure TDACDatasetFeaturesEh.ApplyFilter(Sender: TObject;
  DataSet: TDataSet; IsReopen: Boolean);
var
  Grid: TCustomDBGridEh;
  Data: TCustomDADataSet;
  S: string;
begin
  Grid := TCustomDBGridEh(Sender);
  Data := TCustomDADataSet(DataSet);
  if Grid.STFilter.Local then
  begin
    S := GetExpressionAsFilterString(Grid, GetOneExpressionAsLocalFilterString, nil, False, True);
    Data.Filter := S;
  end else
  begin
    S := GetExpressionAsFilterString(Grid, GetOneExpressionAsSQLWhereString, DateValueToIBSQLStringProc, True);
    Data.FilterSQL := S;
  end;
  Data.Filtered := S <> '';
end;

procedure TDACDatasetFeaturesEh.ApplySorting(Sender: TObject; DataSet: TDataSet; IsReopen: Boolean);
var
  Grid: TCustomDBGridEh;
  I: Integer;
  S, F: string;
begin
  Grid := TCustomDBGridEh(Sender);
  if Grid.SortLocal then
  begin
    for I := 0 to Grid.SortMarkedColumns.Count - 1 do
    begin
      F := Grid.SortMarkedColumns[I].FieldName;
      if F = '' then Continue;
      if Grid.SortMarkedColumns[I].Title.SortMarker = smDownEh then
        F := F + ' DESC';
      S := S + F + ';';
    end;
    TCustomDADataSet(DataSet).IndexFieldNames := S;
  end else ApplySortingForSQLBasedDataSet(Grid, DataSet, True, IsReopen, 'SQL');
end;

initialization

  RegisterDatasetFeaturesEh(TDACDatasetFeaturesEh, TCustomDADataSet);

end.

