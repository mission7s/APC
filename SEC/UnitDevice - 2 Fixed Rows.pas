unit UnitDevice;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitWorkForm, Vcl.Imaging.pngimage,
  WMTools, WMControls, Vcl.ExtCtrls, AdvUtil, Vcl.Grids, AdvObj, BaseGrid,
  AdvGrid, AdvCGrid,
  UnitCommons, UnitConsts;

type
  TfrmDevice = class(TfrmWork)
    acgDeviceList: TAdvColumnGrid;
    procedure FormCreate(Sender: TObject);
    procedure acgDeviceListGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure acgDeviceListDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
  private
    { Private declarations }
    procedure Initialize;
    procedure Finalize;

    procedure InitializeDeviceListGrid;
  public
    { Public declarations }
    procedure SetDeviceStatus(ASource: PSource; AStatus: TDeviceStatus);
  end;

var
  frmDevice: TfrmDevice;

implementation

{$R *.dfm}

procedure TfrmDevice.FormCreate(Sender: TObject);
begin
  inherited;
  Initialize;
end;

procedure TfrmDevice.Initialize;
begin
  InitializeDeviceListGrid;
end;

procedure TfrmDevice.acgDeviceListDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  R: TRect;
begin
  inherited;
  if (gdSelected in State) then
  begin
    with (Sender as TAdvColumnGrid) do
    begin
      R := CellRect(ColCount - 1, ARow);
      R.Left := 0;

      if (LeftCol > 0) then
      begin
        R.Left := -1;
      end;

      Canvas.Brush.Style := bsClear;
      Canvas.Pen.Color := clRed;
      Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom);
//      Canvas.Rectangle(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom);
    end;
  end;
end;

procedure TfrmDevice.acgDeviceListGetCellColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
begin
  inherited;
  if (ARow < CNT_DEVICE_HEADER) then
  begin
    with (Sender as TAdvColumnGrid) do
      AFont.Color := FixedFont.Color;
  end;
end;

procedure TfrmDevice.Finalize;
begin

end;

procedure TfrmDevice.InitializeDeviceListGrid;
var
  I: Integer;
  Column: TGridColumnItem;
  Source: PSourceHandle;
begin
  with acgDeviceList do
  begin
    BeginUpdate;
    try
      FixedRows  := CNT_DEVICE_HEADER;
      RowCount   := GV_SourceList.Count + CNT_DEVICE_HEADER;
      ColCount   := CNT_DEVICE_COLUMNS + GV_DCSList.Count * 2;
      VAlignment := vtaCenter;

      Columns.BeginUpdate;
      try
        Columns.Clear;
        for I := 0 to ColCount - 1 do
        begin
          Column := Columns.Add;
          with Column do
          begin
            HeaderFont.Assign(acgDeviceList.FixedFont);
            Font.Assign(acgDeviceList.Font);

            // Column : No
            if (I = IDX_COL_DEVICE_NO) then
            begin
              Alignment  := taLeftJustify;
              Borders    := [];
              Header     := NAM_COL_DEVICE_NO;
              HeaderAlignment := taCenter;
              ReadOnly   := True;
              VAlignment := vtaCenter;
              Width      := WIDTH_COL_DEVICE_NO;
            end
            // Column : Name
            else if (I = IDX_COL_DEVICE_NAME) then
            begin
              Alignment  := taLeftJustify;
              Header     := NAM_COL_DEVICE_NAME;
              ReadOnly   := True;
              VAlignment := vtaCenter;
              Width      := WIDTH_COL_DEVICE_NAME;
            end
            // Column Status
            else if (((I - CNT_DEVICE_COLUMNS) mod 2) = 0) then
            begin
              Alignment := taLeftJustify;
//              Header    := NAM_COL_DEVICE_STATUS;
              ReadOnly  := True;
              Width     := WIDTH_COL_DEVICE_STATUS;
            end
            // Column Timecode
            else if (((I - CNT_DEVICE_COLUMNS) mod 2) = 1) then
            begin
              Alignment := taLeftJustify;
//              Header    := NAM_COL_DEVICE_TIMECODE;
              ReadOnly  := True;
              Width     := WIDTH_COL_DEVICE_TIMECODE;
            end;
          end;
        end;

        // Merge Column : No
        MergeCells(IDX_COL_DEVICE_NO, 0, 1, CNT_DEVICE_HEADER);
        // Merge Column : Name
        MergeCells(IDX_COL_DEVICE_NAME, 0, 1, CNT_DEVICE_HEADER);

//        CellProperties[IDX_COL_DEVICE_NO, 0].VAlignment := vtaCenter;
//        CellProperties[IDX_COL_DEVICE_NAME, 0].VAlignment := vtaCenter;

        for I := 0 to GV_DCSList.Count - 1 do
        begin
          // Merge Column : DCS
          MergeCells(IDX_COL_DEVICE_NAME + (I * 2) + 1, 0, 2, 1);

          Cells[IDX_COL_DEVICE_NAME + (I * 2) + 1, 0] := String(GV_DCSList[I].Name);

          Cells[IDX_COL_DEVICE_NAME + (I * 2) + 1, 1] := NAM_COL_DEVICE_STATUS;
          Cells[IDX_COL_DEVICE_NAME + (I * 2) + 2, 1] := NAM_COL_DEVICE_TIMECODE;
        end;

        for I := 0 to GV_SourceList.Count - 1 do
        begin
          Cells[IDX_COL_DEVICE_NO, I + CNT_DEVICE_HEADER] := Format('%d', [I + 1]);
          Cells[IDX_COL_DEVICE_NAME, I + CNT_CUESHEET_HEADER + 1] := String(GV_SourceList[I]^.Name);
        end;
      finally
        Columns.EndUpdate;
      end;
    finally
      EndUpdate;
    end;
  end;
end;

procedure TfrmDevice.SetDeviceStatus(ASource: PSource; AStatus: TDeviceStatus);
begin
  if (ASource = nil) then exit;

  case AStatus.EventType of
    ET_PLAYER:
    begin

    end;
  end;
end;

end.
