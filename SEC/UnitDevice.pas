unit UnitDevice;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitWorkForm, Vcl.Imaging.pngimage,
  WMTools, WMControls, Vcl.ExtCtrls, AdvUtil, Vcl.Grids, AdvObj, BaseGrid,
  AdvGrid, AdvCGrid,
  UnitCommons, UnitConsts, UnitSECDLL, UnitMCCDLL;

type
  TfrmDevice = class(TfrmWork)
    acgDeviceList: TAdvColumnGrid;
    procedure FormCreate(Sender: TObject);
    procedure acgDeviceListGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure acgDeviceListDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
  private
  protected
    procedure WMUpdateDeviceCommError(var Message: TMessage); message WM_UPDATE_DEVICE_COMM_ERROR;
    procedure WMUpdateDeviceStatus(var Message: TMessage); message WM_UPDATE_DEVICE_STATUS;
  private
    { Private declarations }
    procedure Initialize;
    procedure Finalize;

    procedure InitializeDeviceListGrid;
    procedure PopulateDeviceListGrid(ASource: PSource; ADCSID: Integer = -1);
//    procedure PopulateDeviceListGrid(AIndex: Integer; AStatus: TDeviceStatus; ADCSID: Word);
  public
    { Public declarations }
    procedure SetDeviceCommError(ADeviceName: String; AStatus: TDeviceStatus);
//    procedure SetDeviceStatus(AIndex: Integer; ASource: PSource; AStatus: TDeviceStatus; ADCSID: Word);
    procedure SetDeviceStatus(ASource: PSource; ASourceHandle: PSourceHandle; AStatus: TDeviceStatus); overload;
    procedure SetDeviceStatus(ADCSIP: String; ADeviceHandle: TDeviceHandle; AStatus: TDeviceStatus); overload;




    function SECSetDeviceCommErrorW(ADeviceStatus: TDeviceStatus; ADeviceName: String): Integer;
    function SECSetDeviceStatusW(ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus): Integer;

    procedure ServerSetDeviceCommErrors(ADeviceStatus: TDeviceStatus; ADeviceName: String);
    procedure ServerSetDeviceStatuses(ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus);

    procedure MCCSetDeviceCommErrors(ADeviceStatus: TDeviceStatus; ADeviceName: String);
    procedure SECSetDeviceCommErrors(ADeviceStatus: TDeviceStatus; ADeviceName: String);

    procedure MCCSetDeviceStatuses(ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus);
    procedure SECSetDeviceStatuses(ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus);
  end;

var
  frmDevice: TfrmDevice;

implementation

uses UnitSEC, UnitChannel;

{$R *.dfm}

procedure TfrmDevice.WMUpdateDeviceCommError(var Message: TMessage);
var
  Source: PSource;
begin
  inherited;

  Source  := PSource(Message.WParam);
  PopulateDeviceListGrid(Source);
end;

procedure TfrmDevice.WMUpdateDeviceStatus(var Message: TMessage);
var
  Source: PSource;
  DCSID: Word;
begin
  inherited;

  Source  := PSource(Message.WParam);
  DCSID   := Message.LParam;

  PopulateDeviceListGrid(Source, DCSID);
end;

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
begin
  inherited;
  with (Sender as TAdvColumnGrid) do
  begin
    if not (gdFixed in State) and
       (((gdSelected in State) or (gdRowSelected in State)) or
        (RowSelect[ARow])) then
    begin
      Canvas.Brush.Style := bsClear;
      Canvas.Pen.Color := COLOR_ROW_SELECT_DEVICE;//clRed;

      Canvas.MoveTo(Rect.Left, Rect.Top);
      Canvas.LineTo(Rect.Right + 1, Rect.Top);

      Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
      Canvas.LineTo(Rect.Right + 1, Rect.Bottom - 1);

//      SetTextColor(Canvas.Handle, clWhite);

      if (ACol = FixedCols) then
      begin
        Canvas.MoveTo(Rect.Left, Rect.Top);
        Canvas.LineTo(Rect.Left, Rect.Bottom);

        if (IsMergedCell(ACol, ARow)) then
        begin
          Canvas.MoveTo(Rect.Right, Rect.Top);
          Canvas.LineTo(Rect.Right, Rect.Bottom);
        end;
      end
      else if (ACol = ColCount - FixedCols) then //or (IsMergedCell(ACol, ARow)) then
      begin
        Canvas.MoveTo(Rect.Right, Rect.Top);
        Canvas.LineTo(Rect.Right, Rect.Bottom);
      end;
    end
    else
    begin
//      SetTextColor(Canvas.Handle, Font.Color);
    end;
  end;

{  with (Sender as TAdvColumnGrid) do
  begin
    if (gdSelected in State) or (gdRowSelected in State) then
    begin
      Canvas.Brush.Style := bsClear;
      Canvas.Pen.Color := clRed;

      Canvas.MoveTo(Rect.Left, Rect.Top);
      Canvas.LineTo(Rect.Right + 1, Rect.Top);

      Canvas.MoveTo(Rect.Left, Rect.Bottom);
      Canvas.LineTo(Rect.Right + 1, Rect.Bottom);

      if (ACol = 0) then
      begin
        Canvas.MoveTo(Rect.Left, Rect.Top);
        Canvas.LineTo(Rect.Left, Rect.Bottom);
      end
      else if (ACol = ColCount - 1) or (IsMergedCell(ACol, ARow)) then
      begin
        Canvas.MoveTo(Rect.Right, Rect.Top);
        Canvas.LineTo(Rect.Right, Rect.Bottom);
      end;
    end;
  end; }
end;

procedure TfrmDevice.acgDeviceListGetCellColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
  RCol, RRow: Integer;
  Source: PSource;
begin
  inherited;

  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    if ((gdSelected in AState) or (gdRowSelected in AState)) or
       (RowSelect[RRow]) then
    begin
      AFont.Color := COLOR_BK_DEVICESTATUS_NORMAL;
      AFont.Color := clWhite;
    end;

    Source := GV_SourceList[RRow - CNT_DEVICE_HEADER];
    if (Source <> nil) then
    begin
      if (Source^.CommSuccess) then
      begin
        if (not Source^.Status.Connected) or
           ((Source^.Status.EventType = ET_PLAYER) and (Source^.Status.Player.PortDown)) then
        begin
          ABrush.Color := COLOR_BK_DEVICESTATUS_DISCONNECT;
          AFont.Color  := COLOR_TX_DEVICESTATUS_DISCONNECT;
        end
        else
        begin
          ABrush.Color := COLOR_BK_DEVICESTATUS_NORMAL;
          if (not (gdSelected in AState) and (not (gdRowSelected in AState))) and
             (not (RowSelect[RRow])) then
          begin
            AFont.Color  := COLOR_TX_DEVICESTATUS_NORMAL;
          end;
        end;
      end
      else
      begin
        ABrush.Color := COLOR_BK_DEVICE_COMM_ERROR;
        AFont.Color  := COLOR_TX_DEVICE_COMM_ERROR;
      end;
    end;
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
      ColCount   := CNT_DEVICE_COLUMNS;
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
            else if (I = IDX_COL_DEVICE_STATUS) then
            begin
              Alignment := taLeftJustify;
              Header    := NAM_COL_DEVICE_STATUS;
              ReadOnly  := True;
              Width     := WIDTH_COL_DEVICE_STATUS;
            end
            // Column Timecode
            else if (I = IDX_COL_DEVICE_TIMECODE) then
            begin
              Alignment := taCenter;
              Header    := NAM_COL_DEVICE_TIMECODE;
              ReadOnly  := True;
              Width     := WIDTH_COL_DEVICE_TIMECODE;
            end
            // Column DCS
            else if (I = IDX_COL_DEVICE_DCS) then
            begin
              Alignment := taLeftJustify;
              Header    := NAM_COL_DEVICE_DCS;
              ReadOnly  := True;
              Width     := WIDTH_COL_DEVICE_DCS;
{            end
            // Column Timecode
            else if (I = IDX_COL_DEVICE_CHANNEL) then
            begin
              Alignment := taLeftJustify;
              Header    := NAM_COL_DEVICE_CHANNEL;
              ReadOnly  := True;
              Width     := WIDTH_COL_DEVICE_CHANNEL; }
            end;
          end;
        end;

        for I := 0 to GV_SourceList.Count - 1 do
        begin
          Cells[IDX_COL_DEVICE_NO, I + CNT_DEVICE_HEADER]         := Format('%d', [I + 1]);
          Cells[IDX_COL_DEVICE_NAME, I + CNT_CUESHEET_HEADER]     := String(GV_SourceList[I]^.Name);
          Cells[IDX_COL_DEVICE_STATUS, I + CNT_CUESHEET_HEADER]   := '';
          Cells[IDX_COL_DEVICE_TIMECODE, I + CNT_CUESHEET_HEADER] := '';
          Cells[IDX_COL_DEVICE_DCS, I + CNT_CUESHEET_HEADER]      := '';
        end;
      finally
        Columns.EndUpdate;
      end;
    finally
      EndUpdate;
    end;
  end;
end;

procedure TfrmDevice.PopulateDeviceListGrid(ASource: PSource; ADCSID: Integer);
var
  R: Integer;
  StatusString: String;
begin
  if (ASource = nil) then exit;

  with acgDeviceList do
  begin
    R := DisplRowIndex(GV_SourceList.IndexOf(ASource) + CNT_CUESHEET_HEADER);

    if (R < FixedRows) or (R > RowCount - 1) then exit;

    if (ASource^.CommSuccess) then
    begin
      case ASource^.Status.EventType of
        ET_PLAYER:
        begin
          with ASource^.Status.Player do
          begin
            if (not ASource^.Status.Connected) then
            begin
              StatusString := 'Disconnected';

  //            RowColor[ARow]     := COLOR_ROW_DEVICE_ERROR;
  //            RowFontColor[ARow] := COLOR_FONT_DEVICE_ERROR;
            end
            else
            begin
  //            RowColor[ARow]     := COLOR_ROW_DEVICE_NORMAL;
  //            RowFontColor[ARow] := COLOR_FONT_DEVICE_NORMAL;

              if (PortDown) then
              begin
                StatusString := 'PortDown';
              end
              else if (CueDone) then
              begin
                if (Play) then
                  StatusString := 'Playing.CueDone'
                else if (Still) then
                  StatusString := 'Paused.CueDone'
                else StatusString := 'CueDone';
              end
              else if (Cue) then
              begin
                if (Play) then
                  StatusString := 'Playing.Cueing'
                else if (Still) then
                  StatusString := 'Paused.Cueing'
                else StatusString := 'CueDone';
              end
              else if (Stop) then StatusString := 'Stop'
              else if (Still) then StatusString := 'Paused'
              else if (Play) then StatusString := 'Playing'
              else if (Jog) then StatusString := 'Jog'
              else if (Shuttle) then StatusString := 'Shuttle'
              else if (FastFoward) then StatusString := 'FastFoward'
              else if (Rewind) then StatusString := 'Rewind'
              else if (Eject) then StatusString := 'Eject'
              else StatusString := 'Stop';
            end;

            Cells[IDX_COL_DEVICE_STATUS, R]   := StatusString;

            case ASource^.SourceType of
              ST_LINE,
              ST_CG: Cells[IDX_COL_DEVICE_TIMECODE, R] := '-';
              else
                Cells[IDX_COL_DEVICE_TIMECODE, R] := TimecodeToString(CurTC);
            end;
          end;
        end;
        ET_SWITCHER:
        begin
          if (not ASource^.Status.Connected) then
          begin
            StatusString := 'Disconnected';
          end
          else
          begin
            StatusString := 'Connected';
          end;

          Cells[IDX_COL_DEVICE_STATUS, R]   := StatusString;

          Cells[IDX_COL_DEVICE_TIMECODE, R] := '-';
        end;
      end;
    end
    else
    begin
      Cells[IDX_COL_DEVICE_STATUS, R]   := 'Comm error';
      Cells[IDX_COL_DEVICE_TIMECODE, R] := '';
    end;

    Cells[IDX_COL_DEVICE_DCS, R] := GetDCSNameByID(ADCSID);

//    Cells[IDX_COL_DEVICE_CHANNEL, R] := String(ASource^.Channel^.Name);

    RepaintRow(R);
  end;
end;

{procedure TfrmDevice.PopulateDeviceListGrid(AIndex: Integer; AStatus: TDeviceStatus; ADCSID: Word);
var
  R: Integer;
  StatusString: String;
begin
  with acgDeviceList do
  begin
    R := DisplRowIndex(AIndex + CNT_CUESHEET_HEADER);

    if (R < FixedRows) or (R > RowCount - 1) then exit;

    case AStatus.EventType of
      ET_PLAYER:
      begin
        with AStatus.Player do
        begin
          if (not Connected) then
          begin
            StatusString := 'Error';

//            RowColor[ARow]     := COLOR_ROW_DEVICE_ERROR;
//            RowFontColor[ARow] := COLOR_FONT_DEVICE_ERROR;
          end
          else
          begin
//            RowColor[ARow]     := COLOR_ROW_DEVICE_NORMAL;
//            RowFontColor[ARow] := COLOR_FONT_DEVICE_NORMAL;

            if (CueDone) then
            begin
              if (Play) then
                StatusString := 'Playing.CueDone'
              else if (Still) then
                StatusString := 'Paused.CueDone'
              else StatusString := 'CueDone';
            end
            else if (Cue) then
            begin
              if (Play) then
                StatusString := 'Playing.Cueing'
              else if (Still) then
                StatusString := 'Paused.Cueing'
              else StatusString := 'CueDone';
            end
            else if (Stop) then StatusString := 'Stop'
            else if (Still) then StatusString := 'Paused'
            else if (Play) then StatusString := 'Playing'
            else if (Jog) then StatusString := 'Jog'
            else if (Shuttle) then StatusString := 'Shuttle'
            else if (FastFoward) then StatusString := 'FastFoward'
            else if (Rewind) then StatusString := 'Rewind'
            else if (Eject) then StatusString := 'Eject'
            else StatusString := 'Stop';
          end;

          Cells[IDX_COL_DEVICE_STATUS, R]   := StatusString;
          Cells[IDX_COL_DEVICE_TIMECODE, R] := TimecodeToString(CurTC);
        end;
      end;
    end;

    Cells[IDX_COL_DEVICE_DCS, R] := GetDCSNameByID(ADCSID);
//    Cells[IDX_COL_DEVICE_CHANNEL, R] := String(ASource^.Channel^.Name);
  end;
end; }

{procedure TfrmDevice.SetDeviceStatus(AIndex: Integer; ASource: PSource; AStatus: TDeviceStatus; ADCSID: Word);
var
  R: Integer;
begin
  if (ASource = nil) then exit;

  R := AIndex + CNT_DEVICE_HEADER;
  if (R > (acgDeviceList.RowCount - 1)) then exit;

  PopulateDeviceListGrid(R, AStatus);

  with acgDeviceList do
  begin
    Cells[IDX_COL_DEVICE_DCS, R]     := GetDCSNameByID(ADCSID);
//    Cells[IDX_COL_DEVICE_CHANNEL, R] := String(ASource^.Channel^.Name);
  end;
end; }

procedure TfrmDevice.SetDeviceCommError(ADeviceName: String; AStatus: TDeviceStatus);
var
  Source: PSource;
begin
  inherited;

  Source := GetSourceByName(ADeviceName);
  if (Source <> nil) then
  begin
    Source^.Status := AStatus;
    PostMessage(Handle, WM_UPDATE_DEVICE_COMM_ERROR, NativeInt(Source), 0);

//    if (GV_SettingOption.OnAirCheckDeviceNotify) then
    begin
      ServerSetDeviceCommErrors(AStatus, ADeviceName);
    end;
  end;
end;

procedure TfrmDevice.SetDeviceStatus(ASource: PSource; ASourceHandle: PSourceHandle; AStatus: TDeviceStatus);
var
  ChannelForm: TfrmChannel;
begin
  if (ASource = nil) then exit;
  if (ASourceHandle = nil) then exit;
  if (ASourceHandle^.Handle <= INVALID_DEVICE_HANDLE) then exit;

  // 메인 이벤트의 장치가 Disconnect가 발생하면 메인/백업 이벤트 교체
  // 메인 이벤트의 장치가 PortDown이 발생하면 메인/백업 이벤트 교체
  // 메인 이벤트의 장치가 Player이고 Stop이 발생하면 발생하면 메인/백업 이벤트 교체
  if (ASource^.CommSuccess) and
     ((ASource^.Status.Connected) and (ASource^.Status.Connected <> AStatus.Connected)) or
     ((ASource^.SourceType in [ST_VSDEC, ST_VSENC, ST_VCR]) and (ASource^.Status.EventType = ET_PLAYER) and
      (((not ASource^.Status.Player.PortDown) and (ASource^.Status.Player.PortDown <> AStatus.Player.PortDown)) or
       ((ASource^.Status.Player.Play) and (AStatus.Player.Stop)))) then
  begin
    ChannelForm := frmSEC.GetChannelFormByID(ASource^.Channel^.ID);
    if (ChannelForm <> nil) then
    begin
      if (ASource^.Status.EventType = ET_PLAYER) then
        Assert(False, GetMainLogStr(lsError, Format('TfrmDevice.SetDeviceStatus main change action, Source = %s, Connected = %s, PortDown = %s, Stop = %s', [String(ASource^.Name), BoolToStr(AStatus.Connected, True), BoolToStr(AStatus.Player.PortDown, True), BoolToStr(AStatus.Player.Stop, True)])))
      else
        Assert(False, GetMainLogStr(lsError, Format('TfrmDevice.SetDeviceStatus main change action, Source = %s, Connected = %s', [String(ASource^.Name), BoolToStr(AStatus.Connected, True)])));

      ChannelForm.MainErrorSourceExchange(ASource);
    end;
  end;

  ASource^.CommSuccess := True;
  ASource^.CommTimeout := 0;
  ASource^.Status := AStatus;

  if (ASourceHandle^.DCS <> nil) then
  begin
    PostMessage(Handle, WM_UPDATE_DEVICE_STATUS, NativeInt(ASource), ASourceHandle^.DCS^.ID);

//    if (GV_SettingOption.OnAirCheckDeviceNotify) then
    begin
      ServerSetDeviceStatuses(ASourceHandle^.DCS^.ID, ASourceHandle^.Handle, AStatus);
    end;
  end;
end;

procedure TfrmDevice.SetDeviceStatus(ADCSIP: String; ADeviceHandle: TDeviceHandle; AStatus: TDeviceStatus);
var
  Source: PSource;
  ChannelForm: TfrmChannel;

  DCS: PDCS;
begin
  inherited;

  Source := GetSourceByIPAndHandle(ADCSIP, ADeviceHandle);
  if (Source = nil) then exit;

  // 메인 이벤트의 장치가 Disconnect가 발생하면 메인/백업 이벤트 교체
  // 메인 이벤트의 장치가 Player이고 PortDown이 발생하면 메인/백업 이벤트 교체
  if (Source^.CommSuccess) and
     ((Source^.Status.Connected) and (Source^.Status.Connected <> AStatus.Connected)) or
     ((Source^.SourceType in [ST_VSDEC, ST_VSENC, ST_VCR]) and (Source^.Status.EventType = ET_PLAYER) and
      (((not Source^.Status.Player.PortDown) and (Source^.Status.Player.PortDown <> AStatus.Player.PortDown)) or
       ((Source^.Status.Player.Play) and (AStatus.Player.Stop)))) then
  begin
    ChannelForm := frmSEC.GetChannelFormByID(Source^.Channel^.ID);
    if (ChannelForm <> nil) then
    begin
      if (Source^.Status.EventType = ET_PLAYER) then
        Assert(False, GetMainLogStr(lsError, Format('TfrmDevice.SetDeviceStatus main change action, Source = %s, Connected = %s, PortDown = %s, Stop = %s', [String(Source^.Name), BoolToStr(AStatus.Connected, True), BoolToStr(AStatus.Player.PortDown, True), BoolToStr(AStatus.Player.Stop, True)])))
      else
        Assert(False, GetMainLogStr(lsError, Format('TfrmDevice.SetDeviceStatus main change action, Source = %s, Connected = %s', [String(Source^.Name), BoolToStr(AStatus.Connected, True)])));

      ChannelForm.MainErrorSourceExchange(Source);
    end;
  end;

  Source^.CommSuccess := True;
  Source^.CommTimeout := 0;
  Source^.Status := AStatus;

  DCS := GetDCSByIP(ADCSIP);
  if (DCS <> nil) then
  begin
    PostMessage(Handle, WM_UPDATE_DEVICE_STATUS, NativeInt(Source), DCS^.ID);

//    if (GV_SettingOption.OnAirCheckDeviceNotify) then
    begin
      ServerSetDeviceStatuses(DCS^.ID, ADeviceHandle, AStatus);
    end;
  end;
end;

function TfrmDevice.SECSetDeviceCommErrorW(ADeviceStatus: TDeviceStatus; ADeviceName: String): Integer;
var
  Source: PSource;
begin
  Result := D_FALSE;

  Source := GetSourceByName(ADeviceName);
  if (Source = nil) then exit;

  Source^.CommSuccess := False;
  Source^.Status := ADeviceStatus;
  PostMessage(Handle, WM_UPDATE_DEVICE_COMM_ERROR, NativeInt(Source), 0);

  Result := D_OK;
end;

function TfrmDevice.SECSetDeviceStatusW(ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus): Integer;
var
  Source: PSource;
  ChannelForm: TfrmChannel;

  DCS: PDCS;
begin
  inherited;
  Result := D_FALSE;

  Source := GetSourceByIDAndHandle(ADCSID, ADeviceHandle);
  if (Source = nil) then exit;

  Source^.CommSuccess := True;
  Source^.CommTimeout := 0;
  Source^.Status := ADeviceStatus;

{  // 메인 이벤트의 장치가 Disconnect가 발생하면 메인/백업 이벤트 교체
  if ((Source.Status.Connected) and (Source.Status.Connected <> ADeviceStatus.Connected)) then
  begin
    ChannelForm := frmSEC.GetChannelFormByID(Source^.Channel^.ID);
    if (ChannelForm <> nil) then
    begin
      ChannelForm.MainErrorSourceExchange(Source);
    end;
  end; }

  PostMessage(Handle, WM_UPDATE_DEVICE_STATUS, NativeInt(Source), ADCSID);

  Result := D_OK;
end;

procedure TfrmDevice.ServerSetDeviceCommErrors(ADeviceStatus: TDeviceStatus; ADeviceName: String);
begin
//  if (frmAllChannels <> nil) then
//    frmAllChannels.SetDeviceCommError(AStatus, ADeviceName);

  MCCSetDeviceCommErrors(ADeviceStatus, ADeviceName);
  SECSetDeviceCommErrors(ADeviceStatus, ADeviceName);
end;

procedure TfrmDevice.ServerSetDeviceStatuses(ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus);
begin
//  if (frmAllChannels <> nil) then
//    frmAllChannels.SetEventStatus(AEventID, AEventStatus);

  MCCSetDeviceStatuses(ADCSID, ADeviceHandle, ADeviceStatus);
  SECSetDeviceStatuses(ADCSID, ADeviceHandle, ADeviceStatus);
end;

procedure TfrmDevice.MCCSetDeviceCommErrors(ADeviceStatus: TDeviceStatus; ADeviceName: String);
var
  I: Integer;
  MCC: PMCC;
begin
  if (not GV_SettingMCC.Use) then exit;
  if (not HasMainControl) then exit;

  for I := 0 to GV_MCCList.Count - 1 do
  begin
    MCC := GV_MCCList[I];
    if (MCC <> nil) and (MCC^.Alive) then
    begin
      frmSEC.MCCEventThread.SetDeviceCommError(MCC, ADeviceStatus, ADeviceName);
    end;
  end;
end;

procedure TfrmDevice.MCCSetDeviceStatuses(ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus);
var
  I: Integer;
  MCC: PMCC;
begin
  if (not GV_SettingMCC.Use) then exit;
  if (not HasMainControl) then exit;

  for I := 0 to GV_MCCList.Count - 1 do
  begin
    MCC := GV_MCCList[I];
    if (MCC <> nil) and (MCC^.Alive) then
    begin
      frmSEC.MCCEventThread.SetDeviceStatus(MCC, ADeviceStatus, ADCSID, ADeviceHandle);
    end;
  end;
end;

procedure TfrmDevice.SECSetDeviceCommErrors(ADeviceStatus: TDeviceStatus; ADeviceName: String);
var
  I: Integer;
  SEC: PSEC;
begin
  if (not HasMainControl) then exit;

  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (SEC <> nil) and (SEC <> GV_SECMine) and (SEC^.Alive) then
    begin
//      SECSetDeviceCommError(SEC^.HostIP, ADeviceStatus, PChar(ADeviceName));
      frmSEC.SECEventThread.SetDeviceCommError(SEC, ADeviceStatus, ADeviceName);
    end;
  end;
end;

procedure TfrmDevice.SECSetDeviceStatuses(ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus);
var
  I: Integer;
  SEC: PSEC;
begin
  if (not HasMainControl) then exit;

  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (SEC <> nil) and (SEC <> GV_SECMine) and (SEC^.Alive) then
    begin
//      SECSetDeviceStatus(SEC^.HostIP, ADCSID, ADeviceHandle, ADeviceStatus);
      frmSEC.SECEventThread.SetDeviceStatus(SEC, ADeviceStatus, ADCSID, ADeviceHandle);
    end;
  end;
end;

end.
