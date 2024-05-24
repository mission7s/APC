unit UnitChannelTimeline;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitWorkForm, WMTimeLine, WMTools,
  Vcl.StdCtrls, Vcl.Imaging.pngimage, WMControls, Vcl.ExtCtrls,
  UnitCommons, UnitConsts;

type
  TfrmChannelTimeline = class(TfrmWork)
    lblChannel: TLabel;
    wmibTimelineZoomIn: TWMImageSpeedButton;
    wmtbTimelineZoom: TWMTrackBar;
    wmtlPlaylist: TWMTimeLine;
    Label2: TLabel;
    lblPlayedTime: TLabel;
    Label3: TLabel;
    lblRemainingTime: TLabel;
    Label5: TLabel;
    lblNextStart: TLabel;
    Label1: TLabel;
    lblPlayingInfo: TLabel;
    Label6: TLabel;
    lblNextInfo: TLabel;
    wmibTimelineZoomOut: TWMImageSpeedButton;
    lblOnAirFlag: TLabel;
    procedure pnlDescMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure wmtbTimelineZoomChange(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
    FChannelID: Word;
    FChannelOnAir: Boolean;

    FTimeLineZoomType: TTimelineZoomType;
    FTimeLineZoomPosition: Integer;

    FChannelForm: TfrmChannelTimeline;

    FPlayedTimeStr: String;
    FRemainingTimeStr: String;
    FNextStartTimeStr: String;

    FTimelineStartDate: TDateTime;
    FTimelineEndDate: TDateTime;

    FFrameNumber: Integer;

    function GetPositionByZoomType(AZoomType: TTimelineZoomType): Integer;
    function GetZoomTypeByPosition(APosition: Integer): TTimelineZoomType;
    procedure UpdateZoomPosition(APosition: Integer);
    procedure SetZoomPosition(Value: Integer);

    procedure Initialize;
    procedure Finalize;

    procedure InitializePlayListTimeLine;
  protected
    procedure WndProc(var Message: TMessage); override;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent; AChannelID: Word; ACombine: Boolean; ALeft, ATop, AWidth, AHeight: Integer); overload;

    procedure TimelineZoomIn;
    procedure TimelineZoomOut;

    procedure CalcuratePlayListTimeLineRange(AStartDate, AEndDate: TDateTime);
    procedure UpdatePlayListTimeLineRange(ATimeLine: TWMTimeLine);
    procedure PopulatePlayListTimeLine(ACompIndex: Integer; ATrack: TTrack);
//    procedure PopulatePlayListTimeLine(ATimeLine: TWMTimeLine);
    procedure DisplayPlayListTimeLine(ATimeLine: TWMTimeLine);
    procedure DeletePlayListTimeLine(ACompIndex: Integer; AData: Pointer);

    procedure ErrorDisplayPlayListTimeLine(AErrorDisplayEnabled: Boolean);

    procedure ClearPlayListTimeLine;

    procedure SetChannelOnAir(AOnAir: Boolean);
    procedure SetChannelTime(APlayedTime, ARemainingTime, ANextStartTime: String; AFrameNumber: Integer);
    procedure SetCueSheetCurr(AItem: PCueSheetItem);
    procedure SetCueSheetNext(AItem: PCueSheetItem);

    property TimelineZoomPosition: Integer read FTimelineZoomPosition write SetZoomPosition;
    property TimelineZoomType: TTimelineZoomType read FTimelineZoomType;
  end;

var
  frmChannelTimeline: TfrmChannelTimeline;

implementation

uses UnitSEC;

{$R *.dfm}

procedure TfrmChannelTimeline.WndProc(var Message: TMessage);
begin
  inherited;

  case Message.Msg of
    WM_UPDATE_CHANNEL_TIME:
    begin
      lblPlayedTime.Caption    := FPlayedTimeStr;
      lblRemainingTime.Caption := FRemainingTimeStr;
      lblNextStart.Caption     := FNextStartTimeStr;

      wmtlPlaylist.FrameNumber := FFrameNumber;
    end;
  end;
end;

procedure TfrmChannelTimeline.Initialize;
begin
  Tag := FChannelID;

  lblChannel.Caption := GetChannelNameByID(FChannelID);

  FPlayedTimeStr    := IDLE_TIME;
  FRemainingTimeStr := IDLE_TIME;
  FNextStartTimeStr := IDLE_TIMECODE;

  lblPlayedTime.Caption     := FPlayedTimeStr;
  lblRemainingTime.Caption  := FRemainingTimeStr;
  lblNextStart.Caption      := FNextStartTimeStr;

  lblPlayingInfo.Caption := '-';
  lblNextInfo.Caption    := '-';

  FFrameNumber := 0;

  wmtbTimelineZoom.Min := 0;
  wmtbTimelineZoom.Max := Round(SecsPerDay * wmtlPlaylist.TimeZoneProperty.FrameRate) + 1;
  wmtbTimelineZoom.Position := wmtbTimelineZoom.Max;
  FTimeLineZoomPosition := wmtbTimelineZoom.Position;

  InitializePlayListTimeLine;
end;

procedure TfrmChannelTimeline.Finalize;
begin

end;

procedure TfrmChannelTimeline.InitializePlayListTimeLine;
begin
  with wmtlPlaylist do
  begin
    with TimeZoneProperty do
    begin
      FrameDayReset := True;
      FrameRate := FrameRate29_97;
      FrameSkip := 600;
      FrameStep := 15;
      RailBarVisible := False;
    end;

    FrameSelectEnabled := False;
    TrackSelectEnabled := False;
    TrackTrimEnabled := False;

    VideoGroupProperty.Count := 0;
    AudioGroupProperty.Count := 0;

    DataGroupProperty.Count  := 3;

    DataCompositions[0].Height := 5;
  end;

  FTimelineStartDate := Date;
  FTimelineEndDate   := Date;

  UpdateZoomPosition(FTimeLineZoomPosition);
end;

procedure TfrmChannelTimeline.FormCreate(Sender: TObject);
begin
  inherited;
  Initialize;
end;

procedure TfrmChannelTimeline.FormDestroy(Sender: TObject);
begin
  inherited;
  Finalize;
end;

procedure TfrmChannelTimeline.FormResize(Sender: TObject);
begin
  inherited;
  with wmtlPlayList.ZoomBarProperty do
  begin
    TimelineZoomPosition := GetPositionByZoomType(FTimelineZoomType);
  end;
end;

constructor TfrmChannelTimeline.Create(AOwner: TComponent; AChannelID: Word; ACombine: Boolean; ALeft, ATop, AWidth, AHeight: Integer);
begin
  inherited Create(AOwner, ACombine, ALeft, ATop, AWidth, AHeight);

  FChannelID := AChannelID;
end;

procedure TfrmChannelTimeline.pnlDescMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  pnlDesc.SetFocus;
end;

function TfrmChannelTimeline.GetPositionByZoomType(AZoomType: TTimelineZoomTYpe): Integer;
var
  RatePerFrame: Word;
begin
  Result := 0;

  with wmtlPlaylist.TimeZoneProperty do
  begin
    case AZoomType of
      zt1Second: Result := Round(1 * FrameRate);
      zt2Seconds: Result := Round(2 * FrameRate);
      zt5Seconds: Result := Round(5 * FrameRate);
      zt10Seconds: Result := Round(10 * FrameRate);
      zt15Seconds: Result := Round(15 * FrameRate);
      zt30Seconds: Result := Round(30 * FrameRate);
      zt1Minute: Result := Round(SecsPerMin * FrameRate);
      zt2Minutes: Result := Round(2 * SecsPerMin * FrameRate);
      zt5Minutes: Result := Round(5 * SecsPerMin * FrameRate);
      zt10Minutes: Result := Round(10 * SecsPerMin * FrameRate);
      zt15Minutes: Result := Round(15 * SecsPerMin * FrameRate);
      zt30Minutes: Result := Round(30 * SecsPerMin * FrameRate);
      zt1Hour: Result := Round(SecsPerHour * FrameRate);
      zt2Hours: Result := Round(2 * SecsPerHour * FrameRate);
      zt6Hours: Result := Round(6 * SecsPerHour * FrameRate);
      zt12Hours: Result := Round(12 * SecsPerHour * FrameRate);
      zt1Day: Result := Round(SecsPerDay * FrameRate);
      ztFit: Result := Round(SecsPerDay * FrameRate) + 1;
    end;
  end;
end;

function TfrmChannelTimeline.GetZoomTypeByPosition(APosition: Integer): TTimelineZoomType;
begin
  Result := ztNone;
  with wmtlPlaylist.TimeZoneProperty do
  begin
    if (APosition <= Round(1 * FrameRate)) then
      Result := zt1Second
    else if (APosition <= Round(2 * FrameRate)) then
      Result := zt2Seconds
    else if (APosition <= Round(5 * FrameRate)) then
      Result := zt5Seconds
    else if (APosition <= Round(10 * FrameRate)) then
      Result := zt10Seconds
    else if (APosition <= Round(15 * FrameRate)) then
      Result := zt15Seconds
    else if (APosition <= Round(30 * FrameRate)) then
      Result := zt30Seconds
    else if (APosition <= Round(SecsPerMin * FrameRate)) then
      Result := zt1Minute
    else if (APosition <= Round(2 * SecsPerMin * FrameRate)) then
      Result := zt2Minutes
    else if (APosition <= Round(5 * SecsPerMin * FrameRate)) then
      Result := zt5Minutes
    else if (APosition <= Round(10 * SecsPerMin * FrameRate)) then
      Result := zt10Minutes
    else if (APosition <= Round(15 * SecsPerMin * FrameRate)) then
      Result := zt15Minutes
    else if (APosition <= Round(30 * SecsPerMin * FrameRate)) then
      Result := zt30Minutes
    else if (APosition <= Round(SecsPerHour * FrameRate)) then
      Result := zt1Hour
    else if (APosition <= Round(2 * SecsPerHour * FrameRate)) then
      Result := zt2Hours
    else if (APosition <= Round(6 * SecsPerHour * FrameRate)) then
      Result := zt6Hours
    else if (APosition <= Round(12 * SecsPerHour * FrameRate)) then
      Result := zt12Hours
    else if (APosition <= Round(SecsPerDay * FrameRate)) then
      Result := zt1Day
    else if (APosition <= Round(SecsPerDay * FrameRate) + 1) then
      Result := ztFit;
  end;
end;

procedure TfrmChannelTimeline.UpdateZoomPosition(APosition: Integer);
var
  SampleTime: Double;
  Frames: Integer;
begin
  if (FTimelineZoomType = ztFit) then
  begin
    wmtlPlaylist.ZoomToFit;
  end
  else
  begin
    with wmtlPlaylist.TimeZoneProperty do
    begin
      BeginUpdate;
      try
//      SampleTime := FrameToSampleTime(APosition, Round(FrameRate));
//      SampleTime := Round(APosition / FrameRate);

        SampleTime := Round(APosition / FrameRate);
        if (SampleTime < 1) then SampleTime := 1;

        Frames := TimecodeToFrame(SecondToTimeCode(SampleTime, FrameRate), FrameRate);

        if (SampleTime >= 1) and (SampleTime < 15) then
        begin
          FrameGap := 12;
          FrameStep := 10;
          FrameSkip := Frames div 2;
        end
        else if (SampleTime >= 15) and (SampleTime < SecsPerMin) then
        begin
          FrameGap := 6;
          FrameStep := 20;
          FrameSkip := Frames div 5;
        end
        else if (SampleTime >= SecsPerMin) and (SampleTime < SecsPerMin * 15) then
        begin
          FrameGap := 24;
          FrameStep := 5;
          FrameSkip := Frames;
        end
        else if (SampleTime >= SecsPerMin * 15) and (SampleTime < SecsPerMin * 30) then
        begin
          FrameGap := 12;
          FrameStep := 12;
          FrameSkip := Frames div 3;
        end
        else if (SampleTime >= SecsPerMin * 30) and (SampleTime < SecsPerHour) then
        begin
          FrameGap := 8;
          FrameStep := 36;
          FrameSkip := Frames div 6;
        end
        else if (SampleTime >= SecsPerHour) then
        begin
          FrameGap := 4;
          FrameStep := 36;
          FrameSkip := Frames div 6;
        end;
      finally
        EndUpdate;
      end;
  //    WMTimeLine.ViewSplitter;
  //    WMTimeLine.ViewAreaRepaint;
    end;
  end;
end;

procedure TfrmChannelTimeline.wmtbTimelineZoomChange(Sender: TObject);
begin
  inherited;
  TimelineZoomPosition := wmtbTimelineZoom.Position;
end;

procedure TfrmChannelTimeline.SetZoomPosition(Value: Integer);
var
  Pos: Integer;
begin
//  if (FZoomPosition <> Value) then
  begin
    FTimelineZoomPosition := Value;
//    wmtlPlaylist.ZoomBarProperty.Position := Value;
    wmtbTimelineZoom.Position := Value;

    FTimelineZoomType := GetZoomTypeByPosition(Value);
    UpdateZoomPosition(Value);
  end;
end;

procedure TfrmChannelTimeline.TimelineZoomIn;
var
  Zoom: Integer;
begin
  inherited;
  with wmtlPlayList.ZoomBarProperty do
  begin
    Zoom := Integer(FTimelineZoomType);
    Dec(Zoom);

    if (Zoom < Integer(Low(TTimelineZoomType)) + 1) then Zoom := Integer(Low(TTimelineZoomType)) + 1;

    FTimelineZoomType := TTimelineZoomType(Zoom);
    TimelineZoomPosition := GetPositionByZoomType(FTimelineZoomType);
  end;
end;

procedure TfrmChannelTimeline.TimelineZoomOut;
var
  Zoom: Integer;
begin
  with wmtlPlayList.ZoomBarProperty do
  begin
    Zoom := Integer(FTimelineZoomType);
    Inc(Zoom);

    if (Zoom > Integer(High(TTimelineZoomType))) then Zoom := Integer(High(TTimelineZoomType));

    FTimelineZoomType := TTimelineZoomType(Zoom);
    TimelineZoomPosition := GetPositionByZoomType(FTimelineZoomType);
  end;
end;

procedure TfrmChannelTimeline.CalcuratePlayListTimeLineRange(AStartDate, AEndDate: TDateTime);
begin
  FTimelineStartDate := AStartDate;
  FTimelineEndDate   := AEndDate;
end;

procedure TfrmChannelTimeline.UpdatePlayListTimeLineRange(ATimeLine: TWMTimeLine);
begin
  with wmtlPlaylist do
  begin
    TimeZoneProperty.FrameStart := ATimeLine.TimeZoneProperty.FrameStart;
    TimeZoneProperty.FrameCount := ATimeLine.TimeZoneProperty.FrameCount;
  end;

  UpdateZoomPosition(FTimelineZoomPosition);
end;

procedure TfrmChannelTimeline.PopulatePlayListTimeLine(ACompIndex: Integer; ATrack: TTrack);
var
  Track: TTrack;
begin
  if (ATrack = nil) then exit;

  with wmtlPlaylist do
  begin
    if (ACompIndex >= 0) and (ACompIndex < DataGroupProperty.Count) then
    begin
      Track := DataCompositions[ACompIndex].Tracks.GetTrackByData(ATrack.Data);
      if (Track = nil) then
        Track := DataCompositions[ACompIndex].Tracks.Add;

      Track.Assign(ATrack);
    end;
  end;
end;

{procedure TfrmChannelTimeline.PopulatePlayListTimeLine(ATimeLine: TWMTimeLine);
var
  I, J: Integer;
  DataComp: TDataComposition;
  Track: TTrack;
begin
  with wmtlPlaylist do
  begin
    BeginUpdateCompositions;
    try
      for I := 0 to DataGroupProperty.Count - 1 do
      begin
        if (I < ATimeLine.DataGroupProperty.Count) then
        begin
          DataComp := ATimeLine.DataCompositions[I];
          for J := 0 to DataComp.Tracks.Count - 1 do
          begin
            Track := DataCompositions[I].Tracks.GetTrackByData(DataComp.Tracks[J].Data);
            if (Track = nil) then
              Track := DataCompositions[I].Tracks.Add;

            Track.Assign(DataComp.Tracks[J]);
          end;
        end;
      end;
    finally
      EndUpdateCompositions;
    end;
  end;
end; }

procedure TfrmChannelTimeline.DisplayPlayListTimeLine(ATimeLine: TWMTimeLine);
begin
//  PopulatePlayListTimeLine(ATimeLine);

  UpdateZoomPosition(FTimelineZoomPosition);
end;

procedure TfrmChannelTimeline.DeletePlayListTimeLine(ACompIndex: Integer; AData: Pointer);
var
  Track: TTrack;
begin
  if (AData = nil) then exit;

  with wmtlPlaylist do
  begin
    if (ACompIndex >= 0) and (ACompIndex < DataGroupProperty.Count) then
    begin
      Track := DataCompositions[ACompIndex].Tracks.GetTrackByData(AData);
      if (Track <> nil) then FreeAndNil(Track);
    end;
  end;

//  UpdateZoomPosition(FTimelineZoomPosition);
end;

procedure TfrmChannelTimeline.ErrorDisplayPlayListTimeLine(AErrorDisplayEnabled: Boolean);
var
  I, J: Integer;
  Track: TTrack;
  Item: PCueSheetItem;
begin
  with wmtlPlaylist do
  begin
    BeginUpdateCompositions;
    try
      for I := 0 to DataGroupProperty.Count - 1 do
      begin
        for J := 0 to DataCompositions[I].Tracks.Count - 1 do
        begin
          Track := DataCompositions[I].Tracks[J];

          Item := Track.Data;
          if (Item <> nil) then
          begin
            if (Item^.EventStatus.State in [esError]) or
               (Item^.MediaStatus in [msNotExist, msShort]) then
            begin
              if (AErrorDisplayEnabled) then
              begin
                Track.Color        := COLOR_BK_EVENTSTATUS_ERROR;
                Track.ColorCaption := COLOR_BK_EVENTSTATUS_ERROR;
              end
              else
              begin
                Track.Color        := GetProgramTypeColorByCode(Item^.ProgramType);
                Track.ColorCaption := GetProgramTypeColorByCode(Item^.ProgramType);
              end;
            end;
          end;
        end;
      end;
    finally
      EndUpdateCompositions;
    end;
  end;
end;

procedure TfrmChannelTimeline.ClearPlayListTimeLine;
var
  I: Integer;
begin
  with wmtlPlaylist do
  begin
    BeginUpdateCompositions;
    try
      for I := 0 to DataGroupProperty.Count - 1 do
      begin
       DataCompositions[I].Tracks.Clear;
      end;
    finally
      EndUpdateCompositions;
    end;
  end;
end;

procedure TfrmChannelTimeline.SetChannelOnAir(AOnAir: Boolean);
begin
  lblOnAirFlag.Caption := OnAirFlagNames[AOnAir];
  if (AOnAir) then
  begin
    lblOnAirFlag.Font.Color := clLime;
  end
  else
  begin
    lblOnAirFlag.Font.Color := clRed;
  end;

  wmtlPlaylist.TimeZoneProperty.RailBarVisible := AOnAir;
end;

procedure TfrmChannelTimeline.SetChannelTime(APlayedTime, ARemainingTime, ANextStartTime: String; AFrameNumber: Integer);
begin
  FPlayedTimeStr    := APlayedTime;
  FRemainingTimeStr := ARemainingTime;
  FNextStartTimeStr := ANextStartTime;

  FFrameNumber := AFrameNumber;

  PostMessage(Handle, WM_UPDATE_CHANNEL_TIME, 0, 0);
end;

procedure TfrmChannelTimeline.SetCueSheetCurr(AItem: PCueSheetItem);
begin
  if (AItem = nil) then
    lblPlayingInfo.Caption := '-'
  else
    lblPlayingInfo.Caption := Format('%s / %s / %s', [AItem^.Title, AItem^.SubTitle, AItem^.MediaId]);
end;

procedure TfrmChannelTimeline.SetCueSheetNext(AItem: PCueSheetItem);
begin
  if (AItem = nil) then
    lblNextInfo.Caption := '-'
  else
    lblNextInfo.Caption := Format('%s / %s / %s', [AItem^.Title, AItem^.SubTitle, AItem^.MediaId]);
end;

end.
