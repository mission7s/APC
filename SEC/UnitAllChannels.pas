unit UnitAllChannels;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitWorkForm, AdvUtil, Vcl.Grids,
  AdvObj, BaseGrid, AdvGrid, AdvCGrid, Vcl.Imaging.pngimage,
  Vcl.ExtCtrls, Vcl.StdCtrls, System.Generics.Collections,
  AdvOfficePager, AdvSplitter,
  WMTools, WMUtils, WMControls, WMTimeLine, {LibXmlParserU, }Xml.VerySimple,
  UnitCommons, UnitDCSDLL, UnitConsts, Vcl.ComCtrls,
  UnitAllChannelsPanel;

type
  TAllChannelsTimerThread = class;

  TfrmAllChannels = class(TfrmWork)
    WMPanel8: TWMPanel;
    wmtlPlaylist: TWMTimeLine;
    pnlChannel: TWMPanel;
    pnlChannelCaption: TWMPanel;
    pnlChannelList: TWMPanel;
    AdvSplitter2: TAdvSplitter;
    WMPanel2: TWMPanel;
    wmibTimelineGotoCurrent: TWMImageSpeedButton;
    wmibTimelineMoveLeft: TWMImageSpeedButton;
    wmibTimelineMoveRight: TWMImageSpeedButton;
    Image1: TImage;
    wmibTimelineZoomIn: TWMImageSpeedButton;
    wmtbTimelineZoom: TWMTrackBar;
    wmibTimelineZoomOut: TWMImageSpeedButton;
    imgTimeline: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure wmtlPlaylistTrackHintEvent(Sender: TObject; Track: TTrack;
      var HintStr: string);
    procedure wmtlPlaylistMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure wmtbTimelineZoomChange(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure wmtlPlaylistTrackCustomDrawEvent(Sender: TObject; ATrack: TTrack;
      ACanvas: TCanvas; ARect: TRect; var ACustomDraw: Boolean);
    procedure wmtlPlaylistDataGroupVertScrollEvent(Sender: TObject;
      ScrollPos: Integer);
  private
    { Private declarations }
    FChannelPanelList: TList<TfrmAllChannelsPanel>;

    FPanelInstance: Pointer;
    FDefPanelProc: Pointer;
    FPanelHandle: HWND;

    FTimelineSpace: Integer;

    FTimelineStartDate: TDateTime;

    FTimeLineDaysPerFrames: Integer;
    FTimeLineMin: Integer;
    FTimeLineMax: Integer;

    FTimeLineZoomType: TTimeLineZoomType;
    FTimeLineZoomPosition: Integer;

    FErrorDisplayEnabled: Boolean;

    FTimerThread: TAllChannelsTimerThread;

    function GetPositionByZoomType(AZoomType: TTimeLineZoomType): Integer;
    function GetZoomTypeByPosition(APosition: Integer): TTimeLineZoomType;
    procedure UpdateZoomPosition(APosition: Integer);
    procedure SetZoomPosition(Value: Integer);

    procedure Initialize;
    procedure Finalize;

    procedure InitializeTimeline;

    function GetTimelinChannelByChannelID(AChannelID: Word): TfrmAllChannelsPanel;

    function GetAllChannelsCompIndexByChannelID(AChannelID: Word): Integer;

    procedure DisplayChannelTime(AChannelID: Word);

    procedure PanelWndProc(var Message: TMessage);
  protected
//    procedure WndProc(var Message: TMessage); override;

    procedure WMUpdateChannelTime(var Message: TMessage); message WM_UPDATE_CHANNEL_TIME;
    procedure WMUpdateEventStatus(var Message: TMessage); message WM_UPDATE_EVENT_STATUS;
    procedure WMUpdateErrorDisplay(var Message: TMessage); message WM_UPDATE_ERROR_DISPLAY;
    procedure WMSetOnAir(var Message: TMessage); message WM_SET_ONAIR;

    procedure WMBeginUpdate(var Message: TMessage); message WM_BEGIN_UPDATE;
    procedure WMInsertCueSheet(var Message: TMessage); message WM_INSERT_CUESHEET;
    procedure WMUpdateCueSheet(var Message: TMessage); message WM_UPDATE_CUESHEET;
    procedure WMDeleteCueSheet(var Message: TMessage); message WM_DELETE_CUESHEET;
    procedure WMClearCueSheet(var Message: TMessage); message WM_CLEAR_CUESHEET;
    procedure WMEndUpdate(var Message: TMessage); message WM_END_UPDATE;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent; ACombine: Boolean; ALeft, ATop, AWidth, AHeight: Integer); overload;
    procedure AdjustAllChannelsChannel;

    procedure PopulatePlayListTimeline(AItem: PCuesheetItem);
    procedure PopulateEventStatusPlayListTimeline(AIndex: Integer; AItem: PCueSheetItem);

    procedure UpdateChangeDatePlayListTimeLine(ADays: Integer);
    procedure RealtimeChangePlayListTimeLine;
    procedure DeletePlayListTimeline(AItem: PCueSheetItem);

    procedure ErrorDisplayPlayListTimeLine(AErrorDisplayEnabled: Boolean);

    procedure ClearPlayListTimeline(AChannelID: Word);

    procedure TimelineGotoCurrent;
    procedure TimelineMoveLeft;
    procedure TimelineMoveRight;

    procedure TimelineZoomIn;
    procedure TimelineZoomOut;

    procedure SetChannelOnAir(AChannelID: Word; AOnAir: Boolean);
    procedure SetChannelTime(AChannelID: Word; APlayedTime, ARemainingTime, ANextStart, ANextDuration: String);

    procedure SetEventStatus(AEventID: TEventID; AStatus: TEventStatus; AIsCurrEvent: Boolean = False);
    procedure SetEventOverall(ADCSIP: String; ADeviceHandle: TDeviceHandle; AOverall: TEventOverall);

    procedure SetMediaStatus(AEventID: TEventID; AStatus: TMediaStatus);

    procedure InputCueSheets(ACueSheetList: TCueSheetList; AChannelID: Word; AIndex: Integer; ACount: Integer); overload;
    procedure InputCueSheets(AIndex: Integer; AItem: PCueSheetItem); overload;

    procedure DeleteCueSheets(AChannelID: Word; ADeleteList: TCueSheetList); overload;
    procedure DeleteCueSheets(AEventID: TEventID); overload;
    procedure DeleteCueSheets(ACueSheetList: TCueSheetList; AChannelID: Word; AFromIndex, AToIndex: Integer); overload;

    procedure ClearCueSheets(AChannelID: Word);

    procedure SetCueSheetCurrs(AEventID: TEventID);
    procedure SetCueSheetNexts(AEventID: TEventID);
    procedure SetCueSheetTargets(AEventID: TEventID);

    property TimelineZoomPosition: Integer read FTimelineZoomPosition write SetZoomPosition;
    property TimelineZoomType: TTimeLineZoomType read FTimelineZoomType;
  end;

  TAllChannelsTimerThread = class(TThread)
  private
    FAllChannelsForm: TfrmAllChannels;

    FCloseEvent: THandle;

    procedure DoCueSheetCheck;
  protected
    procedure Execute; override;
  public
    constructor Create(AAllChannelsForm: TfrmAllChannels);
    destructor Destroy; override;

    procedure Terminate;
  end;

var
  frmAllChannels: TfrmAllChannels;

implementation

uses UnitSEC, System.DateUtils, System.Math;

{$R *.dfm}

procedure TfrmAllChannels.WMUpdateChannelTime(var Message: TMessage);
begin
  RealtimeChangePlayListTimeLine;
end;

procedure TfrmAllChannels.WMUpdateEventStatus(var Message: TMessage);
var
  Index: Integer;
  Item: PCueSheetItem;
begin
  Index := Message.WParam;
  Item := PCueSheetItem(Message.LParam);

  if (Item <> nil) and (Item^.EventStatus.State = esSkipped) then
    DeletePlayListTimeline(Item)
  else
    PopulatePlayListTimeline(Item);
end;

procedure TfrmAllChannels.WMUpdateErrorDisplay(var Message: TMessage);
begin
  FErrorDisplayEnabled := not FErrorDisplayEnabled;
end;

procedure TfrmAllChannels.WMSetOnAir(var Message: TMessage);
var
  ChannelID: Word;
  IsOnAir: Boolean;
begin
  ChannelID := Message.WParam;
  IsOnAir := Boolean(Message.LParam);

  SetChannelOnAir(ChannelID, IsOnAir);
end;

procedure TfrmAllChannels.WMBeginUpdate(var Message: TMessage);
begin
  wmtlPlaylist.BeginUpdateCompositions;
end;

procedure TfrmAllChannels.WMInsertCueSheet(var Message: TMessage);
var
  Index: Integer;
  Item: PCueSheetItem;
begin
  Index := Message.WParam;
  Item := PCueSheetItem(Message.LParam);

  PopulatePlayListTimeline(Item);
end;

procedure TfrmAllChannels.WMUpdateCueSheet(var Message: TMessage);
var
  Index: Integer;
  Item: PCueSheetItem;
begin
  Index := Message.WParam;
  Item := PCueSheetItem(Message.LParam);

  PopulatePlayListTimeline(Item);
end;

procedure TfrmAllChannels.WMDeleteCueSheet(var Message: TMessage);
var
  Index: Integer;
  Item: PCueSheetItem;
begin
  Index := Message.WParam;
  Item := PCueSheetItem(Message.LParam);

  DeletePlayListTimeline(Item);
end;

procedure TfrmAllChannels.WMClearCueSheet(var Message: TMessage);
var
  ChannelID: Word;
begin
  ChannelID := Message.WParam;

  ClearPlayListTimeline(ChannelID);
end;

procedure TfrmAllChannels.WMEndUpdate(var Message: TMessage);
begin
  wmtlPlaylist.EndUpdateCompositions;
end;

//procedure TfrmAllChannels.WndProc(var Message: TMessage);
//var
//  Index: Integer;
//
//  Item: PCueSheetItem;
//  ChannelID: Word;
//
//  EventID: TEventID;
//
//  IsOnAir: Boolean;
//
//  NextStartTime: TDateTime;
//  NextIndex: Integer;
//  CurrIndex: Integer;
//  TargetIndex: Integer;
//
//  CurrentTime: TDateTime;
//  PlayedTime: TDateTime;
//  RemainTime: TDateTime;
//  RemainTargetTime: TDateTime;
//
//  PlayedTimeString: String;
//  RemainTimeString: String;
//  NextStartTimeString: String;
//  NextDurationString: String;
//  RemainTargetTimeString: String;
//
//  SaveStartTime: TEventTime;
//
//  EventStatus: TEventStatus;
//  MediaStatus: TMediaStatus;
//
//  C, R: Integer;
//begin
//  inherited;
//  case Message.Msg of
//    WM_UPDATE_CHANNEL_TIME:
//    begin
//      RealtimeChangePlayListTimeLine;
//    end;
//
//    WM_UPDATE_CURR_EVENT:
//    begin
//{      CurrIndex := GetCueSheetIndexByItem(CueSheetCurr);
//      if (CurrIndex >= 0) then
//      begin
//        if (GV_SettingOption.OnAirEventHighlight) then
//        begin
//          with acgPlaylist do
//          begin
//            R := DisplRowIndex(CurrIndex + CNT_CUESHEET_HEADER);
////            if (R >= FixedRows) or (R <= RowCount - 1) then
//            begin
//              TopRow := R - GV_SettingOption.OnAirEventFixedRow;
//
//              MouseActions.DisjunctRowSelect := False;
//              ClearRowSelect;
//              MouseActions.DisjunctRowSelect := True;
//              SelectRows(R, 1);
//              Row := R;
//            end;
//          end;
//        end;
//
//        CurrIndex := GetNextOnAirMainIndexByItem(CueSheetCurr);
//        OnAirInputEvents(CurrIndex, 1);
//      end;
//
//      if (FChannelTimelineForm <> nil) then
//        FChannelTimelineForm.SetCueSheetCurr(CueSheetCurr);  }
//    end;
//
//    WM_UPDATE_NEXT_EVENT:
//    begin
//{      if (CueSheetCurr <> nil) then
//        NextIndex := GetCueSheetIndexByItem(CueSheetCurr)
//      else
//        NextIndex := GetCueSheetIndexByItem(CueSheetNext);
//
//      if (NextIndex >= 0) then
//      begin
//        if (GV_SettingOption.OnAirEventHighlight) then
//        begin
//          with acgPlaylist do
//          begin
//            R := DisplRowIndex(NextIndex + CNT_CUESHEET_HEADER);
////            if (R >= FixedRows) or (R <= RowCount - 1) then
//            begin
//              TopRow := R - GV_SettingOption.OnAirEventFixedRow;
//
//              MouseActions.DisjunctRowSelect := False;
//              ClearRowSelect;
//              MouseActions.DisjunctRowSelect := True;
//              SelectRows(R, 1);
//              Row := R;
//            end;
//          end;
//        end;
//      end;
//
//      if (FChannelTimelineForm <> nil) then
//        FChannelTimelineForm.SetCueSheetNext(CueSheetNext);  }
//    end;
//
//    WM_UPDATE_TARGET_EVENT:
//    begin
//{      TargetIndex := GetCueSheetIndexByItem(CueSheetTarget);
////      if (TargetIndex >= 0) then
//      begin
//        acgPlaylist.Repaint;
//      end; }
//    end;
//
//    WM_UPDATE_EVENT_STATUS:
//    begin
//      Index := Message.WParam;
//      Item := PCueSheetItem(Message.LParam);
//
//      if (Item <> nil) and (Item^.EventStatus.State = esSkipped) then
//        DeletePlayListTimeline(Item)
//      else
//        PopulatePlayListTimeline(Item);
//    end;
//
//    WM_UPDATE_ERROR_DISPLAY:
//    begin
//      FErrorDisplayEnabled := not FErrorDisplayEnabled;
////      ErrorDisplayPlayListTimeLine;
//    end;
//
//    WM_SET_ONAIR:
//    begin
//      ChannelID := Message.WParam;
//      IsOnAir := Boolean(Message.LParam);
//
//      SetChannelOnAir(ChannelID, IsOnAir);
//    end;
//
//    WM_BEGIN_UPDATE:
//    begin
//      wmtlPlaylist.BeginUpdateCompositions;
//    end;
//
//    WM_INSERT_CUESHEET:
//    begin
//      Index := Message.WParam;
//      Item := PCueSheetItem(Message.LParam);
//
//      PopulatePlayListTimeline(Item);
//    end;
//
//    WM_UPDATE_CUESHEET:
//    begin
//      Index := Message.WParam;
//      Item := PCueSheetItem(Message.LParam);
//
//      PopulatePlayListTimeline(Item);
//    end;
//
//    WM_DELETE_CUESHEET:
//    begin
//      Index := Message.WParam;
//      Item := PCueSheetItem(Message.LParam);
//
//      DeletePlayListTimeline(Item);
//    end;
//
//    WM_CLEAR_CUESHEET:
//    begin
//      ChannelID := Message.WParam;
//
//      ClearPlayListTimeline(ChannelID);
//    end;
//
//    WM_END_UPDATE:
//    begin
//      wmtlPlaylist.EndUpdateCompositions;
//    end;
//  end;
//end;

constructor TfrmAllChannels.Create(AOwner: TComponent; ACombine: Boolean; ALeft, ATop, AWidth, AHeight: Integer);
begin
  inherited Create(AOwner, ACombine, ALeft, ATop, AWidth, AHeight);
end;

procedure TfrmAllChannels.Initialize;
begin
  FChannelPanelList := TList<TfrmAllChannelsPanel>.Create;

  wmtbTimelineZoom.Min := 0;
  wmtbTimelineZoom.Max := Round(SecsPerDay * wmtlPlaylist.TimeZoneProperty.FrameRate) + 1;
  wmtbTimelineZoom.Position := 6;//wmtbTimelineZoom.Max;
  FTimeLineZoomPosition := wmtbTimelineZoom.Position;

  FTimelineSpace := GV_SettingOption.TimelineSpace;

  FErrorDisplayEnabled := True;

  InitializeTimeline;

  FTimerThread := TAllChannelsTimerThread.Create(Self);
end;

procedure TfrmAllChannels.Finalize;
var
  I: Integer;
  ChannelForm: TfrmAllChannelsPanel;
begin
  if (FTimerThread <> nil) then
  begin
    FTimerThread.Terminate;
    FTimerThread.WaitFor;
    FreeAndNil(FTimerThread);
  end;

  for I := GV_ChannelList.Count - 1 downto 0 do
    ClearPlayListTimeline(GV_ChannelList[I]^.ID);

  for I := FChannelPanelList.Count - 1 downto 0 do
  begin
    ChannelForm := FChannelPanelList[I];
    FreeAndNil(ChannelForm);
  end;

  FChannelPanelList.Clear;
  FreeAndNil(FChannelPanelList);

  if FPanelHandle <> 0 then
    SetWindowLong(FPanelHandle, GWL_WNDPROC, LongInt(FDefPanelProc));
  FreeObjectInstance(FPanelInstance);
end;

procedure TfrmAllChannels.InitializeTimeline;
var
  I: Integer;
  Channel: PChannel;
  ChannelForm: TfrmAllChannelsPanel;
begin
  // Timeline
  with wmtlPlaylist do
  begin
    with TimeZoneProperty do
    begin
      FrameDayReset := True;
//      FrameRate := FrameRate29_97;
      FrameRate := FrameRate30_00;
      FrameSkip := 30;
      FrameStep := 15;
      RailBarColor := COLOR_BAR_TIMELINE_RAIL;
      RailBarVisible := True;

      FTimeLineDaysPerFrames := Round(SecsPerDay * TimeZoneProperty.FrameRate);
      FrameCount := FTimeLineDaysPerFrames;
    end;

    FrameSelectEnabled := False;
    TrackSelectEnabled := False;
    TrackTrimEnabled := False;

    VideoGroupProperty.Count := 0;
    AudioGroupProperty.Count := 0;

    DataGroupProperty.Count  := GV_ChannelList.Count;
    for I := 0 to DataGroupProperty.Count - 1 do
    begin
      DataCompositions[I].Height := GV_SettingOption.ChannelTimelineHeight;
    end;
  end;

  FTimelineStartDate := Date;

  FTimeLineMin := 0;
  FTimeLineMax := 0;

  UpdateZoomPosition(FTimeLineZoomPosition);

  // Channel panel
//  pnlChannelCaption.Height := wmtlPlaylist.TimeZoneProperty.Height;

  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    Channel := GV_ChannelList[I];

    ChannelForm := TfrmAllChannelsPanel.Create(pnlChannelList, Channel^.ID);
    ChannelForm.Parent := pnlChannelList;
    ChannelForm.Align := alNone;
    ChannelForm.Height := GV_SettingOption.ChannelTimelineHeight - 2;
    ChannelForm.Tag := Channel^.ID;
    ChannelForm.Show;

    FChannelPanelList.Add(ChannelForm);

    with wmtlPlaylist do
      DataCompositions[I].Tag := Channel^.ID;
  end;
end;

procedure TfrmAllChannels.AdjustAllChannelsChannel;
var
  I: Integer;
  Panel: TfrmAllChannelsPanel;
  PanelTop: Integer;
  L, T, W, H: Integer;
begin
  inherited;

  if (FChannelPanelList = nil) then exit;

  if (FChannelPanelList.Count <= 0) then exit;

//  PanelHeight := 72;//(FVideoGroup.ClientHeight + FVideoCompositionList.Count - 1) div FVideoCompositionList.Count;
  PanelTop := {pnlChannelCaption.Height} - wmtlPlaylist.DataVertScrollPosition;

  for I := 0 to FChannelPanelList.Count - 1 do
  begin
    if (FChannelPanelList[I] = nil) then Continue;

    Panel := FChannelPanelList[I];
    L := 0;
    T := PanelTop;
    W := pnlChannelList.ClientWidth;
    H := Panel.Height;
    Panel.SetBounds(L, T, W, H);
    PanelTop := Panel.Top + Panel.Height + 1;// - 1;
  end;
//  UpdateWindow(Handle);
end;

function TfrmAllChannels.GetPositionByZoomType(AZoomType: TTimeLineZoomType): Integer;
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

function TfrmAllChannels.GetZoomTypeByPosition(APosition: Integer): TTimeLineZoomType;
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

procedure TfrmAllChannels.UpdateZoomPosition(APosition: Integer);
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

//        RealtimeChangePlayListTimeLine;
      finally
        EndUpdate;
      end;
  //    WMTimeLine.ViewSplitter;
  //    WMTimeLine.ViewAreaRepaint;
    end;
  end;
end;

procedure TfrmAllChannels.SetZoomPosition(Value: Integer);
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

procedure TfrmAllChannels.wmtbTimelineZoomChange(Sender: TObject);
begin
  inherited;
  TimelineZoomPosition := wmtbTimelineZoom.Position;
end;

procedure TfrmAllChannels.wmtlPlaylistDataGroupVertScrollEvent(Sender: TObject;
  ScrollPos: Integer);
begin
  inherited;
  AdjustAllChannelsChannel;
end;

procedure TfrmAllChannels.wmtlPlaylistMouseWheel(Sender: TObject;
  Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint;
  var Handled: Boolean);
var
  Zoom: Integer;
  Pos: Integer;
begin
  inherited;

  if (ssCtrl in Shift) then
  begin
    with wmtlPlaylist.ZoomBarProperty do
    begin
      Zoom := Integer(FTimelineZoomType);
      if (WheelDelta > 0) then Inc(Zoom)
      else if (WheelDelta < 0) then Dec(Zoom);

      if (Zoom > Integer(High(TTimeLineZoomType))) then Zoom := Integer(High(TTimeLineZoomType))
      else if (Zoom < Integer(Low(TTimeLineZoomType))) then Zoom := Integer(Low(TTimeLineZoomType));

      FTimelineZoomType := TTimeLineZoomType(Zoom);
      TimelineZoomPosition := GetPositionByZoomType(FTimelineZoomType);
    end;
  end
  else
  begin
    with wmtlPlaylist do
    begin
      Pos := DataVertScroll.ScrollPosition;

      Dec(Pos, WheelDelta);

      DataVertScroll.ScrollPosition := Pos;


{      Pos := FrameNumber;
      if (WheelDelta > 0) then Dec(Pos)
      else if (WheelDelta < 0) then Inc(Pos);

      if (Pos > MaxFrameNumber) then Pos := MaxFrameNumber
      else if (Pos < 0) then Pos := 0;

      FrameNumber := Pos; }
    end;
  end;
  Handled := True;
end;

procedure TfrmAllChannels.wmtlPlaylistTrackCustomDrawEvent(Sender: TObject;
  ATrack: TTrack; ACanvas: TCanvas; ARect: TRect; var ACustomDraw: Boolean);
var
  R: TRect;
  Item: PCueSheetItem;
  BkColor, TxColor, ToColor: TColor;
  BkPicture: TPicture;
  Index: Integer;
  W, H: Integer;
  Caption: String;
  Flags: DWORD;
begin
  inherited;

  if (ATrack = nil) then exit;

  ACustomDraw := True;

  if (ATrack.Data <> nil) then
  begin
    BkPicture := GV_TimelineNormalIcon;

    BkColor := COLOR_BK_EVENTSTATUS_NORMAL;
    TxColor := clWhite;//COLOR_TX_EVENTSTATUS_NORMAL;
    ToColor := COLOR_TO_EVENTSTATUS_NORMAL;

    Caption := '';

    Item := ATrack.Data;
    if (Item <> nil) then
    begin
      BkColor := Item^.BkColor;
      TxColor := clWhite;//Item^.TxColor;
      ToColor := Item^.ToColor;

      case Item^.EventStatus.State of
        esCueing..esPreroll:
          BkPicture := GV_TimelineNextIcon;
        esOnAir:
          BkPicture := GV_TimelineOnairIcon;
        else
          BkPicture := GV_TimelineNormalIcon;
      end;

      Caption := String(Item^.Title);
    end;
  end;

  if (BkColor = COLOR_BK_EVENTSTATUS_NORMAL) then
  begin
    BkColor := $004F3B35;
  end;

{    Item := GetParentCueSheetItemByItem(Item);
    if (Item <> nil) then
    begin
      if (FCueSheetCurr <> nil) and (Item^.GroupNo = FCueSheetCurr.GroupNo) then
      begin
        TopColor  := COLOR_TIMELINE_PLAY_TOP;
        BodyColor := COLOR_TIMELINE_PLAY_BODY;
        BackPicture := GV_TimelinePlayIcon;
      end
      else if (FCueSheetNext <> nil) and (Item^.GroupNo = FCueSheetNext.GroupNo) then
      begin
        TopColor  := COLOR_TIMELINE_NEXT_TOP;
        BodyColor := COLOR_TIMELINE_NEXT_BODY;
        BackPicture := GV_TimelineNextIcon;
      end
      else
      begin
        TopColor  := COLOR_TIMELINE_NORMAL_TOP;
        BodyColor := COLOR_TIMELINE_NORMAL_BODY;
        BackPicture := GV_TimelineNormalIcon;
      end;
    end; }

  R := ARect;
  R.Bottom := ARect.Top + GV_ChannelTimelineTrackTopHeight;

  ACanvas.Brush.Color := ToColor;
  ACanvas.FillRect(R);

  R := ARect;
  R.Top := ARect.Top + GV_ChannelTimelineTrackTopHeight;

  ACanvas.Brush.Color := BkColor;
  ACanvas.FillRect(R);

  W := 0;
  H := 0;
  if (BkPicture <> nil) and (BkPicture.Graphic <> nil) then
  begin
    R := ARect;
    InflateRect(R, -2, -2);

{    R.Right := R.Left + BkPicture.Width;
    R.Top := R.Top + (R.Height - BkPicture.Height) div 2;

    ACanvas.Draw(R.Left, R.Top, BkPicture.Graphic); }

    H := R.Height - GV_ChannelTimelineTrackTopHeight * 2;
    W := Round(BkPicture.Width * (H / BkPicture.Height));

    R.Right := R.Left + W;
    R.Top := R.Top + GV_ChannelTimelineTrackTopHeight;
    R.Bottom := R.Top + H;

    ACanvas.StretchDraw(R, BkPicture.Graphic);

{      W := R.Width;
    H := R.Height;

    if (W < BackPicture.Width) then
    begin
      H := Trunc(BackPicture.Height * W / BackPicture.Width);
      W := Trunc(BackPicture.Width * H / BackPicture.Height);
    end
    else
    begin
      W :=  BackPicture.Width;
    end;

    if (H < BackPicture.Height) then
    begin
      W := Trunc(BackPicture.Width * H / BackPicture.Height);
      H := Trunc(BackPicture.Height * W / BackPicture.Width);
    end
    else
    begin
      H :=  BackPicture.Height;
    end;

    R.Right := R.Left + W;
    R.Top := R.Top + (R.Height - H) div 2;
    R.Bottom := R.Top + H;//(R.Height - H) div 2;

    ACanvas.StretchDraw(R, BackPicture.Graphic); }
  end;

  // Caption
  ACanvas.Font.Assign(ATrack.Font);
  ACanvas.Font.Color := TxColor;
  ACanvas.Brush.Style := bsClear;

  R := ARect;
  InflateRect(R, -2, -2);
//  R.Left := R.Left + W + 2;

  Flags := DT_SINGLELINE or DT_LEFT or DT_VCENTER or DT_END_ELLIPSIS;
  DrawText(ACanvas.Handle, ATrack.Caption, Length(ATrack.Caption), R, Flags);

  R := ARect;
  // Left bar
//  DrawBlendLine(ACanvas, R.Left, R.Top, 1, R.Bottom - R.Top, wmtlPlaylist.CompositionBarProperty.LineVertColor, 0.3);
  DrawBlendLine(ACanvas, R.Left, R.Top, 1, R.Bottom - R.Top, clWhite, 0.3);
//  ACanvas.Pen.Color := wmtlPlaylist.CompositionBarProperty.LineHorzColor;
//  ACanvas.MoveTo(R.Left, R.Top);
//  ACanvas.LineTo(R.Left, R.Bottom);

  // Right Bar
  ACanvas.Pen.Color := wmtlPlaylist.CompositionBarProperty.LineVertColor;
  ACanvas.MoveTo(R.Right - 1, R.Top);
  ACanvas.LineTo(R.Right - 1, R.Bottom);
end;

procedure TfrmAllChannels.wmtlPlaylistTrackHintEvent(Sender: TObject; Track: TTrack;
  var HintStr: string);
var
  ProgItem: PCueSheetItem;
  Item, ParentItem: PCueSheetItem;
  StartTime, EndTime: TEventTime;
begin
  inherited;
{  if (Track.Data <> nil) then
  begin
    Item := Track.Data;

    case Item^.EventMode of
      EM_PROGRAM:
      begin
        HintStr := String(Item^.Title);
        exit;
      end;
      EM_MAIN:
      begin
        StartTime := Item^.StartTime;
        EndTime   := GetEventEndTime(Item^.StartTime, Item^.DurationTC);
      end;
      EM_JOIN:
      begin
        ParentItem := GetParentCueSheetItemByItem(Item);
        if (ParentItem <> nil) then
        begin
          StartTime := ParentItem^.StartTime;
          EndTime   := GetEventEndTime(StartTime, Item^.DurationTC);
        end
        else
          exit;
      end;
      EM_SUB:
      begin
        ParentItem := GetParentCueSheetItemByItem(Item);
        if (ParentItem <> nil) then
        begin
          if (Item^.StartMode = SM_SUBBEGIN) then
          begin
            StartTime := GetEventTimeSubBegin(ParentItem^.StartTime, Item^.StartTime.T);
            EndTime   := GetEventEndTime(StartTime, Item^.DurationTC);
          end
          else
          begin
            StartTime := GetEventTimeSubEnd(ParentItem^.StartTime, ParentItem^.DurationTC, Item^.StartTime.T);
            EndTime   := GetEventEndTime(StartTime, Item^.DurationTC);
          end;
        end
        else
          exit;
      end;
      else
        exit;
    end;

    HintStr := Format('%s'#13#10'%s'#13#10'Start: %s'#13#10'End: %s'#13#10'Duration: %s',
                      [String(Item^.Title),
                       String(Item^.SubTitle),
                       TimecodeToString(StartTime.T),
                       TimecodeToString(EndTime.T),
                       TimecodeToString(Item^.DurationTC)]);

    ProgItem := GetProgramItemByItem(Item);
    if (ProgItem <> nil) then
      HintStr := String(ProgItem^.Title) + #13#10 + HintStr;
  end;  }
end;

procedure TfrmAllChannels.TimelineGotoCurrent;
begin
  FTimelineSpace := GV_SettingOption.TimelineSpace;
//  PostMessage(Handle, WM_UPDATE_CHANNEL_TIME, 0, 0);
end;

procedure TfrmAllChannels.TimelineMoveLeft;
begin
  Inc(FTimelineSpace, GV_SettingOption.TimelineSpaceInterval);
//  if (FTimelineSpace >= wmtlPlaylist.ClientWidth) then FTimelineSpace := wmtlPlaylist.ClientWidth - 19;

//  PostMessage(Handle, WM_UPDATE_CHANNEL_TIME, 0, 0);
end;

procedure TfrmAllChannels.TimelineMoveRight;
begin
  Dec(FTimelineSpace, GV_SettingOption.TimelineSpaceInterval);
//  if (FTimelineSpace < 0) then FTimelineSpace := 0;

//  PostMessage(Handle, WM_UPDATE_CHANNEL_TIME, 0, 0);
end;

procedure TfrmAllChannels.TimelineZoomIn;
var
  Zoom: Integer;
begin
  inherited;
  with wmtlPlayList.ZoomBarProperty do
  begin
    Zoom := Integer(FTimelineZoomType);
    Dec(Zoom);

    if (Zoom < Integer(Low(TTimeLineZoomType)) + 1) then Zoom := Integer(Low(TTimeLineZoomType)) + 1;

    FTimelineZoomType := TTimeLineZoomType(Zoom);
    TimelineZoomPosition := GetPositionByZoomType(FTimelineZoomType);
  end;
end;

procedure TfrmAllChannels.TimelineZoomOut;
var
  Zoom: Integer;
begin
  with wmtlPlayList.ZoomBarProperty do
  begin
    Zoom := Integer(FTimelineZoomType);
    Inc(Zoom);

    if (Zoom > Integer(High(TTimeLineZoomType))) then Zoom := Integer(High(TTimeLineZoomType));

    FTimelineZoomType := TTimeLineZoomType(Zoom);
    TimelineZoomPosition := GetPositionByZoomType(FTimelineZoomType);
  end;
end;

function TfrmAllChannels.GetTimelinChannelByChannelID(AChannelID: Word): TfrmAllChannelsPanel;
var
  I: Integer;
  ChannelForm: TfrmAllChannelsPanel;
begin
  Result := nil;

  for I := 0 to FChannelPanelList.Count - 1 do
  begin
    ChannelForm := FChannelPanelList[I];
    if (ChannelForm <> nil) and (ChannelForm.Tag = AChannelID) then
    begin
      Result := ChannelForm;
      break;
    end;
  end;
end;

function TfrmAllChannels.GetAllChannelsCompIndexByChannelID(AChannelID: Word): Integer;
var
  I: Integer;
begin
  Result := -1;

  with wmtlPlaylist do
    for I := 0 to DataGroupProperty.Count - 1 do
      if (DataCompositions[I].Tag = AChannelID) then
      begin
        Result := I;
        break;
      end;
end;

procedure TfrmAllChannels.DisplayChannelTime(AChannelID: Word);
var
  ChannelForm: TfrmAllChannelsPanel;
begin
  ChannelForm := GetTimelinChannelByChannelID(AChannelID);
  if (ChannelForm <> nil) then
  begin

  end;
end;

procedure TfrmAllChannels.SetChannelOnAir(AChannelID: Word; AOnAir: Boolean);
var
  ChannelForm: TfrmAllChannelsPanel;
begin
  ChannelForm := GetTimelinChannelByChannelID(AChannelID);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.SetChannelOnAir(AOnAir);
  end;
end;

procedure TfrmAllChannels.SetChannelTime(AChannelID: Word; APlayedTime, ARemainingTime, ANextStart, ANextDuration: String);
var
  ChannelForm: TfrmAllChannelsPanel;
begin
  ChannelForm := GetTimelinChannelByChannelID(AChannelID);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.SetChannelTime(APlayedTime, ARemainingTime, ANextStart, ANextDuration);
  end;
end;

procedure TfrmAllChannels.PopulatePlayListTimeline(AItem: PCuesheetItem);
var
  I: Integer;
  CompIndex: Integer;
  Track: TTrack;
begin
  if (AItem = nil) then exit;

  if (AItem^.EventMode <> EM_MAIN) then exit;
  if (AItem^.EventStatus.State = esSkipped) then exit;

  CompIndex := GetAllChannelsCompIndexByChannelID(AItem^.EventID.ChannelID);
  if (CompIndex < 0) then exit;

  with wmtlPlaylist do
  begin
//    BeginUpdateCompositions;
    try
      Track := DataCompositions[CompIndex].Tracks.GetTrackByData(AItem);
      if (Track = nil) then
      begin
        Track := DataCompositions[CompIndex].Tracks.Add;
        Track.Data := AItem;
      end;

      Track.InPoint  := TimecodeToFrame(AItem^.StartTime.T) + Trunc(AItem^.StartTime.D - FTimelineStartDate) * FTimeLineDaysPerFrames; //TimecodeToFrame(AItem^.StartTime.T);// + Trunc(AItem^.StartTime.D - FTimelineStartDate) * FTimeLineDaysPerFrames - 1;
      Track.OutPoint := Track.InPoint + TimecodeToFrame(AItem^.DurationTC);
      Track.Duration := Track.OutPoint - Track.InPoint;

      Track.Color        := GetProgramTypeColorByCode(AItem^.ProgramType);
      Track.ColorCaption := GetProgramTypeColorByCode(AItem^.ProgramType);
    //          Track.ColorSelected  := Track.Color;
    //          Track.ColorHighLight := $000E0607;
    //          Track.ColorShadow    := $000E0607;

      Track.Caption := String(AItem^.Title);

      if (Track.InPoint < FTimeLineMin) then FTimeLineMin := Track.InPoint;
      if (Track.OutPoint > FTimeLineMax) then FTimeLineMax := Track.OutPoint;
    finally
//      EndUpdateCompositions;
    end;
  end;
end;

procedure TfrmAllChannels.PopulateEventStatusPlayListTimeline(AIndex: Integer; AItem: PCueSheetItem);
var
  Item: PCueSheetItem;
  CompIndex: Integer;
  Track: TTrack;

  BackColor, FontColor: TColor;
begin
{  Item := GetCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    with wmtlPlaylist do
    begin
      BeginUpdateCompositions;
      try
        CompIndex := 0;
        case Item^.EventMode of
          EM_PROGRAM: CompIndex := 0;
          EM_MAIN: CompIndex := 1;
          EM_JOIN,
          EM_SUB: CompIndex := GetChildCueSheetIndexByItem(Item) + 1;
          else
            exit;
        end;

        if (CompIndex >= 0) and (CompIndex < DataGroupProperty.Count) then
        begin
          Track := DataCompositions[CompIndex].Tracks.GetTrackByData(Item);
          if (Track <> nil) then
          begin
            if (CueSheetNext <> nil) and (Item^.GroupNo = CueSheetNext^.GroupNo) and
               ((not FChannelOnAir) or ((CueSheetNext^.EventStatus.State in [esLoaded]) and (Item^.EventStatus.State in [esLoaded]))) then
            begin
              BackColor := COLOR_BK_EVENTSTATUS_NEXT;
              FontColor := COLOR_TX_EVENTSTATUS_NEXT;
            end
            else
            begin
              case Item^.EventStatus.State of
                esCueing..esPreroll:
                begin
                  BackColor := COLOR_BK_EVENTSTATUS_CUED;
                  FontColor := COLOR_TX_EVENTSTATUS_CUED;
                end;
                esOnAir:
                begin
                  BackColor := COLOR_BK_EVENTSTATUS_ONAIR;
                  FontColor := COLOR_TX_EVENTSTATUS_ONAIR;
                end;
                esSkipped,
                esFinish..esDone:
                begin
                  BackColor := COLOR_BK_EVENTSTATUS_DONE;
                  FontColor := COLOR_TX_EVENTSTATUS_DONE;
                end;
                esError:
                begin
                  BackColor := COLOR_BK_EVENTSTATUS_ERROR;
                  FontColor := COLOR_TX_EVENTSTATUS_ERROR;
                end;
                else
                begin
                  if (CueSheetTarget <> nil) and (Item^.GroupNo = CueSheetTarget^.GroupNo) then
                  begin
                    BackColor := COLOR_BK_EVENTSTATUS_TARGET;
                    FontColor := COLOR_TX_EVENTSTATUS_TARGET;
                  end
                  else
                  begin
                    BackColor := COLOR_BK_EVENTSTATUS_NORMAL;
                    FontColor := COLOR_TX_EVENTSTATUS_NORMAL;
                  end;
                end;
              end;
            end;

            Track.Color        := BackColor;
            Track.ColorCaption := FontColor;
          end;
        end;
      finally
        EndUpdateCompositions;
      end;
    end;
  end; }
end;

procedure TfrmAllChannels.UpdateChangeDatePlayListTimeLine(ADays: Integer);
var
  I, J: Integer;
  Track: TTrack;
  Duration: Integer;
begin
  inherited;

  with wmtlPlaylist do
  begin
    BeginUpdateCompositions;
    try
      for I := 0 to DataGroupProperty.Count - 1 do
      begin
        for J := 0 to DataCompositions[I].Tracks.Count - 1 do
        begin
          Track := DataCompositions[I].Tracks[J];
          if (Track <> nil) then
          begin
            Duration := Track.Duration;

            Track.InPoint  := Track.InPoint - (ADays * FTimeLineDaysPerFrames);
            Track.OutPoint := Track.InPoint + Duration;
          end;
        end;
      end;
    finally
      EndUpdateCompositions;
    end;
  end;
end;

procedure TfrmAllChannels.RealtimeChangePlayListTimeLine;
var
  I: Integer;
  F: Integer;

  Days: Integer;

  FirstItem: PCueSheetItem;
  LastMainItem: PCueSheetItem;
  LastEndTime: TEventTime;
  SideFrames: Integer;
  SampleTime: Integer;

  CurrentTime: TDateTime;
begin
  inherited;

  CurrentTime := SystemTimeToDateTime(GV_TimeCurrent);

  if (FTimelineStartDate <> Date) then
  begin
    Days := Trunc(CurrentTime - FTimelineStartDate);
    FTimelineStartDate := Date;

    UpdateChangeDatePlayListTimeLine(Days);
  end;

//  wmtlPlaylist.TimeZoneProperty.BeginUpdate;

//  FTimelineStartDate := IncDay(Date, -1);
  with wmtlPlaylist.TimeZoneProperty do
  begin
    SideFrames := Round(FTimelineSpace / (FrameGap / FrameSkip));
    SideFrames := (SideFrames div Round(FrameRate)) * Round(FrameRate) - 1;
//        SampleTime := Round(wmtbTimelineZoom.Position / FrameRate);
//        if (SampleTime < 1) then SampleTime := 1;
//
//        SideFrames := TimecodeToFrame(SecondToTimeCode(SampleTime, FrameRate), FrameRate);
  end;

//  wmtlPlaylist.BeginUpdateCompositions;
  with wmtlPlaylist.TimeZoneProperty do
  begin
    BeginUpdate;
    try
//    TimeZoneProperty.FrameStart := FTimeLineMin - SideFrames;
//    TimeZoneProperty.FrameCount := FTimeLineMax - FTimeLineMin + (SideFrames * 2);


//    FrameStart :=  Trunc(Trunc((Round(SecondOfTheDay(Now) * FrameRate) - SideFrames) / FrameRate) * FrameRate);
//    wmtlPlaylist.FrameNumber := Round((SecondOfTheDay(Now) * FrameRate));

    F := SampleTimeToFrame(SecondOfTheDay(CurrentTime), FrameRate);
    FrameStart :=  F - SideFrames - 1;
    wmtlPlaylist.FrameNumber := F;




//    SetScrollPos(wmtlPlaylist.HorzScrollBar.Handle, SB_HORZ, 0, True);

//    SetScrollPos(wmtlPlaylist.HorzScrollBar.Handle, SB_HORZ, 0, True);
    finally
      EndUpdate;
    end;
  end;
//  wmtlPlaylist.EndUpdateCompositions;
//  wmtlPlaylist.TimeZoneProperty.RailBarVisible := True;
//  wmtlPlaylist.TimeZoneProperty.EndUpdate;
//  wmtlPlaylist.DataGroup.Visible := True;

//  wmtlPlaylist.MarkIn := FTimeLineMin + 1000;
//  wmtlPlaylist.MarkOut := FTimeLineMax - 1000;

//  UpdateZoomPosition(FTimelineZoomPosition);
end;

procedure TfrmAllChannels.DeletePlayListTimeline(AItem: PCueSheetItem);
var
  I: Integer;
  CompIndex: Integer;
  Track: TTrack;
begin
  if (AItem = nil) then exit;

  if (AItem^.EventMode <> EM_MAIN) then exit;

  CompIndex := GetAllChannelsCompIndexByChannelID(AItem^.EventID.ChannelID);
  if (CompIndex < 0) then exit;

  with wmtlPlaylist do
  begin
    Track := DataCompositions[CompIndex].Tracks.GetTrackByData(AItem);

    if (Track <> nil) then
    begin
      FreeAndNil(Track);
    end;
  end;
end;

procedure TfrmAllChannels.ErrorDisplayPlayListTimeLine(AErrorDisplayEnabled: Boolean);
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

procedure TfrmAllChannels.SetEventStatus(AEventID: TEventID; AStatus: TEventStatus; AIsCurrEvent: Boolean = False);
var
  CItem: PCueSheetItem;
  Index: Integer;
  RRow: Integer;
begin
end;

procedure TfrmAllChannels.SetEventOverall(ADCSIP: String; ADeviceHandle: TDeviceHandle; AOverall: TEventOverall);
var
  Item: PCueSheetItem;
  Source: PSource;
  SourceHandles: TSourceHandleList;
  I: Integer;
begin
end;

procedure TfrmAllChannels.SetMediaStatus(AEventID: TEventID; AStatus: TMediaStatus);
begin

end;

procedure TfrmAllChannels.InputCueSheets(ACueSheetList: TCueSheetList; AChannelID: Word; AIndex: Integer; ACount: Integer);
var
  I: Integer;
begin
  if (ACueSheetList = nil) then exit;

  if (AIndex < 0) or (AIndex > ACueSheetList.Count - 1) then exit;

  if (ACount = 0) then
    ACount := ACueSheetList.Count
  else
    ACount := AIndex + ACount;

  for I := AIndex to ACount - 1 do
    PopulatePlayListTimeline(ACueSheetList[I]);
end;

procedure TfrmAllChannels.InputCueSheets(AIndex: Integer; AItem: PCueSheetItem);
begin
  if (AIndex < 0) then exit;
  if (AItem = nil) then exit;

  PopulatePlayListTimeline(AItem);
end;

procedure TfrmAllChannels.DeleteCueSheets(AChannelID: Word; ADeleteList: TCueSheetList);
var
  I: Integer;
begin
  if (ADeleteList = nil) or (ADeleteList.Count <= 0) then exit;

  for I := ADeleteList.Count - 1 downto 0 do
    DeletePlayListTimeline(ADeleteList[I]);
end;

procedure TfrmAllChannels.DeleteCueSheets(AEventID: TEventID);
begin
  // not yet
end;

procedure TfrmAllChannels.DeleteCueSheets(ACueSheetList: TCueSheetList; AChannelID: Word; AFromIndex, AToIndex: Integer);
var
  I: Integer;
begin
  if (ACueSheetList = nil) then exit;

  for I := AToIndex downto AFromIndex do
    DeletePlayListTimeline(ACueSheetList[I]);
end;

procedure TfrmAllChannels.ClearCueSheets(AChannelID: Word);
begin
  ClearPlayListTimeline(AChannelID);
end;

procedure TfrmAllChannels.SetCueSheetCurrs(AEventID: TEventID);
begin
  // not yet
end;

procedure TfrmAllChannels.SetCueSheetNexts(AEventID: TEventID);
begin
  // not yet
end;

procedure TfrmAllChannels.SetCueSheetTargets(AEventID: TEventID);
begin
  // not yet
end;

procedure TfrmAllChannels.ClearPlayListTimeline(AChannelID: Word);
var
  CompIndex: Integer;
begin
  inherited;

  CompIndex := GetAllChannelsCompIndexByChannelID(AChannelID);
  if (CompIndex < 0) then exit;

  with wmtlPlaylist do
  begin
//    BeginUpdateCompositions;
    try
      DataCompositions[CompIndex].Tracks.Clear;
    finally
//      EndUpdateCompositions;
    end;
  end;
end;

procedure TfrmAllChannels.PanelWndProc(var Message: TMessage);
begin
  with Message do
  begin
    case Msg of
      WM_SIZE: AdjustAllChannelsChannel;
    end;
    Result := CallWindowProc(FDefPanelProc, FPanelHandle, Msg, WParam, LParam);
  end;
end;

procedure TfrmAllChannels.FormCreate(Sender: TObject);
begin
  inherited;
  Initialize;

end;

procedure TfrmAllChannels.FormDestroy(Sender: TObject);
begin
  inherited;
  Finalize;
end;

procedure TfrmAllChannels.FormResize(Sender: TObject);
var
  Zoom: Integer;
begin
  inherited;
  with wmtlPlayList.ZoomBarProperty do
  begin
    TimelineZoomPosition := GetPositionByZoomType(FTimelineZoomType);

{    Zoom := Integer(FTimelineZoomType);
    Inc(Zoom);

    if (Zoom > Integer(High(TTimeLineZoomType))) then Zoom := Integer(High(TTimeLineZoomType));

    FTimelineZoomType := TTimeLineZoomType(Zoom);
    TimelineZoomPosition := GetPositionByZoomType(FTimelineZoomType); }
  end;
end;

procedure TfrmAllChannels.FormShow(Sender: TObject);
begin
  inherited;

  if (FTimerThread <> nil) then
    FTimerThread.Start;

  FPanelInstance := MakeObjectInstance(PanelWndProc);

  FPanelHandle := pnlChannel.Handle;
  FDefPanelProc := Pointer(GetWindowLong(FPanelHandle, GWL_WNDPROC));
  SetWindowLong(FPanelHandle, GWL_WNDPROC, LongInt(FPanelInstance));

  AdjustAllChannelsChannel;
end;

{ TAllChannelsTimerThread }

constructor TAllChannelsTimerThread.Create(AAllChannelsForm: TfrmAllChannels);
begin
  FAllChannelsForm := AAllChannelsForm;

  FCloseEvent := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TAllChannelsTimerThread.Destroy;
begin
  Terminate;

  CloseHandle(FCloseEvent);

  inherited Destroy;
end;

procedure TAllChannelsTimerThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FCloseEvent);
end;

procedure TAllChannelsTimerThread.DoCueSheetCheck;
var
  NextStartTime: TDateTime;

  SaveCurr: PCueSheetItem;
  CurrItem: PCueSheetItem;
  CurrIndex: Integer;
  InputIndex: Integer;
begin
{  with FTimelineForm do
  begin
    if (not FChannelOnAir) then exit;

    // 다음 이벤트가 수동모드이고 남은 시작시각이 AutoIncreaseDurationBefore 보다 큰 경우는 대기
    if (CueSheetNext <> nil) and
       (CueSheetNext^.EventMode = EM_MAIN) and
       (CueSheetNext^.StartMode = SM_MANUAL) and
       (CueSheetNext^.EventStatus.State <= esCued) then
    begin
      NextStartTime := EventTimeToDateTime(CueSheetNext^.StartTime);
      if (CompareDateTime(NextStartTime, Now) > 0) then
        exit;
    end;

    if (CueSheetCurr <> nil) then
      CurrIndex := GetCueSheetIndexByItem(CueSheetCurr)
    else
      CurrIndex := 0;

    SaveCurr := CueSheetCurr;
    CurrItem := GetMainItemByInRangeTime(CurrIndex, Now);
    if (SaveCurr <> CurrItem) then
    begin
      CueSheetCurr := CurrItem;
//      InputIndex := GetNextOnAirMainIndexByItem(CueSheetCurr);
//      OnAirInputEvents(InputIndex, 1);
    end;

    if (CueSheetCurr <> nil) then
      CueSheetNext := GetNextMainItemByItem(CueSheetCurr)
    else if (CueSheetNext = nil) then
      CueSheetNext := GetStartOnAirMainItem;
  end;  }
end;

procedure TAllChannelsTimerThread.Execute;
var
  WaitList: array[0..1] of THandle;
begin
  WaitList[0] := GV_TimerExecuteEvent;
//  WaitList[1] := GV_TimerCancelEvent;
  WaitList[1] := FCloseEvent;
  while not Terminated do
  begin
    if (WaitForMultipleObjects(2, @WaitList, False, INFINITE) <> WAIT_OBJECT_0) then
      break; // Terminate thread when GV_TimerCancelEvent is signaled

    DoCueSheetCheck;

    PostMessage(FAllChannelsForm.Handle, WM_UPDATE_CHANNEL_TIME, 0, 0);

{    if (GV_SettingOption.AutoLoadCuesheet) then
    begin
      Inc(FTimelineForm.FAutoLoadIntervalTime);
      if (FTimelineForm.FAutoLoadIntervalTime > (TimecodeToMilliSec(GV_SettingOption.AutoLoadCuesheetInterval) div 1000)) then
        FTimelineForm.FAutoLoadThread.AutoLoad;
    end;

    if (GV_SettingOption.AutoEjectCuesheet) then
    begin
      Inc(FTimelineForm.FAutoEjectIntervalTime);
      if (FTimelineForm.FAutoEjectIntervalTime > (TimecodeToMilliSec(GV_SettingOption.AutoEjectCuesheetInterval) div 1000)) then
        FTimelineForm.FAutoEjectThread.AutoEject;
    end;

    Inc(FTimelineForm.FMediaCheckIntervalTime);
    if (FTimelineForm.FMediaCheckIntervalTime > (TimecodeToMilliSec(GV_SettingOption.MediaCheckInterval) div 1000)) then
      FTimelineForm.FMediaCheckThread.MediaCheck; }
  end;
end;

end.
