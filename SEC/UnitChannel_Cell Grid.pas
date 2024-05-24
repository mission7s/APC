unit UnitChannel;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitWorkForm, AdvUtil, Vcl.Grids,
  AdvObj, BaseGrid, AdvGrid, AdvCGrid, Vcl.Imaging.pngimage,
  Vcl.ExtCtrls, Vcl.StdCtrls, System.Generics.Collections, 
  AdvOfficePager, AdvSplitter,
  WMTools, WMControls, WMTimeLine, LibXmlParserU,
  UnitCommons, UnitDCSDLL, UnitConsts;

type
  TChannelTimerThread = class;

  TfrmChannel = class(TfrmWork)
    acgPlaylist: TAdvColumnGrid;
    wmtlPlaylist: TWMTimeLine;
    WMPanel2: TWMPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    Label7: TLabel;
    WMPanel3: TWMPanel;
    lblPlayedTime: TLabel;
    WMPanel4: TWMPanel;
    lblRemainingTime: TLabel;
    WMPanel5: TWMPanel;
    lblNextStart: TLabel;
    WMPanel6: TWMPanel;
    lblNextDuration: TLabel;
    AdvSplitter1: TAdvSplitter;
    lblPlayListFileName: TLabel;
    WMPanel1: TWMPanel;
    lblOnAirFlag: TLabel;
    Label1: TLabel;
    lblRemainingTargetEvent: TLabel;
    WMPanel7: TWMPanel;
    lblRemainingTargetTime: TLabel;
    wmibFreezeOnAir: TWMImageSpeedButton;
    wmibAssignNext: TWMImageSpeedButton;
    wmibTakeNext: TWMImageSpeedButton;
    wmibIncrease1Second: TWMImageSpeedButton;
    wmibDecrease1Second: TWMImageSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure acgPlaylistCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure acgPlaylistClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure acgPlaylistComboChange(Sender: TObject; ACol, ARow,
      AItemIndex: Integer; ASelection: string);
    procedure acgPlaylistEditCellDone(Sender: TObject; ACol, ARow: Integer);
    procedure acgPlaylistEditChange(Sender: TObject; ACol, ARow: Integer;
      Value: string);
    procedure acgPlaylistGetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: string);
    procedure acgPlaylistKeyPress(Sender: TObject; var Key: Char);
    procedure acgPlaylistSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure acgPlaylistComboDropDown(Sender: TObject; ARow, ACol: Integer);
    procedure acgPlaylistEditingDone(Sender: TObject);
    procedure acgPlaylistDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure acgPlaylistDblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure acgPlaylistCustomCellDraw(Sender: TObject; Canvas: TCanvas; ACol,
      ARow: Integer; AState: TGridDrawState; ARect: TRect; Printing: Boolean);
    procedure acgPlaylistGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure acgPlaylistGetCellBorder(Sender: TObject; ARow, ACol: Integer;
      APen: TPen; var Borders: TCellBorders);
    procedure acgPlaylistGetCellBorderProp(Sender: TObject; ARow, ACol: Integer;
      LeftPen, TopPen, RightPen, BottomPen: TPen);
    procedure wmtlPlaylistTrackHintEvent(Sender: TObject; Track: TTrack;
      var HintStr: string);
    procedure acgPlaylistGetDisplText(Sender: TObject; ACol, ARow: Integer;
      var Value: string);
  private
    { Private declarations }
    FPlayListFileName: String;

    FChannelID: Word;
    FChannelOnAir: Boolean;
    FChannelCueSheet: PChannelCueSheet;

    FLastDisplayNo: Integer;
    FLastSerialNo: Integer;
    FLastGroupNo: Integer;
    FLastCount: Integer;

    FFirstStartTime: TEventTime;

    FIsCellEditing: Boolean;
    FIsCellModified: Boolean;
    FOldCellValue: String;

    FTimeLineDaysPerFrames: Integer;
    FTimeLineMin: Integer;
    FTimeLineMax: Integer;

    FCueSheetCurr: PCueSheetItem;
    FCueSheetNext: PCueSheetItem;
    FCueSheetTarget: PCueSheetItem;

    FTimerThread: TChannelTimerThread;

    function GetCueSheetItemByIndex(AIndex: Integer): PCueSheetItem;
    function GetCueSheetItemByID(AEventID: TEventID): PCueSheetItem;
    function GetCueSheetIndexByItem(AItem: PCueSheetItem): Integer;

    function GetSelectCueSheetItem: PCueSheetItem;
    function GetSelectCueSheetIndex: Integer;

    function GetParentCueSheetItemByIndex(AIndex: Integer): PCueSheetItem;
    function GetParentCueSheetItemByItem(AItem: PCueSheetItem): PCueSheetItem;

    function GetChildCueSheetIndexByItem(AItem: PCueSheetItem; AIncludeComment: Boolean = False): Integer;

    function GetLastChildCueSheetItemByIndex(AIndex: Integer): PCueSheetItem;
    function GetLastChildCueSheetItemByItem(AItem: PCueSheetItem): PCueSheetItem;

    function GetFirstMainItem: PCueSheetItem;
    function GetFirstMainIndex: Integer;
    function GetLastMainItem: PCueSheetItem;
    function GetLastMainIndex: Integer;
    function GetBeforeMainItemByItem(AItem: PCueSheetItem): PCueSheetItem;
    function GetBeforeMainItemByIndex(AIndex: Integer): PCueSheetItem;
    function GetNextMainItemByItem(AItem: PCueSheetItem): PCueSheetItem;
    function GetNextMainItemByIndex(AIndex: Integer): PCueSheetItem;
    function GetNextMainIndexByItem(AItem: PCueSheetItem): Integer;
    function GetNextMainIndexByIndex(AIndex: Integer): Integer;
    function GetNextLoadedMainItemByItem(AItem: PCueSheetItem): PCueSheetItem;

    function GetBeforeMainCountByIndex(AIndex: Integer): Integer;

    function GetMainItemByStartTime(AIndex: Integer; ADateTime: TDateTime): PCueSheetItem;

    // 큐시트 중 DCS에 Input되지 않은 첫번째 메인 이벤트 인덱스를 구함
    function GetStartOnAirMainIndex: Integer;
    // 큐시트 중 DCS에 Input되지 않은 첫번째 메인 이벤트를 구함
    function GetStartOnAirMainItem: PCueSheetItem;

    // 큐시트 중 DCS에 다음 Input될 메인 이벤트를 구함
    function GetNextOnAirMainIndex: Integer;
    // 큐시트 중 DCS에 다음 Input될 메인 이벤트를 구함
    function GetNextOnAirMainItem: PCueSheetItem;

    // 선택된 이벤트의 삭제 목록을 구함
    function GetDeleteCueSheetList(ADeleteList: TCueSheetList): Integer;

    // 선택된 이벤트의 복사 및 잘라내기 목록을 구함
    function GetClipboardCueSheetList(AClipboardCueSheet: TClipboardCueSheet; APasteMode: TPasteMode = pmCut): Integer;

    function MakePlayerEvent(AItem: PCueSheetItem; var AEvent: TEvent): Integer;

    function InputEvent(AItem: PCueSheetItem): Integer;
    function DeleteEvent(AItem: PCueSheetItem): Integer;
    function ClearEvent: Integer;
    function TakeEvent(AItem: PCueSheetItem): Integer;
    function HoldEvent(AItem: PCueSheetItem): Integer;
    function ChangeDurationEvent(AItem: PCueSheetItem; ADuration: TTimecode): Integer;

    procedure Initialize;
    procedure Finalize;

    procedure InitializePlayListGrid;
    procedure InitializePlayListTimeLine;

    procedure PlaylistFileParsing(AFileName: String);

    procedure PopulatePlayListGrid(AIndex: Integer);
    procedure PopulateEventStatusPlayListGrid(AIndex: Integer);
    procedure DisplayPlayListGrid(AIndex: Integer = 0; ACount: Integer = 0);

    procedure PopulatePlayListTimeLine(AIndex: Integer);
    procedure DisplayPlayListTimeLine(AIndex: Integer = 0);
    procedure DeletePlayListTimeLine(AItem: PCueSheetItem);

    procedure ClearCueSheetList;
    procedure ClearPlayListGrid;
    procedure ClearPlayListTimeLine;

    procedure ResetNo(AIndex: Integer; ANo: Integer);
    procedure ResetStartDate(AIndex: Integer; ADays: Integer);
    procedure ResetStartTime(AIndex: Integer); overload;

    procedure ResetStartTime(AIndex: Integer; ASaveEndTime: TEventTime); overload;
    procedure ResetStartTimeByTime(AIndex: Integer; ASaveStartTime: TEventTime; ASaveDurationTC: TTimecode); overload;
    procedure ResetStartTimeByTime(AIndex: Integer; ASaveStartTime: TEventTime); overload;
    procedure ResetStartTimeByTime(AIndex: Integer; ASaveDurationTC: TTimecode); overload;

    procedure ResetStartTimeByDuration(AIndex: Integer; ADuration: TTimecode);

    procedure ResetStartTimePlus(AIndex: Integer; ADurEventTime: TEventTime);
    procedure ResetStartTimeMinus(AIndex: Integer; ADurEventTime: TEventTime);

    procedure ResetChildItems(AIndex: Integer);

    procedure ResetStartDateTimeline(AIndex: Integer);

    procedure CueSheetListQuickSort(L, R: Integer; ACueSheetList: TCueSheetList);

    procedure SetCueSheetItemStatusByIndex(AStartIndex, AEndIndex: Integer; AStatus: TEventStatus = esIdle);

    procedure SetChannelOnAir(AOnAir: Boolean);

    procedure OnAirInputEvents(AIndex, ACount: Integer);
    procedure OnAirDeleteEvents(AFromIndex, AToIndex: Integer);
    procedure OnAirClearEvents;
    procedure OnAirTakeEvent(AIndex: Integer);
    procedure OnAirHoldEvent(AIndex: Integer);
    procedure OnAirChangeDurationEvent(ADuration: TTimecode);
  protected
    procedure WndProc(var Message: TMessage); override;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent; AChannelCueSheet: PChannelCueSheet; ACombine: Boolean; ALeft, ATop, AWidth, AHeight: Integer); overload;

    function IsValidStartDate(AItem: PCueSheetItem; AStartDate: TDate; AIsSelf: Boolean = True): Boolean;
    function IsValidStartTime(AItem: PCueSheetItem; AStartTC: TTimecode; AIsSelf: Boolean = True): Boolean;
    function IsValidDuration(AItem: PCueSheetItem; ADuration: TTimecode; AIsSelf: Boolean = True): Boolean;
    function IsValidInTC(AItem: PCueSheetItem; AInTC: TTimecode; AIsSelf: Boolean = True): Boolean;
    function IsValidOutTC(AItem: PCueSheetItem; AOutTC: TTimecode; AIsSelf: Boolean = True): Boolean;
    function IsValidMediaId(AMediaId: String; AIsSelf: Boolean = True): Boolean;

    procedure NewPlayList;
    procedure OpenPlayList(AFileName: String);
    procedure OpenAddPlayList(AFileName: String);
    procedure SavePlayList;
    procedure SaveAsPlayList(AFileName: String);

    procedure CueSheetListSort(ACueSheetList: TCueSheetList);

    function CheckEditCueSheetPossibleByIndex(AIndex: Integer): Boolean;
    function CheckEditCueSheetPossibleByItem(AItem: PCueSheetItem): Boolean;

    procedure InsertCueSheet(AEventMode: TEventMode);
    procedure ExecuteInsertCueSheetMain(AIndex: Integer; AAddItem: PCueSheetItem);
    procedure ExecuteInsertCueSheetJoin(AIndex: Integer; AParentItem, AAddItem: PCueSheetItem);
    procedure ExecuteInsertCueSheetSub(AIndex: Integer; AParentItem, AAddItem: PCueSheetItem);
    procedure ExecuteInsertCueSheetComment(AIndex: Integer; AAddItem: PCueSheetItem);

    procedure UpdateCueSheet;
    procedure ExecuteUpdateCueSheet(AIndex: Integer; AEventMode: TEventMode; ASaveStartTime: TEventtime; ASaveDurationTC: TTimecode);

    procedure DeleteCueSheet;
    procedure ExecuteDeleteCueSheet(ADeleteList: TCueSheetList);

    procedure CopyToClipboardCueSheet;

    procedure CutToClipboardCueSheet;

    procedure PasteFromClipboardCueSheet;

    procedure ClearClipboardCueSheet;

    procedure StartOnAir;
    procedure FreezeOnAir;
    procedure FinishOnAir;

    procedure AssignNextEvent;
    procedure StartNextEventImmediately;

    procedure IncSecondsOnAirEvent(ASeconds: Integer);

    procedure AssignTargetEvent;

    procedure SetEventStatus(AEventID: TEventID; AStatus: TEventStatus);

    property ChannelCueSheet: PChannelCueSheet read FChannelCueSheet;
    property LastSerialNo: Integer read FLastSerialNo write FLastSerialNo;
    property LastGroupNo: Integer read FLastGroupNo write FLastGroupNo;
//    property LastCount: Integer read FLastCount write FLastCount;

    property PlayListFileName: String read FPlayListFileName;
  end;

  TChannelTimerThread = class(TThread)
  private
    FChannelForm: TfrmChannel;
  protected
    procedure Execute; override;
  public
    constructor Create(AChannelForm: TfrmChannel);
    destructor Destroy; override;
  end;

var
  frmChannel: TfrmChannel;

implementation

uses UnitSEC, System.DateUtils, UnitEditEvent, UnitSelectStartOnAir;

{$R *.dfm}

procedure TfrmChannel.WndProc(var Message: TMessage);
var
  NextStartTime: TDateTime;
  NextIndex: Integer;
  CurrIndex: Integer;

  CurrentTime: TDateTime;
  PlayedTime: TDateTime;
  RemainTime: TDateTime;
  RemainTargetTime: TDateTime;

  PlayedTimeString: String;
  RemainTimeString: String;
  NextStartTimeString: String;
  NextDurationString: String;
  RemainTargetTimeString: String;

  SaveStartTime: TEventTime;

  EventStatus: TEventStatus;

  C, R: Integer;
begin
  inherited;
  case Message.Msg of
    WM_UPDATE_CHANNEL_TIME:
    begin
      // Next Start, Duration time
      if (FCueSheetNext <> nil) then
      begin
        NextStartTime := EventTimeToDateTime(FCueSheetNext^.StartTime);
        // 다음 이벤트가 수동모드일 경우 시작시각 자동 늘림
        if (FCueSheetNext^.EventMode = EM_MAIN) and
           (FCueSheetNext^.StartMode = SM_MANUAL) and
           (FCueSheetNext^.EventStatus <= esCued) and
           (DateTimeToTimecode(NextStartTime - Now) <= GV_SettingTimeParameter.AutoIncreaseDurationBefore) then
        begin
          // 현재 송출중인 이벤트의 길이는 자동으로 늘려줌
          if (FCueSheetCurr <> nil) then
          begin
            FCueSheetCurr.DurationTC := GetPlusTimecode(FCueSheetCurr.DurationTC, GV_SettingTimeParameter.AutoIncreaseDurationAmount);
            CurrIndex := GetCueSheetIndexByItem(FCueSheetCurr);
          end;

          // 다음 이벤트의 시작 시각을 자동으로 늘려줌
          SaveStartTime := FCueSheetNext.StartTime;
          FCueSheetNext.StartTime := GetPlusEventTime(FCueSheetNext.StartTime, TimecodeToEventTime(GV_SettingTimeParameter.AutoIncreaseDurationAmount));
          NextIndex := GetCueSheetIndexByItem(FCueSheetNext);

          ResetStartTimeByTime(NextIndex, SaveStartTime);

          if (FCueSheetCurr <> nil) then
          begin
            OnAirInputEvents(CurrIndex, GV_SettingOption.MaxInputEventCount);

            FLastDisplayNo := GetBeforeMainCountByIndex(CurrIndex);
            DisplayPlayListGrid(CurrIndex);
            DisplayPlayListTimeLine(CurrIndex);
          end
          else
          begin
            OnAirInputEvents(NextIndex, GV_SettingOption.MaxInputEventCount);

            FLastDisplayNo := GetBeforeMainCountByIndex(NextIndex);
            DisplayPlayListGrid(NextIndex);
            DisplayPlayListTimeLine(CurrIndex);
          end;
        end;

        NextStartTimeString := TimecodeToString(FCueSheetNext^.StartTime.T);
        NextDurationString  := TimecodeToString(FCueSheetNext^.DurationTC);
      end
      else
      begin
        NextStartTimeString := IDLE_TIMECODE;
        NextDurationString  := IDLE_TIMECODE;
      end;

      lblNextStart.Caption    := NextStartTimeString;
      lblNextDuration.Caption := NextDurationString;

      // Played, Remaining time
      CurrentTime := Now;

      if (FCueSheetCurr <> nil) then
      begin
        PlayedTime := CurrentTime - EventTimeToDateTime(FCueSheetCurr^.StartTime);
        RemainTime := IncSecond(EventTimeToDateTime(GetEventEndTime(FCueSheetCurr^.StartTime, FCueSheetCurr^.DurationTC)) - CurrentTime);

        PlayedTimeString := FormatDateTime('hh:nn:ss', PlayedTime);
        RemainTimeString := FormatDateTime('hh:nn:ss', RemainTime);
      end
      else
      begin
        if (FCueSheetNext <> nil) then
        begin
          RemainTime := IncSecond(EventTimeToDateTime(FCueSheetNext^.StartTime) - Now);
          RemainTimeString := FormatDateTime('hh:nn:ss', RemainTime);
        end
        else
        begin
          RemainTimeString := IDLE_TIME;
        end;
        PlayedTimeString := IDLE_TIME;
      end;

      lblPlayedTime.Caption    := PlayedTimeString;
      lblRemainingTime.Caption := RemainTimeString;

      // Target event
      if (FCueSheetTarget <> nil) then
      begin
        RemainTargetTime := IncSecond(EventTimeToDateTime(FCueSheetTarget^.StartTime) - CurrentTime);
        RemainTargetTimeString := FormatDateTime('hh:nn:ss', RemainTargetTime)
      end
      else
      begin
        RemainTargetTimeString := IDLE_TIME;
      end;

      lblRemainingTargetTime.Caption := RemainTargetTimeString;

      // Timeline
      with wmtlPlaylist do
      begin
        wmtlPlaylist.FrameNumber := TimecodeToFrame(DateTimeToTimecode(CurrentTime)) +
                                    DaysBetween(FFirstStartTime.D, CurrentTime) * FTimeLineDaysPerFrames;
      end;
    end;

    WM_UPDATE_EVENTSTATUS:
    begin
      CurrIndex := Message.WParam;
      EventStatus := TEventStatus(Message.LParam);

      PopulateEventStatusPlayListGrid(CurrIndex);
    end;

    WM_INVALID_EDIT:
    begin
      C := Message.WParam;
      R := Message.LParam;
            acgPlaylist.Options := acgPlaylist.Options - [goRowSelect];
            acgPlaylist.Options := acgPlaylist.Options + [goEditing];
      acgPlaylist.GotoCell(C, R);
      acgPlaylist.FocusCell(C, R);
//      acgPlaylist.ShowSelection := True;
      acgPlaylist.EditCell(C, R);
      acgPlaylist.SetFocus;
//      acgPlaylist.Col := C;
//      acgPlaylist.Row := R;
//      acgPlaylist.Repaint;
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistCanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
var
  RCol, RRow: Integer;
  Index: Integer;
  Item: PCueSheetItem;
  ParentItem: PCueSheetItem;
begin
  inherited;

  CanEdit := True;
  with (Sender as TAdvColumnGrid) do
  begin
    try
      RCol := RealColIndex(ACol);
      RRow := RealRowIndex(ARow);

      if (RRow < FixedRows) or (RCol < FixedCols) then
      begin
        CanEdit := False;
        exit;
      end;

      CanEdit := (RCol <> IDX_COL_CUESHEET_GROUP) and
                 (RCol <> IDX_COL_CUESHEET_NO) and
                 (RCol <> IDX_COL_CUESHEET_EVENT_MODE) and
                 (RCol <> IDX_COL_CUESHEET_EVENT_STATUS) {and
                 (RCol <> IDX_COL_CUESHEET_DEVICE_STATUS)};

      Index := RRow - FixedRows;

      Item := GetCueSheetItemByIndex(Index);
      if (Item <> nil) then
      begin
        if (Item^.EventStatus > esLoaded) then exit;

        if (Item^.EventMode <> EM_MAIN) then
        begin
          ParentItem := GetParentCueSheetItemByItem(Item);
          if (ParentItem <> nil) then
          begin
            if (ParentItem^.EventStatus > esLoaded) then exit;
          end;
        end;

        with Item^ do
          if (EventMode = EM_COMMENT) then
          begin
            CanEdit := (RCol = IDX_COL_CUESHEET_NO);
            exit;
          end
          else if (EventMode = EM_MAIN) then
          begin
            if (StartMode = SM_ABSOLUTE) then
              CanEdit := (CanEdit)
            else
              CanEdit := (CanEdit) and
                         (RCol <> IDX_COL_CUESHEET_START_DATE) and
                         (RCol <> IDX_COL_CUESHEET_START_TIME);
          end
          else if (EventMode = EM_JOIN) then
          begin
            CanEdit := (CanEdit) and
                       (RCol <> IDX_COL_CUESHEET_START_MODE) and
                       (RCol <> IDX_COL_CUESHEET_START_DATE) and
                       (RCol <> IDX_COL_CUESHEET_START_TIME) and
                       (RCol <> IDX_COL_CUESHEET_DURATON) and
                       (RCol <> IDX_COL_CUESHEET_IN_TC) and
                       (RCol <> IDX_COL_CUESHEET_OUT_TC);
          end
          else
          begin
            CanEdit := (CanEdit) and
                       (RCol <> IDX_COL_CUESHEET_START_DATE);
          end;
      end
      else
        CanEdit := False;

    finally
      if (not CanEdit) then
      begin
        Options := Options + [goRowSelect];
        Options := Options - [goEditing];
      end;
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistClickCell(Sender: TObject; ARow,
  ACol: Integer);
var
  R, C: Integer;
  CurrItem, NextItem: PCueSheetItem;
  NextIndex: Integer;
begin
  inherited;

  with (Sender as TAdvColumnGrid) do
  begin
    if (ACol = IDX_COL_CUESHEET_NO) and (IsNode(ARow)) then
    begin
//      ARow := GetRealRow;

      SetNodeState(Arow, not GetNodeState(ARow));
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistComboChange(Sender: TObject; ACol, ARow,
  AItemIndex: Integer; ASelection: string);
var
  CItem: PCueSheetItem;
  RCol, RRow: Integer;
  Index: Integer;
  SaveInput: TInputType;
  OB: TOutputBkgndType;
  OK: TOutputKeyerType;
begin
  inherited;

  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    if (FChannelCueSheet = nil) then exit;
    if (FChannelCueSheet^.CueSheetList = nil) then exit;

    Index := RRow - FixedRows;

    CItem := GetCueSheetItemByIndex(Index);
    if (CItem <> nil) then
    begin
      with CItem^ do
      begin
        if (RCol = IDX_COL_CUESHEET_START_MODE) then
        begin
          with Columns[IDX_COL_CUESHEET_START_MODE].ComboItems do
            StartMode := TStartMode(Objects[AItemIndex]);
        end
        else if (RCol = IDX_COL_CUESHEET_INPUT) then
        begin
          SaveInput := Input;
          with Columns[IDX_COL_CUESHEET_INPUT].ComboItems do
            Input := TInputType(Objects[AItemIndex]);

          if ((SaveInput in [IT_MAIN, IT_BACKUP]) and (not (Input in [IT_MAIN, IT_BACKUP]))) or
             (not (SaveInput in [IT_MAIN, IT_BACKUP]) and (Input in [IT_MAIN, IT_BACKUP])) then
          begin
            Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems.Clear;

            if (Input in [IT_MAIN, IT_BACKUP]) then
            begin
              for OB := OB_NONE to OB_BOTH do
              begin
                Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems.AddObject(OutputBkgndTypeNames[OB], TObject(OB));
              end;
  //            Output := Integer(GV_SettingEventColumnDefault.OutputBkgnd);
              AllCells[IDX_COL_CUESHEET_OUTPUT, RRow] := OutputBkgndTypeNames[GV_SettingEventColumnDefault.OutputBkgnd];
            end
            else
            begin
              for OK := OK_NONE to OK_OFF do
              begin
                Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems.AddObject(OutputKeyerTypeNames[OK], TObject(OK));
              end;
  //            Output := Integer(GV_SettingEventColumnDefault.OutputKeyer);
              AllCells[IDX_COL_CUESHEET_OUTPUT, RRow] := OutputKeyerTypeNames[GV_SettingEventColumnDefault.OutputKeyer];
            end;
          end;
        end
        else if (RCol = IDX_COL_CUESHEET_OUTPUT) then
        begin
          with Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems do
            Output := Integer(Objects[AItemIndex]);
        end
        else if (RCol = IDX_COL_CUESHEET_TR_TYPE) then
        begin
          with Columns[IDX_COL_CUESHEET_TR_TYPE].ComboItems do
            TransitionType := TTRType(Objects[AItemIndex]);
        end
        else if (RCol = IDX_COL_CUESHEET_TR_RATE) then
        begin
          with Columns[IDX_COL_CUESHEET_TR_RATE].ComboItems do
            TransitionRate := TTRRate(Objects[AItemIndex]);
        end
        else if (RCol = IDX_COL_CUESHEET_PROGRAM_TYPE) then
        begin
          with Columns[IDX_COL_CUESHEET_PROGRAM_TYPE].ComboItems do
            ProgramType := Byte(Objects[AItemIndex]);
        end;
      end;
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistComboDropDown(Sender: TObject; ARow,
  ACol: Integer);
var
  CItem: PCueSheetItem;
  RCol, RRow: Integer;
  Index: Integer;
  SM: TStartMode;
  OB: TOutputBkgndType;
  OK: TOutputKeyerType;
begin
  inherited;
exit;
  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    Index := RRow - FixedRows;

    CItem := GetCueSheetItemByIndex(Index);
    if (CItem <> nil) then
    begin
      if (RCol = IDX_COL_CUESHEET_START_MODE) then
      begin
        with CItem^ do
        begin
          Columns[ACol].ComboItems.Clear;
          case EventMode of
            EM_MAIN:
            begin
              for SM := SM_ABSOLUTE to SM_LOOP do
              begin
                Columns[ACol].ComboItems.AddObject(StartModeNames[SM], TObject(SM));
              end;
            end;
            EM_JOIN:
            begin
              for SM := SM_SUBBEGIN to SM_SUBBEGIN do
              begin
                Columns[ACol].ComboItems.AddObject(StartModeNames[SM], TObject(SM));
              end;
            end;
            EM_SUB:
            begin
              for SM := SM_SUBBEGIN to SM_SUBEND do
              begin
                Columns[ACol].ComboItems.AddObject(StartModeNames[SM], TObject(SM));
              end;
            end;
          end;
        end;
      end
      else if (RCol = IDX_COL_CUESHEET_OUTPUT) then
      begin
        with CItem^ do
        begin
          Columns[ACol].ComboItems.Clear;
          if (Input in [IT_MAIN, IT_BACKUP]) then
          begin
            for OB := OB_NONE to OB_BOTH do
            begin
              Columns[ACol].ComboItems.AddObject(OutputBkgndTypeNames[OB], TObject(OB));
            end;
          end
          else
          begin
            for OK := OK_NONE to OK_OFF do
            begin
              Columns[ACol].ComboItems.AddObject(OutputKeyerTypeNames[OK], TObject(OK));
            end;
          end;
        end;
      end;
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistCustomCellDraw(Sender: TObject;
  Canvas: TCanvas; ACol, ARow: Integer; AState: TGridDrawState; ARect: TRect;
  Printing: Boolean);
begin
  inherited;
  exit;
  with (Sender as TAdvColumnGrid) do
  begin
//      if (GV_ClipboardCueSheetList.IndexOf(CItem) >= 0) then
//      begin
//        ABrush.Style := bsDiagCross;
//      end;

    if (gdSelected in AState) or (gdRowSelected in AState) then
    begin
      Canvas.Brush.Style := bsDiagCross;
      Canvas.Pen.Color := clRed;

      Canvas.MoveTo(ARect.Left, ARect.Top);
      Canvas.LineTo(ARect.Right + 1, ARect.Top);

      Canvas.MoveTo(ARect.Left, ARect.Bottom - 1);
      Canvas.LineTo(ARect.Right + 1, ARect.Bottom - 1);

      SetTextColor(Canvas.Handle, clWhite);

{      if (ACol = 0) then
      begin
        Canvas.MoveTo(Rect.Left, Rect.Top);
        Canvas.LineTo(Rect.Left, Rect.Bottom);
      end
      else if (ACol = ColCount - 1) or (IsMergedCell(ACol, ARow)) then
      begin
        Canvas.MoveTo(Rect.Right, Rect.Top);
        Canvas.LineTo(Rect.Right, Rect.Bottom);
      end; }
    end
    else
    begin
      SetTextColor(Canvas.Handle, Font.Color);
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistDblClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
  inherited;
  with (Sender as TAdvColumnGrid) do
  begin
    if (EditMode) then exit;

    Options := Options - [goRowSelect];
    Options := Options + [goEditing];
    if (IsEditable(ACol, ARow)) then
    begin
//      ShowSelection := True;
      EditCell(ACol, ARow);
    end
    else
    begin
    Options := Options + [goRowSelect];
    Options := Options - [goEditing];
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  R: TRect;
  RRow: Integer;
  Item: PCueSheetItem;
begin
  inherited;
//  exit;
  with (Sender as TAdvColumnGrid) do
  begin
    if ((gdSelected in State) or (gdRowSelected in State)) or
       (RowSelect[ARow]) then
    begin
      Canvas.Brush.Style := bsClear;
      Canvas.Pen.Color := clRed;

      Canvas.MoveTo(Rect.Left, Rect.Top);
      Canvas.LineTo(Rect.Right + 1, Rect.Top);

      Canvas.MoveTo(Rect.Left, Rect.Bottom);
      Canvas.LineTo(Rect.Right + 1, Rect.Bottom);

      SetTextColor(Canvas.Handle, clWhite);

{      if (ACol = 0) then
      begin
        Canvas.MoveTo(Rect.Left, Rect.Top);
        Canvas.LineTo(Rect.Left, Rect.Bottom);
      end
      else if (ACol = ColCount - 1) or (IsMergedCell(ACol, ARow)) then
      begin
        Canvas.MoveTo(Rect.Right, Rect.Top);
        Canvas.LineTo(Rect.Right, Rect.Bottom);
      end; }
    end
    else
    begin
      SetTextColor(Canvas.Handle, Font.Color);
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistEditCellDone(Sender: TObject; ACol,
  ARow: Integer);
var
  RCol, RRow: Integer;
  Value: String;
  Index: Integer;
  CItem: PCueSheetItem;     // Current cuesheet item
  PItem: PCueSheetItem;     // Prior or parent cuesheet item
  CStartTime: TEventTime;   // Current start time
  CEndTime: TEventTime;     // Current end time
  PEndTime: TEventTime;     // Prior or parent end time
  DurTime: TEventTime;      // Duration time of current and modified start time
  SaveTime: TEventTime;     // Saved current start time

  CDurTC: TTimecode;        // Current duration tc
  SaveDurTC: TTimecode;     // Saved current duration tc
  IT: TInputType;

  EnterStartTime: TEventTime;
  EnterDate: TDate;
  EnterTC: TTimecode;

  SaveStartTime: TEventTime;
  SaveDurationTC: TTimecode;

  P: TPoint;
begin
  inherited;
  // Check the cell validate.

//      MessageBeep(MB_ICONWARNING);
//      Application.MessageBox(PChar('111'), PChar(Application.Title), MB_OK or MB_ICONWARNING);

//acgPlaylist.ShowSelection := True;
//ShowMessage('1');
//            acgPlaylist.Options := acgPlaylist.Options + [goRowSelect];
//      acgPlaylist.SetFocus;
//acgPlaylist.ShowSelection := False;
//      exit;

  with (Sender as TAdvColumnGrid) do
  begin
    Options := Options + [goRowSelect];
    Options := Options - [goEditing];
  end;


  if (not FIsCellModified) then exit;

  with (Sender as TAdvColumnGrid) do
  begin
//      ShowSelection := True;
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    Index := RRow - FixedRows;

    CItem := GetCueSheetItemByIndex(Index);
    if (CItem <> nil) then
    begin
      Value := Cells[ACol, ARow];
      with CItem^ do
      begin
        if (RCol = IDX_COL_CUESHEET_NO) then
        begin
          if (EventMode = EM_COMMENT) then
          begin
            if (not IsValidStartTime(CItem, CItem^.StartTime.T)) then
            begin
              StrPCopy(Title, Value);
              exit;
            end;
          end;
        end
        else if (RCol = IDX_COL_CUESHEET_START_MODE) then
        begin
          if (EventMode = EM_SUB) then
          begin
            if (not IsValidStartTime(CItem, CItem^.StartTime.T)) then
            begin
              EditCell(ACol, ARow);
              exit;
            end;
          end;
        end
        else if (RCol = IDX_COL_CUESHEET_START_DATE) then
        begin
          EnterDate := StrToDate(Value);
          if (IsValidStartDate(CItem, EnterDate)) then
          begin
{            if (EventMode = EM_MAIN) then
            begin
              EnterStartTime := StartTime;
              EnterStartTime.D := EnterDate;

//              ResetStartDate(RRow - FixedRows + 1, Trunc(EnterDate - StartTime.D));

//              StartTime.D := EnterDate;
//              ResetStartTime(Index);

              ResetStartTimeByTime(Index, EnterStartTime);
              ResetChildItems(Index);

//              ResetStartDateTimeLine(RRow - FixedRows);
            end;   }

{            EnterStartTime := StartTime;
            EnterStartTime.D := EnterDate; }

            SaveStartTime := StartTime;
            StartTime.D   := EnterDate;

            ResetStartTimeByTime(Index, SaveStartTime);
//            ResetChildItems(Index);
          end
          else
          begin
//            GotoCell(ACol, ARow);
//      Selection := TGridRect(Rect(FixedCols, ARow, ColCount - 1, ARow));
//          RepaintRect(TRect(Selection));

//            Options := Options - [goRowSelect];
//            Options := Options + [goEditing];
//            ShowInplaceEdit;
//            Col := ACol;
//            Row := ARow;
//            PostMessage(Handle, WM_GRIDEDITSHOW, 0, 0);
//            PostMessage(Self.Handle, WM_INVALID_EDIT, ACol, ARow);

//            Col := ColCount - 1;
//            FocusCell(ACol, ARow);
//              Row := ARow;
//              Col := ACol;
//              P := CellRect(ACol, ARow).TopLeft;
//              PostMessage(Handle, WM_LBUTTONDBLCLK, 0, MakeLong(P.X, P.Y));
            EditCell(ACol, ARow);
//      ShowSelection := True;
//      SetFocus;
//            acgPlaylistDrawCell(acgPlaylist, ACol, ARow, Rect(0, 0, 0, 0), [gdSelected]);
//            Options := Options + [goRowSelect];
//            Options := Options - [goEditing];
            exit;
          end;

{          if (EventMode = EM_MAIN) then
          begin
            SaveTime     := StartTime;

            CStartTime   := SaveTime;
            CStartTime.D := StrToDate(Value);

            // Checks whether the entered start time is less than the start time of the previous event.
            PItem := GetBeforeMainItemByItem(CItem);
            if (PItem <> nil) then
            begin
              PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//              ShowMessage(EventTimeToDateTimecodeStr(CurrStartTime));
//              ShowMessage(EventTimeToDateTimecodeStr(BefoStartTime));
              if (CompareEventTime(CStartTime, PEndTime) < 0) then
              begin
                MessageBeep(MB_ICONWARNING);
                MessageBox(Handle, PChar(SStartTimeGreaterThenBeforeEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                EditCell(ACol, ARow);
                exit;
              end;
            end;

            StartTime.D := CStartTime.D;
            ResetStartDate(RRow - FixedRows + 1, Trunc(CStartTime.D - SaveTime.D));
          end; }
        end
        else if (RCol = IDX_COL_CUESHEET_START_TIME) then
        begin
          EnterTC := StringtoTimecode(Value);
          if (IsValidStartTime(CItem, EnterTC)) then
          begin
{            StartTime.T := EnterTC;
            if (EventMode = EM_MAIN) then
            begin

//              ResetStartTime(Index);

              ResetStartTimeByTime(Index, EnterStartTime);
              ResetChildItems(Index);
            end; }
{
            EnterStartTime := StartTime;
            EnterStartTime.T := EnterTC; }

            SaveStartTime := StartTime;
            StartTime.T   := EnterTC;

            ResetStartTimeByTime(Index, SaveStartTime);

//            ResetChildItems(Index);
          end
          else
          begin
            EditCell(ACol, ARow);
            exit;
          end;

{          SaveTime     := StartTime;

          CStartTime   := SaveTime;
          CStartTime.T := StringToTimecode(Value);

          if (EventMode = EM_MAIN) then
          begin
            // Check that the entered timecode is validate.
            if (not IsValidTimecode(CStartTime.T)) then
            begin
              MessageBeep(MB_ICONWARNING);
              MessageBox(Handle, PChar(SInvalidTimeocde), PChar(Application.Title), MB_OK or MB_ICONWARNING);
//                MessageDlg(SInvalidTimeocde, mtWarning, [mbOK], MB_ICONWARNING);
              EditCell(ACol, ARow);
              exit;
            end;

            // Checks whether the entered start time is less than the start time of the previous event.
            PItem := GetBeforeMainItemByItem(CItem);
            if (PItem <> nil) then
            begin
              PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//              ShowMessage(EventTimeToDateTimecodeStr(CurrStartTime));
//              ShowMessage(EventTimeToDateTimecodeStr(BefoStartTime));
              if (CompareEventTime(CStartTime, PEndTime) < 0) then
              begin
                MessageBeep(MB_ICONWARNING);
                MessageBox(Handle, PChar(SStartTimeGreaterThenBeforeEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                EditCell(ACol, ARow);
                exit;
              end;
            end;

            StartTime.T := CStartTime.T;
            DurTime := GetDurEventTime(SaveTime, CStartTime);
//              ShowMessage(EventTimeToDateTimecodeStr(DurTime));
            ResetStartTimePlus(RRow - FixedRows + 1, DurTime);
//            ResetStartTimeMinus(RRow - FixedRows + 1, DurTime);
          end
          else if (EventMode in [EM_JOIN, EM_SUB]) then
          begin
            if (StartMode = SM_SUBBEGIN) then
            begin
              // Checks whether the entered start time is less than the end time of the parent event.
              PItem := GetParentCueSheetItemByItem(CItem);
              if (PItem <> nil) then
              begin
                PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);

                CStartTime.D := 0;
                CEndTime := GetEventEndTime(GetPlusEventTime(PItem^.StartTime, CStartTime), DurationTC);
                if (CompareEventTime(CEndTime, PEndTime) > 0) then
                begin
                  MessageBeep(MB_ICONWARNING);
                  MessageBox(Handle, PChar(SSubStartTimeLessThenParentEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                  EditCell(ACol, ARow);
                  exit;
                end;
              end;

              StartTime.T := CStartTime.T;
            end
            else
            begin
              // Checks whether the entered start time is less than the end time of the parent event.
              PItem := GetParentCueSheetItemByItem(CItem);
              if (PItem <> nil) then
              begin
                PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);

                CStartTime.D := 0;
                CEndTime := GetEventEndTime(GetMinusEventTime(PEndTime, CStartTime), DurationTC);
                if (CompareEventTime(CEndTime, PEndTime) > 0) then
                begin
                  MessageBeep(MB_ICONWARNING);
                  MessageBox(Handle, PChar(SSubStartTimeLessThenParentEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                  EditCell(ACol, ARow);
                  exit;
                end;
              end;

              StartTime.T := CStartTime.T;
            end;
          end; }
        end
        else if (RCol = IDX_COL_CUESHEET_INPUT) then
        begin
          with Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems do
            Output := Integer(Objects[IndexOf(Cells[IDX_COL_CUESHEET_OUTPUT, ARow])]);
        end
        else if (RCol = IDX_COL_CUESHEET_TITLE) then
          StrPCopy(Title, Value)
        else if (RCol = IDX_COL_CUESHEET_SUB_TITLE) then
          StrPCopy(SubTitle, Value)
        else if (RCol = IDX_COL_CUESHEET_SOURCE) then
          StrPCopy(Source, Value)
        else if (RCol = IDX_COL_CUESHEET_MEDIA_ID) then
          StrPCopy(MediaId, Value)
        else if (RCol = IDX_COL_CUESHEET_DURATON) then
        begin
          EnterTC := StringToTimecode(Value);
          if (IsValidDuration(CItem, EnterTC)) then
          begin
{            DurationTC := EnterTC;
            if (EventMode = EM_MAIN) then
            begin
              ResetStartTime(Index);
              ResetChildItems(Index);
            end;
//            DurationTC := EnterTC; }

            SaveDurationTC := DurationTC;
            DurationTC     := EnterTC;

            ResetStartTimeByTime(Index, SaveDurationTC);

//            ResetStartTimeByDuration(Index, EnterTC);
//            ResetChildItems(Index);
          end
          else
          begin
            EditCell(ACol, ARow);
            exit;
          end;
{

          SaveTime.D := 0;
          SaveTime.T := DurationTC;

          CStartTime.D := 0;
          CStartTime.T := StringToTimecode(Value);

          // Check that the entered timecode is validate.
          if (not IsValidTimecode(CStartTime.T)) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInvalidTimeocde), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          // Feture add the media duration validate
          //
          //

          if (EventMode = EM_MAIN) then
          begin
            DurationTC := CStartTime.T;
            DurTime := GetDurEventTime(SaveTime, CStartTime);
            ResetStartTimePlus(RRow - FixedRows + 1, DurTime);
          end
          else if (EventMode in [EM_JOIN, EM_SUB]) then
          begin
            if (StartMode = SM_SUBBEGIN) then
            begin
              // Checks whether the entered end time is less than the end time of the parent event.
              PItem := GetParentCueSheetItemByItem(CItem);
              if (PItem <> nil) then
              begin
                PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//                ShowMessage(EventTimeToDateTimecodeStr(PEndTime));
                CEndTime := GetEventTimeSubBegin(PItem^.StartTime, CItem^.InTC);
                CEndTime := GetEventEndTime(CEndTime, CStartTime.T);
//                ShowMessage(EventTimeToDateTimecodeStr(CEndTime));

                if (CompareEventTime(CEndTime, PEndTime) > 0) then
                begin
                  MessageBeep(MB_ICONWARNING);
                  MessageBox(Handle, PChar(SSubStartTimeLessThenParentEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                  EditCell(ACol, ARow);
                  exit;
                end;
              end;

              DurationTC := CStartTime.T;
            end
            else
            begin
              // Checks whether the entered start time is less than the end time of the parent event.
              PItem := GetParentCueSheetItemByItem(CItem);
              if (PItem <> nil) then
              begin
                PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//                ShowMessage(EventTimeToDateTimecodeStr(PEndTime));

                CEndTime := GetEventTimeSubEnd(PItem^.StartTime, PItem^.DurationTC, CItem^.InTC);
                CEndTime := GetEventEndTime(CEndTime, CStartTime.T);
//                ShowMessage(EventTimeToDateTimecodeStr(CEndTime));

                if (CompareEventTime(CEndTime, PEndTime) > 0) then
                begin
                  MessageBeep(MB_ICONWARNING);
                  MessageBox(Handle, PChar(SSubStartTimeLessThenParentEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                  EditCell(ACol, ARow);
                  exit;
                end;
              end;

              DurationTC := CStartTime.T;
            end;
          end; }
        end
        else if (RCol = IDX_COL_CUESHEET_IN_TC) then
        begin
          EnterTC := StringToTimecode(Value);
          if (IsValidInTC(CItem, EnterTC)) then
          begin
            InTC := EnterTC;
            if (EventMode = EM_MAIN) then
              ResetChildItems(Index);
          end
          else
          begin
            EditCell(ACol, ARow);
            exit;
          end;

{          SaveTime.D := 0;
          SaveTime.T := InTC;

          CStartTime.D := 0;
          CStartTime.T := StringToTimecode(Value);

          // Check that the entered timecode is validate.
          if (not IsValidTimecode(CStartTime.T)) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInvalidTimeocde), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          if (CStartTime.T >= DurationTC) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInTCLessThenDurationTC), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          if (CStartTime.T > OutTC) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInTCLessThenOutTC), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          InTC := CStartTime.T;  }
        end
        else if (RCol = IDX_COL_CUESHEET_OUT_TC) then
        begin
          EnterTC := StringToTimecode(Value);
          if (IsValidOutTC(CItem, EnterTC)) then
          begin
            OutTC := EnterTC;
            if (EventMode = EM_MAIN) then
              ResetChildItems(Index);
          end
          else
          begin
            EditCell(ACol, ARow);
            exit;
          end;

{          SaveTime.D := 0;
          SaveTime.T := OutTC;

          CStartTime.D := 0;
          CStartTime.T := StringToTimecode(Value);

          // Check that the entered timecode is validate.
          if (not IsValidTimecode(CStartTime.T)) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInvalidTimeocde), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          if (CStartTime.T >= DurationTC) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInTCLessThenDurationTC), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          if (CStartTime.T < InTC) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInTCLessThenOutTC), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          OutTC := CStartTime.T; }
        end
        else if (RCol = IDX_COL_CUESHEET_NOTES) then
          StrPCopy(Notes, Value);
      end;
    end;
//      ShowSelection := False;
  end;

  FIsCellEditing  := False;
  FIsCellModified := False;
end;

procedure TfrmChannel.acgPlaylistEditChange(Sender: TObject; ACol,
  ARow: Integer; Value: string);
begin
  inherited;
  if (not FIsCellModified) then
  begin
    with (Sender as TAdvColumnGrid) do
      FOldCellValue := Cells[ACol, ARow];
    FIsCellModified := True;
  end;
end;

procedure TfrmChannel.acgPlaylistEditingDone(Sender: TObject);
var
  RCol, RRow: Integer;
  Value: String;
  CItem: PCueSheetItem;     // Current cuesheet item
  PItem: PCueSheetItem;     // Prior or parent cuesheet item
  CStartTime: TEventTime;   // Current start time
  CEndTime: TEventTime;     // Current end time
  PEndTime: TEventTime;     // Prior or parent end time
  DurTime: TEventTime;      // Duration time of current and modified start time
  SaveTime: TEventTime;     // Saved current start time

  CDurTC: TTimecode;        // Current duration tc
  SaveDurTC: TTimecode;     // Saved current duration tc
  IT: TInputType;

  EnterStartTime: TEventTime;
  EnterDate: TDate;
  EnterTC: TTimecode;
begin
  inherited;
  exit;
  with (Sender as TAdvColumnGrid) do
  begin
    Options := Options + [goRowSelect];
    Options := Options - [goEditing];
  end;

  exit;
  // Check the cell validate.

  if (not FIsCellModified) then exit;

  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(Col);
    RRow := RealRowIndex(Row);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    CItem := GetCueSheetItemByIndex(RRow - FixedRows);
    if (CItem <> nil) then
    begin
      Value := Cells[Col, Row];
      with CItem^ do
      begin
        if (RCol = IDX_COL_CUESHEET_START_MODE) then
        begin
          with Columns[IDX_COL_CUESHEET_START_MODE].ComboItems do
            StartMode := TStartMode(Objects[IndexOf(Value)]);
        end
        else if (RCol = IDX_COL_CUESHEET_START_DATE) then
        begin
          EnterDate := StrToDate(Value);
          if (IsValidStartDate(CItem, EnterDate)) then
          begin
            if (EventMode = EM_MAIN) then
            begin
              ResetStartDate(RRow - FixedRows + 1, Trunc(EnterDate - StartTime.D));
              StartTime.D := EnterDate;
            end;
          end
          else
          begin
            EditCell(Col, Row);
            exit;
          end;

{          if (EventMode = EM_MAIN) then
          begin
            SaveTime     := StartTime;

            CStartTime   := SaveTime;
            CStartTime.D := StrToDate(Value);

            // Checks whether the entered start time is less than the start time of the previous event.
            PItem := GetBeforeMainItemByItem(CItem);
            if (PItem <> nil) then
            begin
              PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//              ShowMessage(EventTimeToDateTimecodeStr(CurrStartTime));
//              ShowMessage(EventTimeToDateTimecodeStr(BefoStartTime));
              if (CompareEventTime(CStartTime, PEndTime) < 0) then
              begin
                MessageBeep(MB_ICONWARNING);
                MessageBox(Handle, PChar(SStartTimeGreaterThenBeforeEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                EditCell(ACol, ARow);
                exit;
              end;
            end;

            StartTime.D := CStartTime.D;
            ResetStartDate(RRow - FixedRows + 1, Trunc(CStartTime.D - SaveTime.D));
          end; }
        end
        else if (RCol = IDX_COL_CUESHEET_START_TIME) then
        begin
          EnterTC := StringtoTimecode(Value);
          if (IsValidStartTime(CItem, EnterTC)) then
          begin
            if (EventMode = EM_MAIN) then
            begin
              EnterStartTime   := StartTime;
              EnterStartTime.T := EnterTC;

              DurTime := GetDurEventTime(StartTime, EnterStartTime);
              ResetStartTimePlus(RRow - FixedRows + 1, DurTime);
            end;

            StartTime.T := EnterTC;
          end
          else
          begin
            EditCell(Col, Row);
            exit;
          end;

{          SaveTime     := StartTime;

          CStartTime   := SaveTime;
          CStartTime.T := StringToTimecode(Value);

          if (EventMode = EM_MAIN) then
          begin
            // Check that the entered timecode is validate.
            if (not IsValidTimecode(CStartTime.T)) then
            begin
              MessageBeep(MB_ICONWARNING);
              MessageBox(Handle, PChar(SInvalidTimeocde), PChar(Application.Title), MB_OK or MB_ICONWARNING);
//                MessageDlg(SInvalidTimeocde, mtWarning, [mbOK], MB_ICONWARNING);
              EditCell(ACol, ARow);
              exit;
            end;

            // Checks whether the entered start time is less than the start time of the previous event.
            PItem := GetBeforeMainItemByItem(CItem);
            if (PItem <> nil) then
            begin
              PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//              ShowMessage(EventTimeToDateTimecodeStr(CurrStartTime));
//              ShowMessage(EventTimeToDateTimecodeStr(BefoStartTime));
              if (CompareEventTime(CStartTime, PEndTime) < 0) then
              begin
                MessageBeep(MB_ICONWARNING);
                MessageBox(Handle, PChar(SStartTimeGreaterThenBeforeEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                EditCell(ACol, ARow);
                exit;
              end;
            end;

            StartTime.T := CStartTime.T;
            DurTime := GetDurEventTime(SaveTime, CStartTime);
//              ShowMessage(EventTimeToDateTimecodeStr(DurTime));
            ResetStartTimePlus(RRow - FixedRows + 1, DurTime);
//            ResetStartTimeMinus(RRow - FixedRows + 1, DurTime);
          end
          else if (EventMode in [EM_JOIN, EM_SUB]) then
          begin
            if (StartMode = SM_SUBBEGIN) then
            begin
              // Checks whether the entered start time is less than the end time of the parent event.
              PItem := GetParentCueSheetItemByItem(CItem);
              if (PItem <> nil) then
              begin
                PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);

                CStartTime.D := 0;
                CEndTime := GetEventEndTime(GetPlusEventTime(PItem^.StartTime, CStartTime), DurationTC);
                if (CompareEventTime(CEndTime, PEndTime) > 0) then
                begin
                  MessageBeep(MB_ICONWARNING);
                  MessageBox(Handle, PChar(SSubStartTimeLessThenParentEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                  EditCell(ACol, ARow);
                  exit;
                end;
              end;

              StartTime.T := CStartTime.T;
            end
            else
            begin
              // Checks whether the entered start time is less than the end time of the parent event.
              PItem := GetParentCueSheetItemByItem(CItem);
              if (PItem <> nil) then
              begin
                PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);

                CStartTime.D := 0;
                CEndTime := GetEventEndTime(GetMinusEventTime(PEndTime, CStartTime), DurationTC);
                if (CompareEventTime(CEndTime, PEndTime) > 0) then
                begin
                  MessageBeep(MB_ICONWARNING);
                  MessageBox(Handle, PChar(SSubStartTimeLessThenParentEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                  EditCell(ACol, ARow);
                  exit;
                end;
              end;

              StartTime.T := CStartTime.T;
            end;
          end; }
        end
        else if (RCol = IDX_COL_CUESHEET_INPUT) then
        begin
          with Columns[IDX_COL_CUESHEET_INPUT].ComboItems do
            Input := TInputType(Objects[IndexOf(Value)]);

          if (Input in [IT_MAIN, IT_BACKUP]) then
          begin
            Output := Integer(GV_SettingEventColumnDefault.OutputBkgnd);
          end
          else
          begin
            Output := Integer(GV_SettingEventColumnDefault.OutputKeyer);
          end;
        end
        else if (RCol = IDX_COL_CUESHEET_OUTPUT) then
        begin
          with Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems do
            Output := Integer(Objects[IndexOf(Value)]);
        end
        else if (RCol = IDX_COL_CUESHEET_TITLE) then
          StrPCopy(Title, Value)
        else if (RCol = IDX_COL_CUESHEET_SUB_TITLE) then
          StrPCopy(SubTitle, Value)
        else if (RCol = IDX_COL_CUESHEET_SOURCE) then
          StrPCopy(Source, Value)
        else if (RCol = IDX_COL_CUESHEET_MEDIA_ID) then
          StrPCopy(MediaId, Value)
        else if (RCol = IDX_COL_CUESHEET_DURATON) then
        begin
          EnterTC := StringToTimecode(Value);
          if (IsValidDuration(CItem, EnterTC)) then
          begin
            if (EventMode = EM_MAIN) then
            begin
              SaveTime.D := 0;
              SaveTime.T := DurationTC;

              CStartTime.D := 0;
              CStartTime.T := EnterTC;

              DurTime := GetDurEventTime(SaveTime, CStartTime);
              ResetStartTimePlus(RRow - FixedRows + 1, DurTime);
            end;
            DurationTC := EnterTC;
          end
          else
          begin
            EditCell(Col, Row);
            exit;
          end;
{

          SaveTime.D := 0;
          SaveTime.T := DurationTC;

          CStartTime.D := 0;
          CStartTime.T := StringToTimecode(Value);

          // Check that the entered timecode is validate.
          if (not IsValidTimecode(CStartTime.T)) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInvalidTimeocde), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          // Feture add the media duration validate
          //
          //

          if (EventMode = EM_MAIN) then
          begin
            DurationTC := CStartTime.T;
            DurTime := GetDurEventTime(SaveTime, CStartTime);
            ResetStartTimePlus(RRow - FixedRows + 1, DurTime);
          end
          else if (EventMode in [EM_JOIN, EM_SUB]) then
          begin
            if (StartMode = SM_SUBBEGIN) then
            begin
              // Checks whether the entered end time is less than the end time of the parent event.
              PItem := GetParentCueSheetItemByItem(CItem);
              if (PItem <> nil) then
              begin
                PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//                ShowMessage(EventTimeToDateTimecodeStr(PEndTime));
                CEndTime := GetEventTimeSubBegin(PItem^.StartTime, CItem^.InTC);
                CEndTime := GetEventEndTime(CEndTime, CStartTime.T);
//                ShowMessage(EventTimeToDateTimecodeStr(CEndTime));

                if (CompareEventTime(CEndTime, PEndTime) > 0) then
                begin
                  MessageBeep(MB_ICONWARNING);
                  MessageBox(Handle, PChar(SSubStartTimeLessThenParentEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                  EditCell(ACol, ARow);
                  exit;
                end;
              end;

              DurationTC := CStartTime.T;
            end
            else
            begin
              // Checks whether the entered start time is less than the end time of the parent event.
              PItem := GetParentCueSheetItemByItem(CItem);
              if (PItem <> nil) then
              begin
                PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//                ShowMessage(EventTimeToDateTimecodeStr(PEndTime));

                CEndTime := GetEventTimeSubEnd(PItem^.StartTime, PItem^.DurationTC, CItem^.InTC);
                CEndTime := GetEventEndTime(CEndTime, CStartTime.T);
//                ShowMessage(EventTimeToDateTimecodeStr(CEndTime));

                if (CompareEventTime(CEndTime, PEndTime) > 0) then
                begin
                  MessageBeep(MB_ICONWARNING);
                  MessageBox(Handle, PChar(SSubStartTimeLessThenParentEndTime), PChar(Application.Title), MB_OK or MB_ICONWARNING);
                  EditCell(ACol, ARow);
                  exit;
                end;
              end;

              DurationTC := CStartTime.T;
            end;
          end; }
        end
        else if (RCol = IDX_COL_CUESHEET_IN_TC) then
        begin
          EnterTC := StringToTimecode(Value);
          if (IsValidInTC(CItem, EnterTC)) then
          begin
            InTC := EnterTC;
          end
          else
          begin
            EditCell(Col, Row);
            exit;
          end;

{          SaveTime.D := 0;
          SaveTime.T := InTC;

          CStartTime.D := 0;
          CStartTime.T := StringToTimecode(Value);

          // Check that the entered timecode is validate.
          if (not IsValidTimecode(CStartTime.T)) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInvalidTimeocde), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          if (CStartTime.T >= DurationTC) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInTCLessThenDurationTC), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          if (CStartTime.T > OutTC) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInTCLessThenOutTC), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          InTC := CStartTime.T;  }
        end
        else if (RCol = IDX_COL_CUESHEET_OUT_TC) then
        begin
          EnterTC := StringToTimecode(Value);
          if (IsValidOutTC(CItem, EnterTC)) then
          begin
            OutTC := EnterTC;
          end
          else
          begin
            EditCell(Col, Row);
            exit;
          end;

{          SaveTime.D := 0;
          SaveTime.T := OutTC;

          CStartTime.D := 0;
          CStartTime.T := StringToTimecode(Value);

          // Check that the entered timecode is validate.
          if (not IsValidTimecode(CStartTime.T)) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInvalidTimeocde), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          if (CStartTime.T >= DurationTC) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInTCLessThenDurationTC), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          if (CStartTime.T < InTC) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SInTCLessThenOutTC), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            EditCell(ACol, ARow);
            exit;
          end;

          OutTC := CStartTime.T; }
        end
        else if (RCol = IDX_COL_CUESHEET_TR_TYPE) then
        begin
          with Columns[IDX_COL_CUESHEET_TR_TYPE].ComboItems do
            TransitionType := TTRType(Objects[IndexOf(Value)]);
        end
        else if (RCol = IDX_COL_CUESHEET_TR_RATE) then
        begin
          with Columns[IDX_COL_CUESHEET_TR_RATE].ComboItems do
            TransitionRate := TTRRate(Objects[IndexOf(Value)]);
        end
        else if (RCol = IDX_COL_CUESHEET_PROGRAM_TYPE) then
        begin
          with Columns[IDX_COL_CUESHEET_PROGRAM_TYPE].ComboItems do
            ProgramType := Byte(Objects[IndexOf(Value)]);
        end
        else if (RCol = IDX_COL_CUESHEET_NOTES) then
          StrPCopy(Notes, Value);
      end;
    end;
  end;

  FIsCellEditing  := False;
  FIsCellModified := False;
end;

procedure TfrmChannel.acgPlaylistGetCellBorder(Sender: TObject; ARow,
  ACol: Integer; APen: TPen; var Borders: TCellBorders);
var
  RCol, RRow: Integer;
begin
  inherited;
  exit;
  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    if (RowSelect[RRow]) then
      Borders := [cbTop, cbBottom];

    if (RCol = IDX_COL_CUESHEET_GROUP) then
    begin
      Borders := Borders + [cbRight];
      exit;
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistGetCellBorderProp(Sender: TObject; ARow,
  ACol: Integer; LeftPen, TopPen, RightPen, BottomPen: TPen);
var
  RCol, RRow: Integer;
begin
  inherited;
  exit;
  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    TopPen.Color := clRed;
    BottomPen.Color := clRed;

    if (RCol = IDX_COL_CUESHEET_GROUP) then
    begin
      RightPen.Color := GridLineColor;
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistGetCellColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
  RCol, RRow: Integer;
  Item: PCueSheetItem;
begin
  inherited;

  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    if (FChannelCueSheet = nil) then exit;
    if (FChannelCueSheet^.CueSheetList = nil) then exit;

    Item := GetCueSheetItemByIndex(RRow - CNT_CUESHEET_HEADER);
    if (Item <> nil) then
    begin
      if {(GV_ClipboardCueSheet.PasteMode = pmCut) and }(GV_ClipboardCueSheet.IndexOf(Item) >= 0) then
      begin
        ABrush.Color := COLOR_BK_CLIPBOARD;
        AFont.Color  := COLOR_TX_CLIPBOARD;
      end
      else if (FCueSheetNext <> nil) and (Item^.GroupNo = FCueSheetNext^.GroupNo) and
              (FCueSheetNext^.EventStatus in [esIdle, esLoaded]) and (Item^.EventStatus <> esError) then
      begin
        ABrush.Color := COLOR_BK_EVENTSTATUS_NEXT;
        AFont.Color  := COLOR_TX_EVENTSTATUS_NEXT;
      end
      else
      begin
        case Item^.EventStatus of
          esCueing..esPreroll:
          begin
            ABrush.Color := COLOR_BK_EVENTSTATUS_CUED;
            AFont.Color  := COLOR_TX_EVENTSTATUS_CUED;
          end;
          esOnAir:
          begin
            ABrush.Color := COLOR_BK_EVENTSTATUS_ONAIR;
            AFont.Color  := COLOR_TX_EVENTSTATUS_ONAIR;
          end;
          esSkipped,
          esFinish..esDone:
          begin
            ABrush.Color := COLOR_BK_EVENTSTATUS_DONE;
            AFont.Color  := COLOR_TX_EVENTSTATUS_DONE;
          end;
          esError:
          begin
            ABrush.Color := COLOR_BK_EVENTSTATUS_ERROR;
            AFont.Color  := COLOR_TX_EVENTSTATUS_ERROR;
          end;
          else
          begin
            if (FCueSheetTarget <> nil) and (Item^.GroupNo = FCueSheetTarget^.GroupNo) then
            begin
              ABrush.Color := COLOR_BK_EVENTSTATUS_TARGET;
              AFont.Color  := COLOR_TX_EVENTSTATUS_TARGET;
            end
            else
            begin
              ABrush.Color := COLOR_BK_EVENTSTATUS_NORMAL;
              AFont.Color  := COLOR_TX_EVENTSTATUS_NORMAL;
            end;
          end;
        end;
      end;
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistGetDisplText(Sender: TObject; ACol,
  ARow: Integer; var Value: string);
var
  Item: PCueSheetItem;
  RCol, RRow: Integer;
begin
  inherited;
exit;
  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < RowCount - 1) then
    begin
      Item := GetCueSheetItemByIndex(RRow - CNT_CUESHEET_HEADER);
      if (Item <> nil) then
      begin
        with Item^ do
        begin
          if (RCol = IDX_COL_CUESHEET_GROUP) then
          begin
            if (EventMode = EM_MAIN) then
            begin
              if (not IsNode(ARow)) then
                AddNode(ARow, ARow + 1)
            end
            else
              CellProperties[0, ARow].NodeLevel := 0; // Because of node show tree bug
          end;

          if (RCol = IDX_COL_CUESHEET_NO) then
          begin
            if (EventMode = EM_COMMENT) then
            begin
              Value := String(Title);
              if (not IsMergedCell(ACol, ARow)) then
              begin
                MergeCells(IDX_COL_CUESHEET_NO, ARow, ColCount - IDX_COL_CUESHEET_NO, 1);
                exit;
              end;
            end
            else
            begin
              if (IsMergedCell(ACol, ARow)) then
                SplitCells(ACol, ARow);

              if (EventMode = EM_MAIN) then
              begin
                Value := Format('%d', [FLastDisplayNo + 1]);
                Inc(FLastDisplayNo);
              end
              else
              begin
                Value := '';
              end;
            end;
          end
          else if (RCol = IDX_COL_CUESHEET_EVENT_MODE) then
          begin
            Value := EventModeShortNames[EventMode];
          end
          else if (RCol = IDX_COL_CUESHEET_START_MODE) then
          begin
            Value := StartModeNames[StartMode];
          end
          else if (RCol = IDX_COL_CUESHEET_START_DATE) then
          begin
            if (EventMode = EM_MAIN) and (EventStatus <> esSkipped) then
              Value := FormatDateTime(FORMAT_DATE, StartTime.D)
            else
              Value := '';
          end
          else if (RCol = IDX_COL_CUESHEET_START_TIME) then
          begin
            if (EventMode in [EM_MAIN, EM_SUB]) and (EventStatus <> esSkipped) then
              Value := TimecodeToString(StartTime.T)
            else
              Value := '';
          end
          else if (RCol = IDX_COL_CUESHEET_INPUT) then
          begin
            Value := InputTypeNames[Input];
          end
          else if (RCol = IDX_COL_CUESHEET_OUTPUT) then
          begin
            if (Input in [IT_MAIN, IT_BACKUP]) then
              Value := OutputBkgndTypeNames[TOutputBkgndType(Output)]
            else
              Value := OutputKeyerTypeNames[TOutputKeyerType(Output)];
          end
          else if (RCol = IDX_COL_CUESHEET_EVENT_STATUS) then
          begin
            Value := EventStatusNames[EventStatus];
          end
          else if (RCol = IDX_COL_CUESHEET_TITLE) then
          begin
            Value := String(Title);
          end
          else if (RCol = IDX_COL_CUESHEET_SUB_TITLE) then
          begin
            Value := String(SubTitle);
          end
          else if (RCol = IDX_COL_CUESHEET_SOURCE) then
          begin
            Value := String(Source);
          end
          else if (RCol = IDX_COL_CUESHEET_MEDIA_ID) then
          begin
            Value := String(MediaId);
          end
          else if (RCol = IDX_COL_CUESHEET_DURATON) then
          begin
            if (EventMode <> EM_JOIN) then
              Value := TimecodeToString(DurationTC)
            else
              Value := '';
          end
          else if (RCol = IDX_COL_CUESHEET_IN_TC) then
          begin
            if (EventMode <> EM_JOIN) then
              Value := TimecodeToString(InTC)
            else
              Value := '';
          end
          else if (RCol = IDX_COL_CUESHEET_OUT_TC) then
          begin
            if (EventMode <> EM_JOIN) then
              Value := TimecodeToString(OutTC)
            else
              Value := '';
          end
          else if (RCol = IDX_COL_CUESHEET_TR_TYPE) then
          begin
            Value := TRTypeNames[TransitionType];
          end
          else if (RCol = IDX_COL_CUESHEET_TR_RATE) then
          begin
            Value := TRRateNames[TransitionRate];
          end
          else if (RCol = IDX_COL_CUESHEET_PROGRAM_TYPE) then
          begin
            Value := GetProgramTypeNameByCode(ProgramType);
          end
          else if (RCol = IDX_COL_CUESHEET_NOTES) then
          begin
            Value := String(Notes);
          end;
        end;
      end;
    end
    else
    begin
      Value := 'End of event';
      if (not IsMergedCell(ACol, ARow)) then
        MergeCells(ACol, RowCount - 1, ColCount - IDX_COL_CUESHEET_NO, 1);
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistGetEditText(Sender: TObject; ACol,
  ARow: Integer; var Value: string);
begin
  inherited;

//  if (not FIsCellEditing) then
//  begin
//    with (Sender as TAdvColumnGrid) do
//      FOldCellValue := Cells[ACol, ARow];
//    FIsCellEditing := True;
//  end;
end;

procedure TfrmChannel.acgPlaylistKeyPress(Sender: TObject; var Key: Char);
var
  CueSheetItem: PCueSheetItem;
  RCol, RRow: Integer;
  SaveInput: TInputType;
  OB: TOutputBkgndType;
  OK: TOutputKeyerType;
begin
  inherited;

  with acgPlaylist do
  begin
    if (Key = #27) then
    begin
      Modified := False;
      RCol := RealColIndex(Col);
      RRow := RealRowIndex(Row);

//      ShowMEssage(Format('%d, %d, %d, %d', [Col, Row, RCol, RRow]));

{      if (FIsCellEditing) then
      begin
        FIsCellEditing  := False;
        FIsCellModified := False;
        Cells[Col, Row] := FOldCellValue;
      end;  }

      if (FIsCellModified) then
      begin
        FIsCellModified := False;
//        PopulatePlayListGrid(RRow - CNT_CUESHEET_HEADER);
        Cells[Col, Row] := FOldCellValue;

//        ShowSelection := False;
      end;

      if (RRow < FixedRows) or (RCol < FixedCols) then exit;

      if (RCol = IDX_COL_CUESHEET_INPUT) then
      begin
        CueSheetItem := GetCueSheetItemByIndex(RRow - FixedRows);
        if (CueSheetItem <> nil) then
          with CueSheetItem^ do
          begin
            SaveInput := Input;

            with Columns[IDX_COL_CUESHEET_INPUT].ComboItems do
              Input := TInputType(Objects[IndexOf(Cells[Col, Row])]);

            if ((SaveInput in [IT_MAIN, IT_BACKUP]) and (not (Input in [IT_MAIN, IT_BACKUP]))) or
               (not (SaveInput in [IT_MAIN, IT_BACKUP]) and (Input in [IT_MAIN, IT_BACKUP])) then
            begin
              Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems.Clear;
              if (Input in [IT_MAIN, IT_BACKUP]) then
              begin
                for OB := OB_NONE to OB_BOTH do
                begin
                  Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems.AddObject(OutputBkgndTypeNames[OB], TObject(OB));
                end;
                Cells[IDX_COL_CUESHEET_OUTPUT, RRow] := OutputBkgndTypeNames[TOutputBkgndType(Output)];
              end
              else
              begin
                for OK := OK_NONE to OK_OFF do
                begin
                  Columns[IDX_COL_CUESHEET_OUTPUT].ComboItems.AddObject(OutputKeyerTypeNames[OK], TObject(OK));
                end;
                Cells[IDX_COL_CUESHEET_OUTPUT, RRow] := OutputKeyerTypeNames[TOutputKeyerType(Output)];
              end;
            end;
          end;
      end;
    end;
  end;
end;

procedure TfrmChannel.acgPlaylistSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
var
  Index: Integer;
  CItem: PCueSheetItem;
  RCol, RRow: Integer;
  SM: TStartMode;
  OB: TOutputBkgndType;
  OK: TOutputKeyerType;
begin
  inherited;

  with (Sender as TAdvColumnGrid) do
  begin
    RCol := RealColIndex(ACol);
    RRow := RealRowIndex(ARow);

    if (RRow < FixedRows) or (RCol < FixedCols) then exit;

    Index := RRow - FixedRows;

    CItem := GetCueSheetItemByIndex(Index);
    if (CItem <> nil) then
    begin
      if (RCol = IDX_COL_CUESHEET_START_MODE) then
      begin
        with CItem^ do
        begin
          Columns[ACol].ComboItems.Clear;
          if (EventMode = EM_MAIN) then
          begin
            for SM := SM_ABSOLUTE to SM_LOOP do
            begin
              Columns[ACol].ComboItems.AddObject(StartModeNames[SM], TObject(SM));
            end;
          end
          else if (EventMode = EM_JOIN) then
          begin
            Columns[ACol].ComboItems.AddObject(StartModeNames[SM_SUBBEGIN], TObject(SM_SUBBEGIN));
          end
          else if (EventMode = EM_SUB) then
          begin
            for SM := SM_SUBBEGIN to SM_SUBEND do
            begin
              Columns[ACol].ComboItems.AddObject(StartModeNames[SM], TObject(SM));
            end;
          end;
        end;
      end
      else if (RCol = IDX_COL_CUESHEET_OUTPUT) then
      begin
        with CItem^ do
        begin
          Columns[ACol].ComboItems.Clear;
          if (Input in [IT_MAIN, IT_BACKUP]) then
          begin
            for OB := OB_NONE to OB_BOTH do
            begin
              Columns[ACol].ComboItems.AddObject(OutputBkgndTypeNames[OB], TObject(OB));
            end;
          end
          else
          begin
            for OK := OK_NONE to OK_OFF do
            begin
              Columns[ACol].ComboItems.AddObject(OutputKeyerTypeNames[OK], TObject(OK));
            end;
          end;
        end;
      end;
    end;
  end;
end;

constructor TfrmChannel.Create(AOwner: TComponent; AChannelCueSheet: PChannelCueSheet; ACombine: Boolean; ALeft, ATop, AWidth, AHeight: Integer);
begin
  inherited Create(AOwner, ACombine, ALeft, ATop, AWidth, AHeight);

  FChannelCueSheet := AChannelCueSheet;
  FChannelID := AChannelCueSheet^.ChannelId;
end;

procedure TfrmChannel.Initialize;
begin
  FChannelOnAir := False;

  FLastDisplayNo := 0;
  FLastSerialNo  := 0;
  FLastGroupNo   := 0;

  FLastCount  := 0;

  FIsCellEditing  := False;
  FIsCellModified := False;

  FCueSheetCurr   := nil;
  FCueSheetNext   := nil;
  FCueSheetTarget := nil;

  lblPlayedTime.Caption           := IDLE_TIME;
  lblRemainingTime.Caption        := IDLE_TIME;
  lblNextStart.Caption            := IDLE_TIMECODE;
  lblNextDuration.Caption         := IDLE_TIMECODE;
  lblRemainingTargetTime.Caption  := IDLE_TIMECODE;

  wmibFreezeOnAir.Caption     := 'Freeze';
  wmibAssignNext.Caption      := 'Assign next';
  wmibTakeNext.Caption        := 'Take next';
  wmibIncrease1Second.Caption := '+1';
  wmibDecrease1Second.Caption := '-1';

  SetChannelOnAir(FChannelOnAir);

  InitializePlayListGrid;
  InitializePlayListTimeLine;
end;

procedure TfrmChannel.Finalize;
begin
  if (FTimerThread <> nil) then
  begin
    FTimerThread.Terminate;
    FTimerThread.WaitFor;
    FreeAndNil(FTimerThread);
  end;

  ClearPlayListTimeLine;
  ClearPlayListGrid;
  ClearCueSheetList;

  // If exist hide rows bug
  with acgPlaylist do
  begin
    acgPlaylist.RowCount := 0;
  end;
end;

function TfrmChannel.GetCueSheetItemByIndex(AIndex: Integer): PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  Result := FChannelCueSheet^.CueSheetList[AIndex];
end;

function TfrmChannel.GetCueSheetItemByID(AEventID: TEventID): PCueSheetItem;
var
  I, CurrIndex: Integer;
  CurrItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := 0 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CurrItem := FChannelCueSheet^.CueSheetList[I];
    if (CurrItem <> nil) then
    begin
      if (IsEqualEventID(CurrItem^.EventID, AEventID)) then
      begin
        Result := CurrItem;
        break;
      end;
    end;
  end;
end;

function TfrmChannel.GetCueSheetIndexByItem(AItem: PCueSheetItem): Integer;
begin
  Result := -1;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  Result := FChannelCueSheet^.CueSheetList.IndexOf(AItem);
end;

function TfrmChannel.GetSelectCueSheetItem: PCueSheetItem;
var
  R: Integer;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  with acgPlaylist do
  begin
    R := RealRowIndex(Row);
    Result := GetCueSheetItemByIndex(R - CNT_CUESHEET_HEADER);
  end;
end;

function TfrmChannel.GetSelectCueSheetIndex: Integer;
begin
  Result := -1;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  with acgPlaylist do
  begin
    Result := RealRowIndex(Row) - CNT_CUESHEET_HEADER;
  end;

  if (Result >= FChannelCueSheet^.CueSheetList.Count) then
    Result := -1;
end;

function TfrmChannel.GetParentCueSheetItemByIndex(AIndex: Integer): PCueSheetItem;
var
  I, CurrIndex: Integer;
  CurrItem, ParentItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  CurrItem := FChannelCueSheet^.CueSheetList[AIndex];

  for I := AIndex downto 0 do
  begin
    ParentItem := FChannelCueSheet^.CueSheetList[I];
    if (ParentItem^.GroupNo = CurrItem^.GroupNo) and (ParentItem^.EventMode = EM_MAIN) then
    begin
      Result := ParentItem;
      break;
    end;
  end;
end;

function TfrmChannel.GetParentCueSheetItemByItem(AItem: PCueSheetItem): PCueSheetItem;
var
  I, CurrIndex: Integer;
  ParentItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AItem = nil) then exit;

  CurrIndex := FChannelCueSheet^.CueSheetList.IndexOf(AItem);
  if (CurrIndex < 0) then exit;

  for I := CurrIndex downto 0 do
  begin
    ParentItem := FChannelCueSheet^.CueSheetList[I];
    if (ParentItem^.EventMode = EM_COMMENT) then continue;

    if (ParentItem^.GroupNo = AItem^.GroupNo) and (ParentItem^.EventMode = EM_MAIN) then
    begin
      Result := ParentItem;
      break;
    end;
//    else if (P^.GroupNo > C^.GroupNo) then break;
  end;
end;

function TfrmChannel.GetChildCueSheetIndexByItem(AItem: PCueSheetItem; AIncludeComment: Boolean): Integer;
var
  I: Integer;
  ParentItem: PCueSheetItem;
  ParentIndex: Integer;
  CurrItem: PCueSheetItem;
  CurrIndex: Integer;
begin
  Result := -1;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AItem = nil) then exit;
  if (not AIncludeComment) and (AItem^.EventMode = EM_COMMENT) then exit;

  if (AItem^.EventMode = EM_MAIN) then
  begin
    Result := 0;
    exit;
  end;

  ParentItem  := GetParentCueSheetItemByItem(AItem);
  if (ParentItem <> nil) then
  begin
    ParentIndex := GetCueSheetIndexByItem(ParentItem);
    CurrIndex := GetCueSheetIndexByItem(AItem);

    if (AIncludeComment) then
      Result := CurrIndex - ParentIndex
    else
    begin
      Result := 0;

      for I := ParentIndex + 1 to CurrIndex do
      begin
//        ShowMessage(IntToStr(ParentIndex));
//        ShowMessage(IntToStr(I));
        CurrItem := FChannelCueSheet^.CueSheetList[I];
        if (CurrItem^.GroupNo = ParentItem^.GroupNo) then
        begin
          if (CurrItem^.EventMode <> EM_COMMENT) then
            Inc(Result);
        end
        else
          break;
      end;
    end;
  end;
end;

function TfrmChannel.GetLastChildCueSheetItemByIndex(AIndex: Integer): PCueSheetItem;
var
  I: Integer;
  PItem, CItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  PItem := FChannelCueSheet^.CueSheetList[AIndex];
  if (PItem <> nil) then
  begin
    for I := AIndex + 1 to FChannelCueSheet^.CueSheetList.Count - 1 do
    begin
      CItem := FChannelCueSheet^.CueSheetList[I];
      if (CItem^.EventMode = EM_COMMENT) then continue;

      if (CItem^.GroupNo <> PItem^.GroupNo) then
      begin
        if (I > 0) then
        begin
          Result := FChannelCueSheet^.CueSheetList[I - 1];
        end;
        break;
      end;
    end;
  end;
end;

function TfrmChannel.GetLastChildCueSheetItemByItem(AItem: PCueSheetItem): PCueSheetItem;
var
  I, CurrIndex: Integer;
  CItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AItem = nil) then exit;

  CurrIndex := FChannelCueSheet^.CueSheetList.IndexOf(AItem);
  if (CurrIndex < 0) then exit;

  for I := CurrIndex + 1 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventMode = EM_COMMENT) then continue;

    if (CItem^.GroupNo <> AItem^.GroupNo) then
    begin
      if (I > 0) then
      begin
        Result := FChannelCueSheet^.CueSheetList[I - 1];
      end;
      break;
    end;
  end;
end;

function TfrmChannel.GetFirstMainItem: PCueSheetItem;
var
  I: Integer;
  PItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := 0 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) then
    begin
      Result := PItem;
      break;
    end;
  end;
end;

function TfrmChannel.GetFirstMainIndex: Integer;
var
  I: Integer;
  PItem: PCueSheetItem;
begin
  Result := -1;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := 0 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) then
    begin
      Result := I;
      break;
    end;
  end;
end;

function TfrmChannel.GetLastMainItem: PCueSheetItem;
var
  I: Integer;
  PItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := FChannelCueSheet^.CueSheetList.Count - 1 downto 0 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) then
    begin
      Result := PItem;
      break;
    end;
  end;
end;

function TfrmChannel.GetLastMainIndex: Integer;
var
  I: Integer;
  PItem: PCueSheetItem;
begin
  Result := -1;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := FChannelCueSheet^.CueSheetList.Count - 1 downto 0 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) then
    begin
      Result := I;
      break;
    end;
  end;
end;

function TfrmChannel.GetBeforeMainItemByItem(AItem: PCueSheetItem): PCueSheetItem;
var
  I, CurrIndex: Integer;
  PItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AItem = nil) then exit;

  CurrIndex := FChannelCueSheet^.CueSheetList.IndexOf(AItem);
  if (CurrIndex < 0) then exit;

  for I := CurrIndex - 1 downto 0 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) and (PItem^.EventStatus <> esSkipped) then
    begin
      Result := PItem;
      break;
    end;
  end;
end;

function TfrmChannel.GetBeforeMainItemByIndex(AIndex: Integer): PCueSheetItem;
var
  I: Integer;
  PItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  for I := AIndex - 1 downto 0 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) and (PItem^.EventStatus <> esSkipped) then
    begin
      Result := PItem;
      break;
    end;
  end;
end;

function TfrmChannel.GetNextMainItemByItem(AItem: PCueSheetItem): PCueSheetItem;
var
  I, CurrIndex: Integer;
  PItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AItem = nil) then exit;

  CurrIndex := FChannelCueSheet^.CueSheetList.IndexOf(AItem);
  if (CurrIndex < 0) then exit;

  for I := CurrIndex + 1 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) and (PItem^.EventStatus <> esSkipped) then
    begin
      Result := PItem;
      break;
    end;
  end;
end;

function TfrmChannel.GetNextMainItemByIndex(AIndex: Integer): PCueSheetItem;
var
  I: Integer;
  PItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  for I := AIndex + 1 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) and (PItem^.EventStatus <> esSkipped) then
    begin
      Result := PItem;
      break;
    end;
  end;
end;

function TfrmChannel.GetNextMainIndexByItem(AItem: PCueSheetItem): Integer;
var
  PItem: PCueSheetItem;
begin
  Result := -1;

  PItem := GetNextMainItemByItem(AItem);
  if (PItem = nil) then exit;

  Result := GetCueSheetIndexByItem(PItem);
end;

function TfrmChannel.GetNextMainIndexByIndex(AIndex: Integer): Integer;
var
  PItem: PCueSheetItem;
begin
  Result := -1;

  PItem := GetNextMainItemByIndex(AIndex);
  if (PItem = nil) then exit;

  Result := GetCueSheetIndexByItem(PItem);
end;

function TfrmChannel.GetNextLoadedMainItemByItem(AItem: PCueSheetItem): PCueSheetItem;
var
  I, CurrIndex: Integer;
  PItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AItem = nil) then exit;

  CurrIndex := FChannelCueSheet^.CueSheetList.IndexOf(AItem);
  if (CurrIndex < 0) then exit;

  for I := CurrIndex + 1 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    if (PItem^.EventMode = EM_MAIN) and (PItem^.EventStatus = esLoaded) then
    begin
      Result := PItem;
      break;
    end;
  end;
end;

function TfrmChannel.GetBeforeMainCountByIndex(AIndex: Integer): Integer;
var
  I: Integer;
  CItem: PCueSheetItem;
begin
  Result := 0;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  Dec(AIndex);
  for I := AIndex downto 0 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventMode = EM_MAIN) then
    begin
      Inc(Result);
    end;
  end;
end;

function TfrmChannel.GetMainItemByStartTime(AIndex: Integer; ADateTime: TDateTime): PCueSheetItem;
var
  I: Integer;
  PItem: PCueSheetItem;
  StartTime: TDateTime;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    PItem := FChannelCueSheet^.CueSheetList[I];
    StartTime := EventTimeToDateTime(PItem^.StartTime);
    if (StartTime >= (ADateTime)) then
    begin
      Result := PItem;
      break;
    end;
  end;
end;

// 큐시트 중 DCS에 Input되지 않은 첫번째 메인 이벤트 인덱스를 구함
function TfrmChannel.GetStartOnAirMainIndex: Integer;
var
  I: Integer;
  CItem: PCueSheetItem;
begin
  Result := -1;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := 0 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventMode = EM_MAIN) and
       (CItem^.EventStatus <> esSkipped) and (CItem^.EventStatus <= esCued) then
    begin
      Result := I;
      break;
    end;
  end;
end;

// 큐시트 중 DCS에 Input되지 않은 첫번째 메인 이벤트를 구함
function TfrmChannel.GetStartOnAirMainItem: PCueSheetItem;
var
  I: Integer;
  CItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := 0 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventMode = EM_MAIN) and (CItem^.EventStatus in [esIdle, esLoaded, esCued]) then
    begin
      Result := CItem;
      break;
    end;
  end;
end;

// 큐시트 중 DCS에 다음 Input될 메인 이벤트를 구함
function TfrmChannel.GetNextOnAirMainIndex: Integer;
var
  I: Integer;
  CItem: PCueSheetItem;
begin
  Result := -1;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := FChannelCueSheet^.CueSheetList.Count - 1 downto 0 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventMode = EM_MAIN) and (CItem^.EventStatus in [esIdle, esLoaded, esCued]) then
    begin
      Result := I;
      break;
    end;
  end;
end;

// 큐시트 중 DCS에 다음 Input될 메인 이벤트를 구함
function TfrmChannel.GetNextOnAirMainItem: PCueSheetItem;
var
  I: Integer;
  CItem: PCueSheetItem;
begin
  Result := nil;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := FChannelCueSheet^.CueSheetList.Count - 1 downto 0 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventMode = EM_MAIN) and (CItem^.EventStatus in [esIdle, esLoaded, esCued]) then
    begin
      Result := CItem;
      break;
    end;
  end;
end;

// 선택된 이벤트의 삭제 목록 및 개수를 구함
function TfrmChannel.GetDeleteCueSheetList(ADeleteList: TCueSheetList): Integer;
var
  I, J: Integer;
  Index: Integer;
  SelectItem, Item: PCueSheetItem;
  SortSelectedRows: TList<Integer>;
begin
  Result := 0;

  if (ADeleteList = nil) then exit;

  with acgPlaylist do
  begin
    if (SelectedRowCount <= 0) then exit;

    SortSelectedRows := TList<Integer>.Create;
    try
      for I := 0 to SelectedRowCount - 1 do
        SortSelectedRows.Add(SelectedRow[I]);

      SortSelectedRows.Sort;
      
      ADeleteList.Clear;
      for I := 0 to SelectedRowCount - 1 do
      begin
        Index := RealRowIndex(SortSelectedRows[I]) - CNT_CUESHEET_HEADER;
        SelectItem := GetCueSheetItemByIndex(Index);
        if (SelectItem <> nil) then
        begin
          case SelectItem^.EventMode of
            EM_MAIN:
            begin
              // Cue or onair evnet not include
              if (SelectItem^.EventStatus in [esCueing..esFinished]) then Continue;

              // 종속된 이벤트 포함
              if (ADeleteList.IndexOf(SelectItem) < 0) then
                ADeleteList.Add(SelectItem);

              for J := Index + 1 to FChannelCueSheet^.CueSheetList.Count - 1 do
              begin
                Item := GetCueSheetItemByIndex(J);
                if (SelectItem^.GroupNo = Item^.GroupNo) and (ADeleteList.IndexOf(Item) < 0) then
                begin
                  ADeleteList.Add(Item);
                end
                else
                  break;
              end;
            end;
            EM_JOIN,
            EM_SUB,
            EM_COMMENT:
            begin
              // Cue or onair parent evnet not include
              Item := GetParentCueSheetItemByItem(SelectItem);
              if (Item <> nil) and (Item^.EventStatus in [esCueing..esFinished]) then Continue;

              // 자신의 이벤트만 포함
              if (ADeleteList.IndexOf(SelectItem) < 0) then
                ADeleteList.Add(SelectItem);
            end;
          end;
        end;
      end;
    finally
      FreeAndNil(SortSelectedRows);
    end;
  end;

  Result := ADeleteList.Count;
end;

function TfrmChannel.GetClipboardCueSheetList(AClipboardCueSheet: TClipboardCueSheet; APasteMode: TPasteMode): Integer;
var
  I, J: Integer;
  Index: Integer;
  SelectItem, Item: PCueSheetItem;
  SortSelectedRows: TList<Integer>;
begin
  Result := 0;

  if (AClipboardCueSheet = nil) then exit;

  with acgPlaylist do
  begin
    if (SelectedRowCount <= 0) then exit;

    AClipboardCueSheet.Clear;
    AClipboardCueSheet.PasteMode := APasteMode;
    AClipboardCueSheet.ChannelID := FChannelID;
    AClipboardCueSheet.PasteIncluded := [];

    SortSelectedRows := TList<Integer>.Create;
    try
      for I := 0 to SelectedRowCount - 1 do
        SortSelectedRows.Add(SelectedRow[I]);

      SortSelectedRows.Sort;
                          
      // 메인 이벤트가 포함되어 있는지 검사
      for I := 0 to SelectedRowCount - 1 do
      begin
        Index := RealRowIndex(SortSelectedRows[I]) - CNT_CUESHEET_HEADER;
        SelectItem := GetCueSheetItemByIndex(Index);
        if (SelectItem <> nil) and (SelectItem^.EventMode = EM_MAIN) and
           (SelectItem^.EventStatus in [esIdle..esLoaded, esDone..esSkipped]) then
        begin
          AClipboardCueSheet.PasteIncluded := AClipboardCueSheet.PasteIncluded + [EM_MAIN];
          break;
        end;
      end;

      for I := 0 to SelectedRowCount - 1 do
      begin
        Index := RealRowIndex(SortSelectedRows[I]) - CNT_CUESHEET_HEADER;
        SelectItem := GetCueSheetItemByIndex(Index);
        if (SelectItem <> nil) then
        begin
          case SelectItem^.EventMode of
            EM_MAIN:
            begin
              // Cue or onair evnet not include
  //            if (SelectItem^.EventStatus in [esCueing..esFinished]) then Continue;

              // 종속된 이벤트 포함
              if (AClipboardCueSheet.IndexOf(SelectItem) < 0) then
                AClipboardCueSheet.Add(SelectItem);

              for J := Index + 1 to FChannelCueSheet^.CueSheetList.Count - 1 do
              begin
                Item := GetCueSheetItemByIndex(J);
                if (SelectItem^.GroupNo = Item^.GroupNo) and (AClipboardCueSheet.IndexOf(Item) < 0) then
                begin
                  AClipboardCueSheet.Add(Item);
                  AClipboardCueSheet.PasteIncluded := AClipboardCueSheet.PasteIncluded + [Item^.EventMode];
                end
                else
                  break;
              end;
            end;
            EM_JOIN,
            EM_SUB,
            EM_COMMENT:
            begin
              // 메인 이벤트가 포함되어 있는 경우 부모이벤트 및 하위 이벤트 모두 포함
              if (EM_MAIN in AClipboardCueSheet.PasteIncluded) then
              begin
                Item := GetParentCueSheetItemByItem(SelectItem);
                // Cue or onair evnet not include
                if (Item <> nil) {and (Item^.EventStatus in [esIdle..esLoaded, esDone..esSkipped])} then
                begin
                  Index := GetCueSheetIndexByItem(Item);
                  for J := Index to FChannelCueSheet^.CueSheetList.Count - 1 do
                  begin
                    Item := GetCueSheetItemByIndex(J);
                    if (SelectItem^.GroupNo = Item^.GroupNo) and (AClipboardCueSheet.IndexOf(Item) < 0) then
                    begin
                      AClipboardCueSheet.Add(Item);
                      AClipboardCueSheet.PasteIncluded := AClipboardCueSheet.PasteIncluded + [Item^.EventMode];
                    end
                    else
                      break;
                  end;
                end
                else if (AClipboardCueSheet.IndexOf(SelectItem) < 0) then
                begin
                  AClipboardCueSheet.Add(SelectItem);
                  AClipboardCueSheet.PasteIncluded := AClipboardCueSheet.PasteIncluded + [SelectItem^.EventMode];
                end;
              end
              else
              begin
                // Cue or onair parent evnet not include
  //              Item := GetParentCueSheetItemByItem(SelectItem);
  //              if (Item <> nil) and (Item^.EventStatus in [esCueing..esFinished]) then Continue;

                // 자신의 이벤트만 포함
                if (AClipboardCueSheet.IndexOf(SelectItem) < 0) then
                begin
                  AClipboardCueSheet.Add(SelectItem);
                  AClipboardCueSheet.PasteIncluded := AClipboardCueSheet.PasteIncluded + [SelectItem^.EventMode];
                end;
              end;
            end;
          end;
        end;
      end;
    finally
      FreeAndNil(SortSelectedRows);
    end;
{    for I := 0 to AClipboardCueSheet.Count - 1 do
      ShowMessage(AClipboardCueSheet[I]^.SubTitle);  }
  end;

  Result := AClipboardCueSheet.Count;
end;

function TfrmChannel.IsValidStartDate(AItem: PCueSheetItem; AStartDate: TDate; AIsSelf: Boolean): Boolean;
var
  ErrorString: String;

  PItem: PCueSheetItem;   // Parent cuesheet item
  PEndTime: TEventTime;   // parent end time

  StartTime: TEventTime;  // Current start time
begin
  Result := True;

  ErrorString := '';
  try
    // Checks whether the entered start time is less than the start time of the previous event.
    PItem := GetBeforeMainItemByItem(AItem);
    if (PItem <> nil) then
    begin
      // Get parent end time
      PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);

      if (AItem^.EventMode = EM_MAIN) then
      begin
        // Get current start time
        StartTime   := AItem^.StartTime;
        StartTime.D := AStartDate;

        if (CompareEventTime(StartTime, PEndTime) < 0) then
        begin
          ErrorString := SStartTimeGreaterThenBeforeEndTime;
          exit;
        end;
      end;
    end;
  finally
    if (ErrorString <> '') then
    begin
      if (AIsSelf) then
        TAdvOfficePage(Parent).AdvOfficePager.ActivePage := TAdvOfficePage(Parent);

      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_OK or MB_ICONWARNING);
      Result := False;
    end;
  end;
end;

function TfrmChannel.IsValidStartTime(AItem: PCueSheetItem; AStartTC: TTimecode; AIsSelf: Boolean): Boolean;
var
  ErrorString: String;

  PItem: PCueSheetItem;   // Parent cuesheet item
  PEndTime: TEventTime;   // parent end time

  StartTime: TEventTime;  // Current start time
  EndTime: TEventTime;    // Current end time
begin
  Result := True;

  ErrorString := '';
  try
    // Check that the entered timecode is validate.
    if (not IsValidTimecode(AStartTC)) then
    begin
      ErrorString := Format(SInvalidTimeocde, ['Start time']);
      exit;
    end;

    // Checks whether the entered start time is less than the start time of the previous event.
    PItem := GetBeforeMainItemByItem(AItem);
    if (PItem <> nil) then
    begin
      // Get parent end time
      PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);

      if (AItem^.EventMode = EM_MAIN) then
      begin
        // Get current start time
        StartTime   := AItem^.StartTime;
        StartTime.T := AStartTC;

        if (CompareEventTime(StartTime, PEndTime) < 0) then
        begin
          ErrorString := SStartTimeGreaterThenBeforeEndTime;
          exit;
        end;
      end
      else if (AItem^.EventMode in [EM_JOIN, EM_SUB]) then
      begin
        // Get current start time
        StartTime.D := 0;
        StartTime.T := AStartTC;

        if (AItem^.StartMode = SM_SUBBEGIN) then
        begin
          // Get current end time
          EndTime := GetEventEndTime(GetPlusEventTime(PItem^.StartTime, StartTime), AItem^.DurationTC);
          if (CompareEventTime(EndTime, PEndTime) > 0) then
          begin
            ErrorString := SSubStartTimeLessThenParentEndTime;
            exit;
          end;
        end
        else
          // Get current end time
          EndTime := GetEventEndTime(GetMinusEventTime(PEndTime, StartTime), AItem^.DurationTC);

          if (CompareEventTime(EndTime, PEndTime) > 0) then
          begin
            ErrorString := SSubStartTimeLessThenParentEndTime;
            exit;
          end
          else if (AStartTC > PItem^.DurationTC) then
          begin
            ErrorString := SSubStartTimeGreaterThenParentStartTime;
            exit;
          end;
      end;
    end;
  finally
    if (ErrorString <> '') then
    begin
      if (AIsSelf) then
        TAdvOfficePage(Parent).AdvOfficePager.ActivePage := TAdvOfficePage(Parent);

      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_OK or MB_ICONWARNING);
      Result := False;
    end;
  end;
end;

function TfrmChannel.IsValidDuration(AItem: PCueSheetItem; ADuration: TTimecode; AIsSelf: Boolean): Boolean;
var
  ErrorString: String;

  PItem: PCueSheetItem; // Parent cuesheet item
  PEndTime: TEventTime; // parent end time

  EndTime: TEventTime;  // Current end time
begin
  Result := True;

  ErrorString := '';
  try
    // Check that the entered timecode is validate.
    if (not IsValidTimecode(ADuration)) then
    begin
      ErrorString := Format(SInvalidTimeocde, ['Duration']);
      exit;
    end;

    if (ADuration < GV_SettingTresholdTime.MinDuration) then
    begin
      ErrorString := Format(SDurationTCGreaterThenMinDuration, ['Duration', TimecodeToString(GV_SettingTresholdTime.MinDuration)]);
      exit;
    end;

    if (ADuration > GV_SettingTresholdTime.MaxDuration) then
    begin
      ErrorString := Format(SDurationTCLessThenMaxDuration, ['Duration', TimecodeToString(GV_SettingTresholdTime.MaxDuration)]);
      exit;
    end;

    // Feture add the media duration validate
    //
    //


    if (AItem <> nil) then
    begin
      if (AItem^.EventMode in [EM_JOIN, EM_SUB]) then
      begin
        // Checks whether the entered end time is less than the end time of the parent event.
        PItem := GetParentCueSheetItemByItem(AItem);
        if (PItem <> nil) then
        begin
          // Get parent end time
          PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);

          if (AItem^.StartMode = SM_SUBBEGIN) then
          begin
            // Get current end time
            EndTime := GetEventTimeSubBegin(PItem^.StartTime, AItem^.StartTime.T);
            EndTime := GetEventEndTime(EndTime, ADuration);

            if (CompareEventTime(EndTime, PEndTime) > 0) then
            begin
              ErrorString := SSubStartTimeLessThenParentEndTime;
              exit;
            end;
          end
          else
          begin
            // Get current end time
            EndTime := GetEventTimeSubEnd(PItem^.StartTime, PItem^.DurationTC, AItem^.StartTime.T);
            EndTime := GetEventEndTime(EndTime, ADuration);

            if (CompareEventTime(EndTime, PEndTime) > 0) then
            begin
              ErrorString := SSubStartTimeLessThenParentEndTime;
              exit;
            end;
          end;
        end;
      end;
    end;
  finally
    if (ErrorString <> '') then
    begin
      if (AIsSelf) then
        TAdvOfficePage(Parent).AdvOfficePager.ActivePage := TAdvOfficePage(Parent);

      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_OK or MB_ICONWARNING);
      Result := False;
    end;
  end;
end;

function TfrmChannel.IsValidInTC(AItem: PCueSheetItem; AInTC: TTimecode; AIsSelf: Boolean): Boolean;
var
  ErrorString: String;
begin
  Result := True;

  ErrorString := '';
  try
    // Check that the entered timecode is validate.
    if (not IsValidTimecode(AInTC)) then
    begin
      ErrorString := Format(SInvalidTimeocde, ['In']);
      exit;
    end;

    if (AItem <> nil) then
    begin
      if (AInTC >= AItem^.DurationTC) then
      begin
        ErrorString := SInTCLessThenDurationTC;
        exit;
      end;

      if (AInTC > AItem^.OutTC) then
      begin
        ErrorString := SInTCLessThenOutTC;
        exit;
      end;
    end;
  finally
    if (ErrorString <> '') then
    begin
      if (AIsSelf) then
        TAdvOfficePage(Parent).AdvOfficePager.ActivePage := TAdvOfficePage(Parent);

      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_OK or MB_ICONWARNING);
      Result := False;
    end;
  end;
end;

function TfrmChannel.IsValidOutTC(AItem: PCueSheetItem; AOutTC: TTimecode; AIsSelf: Boolean): Boolean;
var
  ErrorString: String;
begin
  Result := True;

  ErrorString := '';
  try
    // Check that the entered timecode is validate.
    if (not IsValidTimecode(AOutTC)) then
    begin
      ErrorString := Format(SInvalidTimeocde, ['Out']);
      exit;
    end;

    if (AItem <> nil) then
    begin
      if (AOutTC >= AItem^.DurationTC) then
      begin
        ErrorString := SOutTCLessThenDurationTC;
        exit;
      end;

      if (AOutTC < AItem^.InTC) then
      begin
        ErrorString := SOutTCGreaterThenInTC;
        exit;
      end;
    end;
  finally
    if (ErrorString <> '') then
    begin
      if (AIsSelf) then
        TAdvOfficePage(Parent).AdvOfficePager.ActivePage := TAdvOfficePage(Parent);

      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_OK or MB_ICONWARNING);
      Result := False;
    end;
  end;
end;

function TfrmChannel.IsValidMediaId(AMediaId: String; AIsSelf: Boolean): Boolean;
var
  ErrorString: String;
begin
  Result := True;

  ErrorString := '';
  try
  finally
    if (ErrorString <> '') then
    begin
      if (AIsSelf) then
        TAdvOfficePage(Parent).AdvOfficePager.ActivePage := TAdvOfficePage(Parent);

      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_OK or MB_ICONWARNING);
      Result := False;
    end;
  end;
end;

procedure TfrmChannel.NewPlayList;
begin
//  if (FChannelCueSheet = nil) then exit;
//  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  if (GV_ClipboardCueSheet.Count > 0) and (GV_ClipboardCueSheet.ChannelID = FChannelID) then
    ClearClipboardCueSheet;

  ClearPlayListTimeLine;
  ClearPlayListGrid;
  ClearCueSheetList;

  FLastDisplayNo := 0;
  FLastSerialNo  := 0;
  FLastGroupNo   := 0;

  FLastCount  := 0;

  FTimeLineDaysPerFrames := Round(SecsPerDay * wmtlPlaylist.TimeZoneProperty.FrameRate);
  FTimeLineMin := 0;
  FTimeLineMax := 0;

  StrPCopy(FChannelCueSheet^.OnairDate, FormatDateTime('YYYYMMDD', Date));
  FChannelCueSheet^.OnairFlag := FT_REGULAR;
  FChannelCueSheet^.OnairNo := 0;

  DisplayPlayListGrid(0);
  DisplayPlayListTimeLine(0);

  FPlayListFileName := NEW_CUESHEET_NAME;
  lblPlayListFileName.Caption := FPlayListFileName;

  FCueSheetCurr   := nil;
  FCueSheetNext   := nil;
  FCueSheetTarget := nil;

  acgPlaylist.Repaint;
end;

procedure TfrmChannel.OpenPlayList(AFileName: String);
var
  FirstItem: PCueSheetItem;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  NewPlayList;

  PlaylistFileParsing(AFileName);

  FirstItem := GetFirstMainItem;
  if (FirstItem <> nil) then
    FFirstStartTime := FirstItem^.StartTime;

  FPlayListFileName := AFileName;
  lblPlayListFileName.Caption := FPlayListFileName;

  DisplayPlayListGrid(0, FChannelCueSheet^.CueSheetList.Count);
  DisplayPlayListTimeLine;

  FLastCount   := FChannelCueSheet^.CueSheetList.Count;
  FLastGroupNo := FChannelCueSheet^.CueSheetList[FLastCount - 1]^.GroupNo;
end;

procedure TfrmChannel.OpenAddPlayList(AFileName: String);
var
  LastItem: PCueSheetItem;
  LastEndTime: TEventTime;

  NextItem: PCueSheetItem;
  NextIndex: Integer;

  SaveStartTime: TEventTime;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  PlaylistFileParsing(AFileName);

  LastItem := GetParentCueSheetItemByIndex(FLastCount - 1);
  if (LastItem <> nil) then
  begin
    NextItem := GetNextMainItemByItem(LastItem);
    if (NextItem <> nil) then
    begin
      NextIndex := GetCueSheetIndexByItem(NextItem);
      LastEndTime := GetEventEndTime(LastItem^.StartTime, LastItem^.DurationTC);

      SaveStartTime := NextItem^.StartTime;
      NextItem^.StartTime := LastEndTime;

      ResetStartTimeByTime(NextIndex, SaveStartTime);
    end;
  end;

  DisplayPlayListGrid(FLastCount, FChannelCueSheet^.CueSheetList.Count - FLastCount);
  DisplayPlayListTimeLine(FLastCount);

  FLastCount   := FChannelCueSheet^.CueSheetList.Count;
  FLastGroupNo := FChannelCueSheet^.CueSheetList[FLastCount - 1]^.GroupNo;
end;

procedure TfrmChannel.SavePlayList;
begin
  if (FPlayListFileName <> '') and
     (FPlayListFileName <> NEW_CUESHEET_NAME) then
  begin

  end;
end;

procedure TfrmChannel.SaveAsPlayList(AFileName: String);
begin



end;

procedure TfrmChannel.CueSheetListSort(ACueSheetList: TCueSheetList);
begin
  if (ACueSheetList.Count > 1) then CueSheetListQuickSort(0, pred(ACueSheetList.Count), ACueSheetList);
end;

function TfrmChannel.CheckEditCueSheetPossibleByIndex(AIndex: Integer): Boolean;
var
  CheckIndex: Integer;
  ParentItem: PCueSheetItem;
  ParentIndex: Integer;
begin
  Result := True;

  if (FChannelOnAir) then
  begin
    if (FCueSheetNext <> nil) then
      CheckIndex := GetCueSheetIndexByItem(FCueSheetNext)
    else if (FCueSheetCurr <> nil) then
      CheckIndex := GetCueSheetIndexByItem(FCueSheetCurr)
    else
      CheckIndex := FChannelCueSheet^.CueSheetList.Count - 1;

    ParentItem := GetParentCueSheetItemByIndex(AIndex);
    if (ParentItem <> nil) then
    begin
      ParentIndex := GetCueSheetIndexByItem(ParentItem);
      if (ParentIndex <= CheckIndex) then
      begin
        Result := False;
        exit;
      end;
    end
    else if (AIndex >= 0) and (AIndex <= CheckIndex) then
    begin
      Result := False;
      exit;
    end;
  end;
end;

function TfrmChannel.CheckEditCueSheetPossibleByItem(AItem: PCueSheetItem): Boolean;
var
  CheckIndex: Integer;
  ParentItem: PCueSheetItem;
  ParentIndex: Integer;
  Index: Integer;
begin
  Result := True;

  if (FChannelOnAir) then
  begin
    if (FCueSheetNext <> nil) then
      CheckIndex := GetCueSheetIndexByItem(FCueSheetNext)
    else if (FCueSheetCurr <> nil) then
      CheckIndex := GetCueSheetIndexByItem(FCueSheetCurr)
    else
      CheckIndex := FChannelCueSheet^.CueSheetList.Count - 1;

    Index := GetCueSheetIndexByItem(AItem);

    ParentItem := GetParentCueSheetItemByItem(AItem);
    if (ParentItem <> nil) then
    begin
      ParentIndex := GetCueSheetIndexByItem(ParentItem);
      if (ParentIndex <= CheckIndex) then
      begin
        Result := False;
        exit;
      end;
    end
    else if (Index >= 0) and (Index <= CheckIndex) then
    begin
      Result := False;
      exit;
    end;
  end;
end;

procedure TfrmChannel.InsertCueSheet(AEventMode: TEventMode);
var
  SelectIndex: Integer;
  PItem: PCueSheetItem;
  DurTime: TEventTime;
  I: Integer;
begin
{  SItem := GetSelectCueSheetItem;
  if (SItem <> nil) then
  begin
    if (SItem^.EventMode in [EM_COMMENT]) then
      PItem := GetNextMainItemByItem(SItem)
    else
      PItem := GetParentCueSheetItemByItem(SItem);
  end
  else
    PItem := nil; }

  SelectIndex := GetSelectCueSheetIndex;
  if (not CheckEditCueSheetPossibleByIndex(SelectIndex)) then
  begin
    MessageBeep(MB_ICONWARNING);
    MessageBox(Handle, PChar(SENotPossibleEdit), PChar(Application.Title), MB_OK or MB_ICONWARNING);
    exit;
  end;

  case AEventMode of
    EM_MAIN:
    begin
      // 이전 메인 이벤트를 찾음
      if (SelectIndex < 0) then
        PItem := GetLastMainItem
      else
        PItem := GetBeforeMainItemByIndex(SelectIndex);
    end;
    EM_JOIN,
    EM_SUB:
    begin
      // 부모 이벤트를 찾음
      if (SelectIndex < 0) then
        PItem := GetLastMainItem
      else
        PItem := GetParentCueSheetItemByIndex(SelectIndex);

      if (PItem = nil) then
      begin
        MessageBeep(MB_ICONWARNING);
        MessageBox(Handle, PChar(SENotInsertJoinSubLocationIncorrect), PChar(Application.Title), MB_OK or MB_ICONWARNING);
        exit;
      end;
    end;
    EM_COMMENT:
    begin
      PItem := GetParentCueSheetItemByIndex(SelectIndex);
    end;
    else
      exit;
  end;

  frmEditEvent := TfrmEditEvent.Create(Self, SelectIndex, EM_INSERT, AEventMode, PItem);
  try
    frmEditEvent.ShowModal;
    if (frmEditEvent.ModalResult = mrOK) then
    begin
      case AEventMode of
        EM_MAIN: ExecuteInsertCueSheetMain(SelectIndex, frmEditEvent.AddItem);
        EM_JOIN: ExecuteInsertCueSheetJoin(SelectIndex, PItem, frmEditEvent.AddItem);
        EM_SUB: ExecuteInsertCueSheetSub(SelectIndex, PItem, frmEditEvent.AddItem);
        EM_COMMENT: ExecuteInsertCueSheetComment(SelectIndex, frmEditEvent.AddItem);
        else
          exit;
      end;
    end;
  finally
    FreeAndNil(frmEditEvent);
  end;
end;

procedure TfrmChannel.ExecuteInsertCueSheetMain(AIndex: Integer; AAddItem: PCueSheetItem);
begin
  if (AIndex < 0) then
  begin
    FChannelCueSheet.CueSheetList.Add(frmEditEvent.AddItem);
    AIndex := FChannelCueSheet.CueSheetList.Count - 1;
  end
  else
  begin
    FChannelCueSheet.CueSheetList.Insert(AIndex, frmEditEvent.AddItem);
//                FLastDisplayNo := acgPlaylist.Cells[IDX_COL_CUESHEET_NO,
    FLastDisplayNo := GetBeforeMainCountByIndex(AIndex);

    ResetStartTimePlus(AIndex + 1, TimecodeToEventTime(frmEditEvent.AddItem^.DurationTC));
  end;

//  ResetStartTime(AIndex);
//  ResetChildItems(AIndex);

  if (FChannelOnAir) then
    OnAirInputEvents(AIndex, GV_SettingOption.MaxInputEventCount);

  DisplayPlayListGrid(AIndex, 1);
  DisplayPlayListTimeLine(AIndex);

  FLastCount := FChannelCueSheet.CueSheetList.Count;
end;

procedure TfrmChannel.ExecuteInsertCueSheetJoin(AIndex: Integer; AParentItem, AAddItem: PCueSheetItem);
var
  I: Integer;
  Index: Integer;
  Item: PCueSheetItem;
begin
{  // Join 이벤트의 경우 부모 이벤트의 Join 이벤트의 마지막 위치로 삽입
  Index := FChannelCueSheet.CueSheetList.IndexOf(AParentItem);

  Inc(Index);
  for I := Index to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    Item := FChannelCueSheet^.CueSheetList[I];

    if (Item <> nil) then
    begin
      if (AParentItem^.GroupNo <> Item^.GroupNo) then break;
      if (Item^.EventMode in [EM_JOIN, EM_SUB]) then
      begin
        Inc(Index);
      end
      else break;
    end;
  end; }

  Item := GetCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    // Main 이벤트 위치에서 삽입하는 경우는 그 다음 위치에 삽입
    if (Item^.EventMode = EM_MAIN) then
      Index := AIndex + 1
    else if (Item^.EventMode = EM_JOIN) then
      // Join 이벤트 위치에서 삽입하는 경우는 그 위치에 삽입
      Index := AIndex
    else if (Item^.EventMode in [EM_SUB]) then
    begin
      // Sub 이벤트 위치에서 삽입하는 경우 첫 Sub 이벤트 위치에 삽입
      Index := FChannelCueSheet.CueSheetList.IndexOf(AParentItem);

      Inc(Index);
      for I := Index to FChannelCueSheet^.CueSheetList.Count - 1 do
      begin
        Item := FChannelCueSheet^.CueSheetList[I];

        if (Item <> nil) then
        begin
          if (AParentItem^.GroupNo <> Item^.GroupNo) then break;
          if (Item^.EventMode in [EM_SUB]) then
          begin
            break;
          end;
          Inc(Index);
        end;
      end;
    end
    else
    begin
      // Comment 이벤트 위치에서 삽입하는 경우 이전 Join 이벤트 다음 위치에 삽입
      Index := FChannelCueSheet.CueSheetList.IndexOf(AParentItem);

      Inc(Index);
      for I := Index to FChannelCueSheet^.CueSheetList.Count - 1 do
      begin
        Item := FChannelCueSheet^.CueSheetList[I];

        if (Item <> nil) then
        begin
          if (AParentItem^.GroupNo <> Item^.GroupNo) then break;
          if (Item^.EventMode in [EM_SUB]) then
          begin
            break;
          end;
          if (Index >= AIndex) and (Item^.EventMode in [EM_COMMENT]) then
          begin
            break;
          end;
          Inc(Index);
        end;
      end;
    end;
  end;

  FChannelCueSheet.CueSheetList.Insert(Index, frmEditEvent.AddItem);

  if (FChannelOnAir) then
    OnAirInputEvents(Index, GV_SettingOption.MaxInputEventCount);

  FLastDisplayNo := GetBeforeMainCountByIndex(Index);

  DisplayPlayListGrid(Index, 1);
  DisplayPlayListTimeLine(Index);

  FLastCount := FChannelCueSheet.CueSheetList.Count;
end;

procedure TfrmChannel.ExecuteInsertCueSheetSub(AIndex: Integer; AParentItem, AAddItem: PCueSheetItem);
var
  I: Integer;
  Index: Integer;
  Item: PCueSheetItem;
begin
{  // Sub 이벤트의 경우 다음 이벤트의 이전 위치로 삽입
  Index := FChannelCueSheet.CueSheetList.IndexOf(AParentItem);

  Inc(Index);
  for I := Index to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    Item := FChannelCueSheet^.CueSheetList[I];

    if (Item <> nil) then
    begin
      if (AParentItem^.GroupNo <> Item^.GroupNo) then break;
      if (Item^.EventMode in [EM_JOIN, EM_SUB, EM_COMMENT]) then
      begin
        Inc(Index);
      end
      else break;
    end;
  end; }

  Item := GetCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    // Main, Join 이벤트 위치에서 삽입하는 경우는 첫 Sub 이전 위치에 삽입
    if (Item^.EventMode in [EM_MAIN, EM_JOIN]) then
    begin
      // Sub 이벤트 위치에서 삽입하는 경우 첫 Sub 이벤트 위치에 삽입
      Index := FChannelCueSheet.CueSheetList.IndexOf(AParentItem);

      Inc(Index);
      for I := Index to FChannelCueSheet^.CueSheetList.Count - 1 do
      begin
        Item := FChannelCueSheet^.CueSheetList[I];

        if (Item <> nil) then
        begin
          if (AParentItem^.GroupNo <> Item^.GroupNo) then break;
          if (Item^.EventMode in [EM_SUB]) then
          begin
            break;
          end;
          Inc(Index);
        end;
      end;
    end
    else if (Item^.EventMode = EM_SUB) then
      // Sub 이벤트 위치에서 삽입하는 경우는 그 위치에 삽입
      Index := AIndex
    else
    begin
      // Comment 이벤트 위치에서 삽입하는 경우 첫 Sub 이벤트 보다 아래 위치면 다음 위치에 삽입
      Index := FChannelCueSheet.CueSheetList.IndexOf(AParentItem);

      Inc(Index);
      for I := Index to FChannelCueSheet^.CueSheetList.Count - 1 do
      begin
        Item := FChannelCueSheet^.CueSheetList[I];

        if (Item <> nil) then
        begin
          if (AParentItem^.GroupNo <> Item^.GroupNo) then break;
          if (Item^.EventMode in [EM_SUB]) then
          begin
            break;
          end;
          Inc(Index);
        end;
      end;

      if (Index <= AIndex) then
        Index := AIndex;
    end;
  end;

  FChannelCueSheet.CueSheetList.Insert(Index, frmEditEvent.AddItem);

  if (FChannelOnAir) then
    OnAirInputEvents(Index, GV_SettingOption.MaxInputEventCount);

  FLastDisplayNo := GetBeforeMainCountByIndex(Index);

  DisplayPlayListGrid(Index, 1);
  DisplayPlayListTimeLine(Index);

  FLastCount := FChannelCueSheet.CueSheetList.Count;
end;

procedure TfrmChannel.ExecuteInsertCueSheetComment(AIndex: Integer; AAddItem: PCueSheetItem);
begin
  // Comment 이벤트의 경우 현재 이벤트의 이전 위치에 삽입
  if (AIndex < 0) then
  begin
    FChannelCueSheet.CueSheetList.Add(frmEditEvent.AddItem);
    AIndex := FChannelCueSheet.CueSheetList.Count - 1;
  end
  else
  begin
    FChannelCueSheet.CueSheetList.Insert(AIndex, frmEditEvent.AddItem);

    FLastDisplayNo := GetBeforeMainCountByIndex(AIndex);
  end;

//  if (FChannelOnAir) then
//    OnAirInputEvents(AIndex, GV_SettingOption.MaxInputEventCount);

  DisplayPlayListGrid(AIndex, 1);
  DisplayPlayListTimeLine(AIndex);

  FLastCount := FChannelCueSheet.CueSheetList.Count;
end;

procedure TfrmChannel.UpdateCueSheet;
var
  SelectIndex: Integer;
  SelectItem: PCueSheetItem;
  PEndTime, CEndTime, DurTime: TEventTime;
  I: Integer;

  SaveStartTime: TEventTime;
  SaveDurationTC: TTimecode;
begin
  SelectIndex := GetSelectCueSheetIndex;
  if (SelectIndex < 0) then exit;

  if (not CheckEditCueSheetPossibleByIndex(SelectIndex)) then
  begin
    MessageBeep(MB_ICONWARNING);
    MessageBox(Handle, PChar(SENotPossibleEdit), PChar(Application.Title), MB_OK or MB_ICONWARNING);
    exit;
  end;

  SelectItem  := GetCueSheetItemByIndex(SelectIndex);
  if (SelectItem^.EventMode = EM_MAIN) then
    PEndTime := GetEventEndTime(SelectItem^.StartTime, SelectItem^.DurationTC);

  frmEditEvent := TfrmEditEvent.Create(Self, SelectIndex, EM_UPDATE, SelectItem);
  try
    SaveStartTime  := SelectItem^.StartTime;
    SaveDurationTC := SelectItem^.DurationTC;

    frmEditEvent.ShowModal;
    if (frmEditEvent.ModalResult = mrOK) then
    begin
      ExecuteUpdateCueSheet(SelectIndex, SelectItem^.EventMode, SaveStartTime, SaveDurationTC);
    end;
  finally
    FreeAndNil(frmEditEvent);
  end;
end;

procedure TfrmChannel.wmtlPlaylistTrackHintEvent(Sender: TObject; Track: TTrack;
  var HintStr: string);
var
  Item, ParentItem: PCueSheetItem;
  StartTime, EndTime: TEventTime;
begin
  inherited;
  if (Track.Data <> nil) then
  begin
    Item := Track.Data;

    case Item^.EventMode of
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
  end;
end;

procedure TfrmChannel.ExecuteUpdateCueSheet(AIndex: Integer; AEventMode: TEventMode; ASaveStartTime: TEventtime; ASaveDurationTC: TTimecode);
begin
  FLastDisplayNo := GetBeforeMainCountByIndex(AIndex);

  if (AEventMode = EM_MAIN) then
  begin
    ResetStartTimeByTime(AIndex, ASaveStartTime, ASaveDurationTC);
//    ResetChildItems(AIndex);
  end;

  if (FChannelOnAir) then
    OnAirInputEvents(AIndex, GV_SettingOption.MaxInputEventCount);

  DisplayPlayListGrid(AIndex);
  DisplayPlayListTimeLine(AIndex);

  FLastCount := FChannelCueSheet.CueSheetList.Count - 1;
end;

procedure TfrmChannel.DeleteCueSheet;
var
  DeleteCount: Integer;
  DeleteList: TCueSheetList;

  function GetMessageString: String;
  var
    I: Integer;
    Flag: Boolean;
    SItem: PCueSheetItem;
  begin
    Result := SQDeleteEvent + #13#10;

    with acgPlaylist do
    begin
      Flag := False;
      for I := 0 to DeleteList.Count - 1 do
      begin
        if (I < 5) or
           ((DeleteList.Count > 10) and (I >= DeleteList.Count - 5)) or
           (DeleteList.Count <= 10) then
        begin
          SItem := DeleteList[I];
          if (SItem <> nil) then
          begin
            case SItem^.EventMode of
              EM_COMMENT:
                Result := Result + #13#10 +
                          Format('%s %s %s', [AllCells[IDX_COL_CUESHEET_NO, GetCueSheetIndexByItem(SItem) + CNT_CUESHEET_HEADER],
                                              EventModeShortNames[SItem^.EventMode],
                                              String(SItem^.Title)]);
              else
                Result := Result + #13#10 +
                          Format('%s %s %s-%s', [AllCells[IDX_COL_CUESHEET_NO, GetCueSheetIndexByItem(SItem) + CNT_CUESHEET_HEADER],
                                                 EventModeShortNames[SItem^.EventMode],
                                                 String(SItem^.Title),
                                                 String(SItem^.SubTitle)]);
            end;
{            case DItem^.EventMode of
              EM_MAIN:
                Result := Result + #13#10 +
                          Format('%s %s-%s', [AllCells[IDX_COL_CUESHEET_NO, RRow],
                                              String(DItem^.Title),
                                              String(DItem^.SubTitle)]);
              else
                Result := Result + #13#10 +
                          Format('%s %s-%s', [EventModeShortNames[DItem^.EventMode],
                                              String(DItem^.Title),
                                              String(DItem^.SubTitle)]);
            end; }
          end;
        end;

        if (DeleteList.Count > 10) and (I > 4) and (not Flag) then
        begin
          Result := Result + #13#10#13#10 + '  ：  ' + #13#10;
          Flag := True;
        end;
      end;
    end;
  end;

begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList.Count = 0) then exit;

  DeleteList := TCueSheetList.Create;
  try
    DeleteCount := GetDeleteCueSheetList(DeleteList);
    if (DeleteList.Count = 0) then exit;

    MessageBeep(MB_ICONQUESTION);
    if (MessageBox(Handle, PChar(GetMessageString), PChar(Application.Title), MB_OKCANCEL or MB_ICONQUESTION) = ID_OK) then
    begin
      ExecuteDeleteCueSheet(DeleteList);
    end;
  finally
    DeleteList.Clear;
    FreeAndNil(DeleteList);
  end;
end;

procedure TfrmChannel.ExecuteDeleteCueSheet(ADeleteList: TCueSheetList);
var
  FirstIndex: Integer;
  I: Integer;

  ParentItem: PCueSheetItem;
  ParentStartDate: TDate;
  ParentEndTime: TEventTime;

  DurTime: TEventTime;

  Item: PCueSheetItem;
  Index: Integer;

  StartTime: TEventTime;
begin
  if (ADeleteList = nil) then exit;

  FirstIndex := GetCueSheetIndexByItem(ADeleteList.First);

  DurTime.D := 0;
  DurTime.T := 0;

  ParentItem := nil;
  ParentStartDate := 0;
  for I := FirstIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    Item := GetCueSheetItemByIndex(I);
    if (ADeleteList.IndexOf(Item) >= 0) then
    begin
      if (Item^.EventMode = EM_MAIN) and (Item^.EventStatus in [esIdle..esLoaded, esDone, esSkipped]) then
      begin
        DurTime := GetPlusEventTime(DurTime, TimecodeToEventTime(Item^.DurationTC));
      end;
    end
    else
    begin
      if (Item^.EventMode <> EM_COMMENT) then
      begin
        // If then main event then checks whether the current start time is less than the start time of the previous event.
        if (Item^.EventMode = EM_MAIN) then
        begin
          StartTime := GetMinusEventTime(Item^.StartTime, DurTime);
  //          ShowMessage(EventTimeToDateTimecodeStr(CStartTime));
{          if (ParentItem <> nil) then
          begin
            ParentEndTime := GetEventEndTime(ParentItem^.StartTime, ParentItem^.DurationTC);
            if (CompareEventTime(StartTime, ParentEndTime) < 0) then
            begin
              StartTime := ParentEndTime;
            end;
          end;  }
          ParentItem := Item;

          ParentStartDate := StartTime.D;

          Item^.StartTime := StartTime;
        end
        else
        begin
          Item^.StartTime.D := ParentStartDate;
        end;

{        if (Item^.EventMode = EM_MAIN) then
        begin
          acgPlaylist.AllCells[IDX_COL_CUESHEET_START_DATE, I + CNT_CUESHEET_HEADER] := FormatDateTime(FORMAT_DATE, NextItem^.StartTime.D);
        end;

        if (Item^.EventMode in [EM_MAIN, EM_SUB]) then
        begin
          acgPlaylist.AllCells[IDX_COL_CUESHEET_START_TIME, I + CNT_CUESHEET_HEADER] := TimecodeToString(NextItem^.StartTime.T);
        end; }
      end;
    end;
  end;

  for I := ADeleteList.Count - 1 downto 0 do
  begin
    Item := ADeleteList[I];
    if (Item <> nil) and (Item^.EventStatus in [esIdle..esLoaded, esDone, esSkipped]) then
    begin
      Index := GetCueSheetIndexByItem(Item);
      if (FChannelOnAir) then
      begin
        DeleteEvent(Item);
        if (FCueSheetNext = Item) then
          FCueSheetNext := nil;
      end;

      DeletePlayListTimeLine(Item);
      FChannelCueSheet^.CueSheetList.Remove(Item);

      Dispose(Item);
    end;
  end;

  if (FChannelOnAir) then
    OnAirInputEvents(FirstIndex, GV_SettingOption.MaxInputEventCount);

  FLastDisplayNo := GetBeforeMainCountByIndex(FirstIndex);

  DisplayPlayListGrid(FirstIndex, -ADeleteList.Count);
  DisplayPlayListTimeLine(FirstIndex);

  acgPlaylist.MouseActions.DisjunctRowSelect := False;
  acgPlaylist.ClearRowSelect;
  acgPlaylist.MouseActions.DisjunctRowSelect := True;
  acgPlaylist.SelectRows(FirstIndex + CNT_CUESHEET_HEADER, 1);
  acgPlaylist.Row := FirstIndex + CNT_CUESHEET_HEADER;

  FLastCount := FChannelCueSheet^.CueSheetList.Count;
end;

procedure TfrmChannel.CopyToClipboardCueSheet;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList.Count = 0) then exit;

  GetClipboardCueSheetList(GV_ClipboardCueSheet, pmCopy);
  try
    if (GV_ClipboardCueSheet.Count = 0) then exit;
  finally
    acgPlaylist.Repaint;
  end;
end;

procedure TfrmChannel.CutToClipboardCueSheet;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList.Count = 0) then exit;

  GetClipboardCueSheetList(GV_ClipboardCueSheet, pmCut);
  try
    if (GV_ClipboardCueSheet.Count = 0) then exit;
  finally
    acgPlaylist.Repaint;
  end;
end;

procedure TfrmChannel.PasteFromClipboardCueSheet;
var
  SelectIndex: Integer;
  SelectItem: PCueSheetItem;

  PasteList: TCueSheetList;

  FirstIndex: Integer;

  ParentIndex: Integer;
  ParentItem: PCueSheetItem;
  ParentStartDate: TDate;
  ParentEndTime: TEventTime;
  SaveGroupNo: Integer;

  ChannelForm: TfrmChannel;
  ChannelCueSheet: TChannelCueSheet;

  I, J: Integer;
  Item: PCueSheetItem;

  PasteItem: PCueSheetItem;
  PasteIndex: Integer;
  PasteGap: Integer;

  ResetIndex: Integer;
  InsertIndex: Integer;

  StartTime: TEventTime;

  PastePlus: Boolean;
  PasteDurTime: TEventTime;
  StartDurTime: TEventTime;

  function GetPasteDurationTime: TEventTime;
  var
    I: Integer;
    Item: PCueSheetItem;
    DurStartTime, DurEndTime: TEventTime;
  begin
    Result.D := 0;
    Result.T := 0;

    DurStartTime.D := 0;
    DurStartTime.T := 0;
    for I := 0 to PasteList.Count - 1 do
    begin
      Item := PasteList[I];
      if (Item^.EventMode = EM_MAIN) then
      begin
        DurStartTime := Item^.StartTime;
        break;
      end;
    end;

    DurEndTime.D := 0;
    DurEndTime.T := 0;
    for I := PasteList.Count - 1 downto 0 do
    begin
      Item := PasteList[I];
      if (Item^.EventMode = EM_MAIN) then
      begin
        DurEndTime := GetEventEndTime(Item^.StartTime, Item^.DurationTC);
        break;
      end;
    end;

    Result := GetDurEventTime(DurStartTime, DurEndTime);
  end;

  function GetStartDurationTime(AStartTime: TEventTime): TEventTime;
  var
    I: Integer;
    Item: PCueSheetItem;
    DurStartTime: TEventTime;
  begin
    Result.D := 0;
    Result.T := 0;

    DurStartTime.D := 0;
    DurStartTime.T := 0;
    for I := 0 to PasteList.Count - 1 do
    begin
      Item := PasteList[I];
      if (Item^.EventMode = EM_MAIN) then
      begin
        DurStartTime := Item^.StartTime;
        break;
      end;
    end;

    Result := GetDurEventTime(AStartTime, DurStartTime);
  end;

begin
  if (GV_ClipboardCueSheet.Count = 0) then exit;

  // TMS Grid Bug
  // Mouse wheel top before or end after selectedrowcount = 0
  if (acgPlayList.SelectedRowCount = 0) then
  begin
    acgPlayList.SelectRows(acgPlayList.Row, 1);
  end;

  if (acgPlayList.SelectedRowCount <> 1) then
  begin
    MessageBeep(MB_ICONWARNING);
    MessageBox(Handle, PChar(SEPasteLocationIncorrect), PChar(Application.Title), MB_OK or MB_ICONWARNING);
    exit;
  end;

  SelectIndex := GetSelectCueSheetIndex;
  if (not CheckEditCueSheetPossibleByIndex(SelectIndex)) then
  begin
    MessageBeep(MB_ICONWARNING);
    MessageBox(Handle, PChar(SENotPossibleEdit), PChar(Application.Title), MB_OK or MB_ICONWARNING);
    exit;
  end;

  if (SelectIndex >= 0) then
  begin
    SelectItem := GetCueSheetItemByIndex(SelectIndex);
    ParentItem := GetParentCueSheetItemByItem(SelectItem);

    if (EM_MAIN in GV_ClipboardCueSheet.PasteIncluded) then
    begin
      if (SelectItem <> nil) and (ParentItem <> nil) and (SelectItem <> ParentItem) then
      begin
        MessageBeep(MB_ICONWARNING);
        MessageBox(Handle, PChar(SEPasteHasMainLocationIncorrect), PChar(Application.Title), MB_OK or MB_ICONWARNING);
        exit;
      end;
    end
    else if (GV_ClipboardCueSheet.PasteIncluded <> [EM_COMMENT]) then
    begin
      if (ParentItem = nil) then
      begin
        MessageBeep(MB_ICONWARNING);
        MessageBox(Handle, PChar(SEPasteJoinSubLocationIncorrect), PChar(Application.Title), MB_OK or MB_ICONWARNING);
        exit;
      end;
    end;

    if (SelectItem^.EventMode = EM_COMMENT) then
    begin
      ParentItem := GetNextMainItemByItem(SelectItem);
      if (ParentItem <> nil) then
      begin
        StartTime := ParentItem^.StartTime;
      end
      else
      begin
        StartTime.D := OnAirDateToDate(FChannelCueSheet^.OnairDate);
        StartTime.T := 0;
      end;
    end
    else
      StartTime := SelectItem^.StartTime;
  end;

  if (SelectIndex < 0) and 
     ((not (EM_MAIN in GV_ClipboardCueSheet.PasteIncluded)) and
      (GV_ClipboardCueSheet.PasteIncluded <> [EM_COMMENT])) then
  begin
    MessageBeep(MB_ICONWARNING);
    MessageBox(Handle, PChar(SEPasteHasMainLocationIncorrect), PChar(Application.Title), MB_OK or MB_ICONWARNING);
    exit;
  end;

  PasteList := TCueSheetList.Create;
  try
//    DurTime.D := 0;
//    DurTime.T := 0;

    // Temporary paste item add
    for I := 0 to GV_ClipboardCueSheet.Count - 1 do
    begin
      Item := GV_ClipboardCueSheet[I];

      PasteItem := New(PCueSheetItem);
      Move(Item^, PasteItem^, SizeOf(TCueSheetItem));

      PasteItem^.EventStatus := esIdle;

//      if (PasteItem^.EventMode = EM_MAIN)then
//        DurTime := GetPlusEventTime(DurTime, TimecodeToEventTime(PasteItem^.DurationTC));

      PasteList.Add(PasteItem);
    end;

//    DurTime.D := 0;
//    DurTime.T := 0;

    PasteDurTime := GetPasteDurationTime;

    // If pastmode is cut then delete cut items
    if (GV_ClipboardCueSheet.PasteMode = pmCut) then
    begin
      ChannelForm := frmSEC.GetChannelFormByID(GV_ClipboardCueSheet.ChannelID);
      if (ChannelForm <> nil) then
      begin
        ChannelForm.ExecuteDeleteCueSheet(GV_ClipboardCueSheet.CueSheetList);

        if (ChannelForm = Self) then
        begin
          if (GV_ClipboardCueSheet.IndexOf(SelectItem) >= 0) then
          begin
            for I := SelectIndex to FChannelCueSheet.CueSheetList.Count - 1 do
            begin
              SelectItem := GetCueSheetItemByIndex(I);
              if (SelectItem <> nil) and (GV_ClipboardCueSheet.IndexOf(SelectItem) < 0) then
              begin
                SelectIndex := I;
                StartTime := SelectItem^.StartTime;
                break;
              end;
            end;
          end
          else
          begin
            SelectIndex := GetCueSheetIndexByItem(SelectItem);
            if (SelectIndex >= 0) then
            begin
              if (SelectItem^.EventMode = EM_COMMENT) then
              begin
                StartTime := PasteList.First^.StartTime;
              end
              else
                StartTime := SelectItem^.StartTime;
            end;
          end;
        end;
        GV_ClipboardCueSheet.Clear;
      end;
    end;

    if (SelectIndex < 0) then
    begin
      SelectItem := GetLastMainItem;
      SelectIndex := FChannelCueSheet^.CueSheetList.Count;
      if (SelectItem <> nil) then
      begin
        StartTime := GetEventEndTime(SelectItem^.StartTime, SelectItem^.DurationTC);
      end
      else
      begin
        StartTime.D := OnAirDateToDate(FChannelCueSheet.OnairDate);
        StartTime.T := 0;
      end;
    end;

//    ShowMessage(EventTimeToDateTimecodeStr(StartTime));
    StartDurTime := GetStartDurationTime(StartTime);
//    ShowMessage(EventTimeToDateTimecodeStr(StartDurTime));

    // Add paste items
//    FirstIndex := ChannelForm.GetCueSheetIndexByItem(GV_ClipboardCueSheet.First);

    ParentItem := nil;
    ParentStartDate := 0;

    PasteGap := 0;

    ResetIndex := SelectIndex;
    InsertIndex := SelectIndex;

    SaveGroupNo := -1;

    for I := 0 to PasteList.Count - 1 do
    begin
      PasteItem := PasteList[I];
      PasteIndex := SelectIndex + I;

      FLastSerialNo := FLastSerialNo + 1;

      PasteItem^.EventID.ChannelID := FChannelID;
      StrCopy(PasteItem^.EventID.OnAirDate, FChannelCueSheet.OnairDate);
      PasteItem^.EventID.OnAirFlag := FChannelCueSheet.OnairFlag;
      PasteItem^.EventID.OnAirNo   := FChannelCueSheet.OnairNo;
      PasteItem^.EventID.SerialNo  := FLastSerialNo;

      // If then main event then checks whether the current start time is less than the start time of the previous event.
      if (PasteItem^.EventMode = EM_MAIN) then
      begin
        SaveGroupNo := PasteItem^.GroupNo;
//          StartTime := GetPlusEventTime(PasteItem^.StartTime, DurTime);
//          ShowMessage(EventTimeToDateTimecodeStr(CStartTime));
{        if (ParentItem <> nil) then
        begin
          ParentEndTime := GetEventEndTime(ParentItem^.StartTime, ParentItem^.DurationTC);
          if (CompareEventTime(StartTime, ParentEndTime) < 0) then
          begin
            StartTime := ParentEndTime;
          end;
        end;}

        if (CompareEventTime(StartTime, PasteItem^.StartTime) >= 0) then
          StartTime := GetPlusEventTime(PasteItem^.StartTime, StartDurTime)
        else
          StartTime := GetMinusEventTime(PasteItem^.StartTime, StartDurTime);

        ParentItem := PasteItem;

        ParentStartDate := StartTime.D;

        FLastGroupNo := FLastGroupNo + 1;

        PasteItem^.StartTime := StartTime;
        PasteItem^.GroupNo   := FLastGroupNo;

        ResetIndex := PasteIndex + 1;
      end
      else if (PasteItem^.EventMode in [EM_JOIN, EM_SUB]) then
      begin
        if (EM_MAIN in GV_ClipboardCueSheet.PasteIncluded) then
        begin
          PasteItem^.StartTime.D := ParentStartDate;
          PasteItem^.GroupNo     := FLastGroupNo;
        end
        else
        begin
          ParentItem := GetParentCueSheetItemByItem(SelectItem);
          if (ParentItem <> nil) then
          begin
            ParentIndex := GetCueSheetIndexByItem(ParentItem);
            
            PasteItem^.StartTime.D := ParentItem^.StartTime.D;
            PasteItem^.GroupNo     := ParentItem^.GroupNo;

            if (PasteItem^.EventMode in [EM_JOIN]) then
            begin
              PasteItem^.DurationTC  := ParentItem^.DurationTC;

              // Main 이벤트 위치에서 삽입하는 경우는 그 다음 위치에 삽입
              if (SelectItem^.EventMode = EM_MAIN) then
              begin
                PasteIndex := ParentIndex + PasteGap + 1;
                Inc(PasteGap);
              end
              else if (SelectItem^.EventMode = EM_JOIN) then
              begin
                // Join 이벤트 위치에서 삽입하는 경우는 그 위치에 삽입
                PasteIndex := SelectIndex + PasteGap;
                Inc(PasteGap);
              end
              else if (Item^.EventMode in [EM_SUB]) then
              begin
                // Sub 이벤트 위치에서 삽입하는 경우 첫 Sub 이벤트 위치에 삽입
                PasteIndex := ParentIndex + 1;
                for J := PasteIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
                begin
                  Item := FChannelCueSheet^.CueSheetList[J];

                  if (Item <> nil) then
                  begin
                    if (ParentItem^.GroupNo <> Item^.GroupNo) then break;
                    if (Item^.EventMode in [EM_SUB]) then
                    begin
                      break;
                    end;
                    Inc(PasteIndex);
                  end;
                end;
              end
              else
              begin
                // Comment 이벤트 위치에서 삽입하는 경우 이전 Join 이벤트 다음 위치에 삽입
                PasteIndex := ParentIndex + 1;
                for J := PasteIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
                begin
                  Item := FChannelCueSheet^.CueSheetList[J];

                  if (Item <> nil) then
                  begin
                    if (ParentItem^.GroupNo <> Item^.GroupNo) then break;
                    if (Item^.EventMode in [EM_SUB]) then
                    begin
                      break;
                    end;
                    if (PasteIndex >= SelectIndex + PasteGap) and (Item^.EventMode in [EM_COMMENT]) then
                    begin
                      break;
                    end;
                    Inc(PasteIndex);
                  end;
                end;
              end;
            end
            else
            begin
              // Main, Join 이벤트 위치에서 삽입하는 경우는 첫 Sub 이전 위치에 삽입
              if (SelectItem^.EventMode in [EM_MAIN, EM_JOIN]) then
              begin
                // Sub 이벤트 위치에서 삽입하는 경우 첫 Sub 이벤트 위치에 삽입
                PasteIndex := ParentIndex + 1;
                for J := PasteIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
                begin
                  Item := FChannelCueSheet^.CueSheetList[J];

                  if (Item <> nil) then
                  begin
                    if (ParentItem^.GroupNo <> Item^.GroupNo) then break;
                    if (Item^.EventMode in [EM_SUB]) then
                    begin
                      break;
                    end;
                    Inc(PasteIndex);
                  end;
                end;
              end
              else if (SelectItem^.EventMode = EM_SUB) then
              begin
                // Sub 이벤트 위치에서 삽입하는 경우는 그 위치에 삽입
                PasteIndex := SelectIndex + PasteGap;
                Inc(PasteGap);
              end
              else
              begin
                // Comment 이벤트 위치에서 삽입하는 경우 첫 Sub 이벤트 보다 아래 위치면 다음 위치에 삽입
                PasteIndex := ParentIndex + 1;
                for J := PasteIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
                begin
                  Item := FChannelCueSheet^.CueSheetList[J];

                  if (Item <> nil) then
                  begin
                    if (ParentItem^.GroupNo <> Item^.GroupNo) then break;
                    if (Item^.EventMode in [EM_SUB]) then
                    begin
                      break;
                    end;
                    Inc(PasteIndex);
                  end;
                end;

                if (PasteIndex <= (SelectIndex + PasteGap)) then
                  PasteIndex := SelectIndex + PasteGap; 
              end;
            end;
          end;
        end;
      end
      else
      begin
        if (EM_MAIN in GV_ClipboardCueSheet.PasteIncluded) then
        begin
          if (PasteItem^.GroupNo <> SaveGroupNo) then
          begin
            FLastGroupNo       := FLastGroupNo + 1;
            PasteItem^.GroupNo := FLastGroupNo;
          end
          else
          begin
            PasteItem^.GroupNo := FLastGroupNo;
          end;
        end
        else
        begin
          if (PasteIndex < FChannelCueSheet^.CueSheetList.Count) then
            PasteItem^.GroupNo := SelectItem^.GroupNo
          else
          begin
            FLastGroupNo       := FLastGroupNo + 1;
            PasteItem^.GroupNo := FLastGroupNo;
          end;

            PasteIndex := SelectIndex + PasteGap;
            Inc(PasteGap);
{          // Main 이벤트 위치에서 삽입하는 경우는 그 위치 위에 삽입
          if (SelectItem^.EventMode = EM_MAIN) then
          begin
            PasteIndex := SelectIndex + PasteGap;
            Inc(PasteGap);
          end
          else 
          begin
            // Join, Sub, Comment 이벤트 위치에서 삽입하는 경우는 그 위치에 삽입
            PasteIndex := SelectIndex + PasteGap;
            Inc(PasteGap);
          end;}
        end;
      end;

      if (PasteIndex < FChannelCueSheet^.CueSheetList.Count) then
        FChannelCueSheet^.CueSheetList.Insert(PasteIndex, PasteItem)
      else
        FChannelCueSheet^.CueSheetList.Add(PasteItem);

      if (PasteIndex < InsertIndex)  then
        InsertIndex := PasteIndex;
    end;

    ParentItem := nil;
    ParentStartDate := 0;
    for I := ResetIndex {SelectIndex + PasteList.Count} to FChannelCueSheet^.CueSheetList.Count - 1 do
    begin
      Item := GetCueSheetItemByIndex(I);
      if (Item^.EventMode <> EM_COMMENT) then
      begin
        // If then main event then checks whether the current start time is less than the start time of the previous event.
        if (Item^.EventMode = EM_MAIN) then
        begin
          StartTime := GetPlusEventTime(Item^.StartTime, PasteDurTime);
  //          ShowMessage(EventTimeToDateTimecodeStr(CStartTime));
{          if (ParentItem <> nil) then
          begin
            ParentEndTime := GetEventEndTime(ParentItem^.StartTime, ParentItem^.DurationTC);
            if (CompareEventTime(StartTime, ParentEndTime) < 0) then
            begin
              StartTime := ParentEndTime;
            end;
          end; }
          ParentItem := Item;

          ParentStartDate := StartTime.D;

          Item^.StartTime := StartTime;
        end
        else
        begin
          PasteItem^.StartTime.D := ParentStartDate;
        end;
      end;
    end;

    if (FChannelOnAir) then
      OnAirInputEvents(SelectIndex, GV_SettingOption.MaxInputEventCount);

    FLastDisplayNo := GetBeforeMainCountByIndex(SelectIndex);

    DisplayPlayListGrid(SelectIndex, PasteList.Count);
    DisplayPlayListTimeLine(SelectIndex);

//  acgPlaylist.MouseActions.DisjunctRowSelect := False;
//  acgPlaylist.ClearRowSelect;
//  acgPlaylist.MouseActions.DisjunctRowSelect := True;
//  acgPlaylist.SelectRows(FirstIndex + CNT_CUESHEET_HEADER, IDX_COL_CUESHEET_NO);

    acgPlaylist.MouseActions.DisjunctRowSelect := False;
    acgPlaylist.ClearRowSelect;
    acgPlaylist.MouseActions.DisjunctRowSelect := True;
    acgPlaylist.SelectRows(SelectIndex + CNT_CUESHEET_HEADER, 1);
    acgPlaylist.Row := SelectIndex + CNT_CUESHEET_HEADER;

    FLastCount := FChannelCueSheet^.CueSheetList.Count;
  finally
    PasteList.Clear;
    FreeAndNil(PasteList);
  end;
end;

procedure TfrmChannel.ClearClipboardCueSheet;
begin
  if (GV_ClipboardCueSheet.Count > 0) then
  begin
    GV_ClipboardCueSheet.Clear;

    acgPlayList.Repaint;
  end;
end;

procedure TfrmChannel.PopulatePlayListGrid(AIndex: Integer);
var
  R: Integer;
  CueSheetItem: PCueSheetItem;
begin
  inherited;

  with acgPlaylist do
  begin
    R := AIndex + CNT_CUESHEET_HEADER;

    if (R < FixedRows) then exit;

    CueSheetItem := GetCueSheetItemByIndex(AIndex);
    if (CueSheetItem <> nil) then
    begin
      with CueSheetItem^ do
      begin
        if (EventMode = EM_COMMENT) then
        begin
          AllCells[IDX_COL_CUESHEET_NO, R] := String(Title);
          MergeCells(IDX_COL_CUESHEET_NO, DisplRowIndex(R), ColCount - IDX_COL_CUESHEET_NO, 1);
          exit;
        end
        else
        begin
          SplitCells(IDX_COL_CUESHEET_NO, DisplRowIndex(R));
          if (EventMode = EM_MAIN) then
          begin
            AllCells[IDX_COL_CUESHEET_NO, R]         := Format('%d', [FLastDisplayNo + 1]);
//            AllCells[IDX_COL_CUESHEET_NO, R]         := Format('%d', [GroupNo + 1]);
//            AllCells[IDX_COL_CUESHEET_START_DATE, R] := FormatDateTime(FORMAT_DATE, StartTime.D);
            Inc(FLastDisplayNo);
          end
          else
          begin
            AllCells[IDX_COL_CUESHEET_NO, R]         := '';
//            AllCells[IDX_COL_CUESHEET_START_DATE, R] := '';
          end;
        end;

        AllCells[IDX_COL_CUESHEET_EVENT_MODE, R]   := EventModeShortNames[EventMode];
        AllCells[IDX_COL_CUESHEET_START_MODE, R]   := StartModeNames[StartMode];

        if (EventMode = EM_MAIN) and (EventStatus <> esSkipped) then
          AllCells[IDX_COL_CUESHEET_START_DATE, R] := FormatDateTime(FORMAT_DATE, StartTime.D)
        else
          AllCells[IDX_COL_CUESHEET_START_DATE, R] := '';

        if (EventMode in [EM_MAIN, EM_SUB]) and (EventStatus <> esSkipped) then
          AllCells[IDX_COL_CUESHEET_START_TIME, R] := TimecodeToString(StartTime.T)
        else
          AllCells[IDX_COL_CUESHEET_START_TIME, R] := '';

        AllCells[IDX_COL_CUESHEET_INPUT, R]        := InputTypeNames[Input];

        if (Input in [IT_MAIN, IT_BACKUP]) then
          AllCells[IDX_COL_CUESHEET_OUTPUT, R]     := OutputBkgndTypeNames[TOutputBkgndType(Output)]
        else
          AllCells[IDX_COL_CUESHEET_OUTPUT, R]     := OutputKeyerTypeNames[TOutputKeyerType(Output)];

        AllCells[IDX_COL_CUESHEET_EVENT_STATUS, R]  := EventStatusNames[EventStatus];
//        AllCells[IDX_COL_CUESHEET_DEVICE_STATUS, R] := GetDeviceStatusName(DeviceStatus);

        AllCells[IDX_COL_CUESHEET_TITLE, R]        := String(Title);
        AllCells[IDX_COL_CUESHEET_SUB_TITLE, R]    := String(SubTitle);
        AllCells[IDX_COL_CUESHEET_SOURCE, R]       := String(Source);
        AllCells[IDX_COL_CUESHEET_MEDIA_ID, R]     := String(MediaId);

        if (EventMode <> EM_JOIN) then
        begin
          AllCells[IDX_COL_CUESHEET_DURATON, R]    := TimecodeToString(DurationTC);
          AllCells[IDX_COL_CUESHEET_IN_TC, R]      := TimecodeToString(InTC);
          AllCells[IDX_COL_CUESHEET_OUT_TC, R]     := TimecodeToString(OutTC);
        end
        else
        begin
          AllCells[IDX_COL_CUESHEET_DURATON, R]    := '';
          AllCells[IDX_COL_CUESHEET_IN_TC, R]      := '';
          AllCells[IDX_COL_CUESHEET_OUT_TC, R]     := '';
        end;

        AllCells[IDX_COL_CUESHEET_TR_TYPE, R]      := TRTypeNames[TransitionType];
        AllCells[IDX_COL_CUESHEET_TR_RATE, R]      := TRRateNames[TransitionRate];
        AllCells[IDX_COL_CUESHEET_PROGRAM_TYPE, R] := GetProgramTypeNameByCode(ProgramType);
        AllCells[IDX_COL_CUESHEET_NOTES, R]        := String(Notes);
      end;
    end;
  end;
end;

procedure TfrmChannel.PopulateEventStatusPlayListGrid(AIndex: Integer);
var
  R: Integer;
  CueSheetItem: PCueSheetItem;
begin
  inherited;

  with acgPlaylist do
  begin
    R := AIndex + CNT_CUESHEET_HEADER;

    if (R < FixedRows) then exit;

    CueSheetItem := GetCueSheetItemByIndex(AIndex);
    if (CueSheetItem <> nil) then
    begin
      with CueSheetItem^ do
      begin
        if (EventMode = EM_MAIN) and (EventStatus <> esSkipped) then
          AllCells[IDX_COL_CUESHEET_START_DATE, R] := FormatDateTime(FORMAT_DATE, StartTime.D)
        else
          AllCells[IDX_COL_CUESHEET_START_DATE, R] := '';

        if (EventMode in [EM_MAIN, EM_SUB]) and (EventStatus <> esSkipped) then
          AllCells[IDX_COL_CUESHEET_START_TIME, R] := TimecodeToString(StartTime.T)
        else
          AllCells[IDX_COL_CUESHEET_START_TIME, R] := '';

        AllCells[IDX_COL_CUESHEET_EVENT_STATUS, R]  := EventStatusNames[EventStatus];
//        AllCells[IDX_COL_CUESHEET_DEVICE_STATUS, R] := GetDeviceStatusName(DeviceStatus);
      end;

      RepaintRow(R);
    end;
  end;
end;

procedure TfrmChannel.DisplayPlayListGrid(AIndex: Integer; ACount: Integer);
var
  I, J: Integer;
  OldRowCount: Integer;
  CueSheetItem: PCueSheetItem;
  ParentItem: PCueSheetItem;
  DurTime: TEventTime;
begin
  inherited;
//exit;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  with acgPlaylist do
  begin
    OldRowCount := RowCount;
    RowCount := RowCount + ACount;

    for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
    begin
//      AddRow;
      PopulatePlayListGrid(I);
    end;

//    if (FAddedCount > 0) then
//    begin
//      if (FLastCount > 0) then
//        ParentItem := GetParentCueSheetItemByIndex(FLastCount - 1);
////      else
////        ParentItem := GetCueSheetItemByIndex(FLastCount);
//
//      if (ParentItem <> nil) then
//      begin
//        CueSheetItem := GetNextMainItemByItem(ParentItem);
//        if (CueSheetItem <> nil) then
//        begin
//{          ShowMessage(EventTimeToDateTimecodeStr(ParentItem^.StartTime));
//          ShowMessage(EventTimeToDateTimecodeStr(CueSheetItem^.StartTime));
//          ShowMessage(EventTimeToDateTimecodeStr(DurTime)); }
//          if (CompareEventTime(ParentItem^.StartTime, CueSheetItem^.StartTime) > 0) then
//          begin
//            DurTime := GetDurEventTime(CueSheetItem^.StartTime, GetEventEndTime(ParentItem^.StartTime, ParentItem^.DurationTC));
//            ResetStartTimePlus(FChannelCueSheet^.CueSheetList.IndexOf(CueSheetItem), DurTime)
//          end
//          else
//          begin
//            DurTime := GetDurEventTime(GetEventEndTime(ParentItem^.StartTime, ParentItem^.DurationTC), CueSheetItem^.StartTime);
//            ResetStartTimeMinus(FChannelCueSheet^.CueSheetList.IndexOf(CueSheetItem), DurTime);
//          end;
//        end;
//      end;
//    end;

//    SaveFixedCells := False; // Because of node expand and collpase bug

    RemoveAllNodes;
    I := CNT_CUESHEET_HEADER;//OldRowCount - CNT_CUESHEET_FOOTER;
    J := CNT_CUESHEET_HEADER;//OldRowCount - CNT_CUESHEET_FOOTER;
    while (I < RowCount - CNT_CUESHEET_FOOTER - 1) do
    begin
      CueSheetItem := GetCueSheetItemByIndex(RealRowIndex(I) - CNT_CUESHEET_HEADER);

      if (CueSheetItem <> nil) then
      begin
        if (CueSheetItem^.EventMode = EM_COMMENT) then
        begin
          Inc(I);
          Inc(J);
          Continue;
        end;

        while (J < RowCount - CNT_CUESHEET_FOOTER - 1) and (CueSheetItem^.GroupNo = FChannelCueSheet^.CueSheetList[RealRowIndex(J) - CNT_CUESHEET_HEADER + 1]^.GroupNo) do
          Inc(J);

        if (I <> J) then
          AddNode(I, J - I + 1)
        else
          CellProperties[0, I].NodeLevel := 0; // Because of node show tree bug
      end;
      I := J + 1;
      J := I;
    end;

//    with Columns[IDX_COL_CUESHEET_GROUP] do
//      AutoSize := True;

    ContractAll;

    Cells[IDX_COL_CUESHEET_NO, RowCount - 1] := 'End of event';
    MergeCells(IDX_COL_CUESHEET_NO, RowCount - 1, ColCount - IDX_COL_CUESHEET_NO, 1);
    EndUpdate;
  end;
end;

procedure TfrmChannel.PopulatePlayListTimeLine(AIndex: Integer);
var
  I: Integer;
  Item, ParentItem: PCueSheetItem;
  CompIndex: Integer;
  Track: TTrack;
  SubStartTime: TEventTime;
begin
  Item := GetCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    if (Item^.EventStatus = esSkipped) then exit;

    with wmtlPlaylist do
    begin
      CompIndex := 0;
      case Item^.EventMode of
        EM_MAIN: CompIndex := 0;
        EM_JOIN,
        EM_SUB: CompIndex := GetChildCueSheetIndexByItem(Item);
        else
          exit;
      end;

      if (CompIndex >= 0) and (CompIndex < DataGroupProperty.Count) then
      begin
//        DataCompositions[CompIndex].Tracks.BeginUpdate;
        try
          Track := DataCompositions[CompIndex].Tracks.GetTrackByData(Item);
          if (Track = nil) then
          begin
            Track := DataCompositions[CompIndex].Tracks.Add;
            Track.Data := Item;
          end;

          case Item^.EventMode of
            EM_MAIN:
            begin
              Track.InPoint  := TimecodeToFrame(Item^.StartTime.T) + DaysBetween(FFirstStartTime.D, Item^.StartTime.D) * FTimeLineDaysPerFrames;
              Track.OutPoint := Track.InPoint + TimecodeToFrame(Item^.DurationTC) - 1;
            end;
            EM_JOIN:
            begin
              ParentItem := GetParentCueSheetItemByItem(Item);
              if (ParentItem <> nil) then
              begin
                SubStartTime := GetEventTimeSubBegin(ParentItem^.StartTime, Item^.StartTime.T);

                Track.InPoint  := TimecodeToFrame(SubStartTime.T) + DaysBetween(FFirstStartTime.D, SubStartTime.D) * FTimeLineDaysPerFrames;
                Track.OutPoint := Track.InPoint + TimecodeToFrame(Item^.DurationTC) - 1;
              end;
            end;
            EM_SUB:
            begin
              ParentItem := GetParentCueSheetItemByItem(Item);
              if (ParentItem <> nil) then
              begin
                if (Item^.StartMode = SM_SUBBEGIN) then
                  SubStartTime := GetEventTimeSubBegin(ParentItem^.StartTime, Item^.StartTime.T)
                else
                  SubStartTime := GetEventTimeSubEnd(ParentItem^.StartTime, ParentItem^.DurationTC, Item^.StartTime.T);

                Track.InPoint  := TimecodeToFrame(SubStartTime.T) + DaysBetween(FFirstStartTime.D, SubStartTime.D) * FTimeLineDaysPerFrames;
                Track.OutPoint := Track.InPoint + TimecodeToFrame(Item^.DurationTC) - 1;
              end;
            end;
          end;

          Track.Duration := Track.OutPoint - Track.InPoint;

          Track.Color          := GetProgramTypeColorByCode(Item^.ProgramType);
          Track.ColorCaption   := GetProgramTypeColorByCode(Item^.ProgramType);
//          Track.ColorSelected  := Track.Color;
//          Track.ColorHighLight := $000E0607;
//          Track.ColorShadow    := $000E0607;

          Track.Caption := String(Item^.Title);

          if (FTimeLineMin = 0) or (Track.InPoint < FTimeLineMin) then FTimeLineMin := Track.InPoint;
          if (Track.OutPoint > FTimeLineMax) then FTimeLineMax := Track.OutPoint;
        finally
//          DataCompositions[CompIndex].Tracks.EndUpdate;
        end;
      end;
    end;
  end;
end;

procedure TfrmChannel.DisplayPlayListTimeLine(AIndex: Integer);
var
  I: Integer;
  SideFrames: Integer;
begin
  inherited;
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
//  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  wmtlPlaylist.DataGroup.Visible := False;
  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    PopulatePlayListTimeLine(I);
  end;

  SideFrames := TimecodeToFrame(GV_SettingTimeParameter.AutoIncreaseDurationAmount);
  with wmtlPlaylist do
  begin
    TimeZoneProperty.FrameStart := FTimeLineMin - SideFrames;
    TimeZoneProperty.FrameCount := FTimeLineMax - FTimeLineMin + (SideFrames * 2);
//    FrameNumber := FTimeLineMin - SideFrames;
  end;
  wmtlPlaylist.DataGroup.Visible := True;
end;

procedure TfrmChannel.DeletePlayListTimeLine(AItem: PCueSheetItem);
var
  CompIndex: Integer;
  Track: TTrack;
begin
  if (AItem = nil) then exit;

  with wmtlPlaylist do
  begin
    CompIndex := 0;
    case AItem^.EventMode of
      EM_MAIN: CompIndex := 0;
      EM_JOIN,
      EM_SUB: CompIndex := GetChildCueSheetIndexByItem(AItem);
      else
        exit;
    end;

    if (CompIndex >= 0) and (CompIndex < DataGroupProperty.Count) then
    begin
      DataCompositions[CompIndex].Tracks.BeginUpdate;
      try
        Track := DataCompositions[CompIndex].Tracks.GetTrackByData(AItem);
        if (Track <> nil) then FreeAndNil(Track);
      finally
        DataCompositions[CompIndex].Tracks.EndUpdate;
      end;
    end;
  end;
end;

procedure TfrmChannel.ResetNo(AIndex: Integer; ANo: Integer);
var
  I: Integer;
  CItem: PCueSheetItem;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex >= FChannelCueSheet^.CueSheetList.Count) then exit;

  FLastDisplayNo := ANo;
  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventMode = EM_MAIN) then
    begin
      acgPlaylist.AllCells[IDX_COL_CUESHEET_NO, I + CNT_CUESHEET_HEADER] := Format('%d', [FLastDisplayNo + 1]);
      Inc(FLastDisplayNo);
    end;
  end;
end;

procedure TfrmChannel.ResetStartDate(AIndex: Integer; ADays: Integer);
var
  I: Integer;
  CItem: PCueSheetItem;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex >= FChannelCueSheet^.CueSheetList.Count) then exit;

  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem <> nil) then
    begin
      if (CItem^.EventMode <> EM_COMMENT) then
      begin
        CItem^.StartTime.D := IncDay(CItem^.StartTime.D, ADays);

{        // Update playlist grid
        if (CItem^.EventMode = EM_MAIN) then
        begin
          acgPlaylist.AllCells[IDX_COL_CUESHEET_START_DATE, I + CNT_CUESHEET_HEADER] := FormatDateTime(FORMAT_DATE, CItem^.StartTime.D);
        end; }
      end;
    end;
  end;
end;

procedure TfrmChannel.ResetStartTime(AIndex: Integer);
var
  ParentItem, NextItem: PCueSheetItem;
  ParentEndTime, DurTime: TEventTime;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex >= FChannelCueSheet^.CueSheetList.Count) then exit;

  ParentItem := GetParentCueSheetItemByIndex(AIndex);
  if (ParentItem <> nil) then
  begin


    NextItem := GetNextMainItemByItem(ParentItem);
    if (NextItem <> nil) then
    begin
      ParentEndTime := GetEventEndTime(ParentItem^.StartTime, ParentItem^.DurationTC);
//      ShowMessage(EventTimeToDateTimecodeStr(GetEventEndTime(PItem^.StartTime, PItem^.DurationTC)));
      DurTime := GetDurEventTime(ParentEndTime, NextItem^.StartTime);
//      ShowMessage(EventTimeToDateTimecodeStr(NItem^.StartTime));
      ShowMessage(EventTimeToDateTimecodeStr(DurTime));
      ShowMessage(EventTimeToDateTimecodeStr(ParentEndTime));
      ShowMessage(EventTimeToDateTimecodeStr(NextItem^.StartTime));
//      ResetStartTimePlus(FChannelCueSheet^.CueSheetList.IndexOf(NItem), DurTime);
//exit;
      if (CompareEventTime(ParentEndTime, NextItem^.StartTime) >= 0) then
      begin
//        DurTime := GetDurEventTime(NItem^.StartTime, GetEventEndTime(PItem^.StartTime, PItem^.DurationTC));
//        ShowMessage(EventTimeToDateTimecodeStr(DurTime));
        ResetStartTimePlus(FChannelCueSheet^.CueSheetList.IndexOf(NextItem), DurTime)
      end
      else
      begin
//        DurTime := GetDurEventTime(GetEventEndTime(PItem^.StartTime, PItem^.DurationTC), NItem^.StartTime);
//        ShowMessage(EventTimeToDateTimecodeStr(DurTime));
        ResetStartTimeMinus(FChannelCueSheet^.CueSheetList.IndexOf(NextItem), DurTime);
      end;
    end;
  end;

  if (ParentItem = GetFirstMainItem) then
  begin
    FFirstStartTime := ParentItem^.StartTime;
    FTimeLineMin := 0;
    FTimeLineMax := 0;
  end;

  DisplayPlayListTimeLine(AIndex);
end;

procedure TfrmChannel.ResetStartTime(AIndex: Integer; ASaveEndTime: TEventTime);
var
  Item: PCueSheetItem;
  ItemEndTime: TEventTime;
  DurTime: TEventTime;

  NextItem: PCueSheetItem;
  NextIndex: Integer;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex >= FChannelCueSheet^.CueSheetList.Count) then exit;

  Item := GetCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    if (Item^.EventMode = EM_MAIN) then
    begin
      ItemEndTime := GetEventEndTime(Item^.StartTime, Item^.DurationTC);

      DurTime := GetDurEventTime(ASaveEndTime, ItemEndTime);

      NextItem := GetNextMainItemByItem(Item);
      if (NextItem <> nil) then
      begin
        NextIndex := GetCueSheetIndexByItem(NextItem);
        if (CompareEventTime(ASaveEndTime, ItemEndTime) <= 0) then
        begin
          ResetStartTimePlus(NextIndex, DurTime)
        end
        else
        begin
          ResetStartTimeMinus(NextIndex, DurTime);
        end;
      end;

      if (Item = GetFirstMainItem) then
      begin
        FFirstStartTime := Item^.StartTime;
        FTimeLineMin := 0;
        FTimeLineMax := 0;
      end;

      ResetChildItems(AIndex);
    end;

    DisplayPlayListTimeLine(AIndex);
  end;
end;

procedure TfrmChannel.ResetStartTimeByTime(AIndex: Integer; ASaveStartTime: TEventTime; ASaveDurationTC: TTimecode);
var
  EndTime: TEventTime;
begin
  EndTime := GetEventEndTime(ASaveStartTime, ASaveDurationTC);
  ResetStartTime(AIndex, EndTime);
end;

procedure TfrmChannel.ResetStartTimeByTime(AIndex: Integer; ASaveStartTime: TEventTime);
var
  Item: PCueSheetItem;
  EndTime: TEventTime;
begin
  Item := GetCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    EndTime := GetEventEndTime(ASaveStartTime, Item^.DurationTC);
    ResetStartTime(AIndex, EndTime);
  end;
end;

procedure TfrmChannel.ResetStartTimeByTime(AIndex: Integer; ASaveDurationTC: TTimecode);
var
  Item: PCueSheetItem;
  EndTime: TEventTime;
begin
  Item := GetCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    EndTime := GetEventEndTime(Item^.StartTime, ASaveDurationTC);
    ResetStartTime(AIndex, EndTime);
  end;
end;

procedure TfrmChannel.ResetStartTimeByDuration(AIndex: Integer; ADuration: TTimecode);
var
  Item: PCueSheetItem;
  NextItem: PCueSheetItem;
  NextIndex: Integer;

  DurTime: TEventTime;

  I: Integer;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex >= FChannelCueSheet^.CueSheetList.Count) then exit;

  Item := GetCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    if (Item^.EventMode = EM_MAIN) then
    begin
      DurTime := TimecodeToEventTime(GetDurTimecode(Item^.DurationTC, ADuration));

      NextItem := GetNextMainItemByItem(Item);
      if (NextItem <> nil) then
      begin
        NextIndex := GetCueSheetIndexByItem(NextItem);
        if (Item^.DurationTC <= ADuration) then
        begin
          ResetStartTimePlus(NextIndex, DurTime)
        end
        else
        begin
          ResetStartTimeMinus(NextIndex, DurTime);
        end;
      end;

      Item^.DurationTC := ADuration;
      ResetChildItems(AIndex);
    end
    else
      Item^.DurationTC := ADuration;

    DisplayPlayListTimeLine(AIndex);
  end;
end;

procedure TfrmChannel.ResetStartTimePlus(AIndex: Integer; ADurEventTime: TEventTime);
var
  NextIndex: Integer;
  NextItem: PCueSheetItem;
  NextStartTime: TEventTime;

  I: Integer;

  ParentItem: PCueSheetItem;
  ParentStartDate: TDate;
  ParentEndTime: TEventTime;

  CItem, PItem: PCueSheetItem;
  CStartTime, PEndTime: TEventTime;
  PStartDate: TDate;    // Parent cuesheet start date
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex >= FChannelCueSheet^.CueSheetList.Count) then exit;

  // Get current or next main event index
  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    NextItem := FChannelCueSheet^.CueSheetList[I];
    if (NextItem <> nil) and
       (NextItem^.EventMode = EM_MAIN) and (NextItem^.EventStatus <> esSkipped) then
    begin
      NextIndex := I;
      break;
    end;
  end;

  ParentItem := nil;
  ParentStartDate := 0;

  for I := NextIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    NextItem := FChannelCueSheet^.CueSheetList[I];

    if (NextItem^.EventMode <> EM_COMMENT) then
    begin
      // If then main event then checks whether the current start time is less than the start time of the previous event.
      if (NextItem^.EventMode = EM_MAIN) then
      begin
        NextStartTime := GetPlusEventTime(NextItem^.StartTime, ADurEventTime);
//          ShowMessage(EventTimeToDateTimecodeStr(CStartTime));
        if (ParentItem <> nil) then
        begin
          ParentEndTime := GetEventEndTime(ParentItem^.StartTime, ParentItem^.DurationTC);
//            ShowMessage(EventTimeToDateTimecodeStr(PEndTime));
          if (CompareEventTime(NextStartTime, ParentEndTime) < 0) then
          begin
            NextStartTime := ParentEndTime;
          end;
        end;
        ParentItem := NextItem;

        ParentStartDate := NextStartTime.D;

        NextItem^.StartTime := NextStartTime;
      end
      else
      begin
        NextItem^.StartTime.D := ParentStartDate;
      end;

      if (NextItem^.EventMode = EM_MAIN) then
      begin
        acgPlaylist.AllCells[IDX_COL_CUESHEET_START_DATE, I + CNT_CUESHEET_HEADER] := FormatDateTime(FORMAT_DATE, NextItem^.StartTime.D);
      end;

      if (NextItem^.EventMode in [EM_MAIN, EM_SUB]) then
      begin
        acgPlaylist.AllCells[IDX_COL_CUESHEET_START_TIME, I + CNT_CUESHEET_HEADER] := TimecodeToString(NextItem^.StartTime.T);
      end;
    end;
  end;

  exit;

  // Get next main event index
  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem <> nil) and (CItem^.EventMode = EM_MAIN) then
    begin
      AIndex := I;
      break;
    end;
  end;

  PItem := nil;
  PStartDate := 0;
  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem <> nil) then
    begin
      if (CItem^.EventMode <> EM_COMMENT) then
      begin
        // If then main event then checks whether the current start time is less than the start time of the previous event.
        if (CItem^.EventMode = EM_MAIN) then
        begin
          CStartTime := GetPlusEventTime(CItem^.StartTime, ADurEventTime);
//          ShowMessage(EventTimeToDateTimecodeStr(CStartTime));
          if (PItem <> nil) then
          begin
            PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
//            ShowMessage(EventTimeToDateTimecodeStr(PEndTime));
            if (CompareEventTime(CStartTime, PEndTime) < 0) then
            begin
              CStartTime := PEndTime;
            end;
          end;
          PItem := CItem;

          PStartDate := CStartTime.D;

          CItem^.StartTime := CStartTime;
        end
        else
        begin
          CItem^.StartTime.D := PStartDate;
        end;


  //      acgPlaylist.Cells[IDX_COL_CUESHEET_START_DATE, DispRow] := FormatDateTime(FORMAT_DATE, CItem^.StartTime.D);
  //      acgPlaylist.Cells[IDX_COL_CUESHEET_START_TIME, DispRow] := TimecodeToString(CItem^.StartTime.T);


        if (CItem^.EventMode = EM_MAIN) then
        begin
          acgPlaylist.AllCells[IDX_COL_CUESHEET_START_DATE, I + CNT_CUESHEET_HEADER] := FormatDateTime(FORMAT_DATE, CItem^.StartTime.D);
        end;

        if (CItem^.EventMode in [EM_MAIN, EM_SUB]) then
        begin
          acgPlaylist.AllCells[IDX_COL_CUESHEET_START_TIME, I + CNT_CUESHEET_HEADER] := TimecodeToString(CItem^.StartTime.T);
        end;
      end;
    end;
  end;
end;

procedure TfrmChannel.ResetStartTimeMinus(AIndex: Integer; ADurEventTime: TEventTime);
var
  NextIndex: Integer;
  NextItem: PCueSheetItem;
  NextStartTime: TEventTime;

  I: Integer;

  ParentItem: PCueSheetItem;
  ParentStartDate: TDate;
  ParentEndTime: TEventTime;

  CItem, PItem: PCueSheetItem;
  CStartTime, PEndTime: TEventTime;
  PStartDate: TDate;    // Parent cuesheet start date
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex >= FChannelCueSheet^.CueSheetList.Count) then exit;

  // Get current or next main event index
  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    NextItem := FChannelCueSheet^.CueSheetList[I];
    if (NextItem <> nil) and
       (NextItem^.EventMode = EM_MAIN) and (NextItem^.EventStatus <> esSkipped) then
    begin
      NextIndex := I;
      break;
    end;
  end;

  ParentItem := nil;
  ParentStartDate := 0;

  for I := NextIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    NextItem := FChannelCueSheet^.CueSheetList[I];

    if (NextItem^.EventMode <> EM_COMMENT) then
    begin
      // If then main event then checks whether the current start time is less than the start time of the previous event.
      if (NextItem^.EventMode = EM_MAIN) then
      begin
        NextStartTime := GetMinusEventTime(NextItem^.StartTime, ADurEventTime);
//          ShowMessage(EventTimeToDateTimecodeStr(CStartTime));
        if (ParentItem <> nil) then
        begin
          ParentEndTime := GetEventEndTime(ParentItem^.StartTime, ParentItem^.DurationTC);
//            ShowMessage(EventTimeToDateTimecodeStr(PEndTime));
          if (CompareEventTime(NextStartTime, ParentEndTime) < 0) then
          begin
            NextStartTime := ParentEndTime;
          end;
        end;
        ParentItem := NextItem;

        ParentStartDate := NextStartTime.D;

        NextItem^.StartTime := NextStartTime;
      end
      else
      begin
        NextItem^.StartTime.D := ParentStartDate;
      end;

      if (NextItem^.EventMode = EM_MAIN) then
      begin
        acgPlaylist.AllCells[IDX_COL_CUESHEET_START_DATE, I + CNT_CUESHEET_HEADER] := FormatDateTime(FORMAT_DATE, NextItem^.StartTime.D);
      end;

      if (NextItem^.EventMode in [EM_MAIN, EM_SUB]) then
      begin
        acgPlaylist.AllCells[IDX_COL_CUESHEET_START_TIME, I + CNT_CUESHEET_HEADER] := TimecodeToString(NextItem^.StartTime.T);
      end;
    end;
  end;

  exit;

  PItem := nil;
  PStartDate := 0;
  for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem <> nil) then
    begin
      if (CItem^.EventMode <> EM_COMMENT) then
      begin
        // If then main event then checks whether the current start time is less than the start time of the previous event.
        if (CItem^.EventMode = EM_MAIN) then
        begin
          CStartTime := GetMinusEventTime(CItem^.StartTime, ADurEventTime);
          if (PItem <> nil) then
          begin
            PEndTime := GetEventEndTime(PItem^.StartTime, PItem^.DurationTC);
            if (CompareEventTime(CStartTime, PEndTime) < 0) then
            begin
              CStartTime := PEndTime;
            end;
          end;
          PItem := CItem;

          PStartDate := CStartTime.D;

          CItem^.StartTime := CStartTime;
        end
        else
        begin
          CItem^.StartTime.D := PStartDate;
        end;


  //      acgPlaylist.Cells[IDX_COL_CUESHEET_START_DATE, DispRow] := FormatDateTime(FORMAT_DATE, CItem^.StartTime.D);
  //      acgPlaylist.Cells[IDX_COL_CUESHEET_START_TIME, DispRow] := TimecodeToString(CItem^.StartTime.T);


        if (CItem^.EventMode = EM_MAIN) then
        begin
          acgPlaylist.AllCells[IDX_COL_CUESHEET_START_DATE, I + CNT_CUESHEET_HEADER] := FormatDateTime(FORMAT_DATE, CItem^.StartTime.D);
        end;

        if (CItem^.EventMode in [EM_MAIN, EM_SUB]) then
        begin
          acgPlaylist.AllCells[IDX_COL_CUESHEET_START_TIME, I + CNT_CUESHEET_HEADER] := TimecodeToString(CItem^.StartTime.T);
        end;
      end;
    end;
  end;
end;

procedure TfrmChannel.ResetChildItems(AIndex: Integer);
var
  I: Integer;
  ParentItem: PCueSheetItem;
  ParentIndex: Integer;

  Item: PCueSheetItem;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  ParentItem := GetParentCueSheetItemByIndex(AIndex);
  if (ParentItem <> nil) then
  begin
    ParentIndex := GetCueSheetIndexByItem(ParentItem);
    for I := ParentIndex + 1 to FChannelCueSheet^.CueSheetList.Count - 1 do
    begin
      Item := GetCueSheetItemByIndex(I);
      if (Item <> nil) and (Item^.GroupNo = ParentItem^.GroupNo) then
      begin
        case Item^.EventMode of
          EM_JOIN:
          begin
            Item^.StartTime.D := ParentItem^.StartTime.D;
            Item^.DurationTC  := ParentItem^.DurationTC;
            Item^.InTC        := ParentItem^.InTC;
            Item^.OutTC       := ParentItem^.OutTC;
          end;
          EM_SUB:
          begin
            Item^.StartTime.D := ParentItem^.StartTime.D;
          end;
        end;
      end
      else
        break;
    end;
  end;
end;

procedure TfrmChannel.ResetStartDateTimeline(AIndex: Integer);
var
  I: Integer;
  CItem, PItem: PCueSheetItem;
  SubStartTime: TEventTime;
  CompIndex: Integer;
  Track: TTrack;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex >= FChannelCueSheet^.CueSheetList.Count) then exit;

  with wmtlPlaylist do
  begin
    for I := 0 to DataGroupProperty.Count - 1 do
      DataCompositions[I].Tracks.BeginUpdate;

    try
      for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
      begin
        CItem := FChannelCueSheet^.CueSheetList[I];
        if (CItem <> nil) then
        begin
          with CItem^ do
          begin
            if (EventMode <> EM_COMMENT) then
            begin
              // Update playlist timeline
              case CItem^.EventMode of
                EM_MAIN:
                begin
                  CompIndex := 0;
                  if (CompIndex < DataGroupProperty.Count) then
                  begin
                    Track := DataCompositions[CompIndex].Tracks.GetTrackByData(CItem);
                    if (Track <> nil) then
                    begin
                      Track.InPoint  := TimecodeToFrame(StartTime.T) + Trunc(DaySpan(FFirstStartTime.D, StartTime.D)) * FTimeLineDaysPerFrames;
                      Track.Duration := TimecodeToFrame(DurationTC);
                      Track.OutPoint := Track.InPoint + Track.Duration;
                    end;
                    Inc(CompIndex);
                  end;
                end;
                EM_JOIN:
                begin
                  if (CompIndex < DataGroupProperty.Count) then
                  begin
                    PItem := GetParentCueSheetItemByItem(CItem);
                    if (PItem <> nil) then
                    begin
                      SubStartTime := GetEventTimeSubBegin(PItem^.StartTime, CItem^.StartTime.T);

                      Track := DataCompositions[CompIndex].Tracks.GetTrackByData(CItem);
                      if (Track <> nil) then
                      begin
                        Track.Duration := TimecodeToFrame(DurationTC);
                        Track.InPoint  := TimecodeToFrame(SubStartTime.T) + Trunc(DaySpan(FFirstStartTime.D, SubStartTime.D)) * FTimeLineDaysPerFrames;
                        Track.OutPoint := Track.InPoint + Track.Duration;
                      end;
                      Inc(CompIndex);
                    end;
                  end;
                end;
                EM_SUB:
                begin
                  if (CompIndex < DataGroupProperty.Count) then
                  begin
                    PItem := GetParentCueSheetItemByItem(CItem);
                    if (PItem <> nil) then
                    begin
                      if (CItem^.StartMode = SM_SUBBEGIN) then
                        SubStartTime := GetEventTimeSubBegin(PItem^.StartTime, CItem^.StartTime.T)
                      else
                        SubStartTime := GetEventTimeSubEnd(PItem^.StartTime, PItem^.DurationTC, CItem^.StartTime.T);

                      Track := DataCompositions[CompIndex].Tracks.GetTrackByData(CItem);
                      if (Track <> nil) then
                      begin
                        Track.Duration := TimecodeToFrame(DurationTC);
                        Track.InPoint  := TimecodeToFrame(SubStartTime.T) + Trunc(DaySpan(FFirstStartTime.D, SubStartTime.D)) * FTimeLineDaysPerFrames;
                        Track.OutPoint := Track.InPoint + Track.Duration;
                      end;
                      Inc(CompIndex);
                    end;
                  end;
                end
                else Continue;
              end;
              if (Track <> nil) then
              begin
                if (FTimeLineMin = 0) or (Track.InPoint < FTimeLineMin) then FTimeLineMin := Track.InPoint;
                if (Track.OutPoint > FTimeLineMax) then FTimeLineMax := Track.OutPoint;
              end;
            end;
          end;
        end;
      end;

//      TimeZoneProperty.FrameStart := FTimeLineMin - TIMELINE_SIDE_FRAMES;
//      TimeZoneProperty.FrameCount := FTimeLineMax - FTimeLineMin + (TIMELINE_SIDE_FRAMES * 2);
    finally
      for I := 0 to DataGroupProperty.Count - 1 do
        DataCompositions[I].Tracks.EndUpdate;
    end;
  end;
end;

procedure TfrmChannel.CueSheetListQuickSort(L, R: Integer; ACueSheetList: TCueSheetList);
var
  I, J, P: Integer;
  Save: PCueSheetItem;
  SortList: TCueSheetList;

  function Compare(Item1, Item2: PCueSheetItem): Integer;
  begin
    Result := CompareEventTime(Item1^.StartTime, Item2^.StartTime);
  end;

begin
  SortList := ACueSheetList;
  repeat
    I := L;
    J := R;
    P := (L + R) shr 1;
    repeat
      while Compare(ACueSheetList[I], ACueSheetList[P]) < 0 do
        Inc(I);
      while Compare(ACueSheetList[J], ACueSheetList[P]) > 0 do
        Dec(J);
      if I <= J then
      begin
        Save        := SortList[I];
        SortList[I] := SortList[J];
        SortList[J] := Save;
        if P = I then
          P := J
        else if P = J then
          P := I;
        Inc(I);
        Dec(J);
      end;
    until I > J;
    if L < J then
      CueSheetListQuickSort(L, J, ACueSheetList);
    L := I;
  until I >= R;
end;

procedure TfrmChannel.SetCueSheetItemStatusByIndex(AStartIndex, AEndIndex: Integer; AStatus: TEventStatus);
var
  I: Integer;
  RRow: Integer;
  PSItem, PEItem, CItem: PCueSheetItem;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AStartIndex < 0) or (AStartIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;
  if (AEndIndex < 0) or (AEndIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

{  PSItem := GetParentCueSheetItemByIndex(AStartIndex);
  AStartIndex := GetCueSheetIndexByItem(PSItem);
  if (AStartIndex < 0) then exit;

  PEItem := GetParentCueSheetItemByIndex(AEndIndex);
  AEndIndex := GetLastChildCueSheetItemByIndex(PEItem);
  if (AEndIndex < 0) then exit;





  GetLastMainItem
  EItem := GetLastChildCueSheetItemByIndex(AEndIndex);
  AEndIndex := GetCueSheetIndexByItem(EItem);
  if (AStartIndex < 0) then exit;  }

  for I := AStartIndex to AEndIndex do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventMode <> EM_COMMENT) {and (CItem^.EventStatus <= esPreroll)} then
    begin
      CItem^.EventStatus := AStatus;
      with acgPlayList do
      begin
        RRow := I + CNT_CUESHEET_HEADER;
        AllCells[IDX_COL_CUESHEET_EVENT_STATUS, RRow] := EventStatusNames[AStatus];
      end;
    end;
  end;

  acgPlaylist.Repaint;
end;

procedure TfrmChannel.SetChannelOnAir(AOnAir: Boolean);
begin
  SetChannelOnAirByID(FChannelID, AOnAir);
  FChannelOnAir := AOnAir;

  lblOnAirFlag.Caption := OnAirFlagNames[FChannelOnAir];
  if (FChannelOnAir) then
  begin
    lblOnAirFlag.Font.Color := clLime;
  end
  else
  begin
    lblOnAirFlag.Font.Color := clRed;
  end;

  wmtlPlaylist.TimeZoneProperty.RailBarVisible := FChannelOnAir;
end;

procedure TfrmChannel.OnAirInputEvents(AIndex, ACount: Integer);
var
  I: Integer;
  PItem, CItem: PCueSheetItem;
  PGroupNo: Word;
  GroupCount: Integer;
  R: Integer;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  GroupCount := 0;
  PItem := GetParentCueSheetItemByIndex(AIndex);
  if (PItem <> nil) then
  begin
    PGroupNo := PItem^.GroupNo;
    AIndex := GetCueSheetIndexByItem(PItem);
    for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
    begin
      CItem := FChannelCueSheet^.CueSheetList[I];
      if (CItem^.EventMode = EM_COMMENT) then continue;

{      if (PGroupNo <> CItem^.GroupNo) then
      begin
        // 첫 운행 이벤트가 에러일 경우 자식 이벤트까지만 에러 처리
        if (PItem^.EventStatus = esError) and (FCueSheetNext = nil) then
          break;
      end; }

      if (CItem^.EventStatus <> esSkipped) then
      begin
        R := InputEvent(CItem);

        if (R <> D_OK) then
        begin
          if {(FCueSheetNext = nil) or }(FCueSheetNext = CItem) or
             ((FCueSheetNext <> nil) and (FCueSheetNext^.GroupNo = CItem^.GroupNo)) then
          begin
            CItem^.EventStatus := esError;

//            acgPlaylist.RepaintRow(I + CNT_CUESHEET_HEADER);
          end;
//          break;
        end;
      end;

      if (PGroupNo <> CItem^.GroupNo) then
      begin
        PGroupNo := CItem^.GroupNo;
        Inc(GroupCount);
      end;

      if (GroupCount >= ACount - 1) then break;
    end;
  end;

//  if (PItem^.EventMode = EM_MAIN) and (FCueSheetNext = nil) then
//    FCueSheetNext := PItem;
end;

procedure TfrmChannel.OnAirDeleteEvents(AFromIndex, AToIndex: Integer);
var
  I: Integer;
  PItem, CItem: PCueSheetItem;
  PGroupNo: Word;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AFromIndex < 0) or (AFromIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;
  if (AToIndex < 0) or (AToIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  PItem := GetParentCueSheetItemByIndex(AFromIndex);
  if (PItem = nil) then exit;

  AFromIndex := GetCueSheetIndexByItem(PItem);

  CItem := GetLastChildCueSheetItemByIndex(AToIndex);
  if (CItem = nil) then exit;

  AToIndex := GetCueSheetIndexByItem(CItem);

  for I := AFromIndex to AToIndex do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    DeletePlayListTimeLine(CItem);
    DeleteEvent(CItem);
    if (CItem^.EventStatus <> esSkipped) then
      CItem^.EventStatus := esIdle;

    if (CItem^.EventMode = EM_MAIN) and (FCueSheetNext = CItem) then
      FCueSheetNext := nil;
  end;
end;

procedure TfrmChannel.OnAirClearEvents;
var
  I: Integer;
  PItem, CItem: PCueSheetItem;
  FromIndex, ToIndex: Integer;
begin
  ClearEvent;

  if (FCueSheetCurr <> nil) then
    PItem := GetParentCueSheetItemByItem(FCueSheetCurr)
  else if (FCueSheetNext <> nil) then
    PItem := GetParentCueSheetItemByItem(FCueSheetNext)
  else exit;

  if (PItem = nil) then exit;

  FromIndex := GetCueSheetIndexByItem(PItem);

  CItem := GetLastChildCueSheetItemByIndex(FromIndex + GV_SettingOption.MaxInputEventCount - 1);
  if (CItem <> nil) then
    ToIndex := GetCueSheetIndexByItem(CItem)
  else
    ToIndex := FChannelCueSheet^.CueSheetList.Count - 1;

  for I := FromIndex to ToIndex do
//  for I := 0 to FChannelCueSheet^.CueSheetList.Count - 1 do
  begin
    CItem := FChannelCueSheet^.CueSheetList[I];
    if (CItem^.EventStatus <> esSkipped) then
    begin
      CItem^.EventStatus := esIdle;

      if (CItem^.EventMode <> EM_COMMENT) then
        with acgPlaylist do
          AllCells[IDX_COL_CUESHEET_EVENT_STATUS, RealRowIndex(I) + CNT_CUESHEET_HEADER] := EventStatusNames[CItem^.EventStatus];
    end;

    if (CItem^.EventMode = EM_MAIN) and (FCueSheetNext = CItem) then
      FCueSheetNext := nil;
  end;
end;

procedure TfrmChannel.OnAirTakeEvent(AIndex: Integer);
var
  ParentItem: PCueSheetItem;
  Item: PCueSheetItem;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

{  Parent := GetParentCueSheetItemByIndex(AIndex);

  GroupCount := 0;
  PItem := GetParentCueSheetItemByIndex(AIndex);
  if (PItem <> nil) then
  begin
    PGroupNo := PItem^.GroupNo;
    AIndex := GetCueSheetIndexByItem(PItem);
    for I := AIndex to FChannelCueSheet^.CueSheetList.Count - 1 do
    begin
      CItem := FChannelCueSheet^.CueSheetList[I];
      if (CItem^.EventMode = EM_COMMENT) then continue; }



  Item := GetParentCueSheetItemByIndex(AIndex);
  if (Item <> nil) then
  begin
    TakeEvent(Item);
  end;
end;

procedure TfrmChannel.OnAirHoldEvent(AIndex: Integer);
var
  CItem: PCueSheetItem;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;
  if (AIndex < 0) or (AIndex > FChannelCueSheet^.CueSheetList.Count - 1) then exit;

  CItem := GetParentCueSheetItemByIndex(AIndex);
  if (CItem <> nil) then
  begin
    HoldEvent(CItem);
  end;
end;

procedure TfrmChannel.OnAirChangeDurationEvent(ADuration: TTimecode);
var
  CItem: PCueSheetItem;
begin
  if (FCueSheetCurr = nil) then exit;

  ChangeDurationEvent(FCueSheetCurr, ADuration);
end;

function TfrmChannel.MakePlayerEvent(AItem: PCueSheetItem; var AEvent: TEvent): Integer;
var
  ParentItem: PCueSheetItem;
begin
  Result := D_FALSE;

  FillChar(AEvent, SizeOf(TEvent), #0);
  with AEvent do
  begin
    EventID   := AItem^.EventID;
    StartTime := AItem^.StartTime;
    if (AItem^.StartMode in [SM_SUBBEGIN, SM_SUBEND]) then
    begin
      ParentItem := GetParentCueSheetItemByItem(AItem);
      if (ParentItem <> nil) then
      begin
        if (AItem^.StartMode = SM_SUBBEGIN) then
        begin
//          StartTime := GetPlusEventTime(ParentItem^.StartTime, TimecodeToEventTime(AItem^.StartTime.T));

//          StartTime := ParentItem^.StartTime;
//          StartTime.T := StartTime.T + AItem^.StartTime.T;

          StartTime := GetEventTimeSubBegin(ParentItem^.StartTime, AItem^.StartTime.T);

//            ShowMessage(DateTimeToStr(EventTimeToDateTime(StartTime)));
        end
        else
        begin
//          StartTime := GetMinusEventTime(GetPlusEventTime(ParentItem^.StartTime, TimecodeToEventTime(ParentItem^.DurationTC)), TimecodeToEventTime(AItem^.StartTime.T));

//          StartTime := ParentItem^.StartTime;
//          StartTime.T := StartTime.T + ParentItem^.DurationTC;
//          StartTime.T := StartTime.T - AItem^.StartTime.T - AItem^.DurationTC;

          StartTime := GetEventTimeSubEnd(ParentItem^.StartTime, ParentItem^.DurationTC, AItem^.StartTime.T);

//            ShowMessage(DateTimeToStr(EventTimeToDateTime(StartTime)));
        end;
      end;
    end;
//      ShowMessage(DateTimeToStr(EventTimeToDateTime(StartTime)));
    DurTime             := AItem^.DurationTC;
    TakeEvent           := False;
    ManualEvent         := (AItem^.StartMode = SM_MANUAL);
    EventType           := ET_PLAYER;
    Player.StartTC      := AItem^.InTC;
//    Player.FinishAction := ACueSheet^.FinishAction;
//    StrCopy(Player.ID.BinNo, ACueSheet^.BinNo);
    StrCopy(Player.ID.ID, AItem^.MediaId);
  end;

  Result := D_OK;
end;

function TfrmChannel.InputEvent(AItem: PCueSheetItem): Integer;
var
  I: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;
  Event: TEvent;

  DCSOK: Integer;
begin
  Result := D_FALSE;

  if (AItem^.EventMode = EM_COMMENT) then exit;

  Source := GetSourceByName(String(AItem^.Source));
  if (Source = nil) then exit;

  SourceHandles := Source^.Handles;
  if (SourceHandles = nil) or (SourceHandles.Count <= 0) then exit;

  case Source^.SourceType of
    ST_VSDEC,
    ST_VCR,
    ST_CG,
    ST_LINE:
      Result := MakePlayerEvent(AItem, Event);
    else
      exit;
  end;

  DCSOK := D_FALSE;
  if (Result = D_OK) then
  begin
    for I := 0 to SourceHandles.Count - 1 do
    begin
      Result := DCSInputEvent(SourceHandles[I]^.DCSID, SourceHandles[I]^.Handle, Event);
      if (Result = D_OK) then
      begin
        DCSOK := D_OK;
      end;
    end;

{    if (DCSOK <> D_OK) then
    begin
      AItem^.EventStatus := esError;

      acgPlaylist.RepaintRow(GetCueSheetIndexByItem(AItem) + CNT_CUESHEET_HEADER);
    end; }

    Result := DCSOK;
  end;
end;

function TfrmChannel.DeleteEvent(AItem: PCueSheetItem): Integer;
var
  I: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;
begin
  Result := D_FALSE;

  if (AItem^.EventMode = EM_COMMENT) then exit;

  Source := GetSourceByName(String(AItem^.Source));
  if (Source = nil) then exit;

  SourceHandles := Source^.Handles;
  if (SourceHandles = nil) or (SourceHandles.Count <= 0) then exit;

  for I := 0 to SourceHandles.Count - 1 do
    Result := DCSDeleteEvent(SourceHandles[I]^.DCSID, SourceHandles[I]^.Handle, AItem^.EventID);

  Result := D_OK;
end;

function TfrmChannel.ClearEvent: Integer;
var
  I, J: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;
begin
  Result := D_FALSE;

  for I := 0 to GV_SourceList.Count - 1 do
  begin
    Source := GV_SourceList[I];

    if (Source^.Channel^.ID = FChannelId) then
    begin
      SourceHandles := Source^.Handles;
      if (SourceHandles <> nil) then
      begin
        for J := 0 to SourceHandles.Count - 1 do
          Result := DCSClearEvent(SourceHandles[J]^.DCSID, SourceHandles[J]^.Handle);
      end;
    end;
  end;
end;

function TfrmChannel.TakeEvent(AItem: PCueSheetItem): Integer;
var
  I: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;
  Event: TEvent;
begin
  Result := D_FALSE;

  if (AItem^.EventMode <> EM_MAIN) {or (AItem^.StartMode <> SM_MANUAL) }then exit;

  Source := GetSourceByName(String(AItem^.Source));
  if (Source = nil) then exit;

  SourceHandles := Source^.Handles;
  if (SourceHandles = nil) or (SourceHandles.Count <= 0) then exit;

  for I := 0 to SourceHandles.Count - 1 do
    Result := DCSTakeEvent(SourceHandles[I]^.DCSID, SourceHandles[I]^.Handle, AItem^.EventID, GV_SettingTimeParameter.StandardTimeCorrection);
end;

function TfrmChannel.HoldEvent(AItem: PCueSheetItem): Integer;
var
  I: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;
  Event: TEvent;
begin
  Result := D_FALSE;

  if (AItem^.EventMode <> EM_MAIN) then exit;

  Source := GetSourceByName(String(AItem^.Source));
  if (Source = nil) then exit;

  SourceHandles := Source^.Handles;
  if (SourceHandles = nil) or (SourceHandles.Count <= 0) then exit;

  for I := 0 to SourceHandles.Count - 1 do
    Result := DCSHoldEvent(SourceHandles[I]^.DCSID, SourceHandles[I]^.Handle, AItem^.EventID);
end;

function TfrmChannel.ChangeDurationEvent(AItem: PCueSheetItem; ADuration: TTimecode): Integer;
var
  I: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;
  Event: TEvent;
begin
  Result := D_FALSE;

  if (AItem^.EventMode <> EM_MAIN) {or (AItem^.StartMode <> SM_MANUAL) }then exit;

  Source := GetSourceByName(String(AItem^.Source));
  if (Source = nil) then exit;

  SourceHandles := Source^.Handles;
  if (SourceHandles = nil) or (SourceHandles.Count <= 0) then exit;

  for I := 0 to SourceHandles.Count - 1 do
    Result := DCSChangetDurationEvent(SourceHandles[I]^.DCSID, SourceHandles[I]^.Handle, AItem^.EventID, ADuration);
end;

procedure TfrmChannel.StartOnAir;
var
  I, J: Integer;
  R: Integer;
  Source: PSource;
  OnAirItem: PCueSheetItem;
  OnAirEventID: TEventID;     // DCS onair event id
  OnAirStartTime: TEventTime; // DCS onair event start time
  OnAirDurationTC: TTimecode; // DCS onair event duration tc

  CuedItem: PCueSheetItem;
  CuedEventID: TEventID;      // DCS cued event id
  CuedStartTime: TEventTime;  // DCS cued event start time
  CuedDurationTC: TTimecode;  // DCS cued event duration tc

  DCSIsOnAir: Boolean;        // DCS onair flag
  IsManualStart: Boolean;     // Manual start flag

  CurrItem, NextItem: PCueSheetItem;
  CurrIndex, NextIndex: Integer;
  SkipIndex: Integer;

  StartTime: TEventTime;
  DurationTC: TTimecode;
  EventStatus: TEventStatus;

  StartOnAirTime: TDateTime;

  ErrorString: String;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList.Count <= 0) then exit;

  // Check the aleady onair
//  if (GetChannelOnAirByID(FChannelID)) then
  if (FChannelOnAir) then
  begin
    ErrorString := Format(SChannelAlreadyRunning, [GetChannelNameByID(FChannelID)]);
    MessageBeep(MB_ICONWARNING);
    MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_OK or MB_ICONWARNING);
    exit;
  end;

  // Source channel check
  DCSIsOnAir := False;
  for I := 0 to GV_SourceList.Count - 1 do
  begin
    Source := GV_SourceList[I];
    if (Source <> nil) and
       (Source^.SourceType in [ST_VSDEC, ST_VCR, ST_CG, ST_LINE]) and
       (Source^.Channel <> nil) and (Source^.Channel^.ID = FChannelID) and
       (Source^.Handles <> nil) then
    begin
      for J := 0 to Source^.Handles.Count - 1 do
      begin
        R := DCSGetOnAirEventID(Source^.Handles[J]^.DCSID, Source^.Handles[J]^.Handle, OnAirEventID, CuedEventID);
        if (R = D_OK) then
        begin
          OnAirItem := GetCueSheetItemByID(OnAirEventID);
          CuedItem  := GetCueSheetItemByID(CuedEventID);

          if (OnAirItem <> nil) then
          begin
            R := DCSGetEventInfo(Source^.Handles[J]^.DCSID, Source^.Handles[J]^.Handle, OnAirEventID, OnAirStartTime, OnAirDurationTC);
            if (R = D_OK) then
            begin

            end;
          end;

          if (CuedItem <> nil) then
          begin
            R := DCSGetEventInfo(Source^.Handles[J]^.DCSID, Source^.Handles[J]^.Handle, CuedItem^.EventID, CuedStartTime, CuedDurationTC);
            if (R = D_OK) then
            begin

            end;
          end;

          if ((OnAirItem <> nil) and (OnAirItem^.EventMode = EM_MAIN)) or
             ((CuedItem <> nil) and (CuedItem^.EventMode = EM_MAIN)) then
          begin
            DCSIsOnAir := True;
          end;
        end;
//        else
//          ShowMessage('not on air');
      end;
    end;
  end;

  IsManualStart := True;
  // If DCS is onair then select start onair type
  if (DCSIsOnAir) then
  begin
    CurrItem  := GetCueSheetItemByID(OnAirEventID);
    CurrItem  := GetParentCueSheetItemByItem(CurrItem);
    CurrIndex := GetCueSheetIndexByItem(CurrItem);

    NextItem  := GetCueSheetItemByID(CuedEventID);
    NextItem  := GetParentCueSheetItemByItem(NextItem);
    NextIndex := GetCueSheetIndexByItem(NextItem);

    frmSelectStartOnAir := TfrmSelectStartOnAir.Create(Self, CurrItem, NextItem);
    try
      frmSelectStartOnAir.ShowModal;
      if (frmSelectStartOnAir.ModalResult = mrOk) then
      begin
        case frmSelectStartOnAir.SelectOnAirStart of
          SO_CURRENT:
          begin
            IsManualStart := False;
          end;
          SO_FIRST:
          begin
            IsManualStart := True;
          end;
        end;
      end
      else
        exit;
    finally
      FreeAndNil(frmSelectStartOnAir);
    end;
  end;

  if (IsManualStart) then
  begin
    // Manual start check
    CurrIndex := GetStartOnAirMainIndex;
    CurrItem  := GetCueSheetItemByIndex(CurrIndex);
    if (CurrItem <> nil) then
    begin
      StartOnAirTime := IncMilliSecond(EventTimeToDateTime(CurrItem^.StartTime), TimecodeToMilliSec(GV_SettingTresholdTime.OnAirLockTime));

      if (StartOnAirTime <= Now) then
      begin
        MessageBeep(MB_ICONQUESTION);
        R := MessageBox(Handle, Pchar(SQStartOnAirByManual), PChar(Application.Title), MB_YESNOCANCEL or MB_ICONQUESTION);
        if (R = IDYES) then
        begin
          SetChannelOnAir(True);

          StartTime := CurrItem^.StartTime;

          CurrItem^.StartMode := SM_MANUAL;
          CurrItem^.StartTime := GetPlusEventTime(DateTimeToEventTime(Now), TimecodeToEventTime(GV_SettingTimeParameter.AutoIncreaseDurationAmount));

//          SetSkipCueSheetItemByIndex(0, CurrIndex - 1);

          FFirstStartTime := CurrItem^.StartTime;

          ResetStartTimeByTime(CurrIndex, StartTime);

          OnAirClearEvents;

          FCueSheetNext := CurrItem;

          OnAirInputEvents(CurrIndex, GV_SettingOption.MaxInputEventCount);

          FLastDisplayNo := GetBeforeMainCountByIndex(CurrIndex);
          DisplayPlayListGrid(CurrIndex);

          FTimeLineMin := 0;
          FTimeLineMax := 0;
          DisplayPlayListTimeLine(CurrIndex);

          FCueSheetNext := CurrItem;
          acgPlaylist.Repaint;
        end
        else if (R = IDNO) then
        begin
          SetChannelOnAir(True);

{          if (FCueSheetNext <> nil) then
            NextItem := FCueSheetNext
          else }
            NextItem := GetMainItemByStartTime(CurrIndex, IncMilliSecond(Now, TimecodeToMilliSec(GV_SettingTresholdTime.OnAirLockTime)));

          if (NextItem <> nil) then
          begin
            NextIndex := GetCueSheetIndexByItem(NextItem);
          end
          else
            NextIndex := FLastCount;// FChannelCueSheet^.CueSheetList.Count - 1;
    //      SetOnAirEvent;

          SetCueSheetItemStatusByIndex(CurrIndex, NextIndex - 1, esSkipped);

          FFirstStartTime := CurrItem^.StartTime;

//          ResetStartTime(CurrIndex);

          OnAirClearEvents;

          FCueSheetNext := NextItem;

          OnAirInputEvents(NextIndex, GV_SettingOption.MaxInputEventCount);

          FLastDisplayNo := GetBeforeMainCountByIndex(CurrIndex);
          DisplayPlayListGrid(CurrIndex);

          FTimeLineMin := 0;
          FTimeLineMax := 0;

          DisplayPlayListTimeLine(CurrIndex);

          FCueSheetNext := NextItem;
          acgPlaylist.Repaint;
        end
        else if (R = IDCANCEL) then
          exit;
      end
      else
      begin
        SetChannelOnAir(True);

        NextItem := GetMainItemByStartTime(CurrIndex, IncMilliSecond(Now, TimecodeToMilliSec(GV_SettingTresholdTime.OnAirLockTime)));
        if (NextItem <> nil) then
        begin
          NextIndex := GetCueSheetIndexByItem(NextItem);
        end
        else
          NextIndex := FLastCount;// FChannelCueSheet^.CueSheetList.Count - 1;

        SetCueSheetItemStatusByIndex(CurrIndex, NextIndex - 1, esSkipped);

        FFirstStartTime := CurrItem^.StartTime;

        OnAirClearEvents;

        FCueSheetNext := NextItem;

        OnAirInputEvents(NextIndex, GV_SettingOption.MaxInputEventCount);

        FLastDisplayNo := GetBeforeMainCountByIndex(NextIndex);
        DisplayPlayListGrid(NextIndex);

        FTimeLineMin := 0;
        FTimeLineMax := 0;

        DisplayPlayListTimeLine(NextIndex);

{        OnAirClearEvents;
        OnAirInputEvents(CurrIndex, GV_SettingOption.MaxInputEventCount);

        FLastDisplayNo := GetBeforeMainCountByIndex(CurrIndex);
        DisplayPlayListGrid(CurrIndex);

        FCueSheetNext := CurrItem;
        acgPlaylist.Repaint; }
      end;

      FTimerThread := TChannelTimerThread.Create(Self);
      FTimerThread.Resume;
    end;
  end
  else
  begin
    if (OnAirItem <> nil) or (CuedItem <> nil) then
    begin
      // Check onair event is changed
      for I := 0 to GV_SourceList.Count - 1 do
      begin
        Source := GV_SourceList[I];
        if (Source <> nil) and
           (Source^.SourceType in [ST_VSDEC, ST_VCR, ST_CG, ST_LINE]) and
           (Source^.Channel <> nil) and (Source^.Channel^.ID = FChannelID) and
           (Source^.Handles <> nil) then
        begin
          for J := 0 to Source^.Handles.Count - 1 do
          begin
            R := DCSGetOnAirEventID(Source^.Handles[J]^.DCSID, Source^.Handles[J]^.Handle, OnAirEventID, CuedEventID);
            if (R = D_OK) then
            begin
              OnAirItem := GetCueSheetItemByID(OnAirEventID);
              CuedItem  := GetCueSheetItemByID(CuedEventID);
              break;
            end;
          end;
        end;
      end;

      if (OnAirItem <> CurrItem) or (CuedItem <> NextItem) then
      begin
        MessageBeep(MB_ICONERROR);
        MessageBox(Handle, PChar(SEChangedOnAirEvent), PChar(Application.Title), MB_OK or MB_ICONERROR);
        exit;
      end;

      SetChannelOnAir(True);

      if (CurrItem <> nil) then
      begin
        StartTime := CurrItem.StartTime;
        SkipIndex := CurrIndex;
      end
      else if (NextItem <> nil) then
      begin
        StartTime := NextItem.StartTime;
        SkipIndex := NextIndex;
      end
      else exit;

      if (CurrItem <> nil) then
      begin
        CurrItem^.StartTime   := OnAirStartTime;
        CurrItem^.DurationTC  := OnAirDurationTC;
        CurrItem^.EventStatus := esOnAir;
      end;

      if (NextItem <> nil) then
      begin
        NextItem^.StartTime   := CuedStartTime;
        NextItem^.DurationTC  := CuedDurationTC;
        NextItem^.EventStatus := esCued;
      end;

//      ShowMessage(EventTimeToDateTimecodeStr(CurrItem^.StartTime));

      SetCueSheetItemStatusByIndex(GetStartOnAirMainIndex, SkipIndex - 1, esSkipped);

      ResetStartTimeByTime(CurrIndex, StartTime);

{      for I := SkipIndex to SkipIndex + GV_SettingOption.MaxInputEventCount do
      begin
        CurrItem := GetCueSheetItemByIndex(I);
        if (CurrItem <> nil) then
        begin
          Source := GetSourceByName(String(CurrItem^.Source));
          if (Source <> nil) and (Source^.Handles <> nil) then
          begin
            for J := 0 to Source^.Handles.Count - 1 do
            begin
              R := DCSGetEventInfo(Source^.Handles[J]^.DCSID, Source^.Handles[J]^.Handle, CurrItem^.EventID, StartTime, DurationTC);
              if (R = D_OK) then
              begin
                CurrItem^.StartTime  := StartTime;
                CurrItem^.DurationTC := DurationTC;
              end;

              R := DCSGetEventStatus(Source^.Handles[J]^.DCSID, Source^.Handles[J]^.Handle, CurrItem^.EventID, EventStatus);
              if (R = D_OK) then
              begin
                CurrItem^.EventStatus := EventStatus;
                break;
              end;
            end;
          end;
        end
        else
          break;
      end;  }

      if (CurrItem <> nil) then
        FFirstStartTime := CurrItem^.StartTime
      else if (NextItem <> nil) then
        FFirstStartTime := NextItem^.StartTime;

      FLastDisplayNo := GetBeforeMainCountByIndex(SkipIndex);
      DisplayPlayListGrid(SkipIndex);

      FTimeLineMin := 0;
      FTimeLineMax := 0;

      DisplayPlayListTimeLine(SkipIndex);

      FCueSheetCurr := OnAirItem;
      FCueSheetNext := CuedItem;

      acgPlaylist.Repaint;

      FTimerThread := TChannelTimerThread.Create(Self);
      FTimerThread.Resume;
    end;
  end;

{  exit;

  Start := DateTimeToEventTime(IncSecond(Now, 20));
  ShowMessage(EventTimeToDateTimecodeStr(Start, True));

  P := FChannelCueSheet^.CueSheetList[1];
  DurEventTime := GetDurEventTime(P^.StartTime, Start);

//  ShowMessage(EventTimeToDateTimecodeStr(P^.StartTime, True));
  ShowMessage(EventTimeToDateTimecodeStr(DurEventTime, True));

//  DurEventTime := GetDurEventTime(P^.StartTime, Start);
//  if (P^.StartTime.T < Start.T) then
    ResetStartTimePlus(0, DurEventTime);
//  else
//    ResetStartTimeMinus(0, DurEventTime);

  for I := 0 to FChannelCueSheet.CueSheetList.Count - 1 do
  begin
    InputEvent(FChannelCueSheet.CueSheetList[I]);
  end;

  FTimerThread := TChannelTimerThread.Create(Self);
  FTimerThread.Resume; }
end;

procedure TfrmChannel.FreezeOnAir;
var
  NextStartTime: TDateTime;
  NextIndex: Integer;

  CurrIndex: Integer;
begin
  if (not FChannelOnAir) then exit;

  if (FCueSheetNext <> nil) then
  begin
    NextStartTime := EventTimeToDateTime(FCueSheetNext^.StartTime);
    if (DateTimeToTimecode(NextStartTime - Now) <= GV_SettingTimeParameter.AutoIncreaseDurationBefore) then
    begin
      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(SFreezeOnAirTimeout), PChar(Application.Title), MB_OK or MB_ICONWARNING);
      exit;
    end;

    FCueSheetNext^.StartMode := SM_MANUAL;
    NextIndex := GetCueSheetIndexByItem(FCueSheetNext);

    acgPlaylist.AllCells[IDX_COL_CUESHEET_START_MODE, NextIndex + CNT_CUESHEET_HEADER] := StartModeNames[FCueSheetNext^.StartMode];

    OnAirInputEvents(NextIndex, 1);

    if (FCueSheetCurr <> nil) then
    begin
      CurrIndex := GetCueSheetIndexByItem(FCueSheetCurr);
      OnAirHoldEvent(CurrIndex);
    end;
  end;
end;

procedure TfrmChannel.FinishOnAir;
var
  R: Integer;
  IsFinish: Boolean;
  StartIndex, EndIndex: Integer;
begin
  if (not FChannelOnAir) then exit;

  StartIndex := -1;
  if (FCueSheetCurr <> nil) then
    StartIndex := GetCueSheetIndexByItem(FCueSheetCurr)
  else if (FCueSheetNext <> nil) then
    StartIndex := GetCueSheetIndexByItem(FCueSheetNext);

  IsFinish := False;
  MessageBeep(MB_ICONQUESTION);
  R := MessageBox(Handle, Pchar(SQFinishtOnAirAndPreserveEvent), PChar(Application.Title), MB_YESNOCANCEL or MB_ICONQUESTION);
  if (R = IDYES) then
  begin
    IsFinish := True;
  end
  else if (R = IDNO) then
  begin
    MessageBeep(MB_ICONQUESTION);
    R := MessageBox(Handle, Pchar(SQFinishtOnAirAndClearEvent), PChar(Application.Title), MB_YESNO or MB_ICONQUESTION);
    if (R = IDYES) then
    begin
      OnAirClearEvents;

      IsFinish := True;
    end
    else exit;
  end
  else if (R = IDCANCEL) then
    exit;

  if (IsFinish) then
  begin
    if (StartIndex >= 0) then
    begin
      EndIndex := StartIndex + GV_SettingOption.MaxInputEventCount;
      if (EndIndex >= FChannelCueSheet^.CueSheetList.Count) then
        EndIndex := FChannelCueSheet^.CueSheetList.Count - 1;

      SetCueSheetItemStatusByIndex(StartIndex, EndIndex, esIdle);
    end;

    FCueSheetCurr   := nil;
    FCueSheetNext   := nil;
    FCueSheetTarget := nil;

    SetChannelOnAir(False);

    if (FTimerThread <> nil) then
    begin
      FTimerThread.Terminate;
      FTimerThread.WaitFor;
      FreeAndNil(FTimerThread);
    end;
  end;
end;

procedure TfrmChannel.AssignNextEvent;
var
  I: Integer;
  StartIndex, EndIndex: Integer;
  SItem, CItem: PCueSheetItem;
  SaveStartTime: TEventTime;
begin
  with acgPlaylist do
  begin
    EndIndex := RealRow - CNT_CUESHEET_HEADER;
    CItem := GetParentCueSheetItemByIndex(EndIndex);
    EndIndex := GetCueSheetIndexByItem(CItem);
    if (CItem <> nil) and
       (CItem^.EventMode = EM_MAIN) and
       (CItem^.EventStatus <= esCued) then
    begin
      StartIndex := GetStartOnAirMainIndex;
      if (StartIndex >= 0) then
      begin
{        for I := StartIndex to EndIndex - 1 do
        begin
          CItem := FChannelCueSheet^.CueSheetList[I];
          if (CItem <> nil) and (CItem^.EventStatus <= esPreroll) then
            CItem^.EventStatus := esSkipped;
        end;  }

        SetCueSheetItemStatusByIndex(StartIndex, EndIndex - 1, esSkipped);
        SItem := GetCueSheetItemByIndex(StartIndex);

        SaveStartTime := CItem^.StartTime;
        CItem^.StartTime := SItem^.StartTime;

        ResetStartTimeByTime(EndIndex, SaveStartTime);

        // If onair then delet & input event
        if (FChannelOnAir) then
        begin
          OnAirDeleteEvents(StartIndex, EndIndex - 1);
          OnAirInputEvents(EndIndex, GV_SettingOption.MaxInputEventCount);
        end;

        FLastDisplayNo := GetBeforeMainCountByIndex(StartIndex);
        DisplayPlayListGrid(StartIndex, 0);
        DisplayPlayListTimeLine(StartIndex);
      end;
      FCueSheetNext := CItem;
      acgPlaylist.Repaint;
    end;
  end;
end;

procedure TfrmChannel.StartNextEventImmediately;
var
  CIndex: Integer;
  CItem: PCueSheetItem;

  Source: PSource;
  I: Integer;
  R: Integer;

  CurrIndex: Integer;
  CurrItem: PCueSheetItem;

  NextIndex: Integer;
  NextItem: PCueSheetItem;

  StartTime: TEventTime;
  DurationTC: TTimecode;

  SaveStartTime: TEventTime;
  SaveDurationTC: TTimecode;

//  OnAirIndex: Integer;
//  OnAirItem: PCueSheetItem;
//  OnAirEventID: TEventID;     // DCS onair event id
//  OnAirStartTime: TEventTime; // DCS onair event start time
//  OnAirDurationTC: TTimecode; // DCS onair event duration tc
begin

{ShowMessage(EventTimeToDateTimecodeStr(GetEventEndTime(FChannelCueSheet^.CueSheetList[7]^.StartTime, FChannelCueSheet^.CueSheetList[7]^.DurationTC)));
ShowMessage(EventTimeToDateTimecodeStr(FChannelCueSheet^.CueSheetList[8]^.StartTime));
exit;  }
  if (not FChannelOnAir) then exit;

  if (FCueSheetNext <> nil) and (FCueSheetNext^.EventMode = EM_MAIN) then
  begin
{    // Delete the event after the next event before taking
    CItem := GetNextMainItemByItem(FCueSheetNext);
    if (CItem <> nil) then
    begin
      CIndex := GetCueSheetIndexByItem(CItem);
      for I := CIndex to CIndex + GV_SettingOption.MaxInputEventCount do
      begin
        CItem := GetCueSheetItemByIndex(I);
        if (CItem <> nil) then
          DeleteEvent(CItem)
        else
          break;
      end;
    end; }

    CurrItem  := FCueSheetCurr;
    CurrIndex := GetCueSheetIndexByItem(CurrItem);

    NextItem  := FCueSheetNext;
    NextIndex := GetCueSheetIndexByItem(NextItem);

//    OnAirEventID := FCueSheetNext^.EventID;
    OnAirTakeEvent(GetCueSheetIndexByItem(NextItem));

    Sleep(TimecodeToMilliSec(GV_SettingTimeParameter.StandardTimeCorrection));

    if (CurrItem <> nil) then
    begin
      Source := GetSourceByName(String(CurrItem^.Source));
      if (Source <> nil) and (Source^.Handles <> nil) then
      begin
        for I := 0 to Source^.Handles.Count - 1 do
        begin
          R := DCSGetEventInfo(Source^.Handles[I]^.DCSID, Source^.Handles[I]^.Handle, CurrItem^.EventID, StartTime, DurationTC);
          if (R = D_OK) then
          begin
            CurrItem^.StartTime  := StartTime;
            CurrItem^.DurationTC := DurationTC;

            ResetChildItems(CurrIndex);

            FLastDisplayNo := GetBeforeMainCountByIndex(CurrIndex);
            DisplayPlayListGrid(CurrIndex);
            DisplayPlayListTimeLine(CurrIndex);

            break;
          end;
        end;
      end;
    end;

    Source := GetSourceByName(String(NextItem^.Source));
    if (Source <> nil) and (Source^.Handles <> nil) then
    begin
      for I := 0 to Source^.Handles.Count - 1 do
      begin
        R := DCSGetEventInfo(Source^.Handles[I]^.DCSID, Source^.Handles[I]^.Handle, NextItem^.EventID, StartTime, DurationTC);
        if (R = D_OK) then
        begin
          SaveStartTime  := NextItem^.StartTime;
          SaveDurationTC := NextItem^.DurationTC;

          NextItem^.StartTime  := StartTime;
          NextItem^.DurationTC := DurationTC;

          ResetStartTimeByTime(NextIndex, SaveStartTime, SaveDurationTC);

// Aleady DCS reset start time
//            OnAirInputEvents(GetNextMainIndexByIndex(NextIndex), GV_SettingOption.MaxInputEventCount);

          FLastDisplayNo := GetBeforeMainCountByIndex(NextIndex);
          DisplayPlayListGrid(NextIndex);
          DisplayPlayListTimeLine(NextIndex);

          break;
        end;
      end;
    end;

            OnAirInputEvents(NextIndex, GV_SettingOption.MaxInputEventCount);
  end;
end;

procedure TfrmChannel.IncSecondsOnAirEvent(ASeconds: Integer);
var
  CurrIndex: Integer;
  SaveDurationTC: TTimecode;
begin
  if (not FChannelOnAir) then exit;

  if (ASeconds = 0) then exit;
  if (FCueSheetCurr = nil) then exit;

  CurrIndex := GetCueSheetIndexByItem(FCueSheetCurr);

  SaveDurationTC := FCueSheetCurr^.DurationTC;

  if (ASeconds > 0) then
    FCueSheetCurr^.DurationTC := GetPlusTimecode(FCueSheetCurr^.DurationTC, SecondToTimeCode(ASeconds))
  else
    FCueSheetCurr^.DurationTC := GetMinusTimecode(FCueSheetCurr^.DurationTC, SecondToTimeCode(-ASeconds));

  ResetStartTimeByTime(CurrIndex, SaveDurationTC);

  OnAirChangeDurationEvent(FCueSheetCurr^.DurationTC);

//  OnAirInputEvents(CurrIndex, GV_SettingOption.MaxInputEventCount);

  FLastDisplayNo := GetBeforeMainCountByIndex(CurrIndex);
  DisplayPlayListGrid(CurrIndex);
  DisplayPlayListTimeLine(CurrIndex);
end;

procedure TfrmChannel.AssignTargetEvent;
var
  SelectIndex: Integer;
  ParentItem: PCueSheetItem;
//  ParentRow: Integer;
begin
  if (not FChannelOnAir) then exit;

  with acgPlaylist do
  begin
    SelectIndex := RealRow - CNT_CUESHEET_HEADER;
    ParentItem := GetParentCueSheetItemByIndex(SelectIndex);
    if (ParentItem <> nil) then
    begin
      FCueSheetTarget := ParentItem;

{      ParentRow := GetCueSheetIndexByItem(FCueSheetTarget) + CNT_CUESHEET_HEADER;
      lblRemainingTargetEvent.Caption := Format('Remainig target (%s)', [RealCells[IDX_COL_CUESHEET_NO, ParentRow]]); }

      Repaint;
    end;
  end;
end;

procedure TfrmChannel.SetEventStatus(AEventID: TEventID; AStatus: TEventStatus);
var
  CItem: PCueSheetItem;
  Index: Integer;
  RRow: Integer;
begin
  if (not FChannelOnAir) then exit;

  CItem := GetCueSheetItemByID(AEventID);
  if (CItem <> nil) then
  begin
    CItem^.EventStatus := AStatus;
    if (CItem^.EventMode = EM_MAIN) then
    begin
      if (AStatus = esOnAir) then
      begin
        FCueSheetCurr := CItem;
        FCueSheetNext := GetNextMainItemByItem(CItem);

        if (FCueSheetTarget = FCueSheetCurr) then
          FCueSheetTarget := nil;
      end;

{      if (AStatus = esCued) then
        FCueSheetNext := CItem;  }

      if (FCueSheetNext = nil) and (AStatus in [esLoaded]) then
        FCueSheetNext := CItem;

      if (FCueSheetCurr <> nil) and (AStatus = esDone) then
        FCueSheetCurr := nil;

      // If onair then input next events
      if (FChannelOnAir) and (AStatus in [esError, esSkipped, esDone]) then
      begin
        Index := GetNextOnAirMainIndex;
        OnAirInputEvents(Index, 1);
      end;
    end;

    Index := GetCueSheetIndexByItem(CItem);
    if (Index >= 0) then
    begin
      PostMessage(Handle, WM_UPDATE_EVENTSTATUS, Index, 0);
      exit;
      with acgPlaylist do
      begin
//        if (AStatus > esLoaded) then
//          SendMessage(Handle, WM_CHAR, 27, 0);

        RRow := Index + CNT_CUESHEET_HEADER;
        AllCells[IDX_COL_CUESHEET_EVENT_STATUS, RRow] := EventStatusNames[AStatus];
        RepaintRow(RRow);

{        case AStatus of
          esCueing..esPreroll:
          begin
            RowColor[RRow]     := COLOR_BK_EVENTSTATUS_CUED;
            RowFontColor[RRow] := COLOR_TX_EVENTSTATUS_CUED;
          end;
          esOnAir:
          begin
            RowColor[RRow]     := COLOR_BK_EVENTSTATUS_ONAIR;
            RowFontColor[RRow] := COLOR_TX_EVENTSTATUS_ONAIR;
          end;
          esSkipped,
          esFinish..esDone:
          begin
            RowColor[RRow]     := COLOR_BK_EVENTSTATUS_DONE;
            RowFontColor[RRow] := COLOR_TX_EVENTSTATUS_DONE;
          end;
          esError:
          begin
            RowColor[RRow]     := COLOR_BK_EVENTSTATUS_ERROR;
            RowFontColor[RRow] := COLOR_TX_EVENTSTATUS_ERROR;
          end;
          else
          begin
            RowColor[RRow]     := COLOR_BK_EVENTSTATUS_NORMAL;
            RowFontColor[RRow] := COLOR_TX_EVENTSTATUS_NORMAL;
          end;
        end; }
      end;

//      if (Index = FLastCount - 1) then
//        FCueSheetNext := nil;
    end;
  end;
end;

procedure TfrmChannel.ClearCueSheetList;
var
  I: Integer;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := FChannelCueSheet^.CueSheetList.Count - 1 downto 0 do
    Dispose(FChannelCueSheet^.CueSheetList[I]);

  FChannelCueSheet^.CueSheetList.Clear;
end;

procedure TfrmChannel.ClearPlayListGrid;
begin
  with acgPlaylist do
  begin
    BeginUpdate;
    try
      MouseActions.DisjunctRowSelect := False;
      ClearRowSelect;
      RemoveAllNodes;
      RowCount := CNT_CUESHEET_HEADER + CNT_CUESHEET_FOOTER;
      MouseActions.DisjunctRowSelect := True;
      SelectRows(CNT_CUESHEET_HEADER, 1);
      Row := CNT_CUESHEET_HEADER;
    finally
      acgPlaylist.EndUpdate;
    end;
  end;
end;

procedure TfrmChannel.ClearPlayListTimeLine;
var
  I: Integer;
begin
  inherited;
  with wmtlPlaylist do
  begin
    DataGroup.Visible := False;
    for I := 0 to DataGroupProperty.Count - 1 do
    begin
      DataCompositions[I].Tracks.BeginUpdate;
      try
        DataCompositions[I].Tracks.Clear;
      finally
        DataCompositions[I].Tracks.EndUpdate;
      end;
    end;
    DataGroup.Visible := True;

    Repaint;
  end;
end;

procedure TfrmChannel.FormCreate(Sender: TObject);
begin
  inherited;
  Initialize;

  NewPlayList;

  OpenPlayList('D:\User Data\Git\APC\SEC\Win64\Debug\CueSheet\Test_20171118_정규편성_1.xml');
end;

procedure TfrmChannel.FormDestroy(Sender: TObject);
begin
  inherited;
  Finalize;
end;

procedure TfrmChannel.InitializePlayListGrid;
var
  I, J: Integer;
  Column: TGridColumnItem;
  E: TEventMode;
  S: TStartMode;
  IType: TInputType;
  TrType: TTRType;
  TrRate: TTRRate;
begin
  with acgPlaylist do
  begin
    BeginUpdate;
    try
      FixedRows := CNT_CUESHEET_HEADER;
      RowCount  := CNT_CUESHEET_HEADER + CNT_CUESHEET_FOOTER;
      ColCount  := CNT_CUESHEET_COLUMNS;

      Columns.BeginUpdate;
      try
        Columns.Clear;
        for I := 0 to CNT_CUESHEET_COLUMNS - 1 do
        begin
          Column := Columns.Add;
          with Column do
          begin
            HeaderFont.Assign(acgPlaylist.FixedFont);
            Font.Assign(acgPlaylist.Font);

            // Column : Group
            if (I = IDX_COL_CUESHEET_GROUP) then
            begin
              BorderPen.Color := GridLineColor;
              Borders  := [cbRight];
              Header   := NAM_COL_CUESHEET_GROUP;
              ReadOnly := True;
              Width    := WIDTH_COL_CUESHEET_GROUP;
            end
            // Column : No
            else if (I = IDX_COL_CUESHEET_NO) then
            begin
              Alignment := taLeftJustify;
              Borders   := [];
              Header    := NAM_COL_CUESHEET_NO;
              HeaderAlignment := taCenter;
              ReadOnly  := True;
              Width     := WIDTH_COL_CUESHEET_NO;
            end
            // Column Dropdown List : Event Mode
            else if (I = IDX_COL_CUESHEET_EVENT_MODE) then
            begin
              Alignment := taCenter;
              Header    := NAM_COL_CUESHEET_EVENT_MODE;
              ReadOnly  := True;
              Width     := WIDTH_COL_CUESHEET_EVENT_MODE;

  {            Editor := edComboList;
              for E := EM_MAIN to EM_JOIN do
              begin
                ComboItems.AddObject(EventModeNames[E], TObject(E));
              end; }
            end
            // Column : Event Status
            else if (I = IDX_COL_CUESHEET_EVENT_STATUS) then
            begin
              Alignment := taLeftJustify;
              Header    := NAM_COL_CUESHEET_EVENT_STATUS;
              ReadOnly  := True;
              Width     := WIDTH_COL_CUESHEET_EVENT_STATUS;
            end
            // Column Dropdown List : Start Mode
            else if (I = IDX_COL_CUESHEET_START_MODE) then
            begin
              Header := NAM_COL_CUESHEET_START_MODE;
              Width  := WIDTH_COL_CUESHEET_START_MODE;
              Editor := edComboList;
              for S := SM_ABSOLUTE to SM_SUBEND do
              begin
                ComboItems.AddObject(StartModeNames[S], TObject(S));
              end;
            end
            // Column : Start Date
            else if (I = IDX_COL_CUESHEET_START_DATE) then
            begin
              Header := NAM_COL_CUESHEET_START_DATE;
              Width  := WIDTH_COL_CUESHEET_START_DATE;
  //            Editor := edComboList;
              Editor := edDateEdit;
            end
            // Column : Start Time
            else if (I = IDX_COL_CUESHEET_START_TIME) then
            begin
              Header := NAM_COL_CUESHEET_START_TIME;
              Width  := WIDTH_COL_CUESHEET_START_TIME;
              Editor := edNormal;
  //            EditMask := '!0000-!90-90 !90:00:00:00;1;';
              EditMask := '!99:99:99:99;1; ';
            end
            // Column Dropdown List : Input Type
            else if (I = IDX_COL_CUESHEET_INPUT) then
            begin
              Header := NAM_COL_CUESHEET_INPUT;
              Width  := WIDTH_COL_CUESHEET_INPUT;
              Editor := edComboList;
              for IType := IT_MAIN to IT_AMIXER2 do
              begin
                ComboItems.AddObject(InputTypeNames[IType], TObject(IType));
              end;
            end
            // Column Dropdown List : Output Type
            else if (I = IDX_COL_CUESHEET_OUTPUT) then
            begin
              Header := NAM_COL_CUESHEET_OUTPUT;
              Width  := WIDTH_COL_CUESHEET_OUTPUT;
              Editor := edComboList;
              ComboItems.Clear;
            end
            // Column : Title
            else if (I = IDX_COL_CUESHEET_TITLE) then
            begin
              Header  := NAM_COL_CUESHEET_TITLE;
              Width  := WIDTH_COL_CUESHEET_TITLE;
              Editor := edNormal;
              EditMask := '';
            end
            // Column : Sub Title
            else if (I = IDX_COL_CUESHEET_SUB_TITLE) then
            begin
              Header  := NAM_COL_CUESHEET_SUB_TITLE;
              Width  := WIDTH_COL_CUESHEET_SUB_TITLE;
              Editor := edNormal;
            end
            // Column Dropdown List : Source
            else if (I = IDX_COL_CUESHEET_SOURCE) then
            begin
              Header := NAM_COL_CUESHEET_SOURCE;
              Width  := WIDTH_COL_CUESHEET_SOURCE;
              Editor := edComboList;
              for J := 0 to GV_SourceList.Count - 1 do
              begin
                if (not (GV_SourceList[J]^.SourceType in [ST_ROUTER, ST_MCS])) then
  //                if (GV_SourceList[I]^.DCS <> nil) and (GV_SourceList[I]^.DCS^.Main) then
  //                 ComboItems.AddObject(String(GV_SourceList[I]^.Name), TObject(GV_SourceList[I]^.Handle));
                    ComboItems.Add(String(GV_SourceList[J]^.Name));
              end;
            end
            // Column Dropdown List : Media ID
            else if (I = IDX_COL_CUESHEET_MEDIA_ID) then
            begin
              Header := NAM_COL_CUESHEET_MEDIA_ID;
              Width  := WIDTH_COL_CUESHEET_MEDIA_ID;
              Editor := edEditBtn;
            end
            // Column : Duration TC
            else if (I = IDX_COL_CUESHEET_DURATON) then
            begin
              Header := NAM_COL_CUESHEET_DURATON;
              Width  := WIDTH_COL_CUESHEET_DURATON;
              Editor := edNormal;
              EditMask := '!99:99:99:99;1; ';
            end
            // Column : In TC
            else if (I = IDX_COL_CUESHEET_IN_TC) then
            begin
              Header := NAM_COL_CUESHEET_IN_TC;
              Width  := WIDTH_COL_CUESHEET_IN_TC;
              Editor := edNormal;
              EditMask := '!99:99:99:99;1; ';
            end
            // Column : Out TC
            else if (I = IDX_COL_CUESHEET_OUT_TC) then
            begin
              Header := NAM_COL_CUESHEET_OUT_TC;
              Width  := WIDTH_COL_CUESHEET_OUT_TC;
              Editor := edNormal;
              EditMask := '!99:99:99:99;1; ';
            end
            // Column Dropdown List : Transition Type
            else if (I = IDX_COL_CUESHEET_TR_TYPE) then
            begin
              Header := NAM_COL_CUESHEET_TR_TYPE;
              Width  := WIDTH_COL_CUESHEET_TR_TYPE;
              Editor := edComboList;
              for TrType := TT_FADE to TT_MIX do
              begin
                ComboItems.AddObject(TRTypeNames[TrType], TObject(TrType));
              end;
            end
            // Column Dropdown List : Transition Rate
            else if (I = IDX_COL_CUESHEET_TR_RATE) then
            begin
              Header := NAM_COL_CUESHEET_TR_RATE;
              Width  := WIDTH_COL_CUESHEET_TR_RATE;
              Editor := edComboList;
              for TrRate := TR_CUT to TR_SLOW do
              begin
                ComboItems.AddObject(TRRateNames[TrRate], TObject(TrRate));
              end;
            end
            // Column Dropdown List : Program Type
            else if (I = IDX_COL_CUESHEET_PROGRAM_TYPE) then
            begin
              Header := NAM_COL_CUESHEET_PROGRAM_TYPE;
              Width  := WIDTH_COL_CUESHEET_PROGRAM_TYPE;
              Editor := edComboList;
              for J := 0 to GV_ProgramTypeList.Count - 1 do
              begin
                ComboItems.AddObject(GV_ProgramTypeList[J]^.Name, TObject(GV_ProgramTypeList[J]^.Code));
              end;
            end
            // Column : Notes
            else if (I = IDX_COL_CUESHEET_NOTES) then
            begin
              Header := NAM_COL_CUESHEET_NOTES;
              Width  := WIDTH_COL_CUESHEET_NOTES;
              Editor := edNormal;
              EditMask := '';
            end;
          end;
        end;

//        DisplayPlayListGrid;
      finally
        Columns.EndUpdate;
      end;
    finally
      EndUpdate;
    end;
  end;
end;

procedure TfrmChannel.InitializePlayListTimeLine;
begin
  FTimeLineDaysPerFrames := Round(SecsPerDay * wmtlPlaylist.TimeZoneProperty.FrameRate);

  FTimeLineMin := 0;
  FTimeLineMax := 0;

  with wmtlPlaylist do
  begin
    with TimeZoneProperty do
    begin
      FrameRate := FrameRate29_97;
      FrameSkip := 600;
      FrameStep := 15;
      RailBarVisible := False;
    end;
    VideoGroupProperty.Count := 0;
    AudioGroupProperty.Count := 0;

    DataGroupProperty.Count  := 2;
  end;
end;

procedure TfrmChannel.PlaylistFileParsing(AFileName: String);
var
  XmlParser: TXmlParser;
  CueSheetItem: PCueSheetItem;

  procedure ScanElement;
  var
    I: Integer;
    ContentStr: AnsiString;
  begin
    while XmlParser.Scan do
    begin
      case XmlParser.CurPartType OF
        ptStartTag,
        ptEmptyTag:
        begin
          if (XmlParser.CurAttr.Count > 0) then
          begin
            if (XmlParser.CurName = XML_EVENT) then
            begin
              CueSheetItem := New(PCueSheetItem);
              FillChar(CueSheetItem^, SizeOf(TCueSheetItem), #0);
              with CueSheetItem^ do
              begin
                EventID.ChannelID := FChannelID;
                StrCopy(EventID.OnAirDate, FChannelCueSheet^.OnairDate);
                EventID.OnAirFlag := FChannelCueSheet^.OnairFlag;
                EventID.OnAirNo   := FChannelCueSheet^.OnairNo;

                for I := 0 to XmlParser.CurAttr.Count - 1 do
                begin
                  if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_SERIAL_NO) then
                  begin
                    // Play list를 처음 open 했을 경우에만 큐시트의 SerialNo로 저장
                    if (FLastCount <= 0) then
                      EventID.SerialNo := StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0)
                    else
                    begin
                      FLastSerialNo := FLastSerialNo + 1;
                      EventID.SerialNo := FLastSerialNo;
                    end;
                  end
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_GROUP_NO) then
                  begin
                    // Play list를 처음 open 했을 경우에만 큐시트의 GroupNo로 저장
                    GroupNo := StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0);
                    if (FLastCount > 0) then
                    begin
                      GroupNo := GroupNo + FLastGroupNo;
                    end;
                  end
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_EVENT_MODE) then
                    EventMode := TEventMode(StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0))
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_START_MODE) then
                    StartMode := TStartMode(StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0))
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_START_TIME) then
                    StartTime := DateTimecodeStrToEventTime(TNvpNode(XmlParser.CurAttr[I]).Value)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_INPUT) then
                    Input := TInputType(StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0))
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_OUTPUT) then
                    Output := StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_TITLE) then
                    StrPCopy(Title, TNvpNode(XmlParser.CurAttr[I]).Value)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_SUB_TITLE) then
                    StrPCopy(SubTitle, TNvpNode(XmlParser.CurAttr[I]).Value)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_SOURCE) then
                    StrPCopy(Source, TNvpNode(XmlParser.CurAttr[I]).Value)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_MEDIA_ID) then
                    StrPCopy(MediaId, TNvpNode(XmlParser.CurAttr[I]).Value)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_DURATION_TC) then
                    DurationTC := StringToTimecode(TNvpNode(XmlParser.CurAttr[I]).Value)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_IN_TC) then
                    InTC := StringToTimecode(TNvpNode(XmlParser.CurAttr[I]).Value)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_OUT_TC) then
                    OutTC := StringToTimecode(TNvpNode(XmlParser.CurAttr[I]).Value)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_TRANSITION_TYPE) then
                    TransitionType := TTRType(StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0))
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_TRANSITION_RATE) then
                    TransitionRate := TTRRate(StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0))
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_PROGRAM_TYPE) then
                    ProgramType := StrToIntDef(TNvpNode(XmlParser.CurAttr[I]).Value, 0)
                  else if (TNvpNode(XmlParser.CurAttr[I]).Name = XML_NOTES) then
                    StrPCopy(Notes, TNvpNode(XmlParser.CurAttr[I]).Value);
                end;
              end;
              FChannelCueSheet^.CueSheetList.Add(CueSheetItem);
            end;

          end;

          if (XmlParser.CurPartType = ptStartTag) then  // Recursion
            ScanElement;
        end;
        ptEndTag: break;
        ptContent,
        ptCData:
        begin
          // Play list를 처음 open 했을 경우에만 채널의 큐시트 정보를 저장
          if (FLastCount <= 0) then
          begin
            if (Length(XmlParser.CurContent) > 40) then
              ContentStr := Copy (XmlParser.CurContent, 1, 40) + #133
            else
              ContentStr := XmlParser.CurContent;

  {          if (XmlParser.CurName = XML_CHANNEL_ID) then
            begin
              FChannelCueSheet^.ChannelId := StrToIntDef(ContentStr, 0);
            end
            else} if (XmlParser.CurName = XML_ONAIR_DATE) then
            begin
              StrPCopy(FChannelCueSheet^.OnairDate, ContentStr);
            end
            else if (XmlParser.CurName = XML_ONAIR_FLAG) then
            begin
              if (Length(ContentStr) > 0) then
                FChannelCueSheet^.OnairFlag := TOnAirFlagType(Ord(ContentStr[1]))
              else
                FChannelCueSheet^.OnairFlag := FT_REGULAR;
            end
            else if (XmlParser.CurName = XML_ONAIR_NO) then
            begin
              FChannelCueSheet^.OnairNo := StrToIntDef(ContentStr, 0);
            end;
          end;
        end;
      end;
    end;
  end;

begin
  XmlParser := TXmlParser.Create;
  try
    XmlParser.LoadFromFile(AFileName);
    XmlParser.Normalize := True;
    XmlParser.StartScan;

    ScanElement;
  finally
    FreeAndNil(XmlParser);
  end;
end;

{ TChannelTimerThread }

constructor TChannelTimerThread.Create(AChannelForm: TfrmChannel);
begin
  FchannelForm := AChannelForm;

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TChannelTimerThread.Destroy;
begin
  inherited Destroy;
end;

procedure TChannelTimerThread.Execute;
var
  WaitList: array[0..1] of THandle;
begin
  WaitList[0] := GV_TimerExecuteEvent;
  WaitList[1] := GV_TimerCancelEvent;
  while not Terminated do
  begin
    if (WaitForMultipleObjects(2, @WaitList[0], False, INFINITE) <> WAIT_OBJECT_0) then
      break; // Terminate thread when GV_TimerCancelEvent is signaled

    PostMessage(FChannelForm.Handle, WM_UPDATE_CHANNEL_TIME, 0, 0);
  end;
end;

end.
