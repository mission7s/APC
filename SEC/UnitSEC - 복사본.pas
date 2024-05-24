unit UnitSEC;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitNormalForm, Vcl.Imaging.pngimage,
  WMTools, WMControls, Vcl.ExtCtrls, System.Actions, Vcl.ActnList, System.SyncObjs,
  Vcl.XPStyleActnCtrls, Vcl.ActnMan, Vcl.ToolWin, Vcl.ActnCtrls, Vcl.ActnMenus,
  Vcl.ActnColorMaps, AdvOfficePager, AdvOfficePagerStylers, Vcl.StdCtrls,
  AdvSplitter,
  UnitCommons, UnitDCSDLL, UnitMCCDLL, UnitConsts, UnitUDPIn, UnitUDPOut, UnitTypeConvert,
  UnitChannelTimeline, UnitChannel, UnitDevice, AdvUtil, Vcl.Grids, AdvObj, BaseGrid, AdvGrid,
  AdvCGrid, AdvScrollBox;

type
  TTimerThread = class;
  TCrossCheckThread = class;
  TMCCCheckThread = class;

  TfrmSEC = class(TfrmNormal)
    aopStyler: TAdvOfficePagerOfficeStyler;
    XPColorMap: TXPColorMap;
    ActionMainMenuBar1: TActionMainMenuBar;
    actManager: TActionManager;
    actHelp: TAction;
    actAbout: TAction;
    actFileClose: TAction;
    actViewTruePeak: TAction;
    actViewMomentary: TAction;
    actViewShortTerm: TAction;
    actViewIntegrate: TAction;
    actSettingsDeviceAndMeasure: TAction;
    actSettingsChannelRouting: TAction;
    actEventInsertMainEvent: TAction;
    lblCurrentTime: TLabel;
    aoPagerMain: TAdvOfficePager;
    aopAllChannel: TAdvOfficePage;
    aoPagerDevice: TAdvOfficePager;
    aopDeviceMedia: TAdvOfficePage;
    aopDeviceCG: TAdvOfficePage;
    aopDeviceInput: TAdvOfficePage;
    AdvSplitter1: TAdvSplitter;
    aopDevice: TAdvOfficePage;
    actFileNewPlaylist: TAction;
    actFileOpenPlaylist: TAction;
    actControlStartOnair: TAction;
    actEventInsertJoinEvent: TAction;
    actEventInsertSubEvent: TAction;
    actEventInsertCommentEvent: TAction;
    actFileSavePlaylist: TAction;
    actEventUpdateEvent: TAction;
    actEventDeleteEvent: TAction;
    actFileOpenAddPlayList: TAction;
    actFileSaveAsPlayList: TAction;
    actControlFreezeOnAir: TAction;
    actControlFinishOnAir: TAction;
    actControlAssignNextEvent: TAction;
    actControlStartNextEventImmediately: TAction;
    actEditCopyEvent: TAction;
    actEditCutEvent: TAction;
    actEditPasteEvent: TAction;
    actEditClearClipboard: TAction;
    actEditFindEvent: TAction;
    actEditReplaceEvent: TAction;
    actEventInspectEvent: TAction;
    actControlIncrease1Second: TAction;
    actControlDecrease1Second: TAction;
    actControlAssignTargetEvent: TAction;
    actEditTimelineZoomIn: TAction;
    actEditTimelineZoomOut: TAction;
    actEventInsertProgram: TAction;
    aopAllChannels: TAdvScrollBox;
    actAllChannelsTimelineZoomIn: TAction;
    actAllChannelsTimelineZoomOut: TAction;
    actGotoCurrentEvent: TAction;
    pnlActive: TWMPanel;
    pnlSECName: TWMPanel;
    Label6: TLabel;
    Label7: TLabel;
    procedure actFileCloseExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure actFileNewPlaylistExecute(Sender: TObject);
    procedure actFileOpenPlaylistExecute(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actControlStartOnairExecute(Sender: TObject);
    procedure actEventInsertMainEventExecute(Sender: TObject);
    procedure actEventInsertJoinEventExecute(Sender: TObject);
    procedure actEventInsertSubEventExecute(Sender: TObject);
    procedure actEventInsertCommentEventExecute(Sender: TObject);
    procedure aoPagerMainChange(Sender: TObject);
    procedure actEventUpdateEventExecute(Sender: TObject);
    procedure actEventDeleteEventExecute(Sender: TObject);
    procedure actFileOpenAddPlayListExecute(Sender: TObject);
    procedure actFileSavePlaylistExecute(Sender: TObject);
    procedure actFileSaveAsPlayListExecute(Sender: TObject);
    procedure actControlFreezeOnAirExecute(Sender: TObject);
    procedure actControlFinishOnAirExecute(Sender: TObject);
    procedure actControlAssignNextEventExecute(Sender: TObject);
    procedure actControlStartNextEventImmediatelyExecute(Sender: TObject);
    procedure actEditCutEventExecute(Sender: TObject);
    procedure actEditClearClipboardExecute(Sender: TObject);
    procedure actEditPasteEventExecute(Sender: TObject);
    procedure actEditCopyEventExecute(Sender: TObject);
    procedure actControlIncrease1SecondExecute(Sender: TObject);
    procedure actControlDecrease1SecondExecute(Sender: TObject);
    procedure actControlAssignTargetEventExecute(Sender: TObject);
    procedure actEditTimelineZoomInExecute(Sender: TObject);
    procedure actEditTimelineZoomOutExecute(Sender: TObject);
    procedure lblCurrentTimeMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure actEventInspectEventExecute(Sender: TObject);
    procedure actEventInsertProgramExecute(Sender: TObject);
    procedure aopAllChannelsResize(Sender: TObject);
    procedure actAllChannelsTimelineZoomInExecute(Sender: TObject);
    procedure actAllChannelsTimelineZoomOutExecute(Sender: TObject);
    procedure actGotoCurrentEventExecute(Sender: TObject);
  private
    { Private declarations }
    FTimerThread: TTimerThread;
    FMCCCheckThread: TMCCCheckThread;
    FCrossCheckThread: TCrossCheckThread;

    FChannelTimelineHeight: Integer;

    FOldScrollBoxWndProc: TWndMethod;
    procedure NewScrollBoxWndProc(var M: TMessage);

    procedure Initialize;
    procedure Finalize;

    procedure InitializeAllChannelPage;
    procedure FinalizeAllChannelPage;

    procedure InitializeChannelPage;
    procedure FinalizeChannelPage;

    procedure InitializeDevicePage;
    procedure FinalizeDevicePage;

    procedure DisplayActivate;

    procedure DeviceOpen;
    procedure DeviceClose;

    procedure TimerThreadEvent(Sender: TObject);

    procedure InsertEvent(AEventMode: TEventMode);
    procedure UpdateEvent;
    procedure DeleteEvent;

    procedure InspectEvent;

    procedure CutToClipboardEvent;
    procedure CopyToClipboardEvent;
    procedure PasteFromClipboardEvent;

    procedure ClearClipboardEvent;

    procedure TimelineZoomIn;
    procedure TimelineZoomOut;

    procedure MCCCheck;

    procedure ScreenOnActiveControlChange(Sender: TObject);
  protected
    procedure WndProc(var Message: TMessage); override;
  public
    { Public declarations }

    function GetChannelFormByID(AChannelID: Byte): TfrmChannel;
    function GetChannelTimelineFormByID(AChannelID: Byte): TfrmChannelTimeline;

    // 0X00 System Control
    function SECIsAlive(var AIsAlive: Boolean): Integer;                    // 0X00.00
    function SECIsMain(var AIsMain: Boolean): Integer;                      // 0X00.01
    function SECSetAlive(ABuffer: AnsiString): Integer;                     // 0X00.10
    function SECSetMain(ABuffer: AnsiString): Integer;                      // 0X00.11
  end;

  TTimerThread = class(TThread)
  private
    FTimerEnabledFlag: THandle;
    FCancelFlag: THandle;
    FTimerProc: TNotifyEvent; // method to call
    FInterval: Cardinal;
    procedure SetEnabled(AEnable: Boolean);
    function GetEnabled: Boolean;
    procedure SetInterval(AInterval: Cardinal);
  protected
    procedure Execute; override;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Terminate;
    property Enabled: Boolean read GetEnabled write SetEnabled;
    property Interval: Cardinal read FInterval write SetInterval;
    // Note: OnTimerEvent is executed in TTimerThread thread
    property OnTimerEvent: TNotifyEvent read FTimerProc write FTimerProc;
  end;

  TMCCCheckThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FCloseEvent: THandle;
  protected
    procedure DoMCCCheck;
    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;
  end;

  TCrossCheckThread = class(TThread)
  private
    { Private declarations }
    FSECForm: TfrmSEC;

    FUDPIn: TUDPIn;
    FUDPOut: TUDPOut;

    FIsCommand: Boolean;
    FCMD1, FCMD2: Byte;

    FSyncMsgEvent: THandle;
    FRecvBuffer: AnsiString;
    FRecvData: AnsiString;
    FLastResult: Integer;

    FNumCrossCheck: Word;

    function SECGetIsAlive(AHostIP: AnsiString; var AIsAlive: Boolean): Integer;
    function SECGetIsMain(AHostIP: AnsiString; var AIsMain: Boolean): Integer;

    function SECSetAllive(AHostIP: AnsiString; ASECID: Word; AAlive: Boolean): Integer;
    function SECSetMain(AHostIP: AnsiString; AMainSECID: Word): Integer;

    function SendCommand(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
    function SendResponse(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
    function SendAck(AHostIP: AnsiString; APort: Word): Integer;
    function SendNak(AHostIP: AnsiString; APort: Word; ANakError: Byte): Integer;
    function SendError(AHostIP: AnsiString; APort: Word; AErrorCode: Integer): Integer;

    procedure UDPInRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);

    procedure DoMainCheck;
    procedure DoCrossCheck;
    procedure DoSubAliveCheck;
  protected
    procedure Execute; override;
  public
    constructor Create(ASECForm: TfrmSEC);
    destructor Destroy; override;

    function TransmitCommand(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
    function TransmitResponse(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
    function TransmitAck(AHostIP: AnsiString; APort: Word): Integer;
    function TransmitNak(AHostIP: AnsiString; APort: Word; ANakError: Byte): Integer;
    function TransmitError(AHostIP: AnsiString; APort: Word; AErrorCode: Integer): Integer;
  end;

var
  frmSEC: TfrmSEC;
  procedure DeviceStatusNotify(ADCSIP: PChar; ADeviceHandle: TDeviceHandle; AStatus: TDeviceStatus); stdcall;
  procedure EventStatusNotify(ADCSIP: PChar; AEventID: TEventID; AStatus: TEventStatus); stdcall;
  procedure EventOverallNotify(ADCSIP: PChar; ADeviceHandle: TDeviceHandle; AOverall: TEventOverall); stdcall;
    procedure SECUpdateActnMenusProc;


implementation

uses UnitStart, UnitEditEvent;

{$R *.dfm}

procedure DeviceStatusNotify(ADCSIP: PChar; ADeviceHandle: TDeviceHandle; AStatus: TDeviceStatus);
var
  I, J: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;

  DCSID: Word;
  DeviceForm: TfrmDevice;
begin
  if (frmSEC <> nil) and (frmSEC.HandleAllocated) then
    with frmSEC do
    begin
      if (aopDevice.ControlCount > 0) then
      begin
        DeviceForm := TfrmDevice(aopDevice.Controls[0]);
        if (DeviceForm <> nil) then
          DeviceForm.SetDeviceStatus(String(ADCSIP), ADeviceHandle, AStatus);
      end;
    end;
exit;

{  Source := nil;
  for I := 0 to GV_SourceList.Count - 1 do
  begin
    SourceHandles := GV_SourceList[I]^.Handles;
    if (SourceHandles <> nil) then
    begin
      for J := 0 to SourceHandles.Count - 1 do
        if (SourceHandles[J].Handle = ADeviceHandle) then
        begin
          Source := GV_SourceList[I];
          DCSID := SourceHandles[J].DCSID;
          break;
        end;
    end;
    if (Source <> nil) then break;
  end;

  if (Source = nil) then exit;

  if (frmSEC <> nil) and (frmSEC.HandleAllocated) then
    with frmSEC do
    begin
      if (aopDevice.ControlCount > 0) then
      begin
        DeviceForm := TfrmDevice(aopDevice.Controls[0]);
        if (DeviceForm <> nil) then
          DeviceForm.SetDeviceStatus(I, Source, AStatus, DCSID);
      end;
    end;
  exit;

  // None broadcat
  Source := nil;
  for I := 0 to GV_SourceList.Count - 1 do
  begin
    SourceHandles := GV_SourceList[I]^.Handles;
    if (SourceHandles <> nil) then
    begin
      for J := 0 to SourceHandles.Count - 1 do
        if (String(SourceHandles[J].DCSIP) = String(ADCSIP)) and
           (SourceHandles[J].Handle = ADeviceHandle) then
        begin
          Source := GV_SourceList[I];
          DCSID := SourceHandles[J].DCSID;
          break;
        end;
    end;
    if (Source <> nil) then break;
  end;

  if (Source = nil) then exit;

  if (frmSEC <> nil) and (frmSEC.HandleAllocated) then
    with frmSEC do
    begin
      if (aopDevice.ControlCount > 0) then
      begin
        DeviceForm := TfrmDevice(aopDevice.Controls[0]);
        if (DeviceForm <> nil) then
          DeviceForm.SetDeviceStatus(I, Source, AStatus, DCSID);
      end;
    end; }
end;

procedure EventStatusNotify(ADCSIP: PChar; AEventID: TEventID; AStatus: TEventStatus);
var
  ChannelForm: TfrmChannel;
begin
  if (frmSEC <> nil) and (frmSEC.HandleAllocated) then
    with frmSEC do
    begin
      ChannelForm := GetChannelFormByID(AEventID.ChannelID);
      if (ChannelForm <> nil) then
        ChannelForm.SetEventStatus(AEventID, AStatus);
    end;
end;

procedure EventOverallNotify(ADCSIP: PChar; ADeviceHandle: TDeviceHandle; AOverall: TEventOverall);
var
  I: Integer;
  ChannelForm: TfrmChannel;
begin
{  if (frmSEC <> nil) and (frmSEC.HandleAllocated) then
    with frmSEC do
    begin
      ChannelForm := GetChannelFormByID(AOverall.ChannelID);
      if (ChannelForm <> nil) then
        ChannelForm.SetEventOverall(String(ADCSIP), ADeviceHandle, AOverall);
    end; }
end;

procedure TfrmSEC.ScreenOnActiveControlChange(Sender: TObject);
var
  I: Integer;
  ChannelTimelineForm: TfrmChannelTimeline;
begin
  if (Screen.ActiveControl = nil) then exit;
  if (Screen.ActiveControl.Owner = nil) then exit;
  if not (Screen.ActiveControl.Owner is TfrmChannelTimeline) then exit;

  for I := 0 to Screen.FormCount - 1 do
  begin
    if (Screen.Forms[I] is TfrmChannelTimeline) then
    begin
      with (Screen.Forms[I] as TfrmChannelTimeline) do
      begin
//        pnlDesc.Color := $00251C19;
        pnlDesc.ColorHighLight := $00694F44;
        pnlDesc.ColorShadow    := $00694F44;
      end;
    end;
  end;

  ChannelTimelineForm := (Screen.ActiveControl.Owner as TfrmChannelTimeline);
//  ChannelTimelineForm.WMPanel.Color:= $00AD5830;
  ChannelTimelineForm.pnlDesc.ColorHighLight := $00CC9933;//clRed;
  ChannelTimelineForm.pnlDesc.ColorShadow    := $00CC9933;//clRed;
end;

procedure TfrmSEC.WndProc(var Message: TMessage);
begin
  inherited;

  case Message.Msg of
    WM_UPDATE_CURRENT_TIME:
    begin
      lblCurrentTime.Caption := FormatDateTime('YYYY-MM-DD hh:nn:ss', Now);
    end;
    WM_UPDATE_ACTIVATE:
    begin
      DisplayActivate;
    end;
    WM_EXECUTE_MCC_CHECK:
    begin
      MCCCheck;
    end;
  end;
end;

procedure TfrmSEC.NewScrollBoxWndProc(var M: TMessage);
var
  TrackPos: Integer;
begin
  case M.Msg of
    WM_VSCROLL:
      begin
        if (TWMVScroll(M).ScrollCode = SB_THUMBPOSITION) then exit;
        if (TWMVScroll(M).ScrollCode = SB_THUMBTRACK) then
        begin
          TrackPos := TWMVScroll(M).Pos;
          if FChannelTimelineHeight > 0 then
            TrackPos := Round(TrackPos / FChannelTimelineHeight) * FChannelTimelineHeight;
          aopAllChannels.VertScrollBar.Position := TrackPos;
        end;
//        Caption := IntToStr(TWMVScroll(M).Pos);// 'Vret';
      end;
  end;

  FOldScrollBoxWndProc(M);
end;

function TfrmSEC.GetChannelFormByID(AChannelID: Byte): TfrmChannel;
var
  I: Integer;
begin
  Result := nil;
  for I := 1 to aoPagerMain.AdvPageCount - 1 do
  begin
    if (aoPagerMain.AdvPages[I].Tag = AChannelID) then
    begin
      if (aoPagerMain.AdvPages[I].ControlCount > 0) then
        Result := TfrmChannel(aoPagerMain.AdvPages[I].Controls[0]);
      break;
    end;
  end;
end;

function TfrmSEC.GetChannelTimelineFormByID(AChannelID: Byte): TfrmChannelTimeline;
var
  I: Integer;
begin
  Result := nil;

  if (aopAllChannels.ControlCount > 0) then
  begin
    for I := 0 to aopAllChannels.ControlCount - 1 do
    begin
      if (aopAllChannels.Controls[I].Tag = AChannelID) then
      begin
        Result := TfrmChannelTimeline(aopAllChannels.Controls[I]);
        break;
      end;
    end;
  end;
end;

procedure TfrmSEC.actAllChannelsTimelineZoomInExecute(Sender: TObject);
var
  I: Integer;
  ChannelTimelineForm: TfrmChannelTimeline;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <> 0) then exit;

  if (Screen.ActiveControl = nil) then exit;
  if (Screen.ActiveControl.Owner = nil) then exit;
  if not (Screen.ActiveControl.Owner is TfrmChannelTimeline) then exit;

  ChannelTimelineForm := (Screen.ActiveControl.Owner as TfrmChannelTimeline);
  if (ChannelTimelineForm <> nil) then
    ChannelTimelineForm.TimelineZoomIn;
end;

procedure TfrmSEC.actAllChannelsTimelineZoomOutExecute(Sender: TObject);
var
  I: Integer;
  ChannelTimelineForm: TfrmChannelTimeline;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <> 0) then exit;

  if (Screen.ActiveControl = nil) then exit;
  if (Screen.ActiveControl.Owner = nil) then exit;
  if not (Screen.ActiveControl.Owner is TfrmChannelTimeline) then exit;

  ChannelTimelineForm := (Screen.ActiveControl.Owner as TfrmChannelTimeline);
  if (ChannelTimelineForm <> nil) then
    ChannelTimelineForm.TimelineZoomOut;
end;

procedure TfrmSEC.actControlAssignNextEventExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.AssignNextEvent;
end;

procedure TfrmSEC.actControlAssignTargetEventExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.AssignTargetEvent;
end;

procedure TfrmSEC.actControlDecrease1SecondExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.IncSecondsOnAirEvent(-1);
end;

procedure TfrmSEC.actControlFinishOnAirExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.FinishOnAir;
end;

procedure TfrmSEC.actControlFreezeOnAirExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.FreezeOnAir;
end;

procedure TfrmSEC.actControlIncrease1SecondExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.IncSecondsOnAirEvent(1);
end;

procedure TfrmSEC.actControlStartNextEventImmediatelyExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.StartNextEventImmediately;
end;

procedure TfrmSEC.actControlStartOnairExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.StartOnAir;
end;

procedure TfrmSEC.actEditClearClipboardExecute(Sender: TObject);
begin
  inherited;
  ClearClipboardEvent;
end;

procedure TfrmSEC.actEditCopyEventExecute(Sender: TObject);
begin
  inherited;
  CopyToClipboardEvent;
end;

procedure TfrmSEC.actEditCutEventExecute(Sender: TObject);
begin
  inherited;
  CutToClipboardEvent;
end;

procedure TfrmSEC.actEventDeleteEventExecute(Sender: TObject);
begin
  inherited;
  DeleteEvent;
end;

procedure TfrmSEC.actEventInsertCommentEventExecute(Sender: TObject);
begin
  inherited;
  InsertEvent(EM_COMMENT);
end;

procedure TfrmSEC.actEventInsertJoinEventExecute(Sender: TObject);
begin
  inherited;
  InsertEvent(EM_JOIN);
end;

procedure TfrmSEC.actEventInsertMainEventExecute(Sender: TObject);
begin
  inherited;
  InsertEvent(EM_MAIN);
end;

procedure TfrmSEC.actEventInsertSubEventExecute(Sender: TObject);
begin
  inherited;
  InsertEvent(EM_SUB);
end;

procedure TfrmSEC.actEventInspectEventExecute(Sender: TObject);
begin
  inherited;
  InspectEvent;
end;

procedure TfrmSEC.actEventInsertProgramExecute(Sender: TObject);
begin
  inherited;
  InsertEvent(EM_PROGRAM);
end;

procedure TfrmSEC.actEditPasteEventExecute(Sender: TObject);
begin
  inherited;
  PasteFromClipboardEvent;
end;

procedure TfrmSEC.actEditTimelineZoomInExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.TimelineZoomIn;
end;

procedure TfrmSEC.actEditTimelineZoomOutExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.TimelineZoomOut;
end;

procedure TfrmSEC.actEventUpdateEventExecute(Sender: TObject);
begin
  inherited;
  UpdateEvent;
end;

procedure TfrmSEC.actFileNewPlaylistExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    if (ChannelForm.ChannelOnAir) then
    begin
      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(SENoCreateCuesheetWhileChannelOnair), PChar(Application.Title), MB_OK or MB_ICONWARNING);
      exit;
    end;

    ChannelForm.NewPlayList;
  end;
end;

procedure TfrmSEC.actFileOpenAddPlayListExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  with TOpenDialog.Create(Self) do
    try
      Title := 'Open sdd playlist file';
//      Filename := ExtractFileName(FFileName);
//      InitialDir := ExtractFilePath(FFileName);
      Filter  := 'Playlist file|*.xml|Any file(*.*)|*.*';
      if (Execute) then
      begin
        ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
        if (ChannelForm <> nil) then
          ChannelForm.OpenAddPlaylist(FileName);
      end;
    finally
      Free;
    end;
end;

procedure TfrmSEC.actFileOpenPlaylistExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  with TOpenDialog.Create(Self) do
    try
      Title := 'Open playlist file';
//      Filename := ExtractFileName(FFileName);
//      InitialDir := ExtractFilePath(FFileName);
      Filter  := 'Playlist file|*.xml|Any file(*.*)|*.*';
      if (Execute) then
      begin
        ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
        if (ChannelForm <> nil) then
        begin
          if (ChannelForm.ChannelOnAir) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SENoOpenCuesheetWhileChannelOnair), PChar(Application.Title), MB_OK or MB_ICONWARNING);
            exit;
          end;

          ChannelForm.OpenPlaylist(FileName);
        end;
      end;
    finally
      Free;
    end;
end;

procedure TfrmSEC.actFileSaveAsPlayListExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    with TSaveDialog.Create(Self) do
      try
        Title := 'Save as playlist file';
        Filename := ExtractFileName(ChannelForm.PlayListFileName);
        InitialDir := ExtractFilePath(ChannelForm.PlayListFileName);
        Options := Options + [ofOverwritePrompt];
        Filter  := 'Playlist file|*.xml|Any file(*.*)|*.*';
        if (Execute) then
        begin
          ChannelForm.SaveAsPlaylist(FileName, Date);
        end;
      finally
        Free;
      end;
end;

procedure TfrmSEC.actFileSavePlaylistExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    if (ChannelForm.PlayListFileName <> '') then
      if (ChannelForm.PlayListFileName = NEW_CUESHEET_NAME) then
        actFileSaveAsPlayList.Execute
      else
        ChannelForm.SavePlaylist;
  end;
end;

procedure TfrmSEC.actGotoCurrentEventExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.GotoCurrentEvent;
  end;
end;

procedure TfrmSEC.aoPagerMainChange(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.acgPlaylist.SetFocus;
end;

procedure TfrmSEC.aopAllChannelsResize(Sender: TObject);
var
  I: Integer;
  ChannelTop: Integer;
  SumHeight: Integer;
  ScrollRange: Integer;
  ChannelTimelineForm: TfrmChannelTimeline;
begin
  inherited;

  if (aopAllChannels.ControlCount > 0) then
  begin
    FChannelTimelineHeight := 0;
    SumHeight := 0;
    for I := 0 to GV_ChannelList.Count - 1 do
    begin
      Inc(SumHeight, GV_ChannelTimelineMinHeight);
      if (I >= GV_ChannelMinCount - 1) and
         (SumHeight + GV_ChannelTimelineMinHeight > aopAllChannels.Height) then
      begin
        FChannelTimelineHeight := aopAllChannels.Height div (I + 1);
        break;
      end;
    end;

    if (FChannelTimelineHeight = 0) and (SumHeight < aopAllChannels.Height) then
    begin
      FChannelTimelineHeight := (aopAllChannels.Height div GV_ChannelList.Count);
    end;

    if (FChannelTimelineHeight > 0) then
    begin
      ScrollRange := aopAllChannels.VertScrollBar.Range;
      ChannelTop := 0 - aopAllChannels.VertScrollBar.Position;
      for I := 0 to aopAllChannels.ControlCount - 1 do
      begin
        ChannelTimelineForm := TfrmChannelTimeline(aopAllChannels.Controls[I]);
        if (ChannelTimelineForm <> nil) then
        begin
          ChannelTimelineForm.Top := ChannelTop;
          ChannelTimelineForm.Height := FChannelTimelineHeight;// - GV_Channels[I].WMTitleBar.Height;
          ChannelTimelineForm.Width := aopAllChannels.ClientWidth;
          ChannelTimelineForm.Realign;

          ChannelTop := ChannelTop + FChannelTimelineHeight;
        end;
      end;

      aopAllChannels.VertScrollBar.Increment := FChannelTimelineHeight;
      aopAllChannels.VertScrollBar.ThumbSize := FChannelTimelineHeight;//sbChannel.Height;//sbChannel.VertScrollBar.Range div FChannelHeight;
      aopAllChannels.Realign;

      if (ScrollRange > 0) then
        aopAllChannels.VertScrollBar.Position := Round(aopAllChannels.VertScrollBar.Position * (aopAllChannels.VertScrollBar.Range / ScrollRange));
    end;
  end;
end;

procedure TfrmSEC.FormCreate(Sender: TObject);
begin
  inherited;
  Initialize;
end;

procedure TfrmSEC.FormDestroy(Sender: TObject);
begin
  inherited;
  Finalize;
end;

procedure SECUpdateActnMenusProc;
begin
//  ShowMessage(frmSEC.ActionMainMenuBar1.ColorMap.Name);
  with frmSEC.XPColorMap do
  begin
    ShadowColor := cl3DDkShadow;
    Color := $00251C19;
    DisabledColor := clGray;
    DisabledFontColor := clMedGray;
    DisabledFontShadow := clBlack;
    FontColor := $00FFBDAD;
    HighlightColor := $00EEF6F7;
    HotColor := clWhite;
    HotFontColor := clWhite;
    MenuColor := $FFFFFFFF;
    FrameTopLeftInner := $005A453F;
    FrameTopLeftOuter := $00251C19;
    FrameBottomRightInner := $005A453F;
    FrameBottomRightOuter := $00251C19;
    BtnFrameColor := $005A453F;
    BtnSelectedColor := $00614A43;
    BtnSelectedFont := clWhite;
    SelectedColor := $00614A43;
    SelectedFontColor := clWhite;
    UnusedColor := $00EFD3C6;
  end;
  frmSEC.ActionMainMenuBar1.ColorMap := frmSEC.XPColorMap;
  frmSEC.ActionMainMenuBar1.Color := $00251C19;
  frmSEC.ActionMainMenuBar1.Font.Name := 'Century Gothic';
  frmSEC.ActionMainMenuBar1.Font.Color := $00FFBDAD;
  frmSEC.ActionMainMenuBar1.Font.Size := 9;
//  ShowMessage(Format('%x', [Integer(frmSEC.XPColorMap.FontColor)]));
end;

procedure TfrmSEC.Initialize;
var
  R: Integer;
  I: Integer;
  OfficePage: TAdvOfficePage;
  Channel: PChannel;
  ChannelForm: TfrmChannel;
begin
  Screen.OnActiveControlChange := ScreenOnActiveControlChange;
  UpdateActnMenusProc := SECUpdateActnMenusProc;

  GV_LogCS := TCriticalSection.Create;

  GV_ChannelList := TChannelList.Create;
  GV_SECList := TSECList.Create;
  GV_MCCList := TMCCList.Create;
  GV_DCSList := TDCSList.Create;
  GV_SourceList := TSourceList.Create;
  GV_MCSList := TMCSList.Create;

  GV_ProgramTypeList := TProgramTypeList.Create;

  FillChar(GV_NullEventID, SizeOf(TEventID), #0);

  GV_ClipboardCueSheet := TClipboardCueSheet.Create;

  GV_TimerExecuteEvent := CreateEvent(nil, True, False, nil);
  GV_TimerCancelEvent  := CreateEvent(nil, True, False, nil);

  LoadConfig;

  with GV_SettingMCC do
  begin
    R := MCCInitialize(NotifyPort, InPort, OutPort);
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_MCCInitializeFailed, [R, NotifyPort, InPort, OutPort]));
  end;

  with GV_SettingDCS do
  begin
    R := DCSInitialize(NotifyPort, InPort, OutPort);
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_DCSInitializeFailed, [R, NotifyPort, InPort, OutPort]));

    R := DCSDeviceStatusNotify(@DeviceStatusNotify);
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_DCSSetDeviceStatusNotifyFailed, [R]));

    R := DCSEventStatusNotify(@EventStatusNotify);
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_DCSSetEventStatusNotifyFailed, [R]));

    R := DCSEventOverallNotify(@EventOverallNotify);
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_DCSSetEventOverallNotifyFailed, [R]));
  end;

  DeviceOpen;

  InitializeAllChannelPage;
  InitializeChannelPage;
  InitializeDevicePage;

  FOldScrollBoxWndProc := aopAllChannels.WindowProc;
  aopAllChannels.WindowProc := NewScrollBoxWndProc;

  FCrossCheckThread := TCrossCheckThread.Create(Self);
  FCrossCheckThread.DoMainCheck;
  FCrossCheckThread.Start;

  if (GV_SettingMCC.Use) then
  begin
    FMCCCheckThread := TMCCCheckThread.Create(Self);
    FMCCCheckThread.Resume;
  end;

  FTimerThread := TTimerThread.Create;
  FTimerThread.Interval := 1000;
  FTimerThread.OnTimerEvent := TimerThreadEvent;
  FTimerThread.Enabled := True;

//  Maximize;
end;

procedure TfrmSEC.Finalize;
var
  R: Integer;
begin
  SetEvent(GV_TimerCancelEvent);

  FTimerThread.Terminate;
  FTimerThread.WaitFor;
  FreeAndNil(FTimerThread);

  if (FMCCCheckThread <> nil) then
  begin
    FMCCCheckThread.Terminate;
    FMCCCheckThread.WaitFor;
    FreeAndNil(FMCCCheckThread);
  end;

  FCrossCheckThread.Terminate;
  FCrossCheckThread.WaitFor;
  FreeAndNil(FCrossCheckThread);

  CloseHandle(GV_TimerExecuteEvent);
  CloseHandle(GV_TimerCancelEvent);

  DeviceClose;

  FinalizeDevicePage;
  FinalizeChannelPage;
  FinalizeAllChannelPage;

  R := DCSFinalize;
  if (R <> D_OK) then
    Assert(False, GetMainLogStr(lsError, @LSE_DCSFinalizeFailed, [R]));

  R := MCCFinalize;
  if (R <> D_OK) then
    Assert(False, GetMainLogStr(lsError, @LSE_MCCFinalizeFailed, [R]));

  SaveConfig;

  ClearProgramTypeList;

  ClearMCSList;
  ClearSourceList;
  ClearDCSList;
  ClearMCCList;
  ClearSECList;
  ClearChannelList;

  FreeAndNil(GV_ClipboardCueSheet);

  FreeAndNil(GV_ProgramTypeList);

  FreeAndNil(GV_MCSList);
  FreeAndNil(GV_SourceList);
  FreeAndNil(GV_DCSList);
  FreeAndNil(GV_MCCList);
  FreeAndNil(GV_SECList);
  FreeAndNil(GV_ChannelList);

  FreeAndNil(GV_LogCS);
end;

procedure TfrmSEC.InitializeAllChannelPage;
var
  I: Integer;
  Channel: PChannel;
  ChannelTimelineForm: TfrmChannelTimeline;
begin
  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    Channel := GV_ChannelList[I];

    ChannelTimelineForm := TfrmChannelTimeline.Create(aopAllChannels, Channel^.ID, True, 0, 0, aopAllChannels.ClientWidth, GV_ChannelTimelineMinHeight);
    ChannelTimelineForm.Parent := aopAllChannels;
    ChannelTimelineForm.Align := alNone;
    ChannelTimelineForm.Caption := Channel^.Name;
    ChannelTimelineForm.Tag := Channel^.ID;
    ChannelTimelineForm.Show;
  end;
end;

procedure TfrmSEC.FinalizeAllChannelPage;
var
  I: Integer;
  ChannelTimelineForm: TfrmChannelTimeline;
begin
  for I := aopAllChannels.ControlCount - 1 downto 0 do
  begin
    ChannelTimelineForm := TfrmChannelTimeline(aopAllChannels.Controls[I]);
    if (ChannelTimelineForm <> nil) then
    begin
      FreeAndNil(ChannelTimelineForm);
    end;
  end;
end;

procedure TfrmSEC.InitializeChannelPage;
var
  I: Integer;
  OfficePage: TAdvOfficePage;
  Channel: PChannel;
  ChannelForm: TfrmChannel;
begin
  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    Channel := GV_ChannelList[I];
    with aoPagerMain do
    begin
      OfficePage := TAdvOfficePage.Create(Self);
      OfficePage.AdvOfficePager := aoPagerMain;
      OfficePage.Caption := Channel^.Name;
      OfficePage.Tag := Channel^.ID;

      ChannelForm := TfrmChannel.Create(OfficePage, Channel^.ID, GetChannelTimelineFormByID(Channel^.ID), True, 0, 0, OfficePage.ClientWidth, OfficePage.ClientHeight);
      ChannelForm.Parent := OfficePage;
      ChannelForm.Align := alClient;
      ChannelForm.Show;
    end;
  end;
end;

procedure TfrmSEC.FinalizeChannelPage;
var
  I: Integer;
  OfficePage: TAdvOfficePage;
  ChannelForm: TfrmChannel;
begin
  for I := aoPagerMain.AdvPageCount - 1 downto 0 do
  begin
    if (I > 0) then
    begin
      OfficePage := aoPagerMain.AdvPages[I];
      if (OfficePage <> nil) then
      begin
        if (OfficePage.ControlCount > 0) then
        begin
          ChannelForm := TfrmChannel(OfficePage.Controls[0]);
          if (ChannelForm <> nil) then
          begin
            FreeAndNil(ChannelForm);
          end;
        end;
        FreeAndNil(OfficePage);
      end;
    end;
  end;
end;

procedure TfrmSEC.InitializeDevicePage;
var
  DeviceForm: TfrmDevice;
begin
  with aopDevice do
  begin
    DeviceForm := TfrmDevice.Create(aopDevice, True, 0, 0, aopDevice.ClientWidth, aopDevice.ClientHeight);
    DeviceForm.Parent := aopDevice;
    DeviceForm.Align := alClient;
    DeviceForm.Show;
  end;
end;

procedure TfrmSEC.FinalizeDevicePage;
var
  DeviceForm: TfrmDevice;
begin
  if (aopDevice.ControlCount > 0) then
  begin
    DeviceForm := TfrmDevice(aopDevice.Controls[0]);
    if (DeviceForm <> nil) then
    begin
      FreeAndNil(DeviceForm);
    end;
  end;
end;

procedure TfrmSEC.DisplayActivate;
begin
  if (GV_SECMine <> nil) and (GV_SECMine^.Main) then
    pnlActive.Caption := 'Main'
  else
    pnlActive.Caption := 'Sub';

  if (GV_SECMine <> nil) then
    pnlSECName.Caption := String(GV_SECMine^.Name)
  else
    pnlSECName.Caption := '';
end;

procedure TfrmSEC.DeviceOpen;
begin
  frmStart := TfrmStart.Create(Self);
  try
    frmStart.ShowModal;
  finally
    FreeAndNil(frmStart);
  end;
end;

procedure TfrmSEC.DeviceClose;
var
  I, J: Integer;
  R: Integer;
  SourceHandle: PSourceHandle;
begin
//exit;
  for I := GV_SourceList.Count - 1 downto 0 do
  begin
    if (GV_SourceList[I]^.Handles <> nil) then
    begin
      for J := 0 to GV_SourceList[I]^.Handles.Count - 1 do
      begin
        SourceHandle := GV_SourceList[I]^.Handles[J];

        if (SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
        begin
          R := DCSClose(SourceHandle^.DCSID, SourceHandle^.Handle);
          if (R <> D_OK) then
//            ShowMessage(Format('Close Failed ID=%d, Handle=%d, Name=%s', [SourceHandle^.DCSID, SourceHandle^.Handle, GV_SourceList[I]^.Name]))
            Assert(False, GetMainLogStr(lsError, @LSE_DCSCloseDeviceFailed, [R, SourceHandle^.DCSID, String(GV_SourceList[I]^.Name), SourceHandle^.Handle]))
          else
            Assert(False, GetMainLogStr(lsNormal, @LS_DCSCloseDeviceSuccess, [SourceHandle^.DCSID, String(GV_SourceList[I]^.Name), SourceHandle^.Handle]));
        end;
      end;
    end;
  end;
end;

procedure TfrmSEC.actFileCloseExecute(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TfrmSEC.TimerThreadEvent(Sender: TObject);
begin
  PostMessage(Handle, WM_UPDATE_CURRENT_TIME, 0, 0);
//  PostMessage(Handle, WM_UPDATE_DEVICESTATUS, 0, 0);
end;

procedure TfrmSEC.InsertEvent(AEventMode: TEventMode);
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.InsertCueSheet(AEventMode);
  end;
end;

procedure TfrmSEC.lblCurrentTimeMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  inherited;
  if WMTitleBar.MoveEnabled then
  begin
    ReleaseCapture;
    SendMessage(Handle, WM_SYSCOMMAND, $F012, 0);
  end;
end;

procedure TfrmSEC.UpdateEvent;
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.UpdateCueSheet;
  end;
end;

procedure TfrmSEC.DeleteEvent;
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.DeleteCueSheet;
  end;
end;

procedure TfrmSEC.InspectEvent;
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.InspectCueSheet;
  end;
end;

procedure TfrmSEC.CutToClipboardEvent;
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.CutToClipboardCueSheet;
  end;
end;

procedure TfrmSEC.CopyToClipboardEvent;
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.CopyToClipboardCueSheet;
  end;
end;

procedure TfrmSEC.PasteFromClipboardEvent;
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.PasteFromClipboardCueSheet;
  end;
end;

procedure TfrmSEC.ClearClipboardEvent;
var
  I: Integer;
  ChannelForm: TfrmChannel;
begin
  inherited;

//  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(GV_ClipboardCueSheet.ChannelID);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.ClearClipboardCueSheet;
  end;
end;

procedure TfrmSEC.TimelineZoomIn;
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.TimelineZoomIn;
  end;
end;

procedure TfrmSEC.TimelineZoomOut;
var
  ChannelForm: TfrmChannel;
begin
  inherited;
  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    ChannelForm.TimelineZoomOut;
  end;
end;

procedure TfrmSEC.MCCCheck;
var
  I, J, R: Integer;
  SaveOpened: Boolean;
  ChannelForm: TfrmChannel;
begin
  for I := 0 to GV_MCCList.Count - 1 do
  begin
    SaveOpened := GV_MCCList[I]^.Opened;

    R := MCCOpen(GV_MCCList[I]^.ID, GV_MCCList[I]^.HostIP, GV_MCCList[I]^.Name);

    if ((R = D_OK) and (not SaveOpened)) or
       ((R <> D_OK) and (SaveOpened)) then
    begin
      for J := 0 to 0 do //GV_ChannelList.Count - 1 do
      begin
        ChannelForm := GetChannelFormByID(GV_ChannelList[J]^.ID);
        if (ChannelForm <> nil) then
        begin
          GV_MCCList[I]^.Opened := (R = D_OK);
          ChannelForm.UpdateMCCCheck(GV_MCCList[I]^.ID, GV_MCCList[I]^.Opened);

//          PostMessage(ChannelForm.Handle, WM_UPDATE_MCC_CHECK, GV_MCCList[I]^.ID, NativeInt(GV_MCCList[I]^.Opened));
        end;
      end;
    end;

    if (R = D_OK) then
    begin
      Assert(False, GetMainLogStr(lsNormal, @LS_MCCCheckSuccess, [GV_MCCList[I]^.ID, GV_MCCList[I]^.Name]));
      GV_MCCList[I]^.Opened := True;
    end
    else
    begin
      Assert(False, GetMainLogStr(lsError, @LSE_MCCCheckFailed, [R, GV_MCCList[I]^.ID, GV_MCCList[I]^.Name]));
      GV_MCCList[I]^.Opened := False;
    end;
  end;
end;

function TfrmSEC.SECIsAlive(var AIsAlive: Boolean): Integer;
begin
  Result := D_FALSE;

  AIsAlive := True;

  Result := D_OK;
end;

function TfrmSEC.SECIsMain(var AIsMain: Boolean): Integer;
begin
  Result := D_FALSE;

  if (GV_SECMine <> nil) then
    AIsMain := GV_SECMine^.Main
  else
    AIsMain := False;

  Result := D_OK;
end;

function TfrmSEC.SECSetAlive(ABuffer: AnsiString): Integer;
var
  SECID: Word;
  SEC: PSEC;
  Allive: Boolean;
begin
  Result := D_FALSE;

  SECID := PAnsiCharToWord(@ABuffer[1]);
  Allive := PAnsiCharToBool(@ABuffer[3]);

  SEC := GetSECByID(SECID);
  if (SEC <> nil) then
  begin
    SEC^.Alive := True;

    Result := D_OK;
  end;
  FCrossCheckThread.FNumCrossCheck := 0;
end;

function TfrmSEC.SECSetMain(ABuffer: AnsiString): Integer;
var
  SECID: Word;
  SEC: PSEC;
begin
  Result := D_FALSE;

  SECID := PAnsiCharToWord(@ABuffer[1]);

  SEC := GetSECByID(SECID);
  if (SEC <> nil) then
  begin
    GV_SECMain := SEC;
    GV_SECMain^.Main := True;

    Result := D_OK;
  end;
  FCrossCheckThread.FNumCrossCheck := 0;
end;

{ TTimerThread }

constructor TTimerThread.Create;
begin
  FTimerEnabledFlag := CreateEvent(nil, True, False, nil);
  FCancelFlag := CreateEvent(nil, True, False, nil);
  FTimerProc := nil;
  FInterval := 1000;
  FreeOnTerminate := False; // Main thread controls for thread destruction
  inherited Create(False);
end;

destructor TTimerThread.Destroy; // Call TTimerThread.Free to cancel the thread
begin
  Terminate;
  if GetCurrentThreadID = MainThreadID then
  begin
    OutputDebugString('TTimerThread.Destroy :: MainThreadID (Waitfor)');
    Waitfor; // Synchronize
  end;
  CloseHandle(FCancelFlag);
  CloseHandle(FTimerEnabledFlag);
  OutputDebugString('TTimerThread.Destroy');
  inherited;
end;

procedure TTimerThread.Terminate;
begin
  inherited Terminate;
  ResetEvent(FTimerEnabledFlag); // Stop timer event
  SetEvent(FCancelFlag); // Set cancel flag
end;

procedure TTimerThread.SetEnabled(AEnable: Boolean);
begin
  if AEnable then
    SetEvent(FTimerEnabledFlag)
  else
    ResetEvent(FTimerEnabledFlag);
end;

function TTimerThread.GetEnabled: Boolean;
begin
  Result := WaitForSingleObject(FTimerEnabledFlag, 0) = WAIT_OBJECT_0 // Signaled
end;

procedure TTimerThread.SetInterval(AInterval: Cardinal);
begin
  FInterval := AInterval;
end;

procedure TTimerThread.Execute;
var
  WaitList: array[0..1] of THandle;
  WaitInterval, LastProcTime: Int64;
  Frequency, StartCount, StopCount: Int64; // minimal stop watch
begin
  QueryPerformanceFrequency(Frequency); // this will never return 0 on Windows XP or later
  WaitList[0] := FTimerEnabledFlag;
  WaitList[1] := FCancelFlag;
  LastProcTime := 0;

  if (Enabled) then
    FTimerProc(Self);

  while not Terminated do
  begin
    ResetEvent(GV_TimerExecuteEvent);
    if (WaitForMultipleObjects(2, @WaitList[0], False, INFINITE) <> WAIT_OBJECT_0) then
      break; // Terminate thread when FCancelFlag is signaled

    WaitInterval := FInterval - LastProcTime;
    if (WaitInterval < 0) then
      WaitInterval := 0;
    if (WaitForSingleObject(FCancelFlag, WaitInterval) <> WAIT_TIMEOUT) then
      break;

    if (Enabled) then
    begin
      QueryPerformanceCounter(StartCount);
      if not Terminated then
      begin
        if Assigned(FTimerProc) then
          FTimerProc(Self);
        SetEvent(GV_TimerExecuteEvent);
      end;
      QueryPerformanceCounter(StopCount);
      // Interval adjusted for FTimerProc execution time
      LastProcTime := 1000 * (StopCount - StartCount) div Frequency; // ElapsedMilliSeconds
    end;
  end;
end;

{ TMCCCheckThread }

constructor TMCCCheckThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  GV_MCCCheckExecuteEvent := CreateEvent(nil, True, True, nil);
  FCloseEvent := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TMCCCheckThread.Destroy;
begin
  Terminate;

  CloseHandle(GV_MCCCheckExecuteEvent);
  CloseHandle(FCloseEvent);

  inherited Destroy;
end;

procedure TMCCCheckThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FCloseEvent);
end;

procedure TMCCCheckThread.DoMCCCheck;
var
  I, J, R: Integer;
  SaveOpened: Boolean;
  ChannelForm: TfrmChannel;
begin
{  for I := 0 to GV_MCCList.Count - 1 do
  begin
    SaveOpened := GV_MCCList[I]^.Opened;

    R := MCCOpen(GV_MCCList[I]^.ID, GV_MCCList[I]^.HostIP, GV_MCCList[I]^.Name);

    if ((R = D_OK) and (not SaveOpened)) or
       ((R <> D_OK) and (SaveOpened)) then
    begin
      for J := 0 to 0 do //GV_ChannelList.Count - 1 do
      begin
        ChannelForm := FSEC.GetChannelFormByID(GV_ChannelList[J]^.ID);
        if (ChannelForm <> nil) then
        begin
          GV_MCCList[I]^.Opened := (R = D_OK);
          PostMessage(ChannelForm.Handle, WM_UPDATE_MCC_CHECK, GV_MCCList[I]^.ID, NativeInt(GV_MCCList[I]^.Opened));
        end;
      end;
    end;

    if (R = D_OK) then
    begin
      Assert(False, GetMainLogStr(lsNormal, @LS_MCCCheckSuccess, [GV_MCCList[I]^.ID, GV_MCCList[I]^.Name]));
      GV_MCCList[I]^.Opened := True;
    end
    else
    begin
      Assert(False, GetMainLogStr(lsError, @LSE_MCCCheckFailed, [R, GV_MCCList[I]^.ID, GV_MCCList[I]^.Name]));
      GV_MCCList[I]^.Opened := False;
    end;
  end; }

//  PostMessage(FSEC.Handle, WM_EXECUTE_MCC_CHECK, 0, 0);

  FSEC.MCCCheck;
end;

procedure TMCCCheckThread.Execute;
var
  R: Cardinal;
begin
  while not Terminated do
  begin
    R := WaitForSingleObject(FCloseEvent, TimecodeToMilliSec(GV_SettingTimeParameter.MCCCheckInterval));
    if (R = WAIT_OBJECT_0) then break;

    R := WaitForSingleObject(GV_MCCCheckExecuteEvent, INFINITE);

    DoMCCCheck;
  end;
end;

{ TCrossCheckThread }

constructor TCrossCheckThread.Create(ASECForm: TfrmSEC);
begin
  FSECForm := ASECForm;

  with GV_SettingSEC do
  begin
    FUDPIn  := TUDPIn.Create(CrossPort);
    FUDPOut := TUDPOut.Create(CrossPort);
  end;

  FUDPIn.OnUDPInRead := UDPInRead;
  FUDPIn.Start;
  while not FUDPIn.Started do
    Sleep(30);

  FUDPOut.Start;
  while not FUDPOut.Started do
    Sleep(30);

  FIsCommand := False;
  FCMD1 := $00;
  FCMD2 := $00;

  FSyncMsgEvent := CreateEvent(nil, True, False, nil);

  FNumCrossCheck := 0;

  FreeOnTerminate := False;

  inherited Create(True);
end;

destructor TCrossCheckThread.Destroy;
begin
  CloseHandle(FSyncMsgEvent);

  FUDPIn.Close;
  FUDPIn.Terminate;
  FUDPIn.WaitFor;
  FreeAndNil(FUDPIn);

  FUDPOut.Close;
  FUDPOut.Terminate;
  FUDPOut.WaitFor;
  FreeAndNil(FUDPOut);

  inherited Destroy;
end;

procedure TCrossCheckThread.UDPInRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);
var
  ByteCount: Integer;
  CMD1, CMD2: Byte;
  Device: TDeviceHandle;

  R: Integer;
  IsAlive: Boolean;
  IsMain: Boolean;
  EventCount: Integer;
  Event: TEvent;
  SendBuffer: AnsiString;
begin
  if (ADataSize <= 0) then exit;
  FRecvBuffer := FRecvBuffer + AData;

  if (Length(FRecvBuffer) < 1) then exit;

//  if (GV_DCSMine <> nil) and (GV_DCSMine^.Main) then
  if (not FIsCommand) then
  begin
    case Ord(FRecvBuffer[1]) of
      $02:
      begin
        if (Length(FRecvBuffer) < 2) then exit;
        ByteCount := Ord(FRecvBuffer[2]);
        if (Length(FRecvBuffer) >= ByteCount + 3) then
        begin
          if (CheckSum(FRecvBuffer)) then
          begin
            CMD1 := Ord(FRecvBuffer[3]);
            CMD2 := Ord(FRecvBuffer[4]);

            FRecvData := System.Copy(FRecvBuffer, 5, ByteCount - 2);

            case CMD1 of
              $00: // System Control (0X00)
              begin
                case CMD2 of
                  $00: // Is Alive
                  begin
                    R := FSECForm.SECIsAlive(IsAlive);
                    if (R = D_OK) then
                      SendBuffer := BoolToAnsiString(IsAlive);
                  end;
                  $01: // Is Main
                  begin
                    R := FSECForm.SECIsMain(IsMain);
                    if (R = D_OK) then
                      SendBuffer := BoolToAnsiString(IsMain);
                  end;
                  $10: // Set Allive
                  begin
                    R := FSECForm.SECSetAlive(FRecvData);
                  end;
                  $11: // Set Main
                  begin
                    R := FSECForm.SECSetMain(FRecvData);
                  end;
                end;
              end;
            end;

            if (R = D_OK) then
            begin
              case CMD1 of
                $00: // System Control (0X00)
                begin
                  case CMD2 of
                    $00,
                    $01: TransmitResponse(ABindingIP, FUDPOut.Port, CMD1, CMD2, SendBuffer, Length(SendBuffer));
                    $10,
                    $11: TransmitAck(ABindingIP, FUDPOut.Port);
                  end;
                end;
              end;
            end
            else
            begin
              // Sending error code
  //            SendBuffer := AnsiChar(D_ERR) + IntToAnsiString(R);
  //            FUDPOut.Send(ABindingIP, FUDPOut.Port, SendBuffer);
              TransmitError(ABindingIP, FUDPOut.Port, R);
            end;
          end
          else
          begin
  //          FUDPOut.Send(ABindingIP, AnsiChar(D_NAK) + IntToAnsiString(E_NAK_CHECKSUM));
            TransmitNak(ABindingIP, FUDPOut.Port, E_NAK_CHECKSUM);
          end;
        end;

        if (Length(FRecvBuffer) > ByteCount + 3) then
          FRecvBuffer := Copy(FRecvBuffer, ByteCount + 4, Length(FRecvBuffer))
        else
          FRecvBuffer := '';
      end;
      else
      begin
  //      FUDPOut.Send(ABindingIP, AnsiChar(D_ERR) + IntToAnsiString(E_UNDEFIND_COMMAND));
        TransmitError(ABindingIP, FUDPOut.Port, E_UNDEFIND_COMMAND);
        FRecvBuffer := '';
      end;
    end;
  end
  else
  begin
    case Ord(FRecvBuffer[1]) of
      $02:
        begin
          if (Length(FRecvBuffer) < 2) then exit;
          ByteCount := Ord(FRecvBuffer[2]);
          if (Length(FRecvBuffer) = ByteCount + 3) then
          begin
            if (CheckSum(FRecvBuffer)) then
            begin
              CMD1 := Ord(FRecvBuffer[3]);
              CMD2 := Ord(FRecvBuffer[4]);

              FRecvData := System.Copy(FRecvBuffer, 5, ByteCount - 2);

              if (CMD1 = FCMD1) and (CMD2 = FCMD2 + $80)  then
                FLastResult := D_OK
              else
                FLastResult := D_FALSE;
            end
            else
              FLastResult := E_NAK_CHECKSUM;

            SetEvent(FSyncMsgEvent);
            FRecvBuffer := '';
          end
          else if ((ByteCount <= 0) or (Length(FRecvBuffer) > ByteCount + 3)) then
          begin
            FLastResult := D_FALSE;
            SetEvent(FSyncMsgEvent);
            FRecvBuffer := '';
          end;
        end;
      $04: // ACK
        begin
          FLastResult := D_OK;
          SetEvent(FSyncMsgEvent);
          FRecvBuffer := '';
        end;
      else
      begin
        FLastResult := D_FALSE;
        SetEvent(FSyncMsgEvent);
        FRecvBuffer := '';
      end;
    end;
  end;
end;

function TCrossCheckThread.SendCommand(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
var
  Buffer: AnsiString;
  CheckSum: Byte;
  I: integer;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_STX) + AnsiChar($02 + ADataSize) + AnsiChar(ACMD1) + AnsiChar(ACMD2) + AData;

  CheckSum := ACMD1 + ACMD2;
  for I := 1 to ADataSize do
    CheckSum := CheckSum + Ord(AData[I]);

  CheckSum := 0 - CheckSum;
  Buffer := Buffer + AnsiChar(CheckSum);

  FUDPOut.Send(AHostIP, APort, Buffer);

  Result := D_OK;
end;

function TCrossCheckThread.SendResponse(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
var
  Buffer: AnsiString;
  CheckSum: Byte;
  I: integer;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_STX) + AnsiChar($02 + ADataSize) + AnsiChar(ACMD1) + AnsiChar(ACMD2) + AData;

  CheckSum := ACMD1 + ACMD2;
  for I := 1 to ADataSize do
    CheckSum := CheckSum + Ord(AData[I]);

  CheckSum := 0 - CheckSum;
  Buffer := Buffer + AnsiChar(CheckSum);

  FUDPOut.Send(AHostIP, APort, Buffer);

  Result := D_OK;
end;

function TCrossCheckThread.SendAck(AHostIP: AnsiString; APort: Word): Integer;
begin
  Result := D_FALSE;

  FUDPOut.Send(AHostIP, APort, AnsiChar(D_ACK));

  Result := D_OK;
end;

function TCrossCheckThread.SendNak(AHostIP: AnsiString; APort: Word; ANakError: Byte): Integer;
var
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_NAK) + AnsiChar(ANakError);

  FUDPOut.Send(AHostIP, APort, Buffer);

  Result := D_OK;
end;

function TCrossCheckThread.SendError(AHostIP: AnsiString; APort: Word; AErrorCode: Integer): Integer;
var
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_ERR) + IntToAnsiString(AErrorCode);

  FUDPOut.Send(AHostIP, APort, Buffer);

  Result := D_OK;
end;

function TCrossCheckThread.TransmitCommand(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
var
  R: DWORD;
begin
  Result := S_FALSE;

  ResetEvent(FSyncMsgEvent);
  FRecvBuffer := '';
  FRecvData := '';

  FIsCommand := True;
  FCMD1 := ACMD1;
  FCMD2 := ACMD2;

//  FCriticalSection.Enter;
  try
    if (SendCommand(AHostIP, APort, ACMD1, ACMD2, AData, ADataSize) = NOERROR) then
    begin
      R := WaitForSingleObject(FSyncMsgEvent, TIME_OUT);
      case R of
        WAIT_OBJECT_0:
          begin
            Result := FLastResult;
          end;
        else Result := E_TIMEOUT;
       end;
    end;
  finally
//    FCriticalSection.Leave;

    FIsCommand := False;
  end;
end;

function TCrossCheckThread.TransmitResponse(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
begin
  ACMD2 := ACMD2 + $80;
  Result := SendResponse(AHostIP, APort, ACMD1, ACMD2, AData, ADataSize);
end;

function TCrossCheckThread.TransmitAck(AHostIP: AnsiString; APort: Word): Integer;
begin
  Result := SendAck(AHostIP, APort);
end;

function TCrossCheckThread.TransmitNak(AHostIP: AnsiString; APort: Word; ANakError: Byte): Integer;
begin
  Result := SendNak(AHostIP, APort, ANakError);
end;

function TCrossCheckThread.TransmitError(AHostIP: AnsiString; APort: Word; AErrorCode: Integer): Integer;
begin
  Result := SendError(AHostIP, APort, AErrorCode);
end;

function TCrossCheckThread.SECGetIsAlive(AHostIP: AnsiString; var AIsAlive: Boolean): Integer;
var
  Buffer: AnsiString;
begin
  AIsAlive := False;

  Buffer := '';
  Result := TransmitCommand(AHostIP, FUDPOut.Port, $00, $00, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    AIsAlive := PAnsiCharToBool(@FRecvData[1]);
  end;
end;

function TCrossCheckThread.SECGetIsMain(AHostIP: AnsiString; var AIsMain: Boolean): Integer;
var
  Buffer: AnsiString;
begin
  AIsMain := False;

  Buffer := '';
  Result := TransmitCommand(AHostIP, FUDPOut.Port, $00, $01, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    AIsMain := PAnsiCharToBool(@FRecvData[1]);
  end;
end;

function TCrossCheckThread.SECSetAllive(AHostIP: AnsiString; ASECID: Word; AAlive: Boolean): Integer;
var
  Buffer: AnsiString;
begin
  Buffer := WordToAnsiString(ASECID) + BoolToAnsiString(AAlive);
  Result := TransmitCommand(AHostIP, FUDPOut.Port, $00, $10, Buffer, Length(Buffer));
end;

function TCrossCheckThread.SECSetMain(AHostIP: AnsiString; AMainSECID: Word): Integer;
var
  Buffer: AnsiString;
begin
  Buffer := WordToAnsiString(AMainSECID);
  Result := TransmitCommand(AHostIP, FUDPOut.Port, $00, $11, Buffer, Length(Buffer));
end;

procedure TCrossCheckThread.Execute;
begin
  { Place thread code here }

//  DoMainCheck;

  while not Terminated do
  begin
//    DoCrossCheck;
    Sleep(TimecodeToMilliSec(GV_SettingSEC.CrossCheckInterval));
  end;
end;

procedure TCrossCheckThread.DoMainCheck;
var
  I, J, K: Integer;
  SEC: PSEC;

  R: Integer;
  IsMain: Boolean;
begin
  // SEC 실행 시 현재 운행중인 Main SEC가 있으면
  // 현재 SEC를 Sub로 전환하고 채널 운행 모니터링 정보를 수신한다.

  // Another SEC main check
  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (GV_SECMine <> SEC) then
    begin
      R := SECGetIsMain(SEC^.HostIP, IsMain);
      if (R = D_OK) and (IsMain) then
      begin
        GV_SECMain := SEC;
        GV_SECMain^.Main := IsMain;

        GV_SECMine^.Main := False;

{        // 채널 모니터링 정보를 수신
        for J := 0 to GV_DeviceList.Count - 1 do
        begin
          DeviceHandle := GV_DeviceList[J]^.Handle;

          R := DCSGetEventCount(DCS^.HostIP, DeviceHandle, EventCount);
          if (R = D_OK) and (EventCount > 0) then
          begin
            D := GetDeviceThreadByHandle(DeviceHandle);
            if (D <> nil) then
            begin
              for K := 0 to EventCount - 1 do
              begin
                R := DCSGetEvent(DCS^.HostIP, DeviceHandle, K, Event);
                if (R = D_OK) then
                  D.InputEvent('', Event);
              end;
            end;
          end;
        end; }

        break;
      end;
    end;
  end;

  // Channel re active
  PostMessage(FSECForm.Handle, WM_UPDATE_ACTIVATE, 0, 0);
end;

procedure TCrossCheckThread.DoCrossCheck;
var
  I: Integer;
  SEC: PSEC;

  R: Integer;
  IsAlive: Boolean;

  NextIndex: Integer;
  NextSEC: PSEC;
begin
  if (GV_SECMine <> nil) and (GV_SECMine^.Main) then exit;

  // Main SEC alive check
  for I := 0 to GV_DCSList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (GV_SECMine <> SEC) and (SEC^.Main) then
    begin
      R := SECGetIsAlive(SEC^.HostIP, IsAlive);
      if (R <> D_OK) or (not IsAlive) then
      begin
        Inc(FNumCrossCheck);
      end
      else
        FNumCrossCheck := 0;

      break;
    end;
  end;

  if (FNumCrossCheck >= GV_SettingSEC.NumCrossCheck) then
  begin
    NextIndex := GV_SECList.IndexOf(GV_SECMain) + 1;
    if (NextIndex < 0) or (NextIndex > GV_SECList.Count - 1) then
      NextIndex := 0;

    // 다음 메인 SEC 설정
    NextSEC := GV_SECList[NextIndex];

    // Send not main another SEC
    for I := 0 to GV_SECList.Count - 1 do
    begin
      SEC := GV_SECList[I];
      if (GV_SECMain <> SEC) and
         (GV_SECMine <> SEC) then
      begin
        R := SECSetMain(SEC^.HostIP, NextSEC^.ID);
      end;
    end;

    if (GV_SECMine = NextSEC) then
    begin
      GV_SECMine^.Main := True;

      // Main/Sub 전환
      if (GV_SECMain <> nil) then
      begin
        GV_SECMain^.Main := False;
        GV_SECMain := GV_SECMine;
      end;

      // Channel re active
      PostMessage(FSECForm.Handle, WM_UPDATE_ACTIVATE, 0, 0);
    end;

    FNumCrossCheck := 0;
  end;
end;

// Sub SEC system alive check
procedure TCrossCheckThread.DoSubAliveCheck;
var
  I: Integer;
  SEC: PSEC;

  R: Integer;
  IsAlive: Boolean;

  NextIndex: Integer;
  NextSEC: PSEC;
begin
{  if (GV_SECMine <> nil) and (GV_SECMine^.Main) then exit;

  // Another SEC alive check
  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (GV_SECMine <> SEC) and (not (SEC^.Main)) then
    begin
      SaveAlive := SEC^.Alive;

      R := SECGetIsAlive(SEC^.HostIP, IsAlive);

      if (IsAlive <> SaveAlive) and (IsAlive) then
      begin
        SECBeginUpdate()
        MCCBeginUpdate(Message.WParam, ChannelID);
        try
          MCCSetOnAir(Message.WParam, ChannelID, ChannelOnAir);
          if (ChannelOnAir) then
          begin
            for I := 0 to FCueSheetList.Count - 1 do
              MCCInputCueSheet(Message.WParam, I, FCueSheetList[I]^);

            if (CueSheetNext <> nil) then
              MCCSetCueSheetNext(Message.WParam, CueSheetNext^.EventID)
            else
              MCCSetCueSheetNext(Message.WParam, GV_NullEventID);

            if (CueSheetCurr <> nil) then
              MCCSetCueSheetCurr(Message.WParam, CueSheetCurr^.EventID)
            else
              MCCSetCueSheetCurr(Message.WParam, GV_NullEventID);

            if (CueSheetTarget <> nil) then
              MCCSetCueSheetTarget(Message.WParam, CueSheetTarget^.EventID)
            else
              MCCSetCueSheetTarget(Message.WParam, GV_NullEventID);
          end;
        finally
          MCCEndUpdate(Message.WParam, ChannelID);
        end;
      end;


      if (R = D_OK) and (not IsAlive) then
      begin
        Inc(FNumCrossCheck);
      end
      else
        FNumCrossCheck := 0;

      break;
    end;
  end;

  if (FNumCrossCheck >= GV_SettingSEC.NumCrossCheck) then
  begin
    NextIndex := GV_SECList.IndexOf(GV_SECMain) + 1;
    if (NextIndex < 0) or (NextIndex > GV_SECList.Count - 1) then
      NextIndex := 0;

    // 다음 메인 SEC 설정
    NextSEC := GV_SECList[NextIndex];

    // Send not main another SEC
    for I := 0 to GV_SECList.Count - 1 do
    begin
      SEC := GV_SECList[I];
      if (GV_SECMain <> SEC) and
         (GV_SECMine <> SEC) then
      begin
        R := SECSetMain(SEC^.HostIP, NextSEC^.ID);
      end;
    end;

    if (GV_SECMine = NextSEC) then
    begin
      GV_SECMine^.Main := True;

      // Main/Sub 전환
      if (GV_SECMain <> nil) then
      begin
        GV_SECMain^.Main := False;
        GV_SECMain := GV_SECMine;
      end;

      // Channel re active
      PostMessage(FSECForm.Handle, WM_UPDATE_ACTIVATE, 0, 0);
    end;

    FNumCrossCheck := 0;
  end;  }
end;

end.
