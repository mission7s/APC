unit UnitSEC;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitNormalForm, Vcl.Imaging.pngimage, Vcl.ClipBrd,
  WMTools, WMControls, Vcl.ExtCtrls, System.Actions, Vcl.ActnList, System.SyncObjs, Winapi.MMSystem,
  Vcl.XPStyleActnCtrls, Vcl.ActnMan, Vcl.ToolWin, Vcl.ActnCtrls, Vcl.ActnMenus, Vcl.Menus,
  Vcl.ActnColorMaps, AdvOfficePager, AdvOfficePagerStylers, Vcl.StdCtrls, Generics.Collections,
  System.DateUtils, AdvSplitter,
  UnitCommons, UnitDCSDLL, UnitMCCDLL, UnitSECDLL, UnitConsts, UnitUDPIn, UnitUDPOut, UnitTypeConvert,
  UnitAllChannels, UnitChannel, UnitDevice, UnitWarningDialog, UnitAlarmThread,
  AdvUtil, Vcl.Grids, AdvObj, BaseGrid, AdvGrid,
  AdvCGrid, AdvScrollBox, System.ImageList, Vcl.ImgList, Vcl.StdStyleActnCtrls;

type
  TCrossCheckThread = class;

//  TTimerThread = class;
  TMCCCheckThread = class;
  TDCSCheckThread = class;
  TSECCheckThread = class;

  TDCSEventThread = class;
  TSECEventThread = class;
  TMCCEventThread = class;

  TDCSMediaThread = class;
  TDCSDeviceThread = class;


  TfrmSEC = class(TfrmNormal)
    aopStyler: TAdvOfficePagerOfficeStyler;
    actMainMenuSEC: TActionMainMenuBar;
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
    actAllChannelsTimelineZoomIn: TAction;
    actAllChannelsTimelineZoomOut: TAction;
    actControlGotoCurrentEvent: TAction;
    actEditTimelineMoveLeft: TAction;
    actEditTimelineMoveRight: TAction;
    actEditTimelineGotoCurrent: TAction;
    actEventSourceExchange: TAction;
    Action3: TAction;
    actAllChannelsTimelineMoveLeft: TAction;
    actAllChannelsTimelineMoveRight: TAction;
    actAllChannelsTimelineGotoCurrent: TAction;
    actEditSelectMedia: TAction;
    Image1: TImage;
    imgSECMain: TImage;
    lblSECName: TLabel;
    imgSECSub: TImage;
    iLstManager: TImageList;
    actViewTimeline: TAction;
    XPColorMap: TXPColorMap;
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
    procedure actAllChannelsTimelineZoomInExecute(Sender: TObject);
    procedure actAllChannelsTimelineZoomOutExecute(Sender: TObject);
    procedure actControlGotoCurrentEventExecute(Sender: TObject);
    procedure actEditTimelineGotoCurrentExecute(Sender: TObject);
    procedure actEditTimelineMoveLeftExecute(Sender: TObject);
    procedure actEditTimelineMoveRightExecute(Sender: TObject);
    procedure actEventSourceExchangeExecute(Sender: TObject);
    procedure actAllChannelsTimelineMoveLeftExecute(Sender: TObject);
    procedure actAllChannelsTimelineMoveRightExecute(Sender: TObject);
    procedure actAllChannelsTimelineGotoCurrentExecute(Sender: TObject);
    procedure actEditSelectMediaExecute(Sender: TObject);
    procedure FormShortCut(var Msg: TWMKey; var Handled: Boolean);
    procedure actMainMenuSECGetControlClass(Sender: TCustomActionBar;
      AnItem: TActionClient; var ControlClass: TCustomActionControlClass);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    FSaveMenuFont: TFont; // will hold initial main menu bar's font settings
    FSaveColorMap: TXPColorMap;

    FTimerID: UINT;

//    FTimerThread: TTimerThread;
    FMCCCheckThread: TMCCCheckThread;

    FDCSCheckThread: TDCSCheckThread;

    FSECCheckThread: TSECCheckThread;

    FDCSEventThread: TDCSEventThread;
    FSECEventThread: TSECEventThread;
    FMCCEventThread: TMCCEventThread;

    FDCSMediaThread: TDCSMediaThread;

    FDCSDeviceThread: TDCSDeviceThread;

    FCrossCheckThread: TCrossCheckThread;

    FChannelTimelineHeight: Integer;

    FIsEditing: Boolean;

    FCurrentTime: TDateTime;


    FWarningDialogDeviceCheck: TfrmWarningDialog;

    procedure DisplayStartCheck(ACheckStr: String);

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
    procedure SECCheck;

//    procedure SECMainCheck;
//    procedure SECMainChange;
//    procedure SECMainDeviceControlBy;
//    procedure SECMainAliveCheck;

    procedure DCSCheck;

	  procedure ApplicationMessage(var Msg: TMsg; var Handled: Boolean);

    procedure ScreenOnActiveControlChange(Sender: TObject);
  protected
    procedure CreateParams(var Params: TCreateParams); override;
//    procedure WndProc(var Message: TMessage); override;

    procedure WMSettingChange(var Message: TWMSettingChange); message WM_SETTINGCHANGE;

    procedure WMUpdateCurrentTime(var Message: TMessage); message WM_UPDATE_CURRENT_TIME;
    procedure WMUpdateActive(var Message: TMessage); message WM_UPDATE_ACTIVATE;
    procedure WMWarningDisplayDeviceCheck(var Message: TMessage); message WM_WARNING_DISPLAY_DEVICE_CHECK;
  public
    { Public declarations }

    function GetChannelFormByID(AChannelID: Word): TfrmChannel;
    function GetDeviceForm: TfrmDevice;

    procedure EnableShortCutAction(AEnabled: Boolean);

    // 0X00 System Control
    function SECIsAliveW(var AIsAlive: Boolean): Integer;                    // 0X00.10
    function SECIsMainW(var AIsMain: Boolean): Integer;                      // 0X00.11
    function SECSetMainW(ABuffer: AnsiString): Integer;                      // 0X00.12


    // 0X00 System Control
{    function SECIsAlive(var AIsAlive: Boolean): Integer;                    // 0X00.00
    function SECIsMain(var AIsMain: Boolean): Integer;                      // 0X00.01

    function SECGetChannelOnAir(ABuffer: AnsiString; AChannelOnAir: Boolean): Integer;  // 0X00.20

    function SECSetAlive(ABuffer: AnsiString): Integer;                     // 0X00.10
    function SECSetMain(ABuffer: AnsiString): Integer;                      // 0X00.11

    function SECSetChannelOnAir(ABuffer: AnsiString): Integer;              // 0X00.30  }

    property MCCCheckThread: TMCCCheckThread read FMCCCheckThread;
    property DCSCheckThread: TDCSCheckThread read FDCSCheckThread;
    property SECCheckThread: TSECCheckThread read FSECCheckThread;

{    procedure SECCheckSetExecute;
    procedure SECCheckSetComplete;

    procedure SECCheckResetExecute;
    procedure SECCheckResetComplete;

    procedure SECCheckSetExecute;
    procedure SECCheckSetExecute;  }

    property DCSEventThread: TDCSEventThread read FDCSEventThread;
    property SECEventThread: TSECEventThread read FSECEventThread;
    property MCCEventThread: TMCCEventThread read FMCCEventThread;

    property DCSMediaThread: TDCSMediaThread read FDCSMediaThread;

    property DCSDeviceThread: TDCSDeviceThread read FDCSDeviceThread;
  end;

{  TTimerThread = class(TThread)
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
  end; }

  TCrossCheckThread = class(TThread)
  private
    { Private declarations }
    FSEC: TfrmSEC;

    FUDPIn: TUDPIn;

    FIsCommand: Boolean;
    FCMD1, FCMD2: Byte;

    FSyncMsgEvent: THandle;
    FRecvBuffer: AnsiString;
    FRecvData: AnsiString;
    FLastResult: Integer;

    FCommandCritSec: TCriticalSection;
    FInCritSec: TCriticalSection;

    FNumCrossCheck: Word;

    function SendCommand(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
    function SendResponse(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
    function SendAck(AHostIP: AnsiString; APort: Word): Integer;
    function SendNak(AHostIP: AnsiString; APort: Word; ANakError: Byte): Integer;
    function SendError(AHostIP: AnsiString; APort: Word; AErrorCode: Integer): Integer;

    procedure UDPInRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);

    procedure MainDeviceControlBy;
    procedure MainChange;

    procedure DoMainCheck;
    procedure DoCrossCheck;
  protected
    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    function TransmitCommand(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
    function TransmitResponse(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
    function TransmitAck(AHostIP: AnsiString; APort: Word): Integer;
    function TransmitNak(AHostIP: AnsiString; APort: Word; ANakError: Byte): Integer;
    function TransmitError(AHostIP: AnsiString; APort: Word; AErrorCode: Integer): Integer;

    function SECIsAlive(AHostIP: AnsiString; var AIsAlive: Boolean): Integer;
    function SECIsMain(AHostIP: AnsiString; var AIsMain: Boolean): Integer;
    function SECSetMain(AHostIP: AnsiString; AMainSECID: Word): Integer;
  end;

  TMCCCheckThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FExecuteEvent: THandle;
    FCompleteEvent: THandle;

    FEventClose: THandle;
  protected
    procedure DoStartCheck;
    procedure DoMCCCheck;

    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;

    procedure SetExecute;
    procedure SetComplete;

    procedure ResetExecute;
    procedure ResetComplete;

    procedure WaitExecute;
    procedure WaitComplete;
  end;

  TSECCheckThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FEventClose: THandle;
  protected
    procedure DoSECCheck;
    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;
  end;

  TDCSCheckThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FExecuteEvent: THandle;
    FCompleteEvent: THandle;

    FEventClose: THandle;
  protected
    procedure DoDCSCheck;
    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;

    procedure SetExecute;
    procedure SetComplete;

    procedure ResetExecute;
    procedure ResetComplete;

    procedure WaitExecute;
    procedure WaitComplete;
  end;

  TDCSCommandType = (ctNone, ctControlBy, {ctDelete, }ctClear, ctTake, ctHold, ctChangeDuration,
                     ctOnAirEventID, ctEventInfo, ctExist, ctSize,
                     ctSetXpt);

  TDCSCommand = record
//    DCSID: Word;
//    DeviceHandle: TDeviceHandle;
    SourceHandle: PSourceHandle;
    CommandType: TDCSCommandType;
		case Integer of
		  0:
      (
        EventID: TEventID;
        Timecode: TTimecode;
      );
		  1:
      (
        MediaID: array[0..MEDIAID_LEN] of Char;
      );
		  2:
      (
        ChannelID: Word;
      );
      3:
      (
        XptInput: Integer;
      );
  end;
  PDCSCommand = ^TDCSCommand;

  TDCSEventType = (etNone, etInput, etDelete);
  TDCSEvent = record
//    DCSID: Word;
//    DeviceHandle: TDeviceHandle;
    SourceHandle: PSourceHandle;
    EventType: TDCSEventType;
		case Integer of
		  0:
      (
        Event: TEvent;
      );
		  1:
      (
        EventID: TEventID;
      );
		  2:
      (
        ChannelID: Word;
      );
  end;
  PDCSEvent = ^TDCSEvent;
  TDCSEventList = TList<PDCSEvent>;

  TDCSMediaCheck = record
//    DCSID: Word;
//    DeviceHandle: TDeviceHandle;
    SourceHandle: PSourceHandle;
    EventID: TEventID;
    MediaID: array[0..MEDIAID_LEN] of Char;
    InTC: TTimecode;
    DurationTC: TTimecode;
  end;
  PDCSMediaCheck = ^TDCSMediaCheck;
  TDCSMediaCheckList = TList<PDCSMediaCheck>;

  TDCSEventThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FEventLock: TCriticalSection;
    FEventQueue: TDCSEventList;

    FMediaCheckLock: TCriticalSection;
    FMediaCheckQueue: TDCSMediaCheckList;
    FMediaCheckIndex: Integer;

    FInputIndex: Integer;
    FInputCount: Integer;

    FEventCommand: THandle;
    FEventProc: THandle;
    FEventMediaCheck: THandle;

    FEventCompleted: THandle;

    FEventLoopProcStop: Boolean;

    FCommand: TDCSCommand;

    FCMD_Result: Integer;
    FCMD_OnAirEventID: TEventID;
    FCMD_NextEventID: TEventID;
    FCMD_StartTime: TEventTime;
    FCMD_DurationTC: TTimecode;
    FCMD_Exist: Boolean;



    FDeleteEvent: THandle;
    FClearEvent: THandle;
    FTakeEvent: THandle;
    FHoldEvent: THandle;
    FChangeDurationEvent: THandle;
    FEventClose: THandle;

    function GetEventByID(AID:Word; AHandle: TDeviceHandle; AEventID: TEventID): PDCSEvent;

    procedure ClearEventQueueByChannelID(ASourceHandle: PSourceHandle; AChannelID: Word);
    procedure ClearEventQueue;

    procedure ClearMediaCheckQueueBySourcdeHandle(ASourceHandle: PSourceHandle);
    procedure ClearMediaCheckQueueByChannelID(ASourceHandle: PSourceHandle; AChannelID: Word);
    procedure ClearMediaCheckQueue;

    procedure EventQueueCheck;
    procedure MediaCheckQueueCheck;

    procedure SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType); overload;
    procedure SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType; AEventID: TEventID; ATimecode: TTimecode = 0); overload;
    procedure SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType; AChannelID: Word); overload;
    procedure SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType; AMediaID: String); overload;
    procedure SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType; AXptInput: Integer); overload;

    procedure ResetCommand;

    procedure DoEventCommand;
    procedure DoEventProc;
    procedure DoEventMediaCheck;
    procedure DoEventDeviceCheck;

//    procedure DoEventDelete;
//    procedure DoEventTake;
//    procedure DoEventHold;
//    procedure DoEventChangeDuration;
  protected
    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;

    procedure ResetMediaCheckIndex;

    function SetControlBy(ASourceHandle: PSourceHandle): Integer;

    function GetExist(ASourceHandle: PSourceHandle; AMediaID: String; var AExist: Boolean): Integer;
    function GetSize(ASourceHandle: PSourceHandle; AMediaID: String; var ADuration: TTimecode): Integer;

    function SetXpt(ASourceHandle: PSourceHandle; AInput: Integer): Integer;

    function InputEvent(ASourceHandle: PSourceHandle; AEvent: TEvent): Integer;
    function DeleteEvent(ASourceHandle: PSourceHandle; AEventID: TEventID): Integer;
    function ClearEvent(ASourceHandle: PSourceHandle; AChannelID: Word): Integer;
    function MediaCheck(ASourceHandle: PSourceHandle; AEventID: TEventID; AMediaID: String; AInTC, ADurationTC: TTimecode): Integer;

    function TakeEvent(ASourceHandle: PSourceHandle; AEventID: TEventID; ADelayTime: TTimecode): Integer;
    function HoldEvent(ASourceHandle: PSourceHandle; AEventID: TEventID): Integer;
    function ChangeDurationEvent(ASourceHandle: PSourceHandle; AEventID: TEventID; ADurTimecode: TTimecode): Integer;

    function GetOnAirEventID(ASourceHandle: PSourceHandle; var AOnAirEventID: TEventID; var ANextEventID: TEventID): Integer;
    function GetEventInfo(ASourceHandle: PSourceHandle; AEventID: TEventID; var AStartTime: TEventTime; var ADurationTC: TTimecode): Integer;
  end;

  TSECCommandType = (sctNone, sctDeviceError, sctDeviceStatus, sctEventStatus, sctMediaStatus,
                     sctOnAir, sctCurr, sctNext, sctTarget, sctClear, sctBeginUpdate, sctEndUpdate);

  TSECCommand = record
    SEC: PSEC;
    CommandType: TSECCommandType;
		case Integer of
		  0:
      (
        DeviceStatus: TDeviceStatus;
        case Integer of
          0:
          (
            DeviceName: array[0..DEVICENAME_LEN] of Char;
          );
          1:
          (
            DCSID: Word;
            DeviceHandle: TDeviceHandle;
          );
      );
		  1:
      (
        EventID: TEventID;
        case Integer of
          0:
          (
            EventStatus: TEventStatus;
          );
          1:
          (
            MediaStatus: TMediaStatus;
          );
      );
		  2:
      (
        ChannelID: Word;
        case Integer of
          0:
          (
            IsOnAir: Boolean;
          );
          1:
          (
            Index: Integer;
            Count: Integer;
          );
      );
  end;
  PSECCommand = ^TSECCommand;

  TSECEventType = (setNone, setInput, setDelete);
  TSECEvent = record
    SEC: PSEC;
    EventType: TSECEventType;
		case Integer of
		  0:
      (
        Index: Integer;
        Item: TCueSheetItem;
      );
		  1:
      (
        EventID: TEventID;
      );
		  2:
      (
        ChannelID: Word;
      );
  end;
  PSECEvent = ^TSECEvent;
  TSECEventList = TList<PSECEvent>;

  TSECEventThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FEventLock: TCriticalSection;
    FEventQueue: TSECEventList;

    FInputIndex: Integer;
    FInputCount: Integer;

    FEventCommand: THandle;
    FEventProc: THandle;

    FEventCompleted: THandle;

    FEventLoopProcStop: Boolean;

    FCommand: TSECCommand;

    FCMD_Result: Integer;
    FCMD_OnAirEventID: TEventID;
    FCMD_NextEventID: TEventID;
    FCMD_StartTime: TEventTime;
    FCMD_DurationTC: TTimecode;
    FCMD_Exist: Boolean;

    FEventClose: THandle;

    FIsCancel: Boolean;

    procedure ClearEventQueueByChannelID(ASEC: PSEC; AChannelID: Word);
    procedure ClearEventQueue;

    procedure EventQueueCheck;

    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType); overload;

    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; ADeviceStatus: TDeviceStatus; ADeviceName: PChar); overload;
    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; ADeviceStatus: TDeviceStatus; ADCSID: Word; ADeviceHandle: TDeviceHandle); overload;

    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AEventID: TEventID); overload;
    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AEventID: TEventID; AEventStatus: TEventStatus); overload;
    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AEventID: TEventID; AMediaStatus: TMediaStatus); overload;

    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AChannelID: Word); overload;
    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AChannelID: Word; AIsOnAir: Boolean); overload;
    procedure SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AChannelID: Word; AIndex: Integer; ACount: Integer = 1); overload;

    procedure ResetCommand;

    procedure DoEventCommand;
    procedure DoEventProc;
  protected
    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;

    function SetDeviceCommError(ASEC: PSEC; ADeviceStatus: TDeviceStatus; ADeviceName: String): Integer;
    function SetDeviceStatus(ASEC: PSEC; ADeviceStatus: TDeviceStatus; ADCSID: Word; ADeviceHandle: TDeviceHandle): Integer;

    function SetEventStatus(ASEC: PSEC; AEventID: TEventID; AEventStatus: TEventStatus): Integer;
    function SetMediaStatus(ASEC: PSEC; AEventID: TEventID; AMediaStatus: TMediaStatus): Integer;

    function SetOnAir(ASEC: PSEC; AChannelID: Word; AIsOnAir: Boolean): Integer;
    function SetCueSheetCurr(ASEC: PSEC; AEventID: TEventID): Integer;
    function SetCueSheetNext(ASEC: PSEC; AEventID: TEventID): Integer;
    function SetCueSheetTarget(ASEC: PSEC; AEventID: TEventID): Integer;

    function InputCueSheet(ASEC: PSEC; AIndex: Integer; AItem: TCueSheetItem): Integer;
    function DeleteCueSheet(ASEC: PSEC; AEventID: TEventID): Integer;
    function ClearCueSheet(ASEC: PSEC; AChannelID: Word): Integer;
  end;

  TMCCCommandType = (mctNone, mctDeviceError, mctDeviceStatus, mctEventStatus, mctMediaStatus,
                     mctOnAir, mctCurr, mctNext, mctTarget, mctClear, mctBeginUpdate, mctEndUpdate);

  TMCCCommand = record
    MCC: PMCC;
    CommandType: TMCCCommandType;
		case Integer of
		  0:
      (
        DeviceStatus: TDeviceStatus;
        case Integer of
          0:
          (
            DeviceName: array[0..DEVICENAME_LEN] of Char;
          );
          1:
          (
            DCSID: Word;
            DeviceHandle: TDeviceHandle;
          );
      );
		  1:
      (
        EventID: TEventID;
        case Integer of
          0:
          (
            EventStatus: TEventStatus;
          );
          1:
          (
            MediaStatus: TMediaStatus;
          );
      );
		  2:
      (
        ChannelID: Word;
        case Integer of
          0:
          (
            IsOnAir: Boolean;
          );
          1:
          (
            Index: Integer;
            Count: Integer;
          );
      );
  end;
  PMCCCommand = ^TMCCCommand;

  TMCCEventType = (metNone, metInput, metDelete);
  TMCCEvent = record
    MCC: PMCC;
    EventType: TMCCEventType;
		case Integer of
		  0:
      (
        Index: Integer;
        Item: TCueSheetItem;
      );
		  1:
      (
        EventID: TEventID;
      );
		  2:
      (
        ChannelID: Word;
      );
  end;
  PMCCEvent = ^TMCCEvent;
  TMCCEventList = TList<PMCCEvent>;

  TMCCEventThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FEventLock: TCriticalSection;
    FEventQueue: TMCCEventList;

    FInputIndex: Integer;
    FInputCount: Integer;

    FEventCommand: THandle;
    FEventProc: THandle;

    FEventCompleted: THandle;

    FEventLoopProcStop: Boolean;

    FCommand: TMCCCommand;

    FCMD_Result: Integer;
    FCMD_OnAirEventID: TEventID;
    FCMD_NextEventID: TEventID;
    FCMD_StartTime: TEventTime;
    FCMD_DurationTC: TTimecode;
    FCMD_Exist: Boolean;

    FEventClose: THandle;

    FIsCancel: Boolean;

    procedure ClearEventQueueByChannelID(AMCC: PMCC; AChannelID: Word);
    procedure ClearEventQueue;

    procedure EventQueueCheck;

    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType); overload;

    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; ADeviceStatus: TDeviceStatus; ADeviceName: PChar); overload;
    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; ADeviceStatus: TDeviceStatus; ADCSID: Word; ADeviceHandle: TDeviceHandle); overload;

    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AEventID: TEventID); overload;
    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AEventID: TEventID; AEventStatus: TEventStatus); overload;
    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AEventID: TEventID; AMediaStatus: TMediaStatus); overload;

    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AChannelID: Word); overload;
    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AChannelID: Word; AIsOnAir: Boolean); overload;
    procedure SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AChannelID: Word; AIndex: Integer; ACount: Integer = 1); overload;

    procedure ResetCommand;

    procedure DoEventCommand;
    procedure DoEventProc;
  protected
    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;

    function SetDeviceCommError(AMCC: PMCC; ADeviceStatus: TDeviceStatus; ADeviceName: String): Integer;
    function SetDeviceStatus(AMCC: PMCC; ADeviceStatus: TDeviceStatus; ADCSID: Word; ADeviceHandle: TDeviceHandle): Integer;

    function SetEventStatus(AMCC: PMCC; AEventID: TEventID; AEventStatus: TEventStatus): Integer;
    function SetMediaStatus(AMCC: PMCC; AEventID: TEventID; AMediaStatus: TMediaStatus): Integer;

    function SetOnAir(AMCC: PMCC; AChannelID: Word; AIsOnAir: Boolean): Integer;
    function SetCueSheetCurr(AMCC: PMCC; AEventID: TEventID): Integer;
    function SetCueSheetNext(AMCC: PMCC; AEventID: TEventID): Integer;
    function SetCueSheetTarget(AMCC: PMCC; AEventID: TEventID): Integer;

    function InputCueSheet(AMCC: PMCC; AIndex: Integer; AItem: TCueSheetItem): Integer;
    function DeleteCueSheet(AMCC: PMCC; AEventID: TEventID): Integer;
    function ClearCueSheet(AMCC: PMCC; AChannelID: Word): Integer;
  end;

  TDCSMediaThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FMediaCheckEvent: THandle;
    FCloseEvent: THandle;
  protected
    procedure DoMediaCheck;

    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;
    procedure MediaCheck;
  end;

  TDCSDeviceThread = class(TThread)
  private
    FSEC: TfrmSEC;

    FDeviceCheckEvent: THandle;
    FCloseEvent: THandle;

    FWarningCommDialog: TfrmWarningDialog;
  protected
    procedure DoDeviceCheck;

    procedure Execute; override;
  public
    constructor Create(ASEC: TfrmSEC);
    destructor Destroy; override;

    procedure Terminate;
    procedure DeviceCheck;
  end;

var
  frmSEC: TfrmSEC;

  ServerSysRecvBuffer: AnsiString;
  ServerSysCritSec: TCriticalSection;

  ServerRecvBuffer: AnsiString;
  ServerCritSec: TCriticalSection;

  procedure ServerSysRead(ABindingIP: PAnsiChar; AData: PAnsiChar; ADataSize: Integer); stdcall;
  procedure ServerRead(ABindingIP: PAnsiChar; AData: PAnsiChar; ADataSize: Integer); stdcall;

  // 0X00 System Control
  function ServerSysIsAlive(var AIsalive: Boolean): Integer;            // 0X00.00
  function ServerSysIsMain(var AIsMain: Boolean): Integer;              // 0X00.01
  function ServerSysSetMain(ABuffer: AnsiString): Integer;              // 0X00.10

  function ServerOpen(ABuffer: AnsiString): Integer;                    // 0X00.00
  function ServerClose(ABuffer: AnsiString): Integer;                   // 0X00.01
  function ServerIsAlive(ABuffer: AnsiString): Integer;                 // 0X00.10
  function ServerIsMain(ABuffer: AnsiString; var AIsMain: Boolean): Integer;                  // 0X00.11
  function ServerSetMain(ABuffer: AnsiString): Integer; // 0X00.12

  // 0X20 Preset/Select Commands

  // 0X30 Sense Queries

    // 0X40 CueSheet Control
  function ServerBeginUpdate(ABuffer: AnsiString): Integer;             // 0X40.00
  function ServerEndUpdate(ABuffer: AnsiString): Integer;               // 0X40.01

  function ServerSetDeviceCommError(ABuffer: AnsiString): Integer;      // 0X40.02
  function ServerSetDeviceStatus(ABuffer: AnsiString): Integer;         // 0X40.03

  function ServerSetOnAir(ABuffer: AnsiString): Integer;                // 0X40.10
  function ServerSetEventStatus(ABuffer: AnsiString): Integer;          // 0X40.11
  function ServerSetMediaStatus(ABuffer: AnsiString): Integer;          // 0X40.12
  function ServerSetTimelineRange(ABuffer: AnsiString): Integer;        // 0X40.13

  function ServerInputCueSheet(ABuffer: AnsiString): Integer;           // 0X40.20
  function ServerDeleteCueSheet(ABuffer: AnsiString): Integer;          // 0X40.21
  function ServerClearCueSheet(ABuffer: AnsiString): Integer;           // 0X40.22

  function ServerSetCueSheetCurr(ABuffer: AnsiString): Integer;         // 0X40.30
  function ServerSetCueSheetNext(ABuffer: AnsiString): Integer;         // 0X40.31
  function ServerSetCueSheetTarget(ABuffer: AnsiString): Integer;       // 0X40.32

  procedure DeviceStatusNotify(ADCSIP: PChar; AHandle: TDeviceHandle; AStatus: TDeviceStatus); stdcall;
  procedure EventStatusNotify(ADCSIP: PChar; AEventID: TEventID; AStatus: TEventStatus); stdcall;
  procedure EventOverallNotify(ADCSIP: PChar; AHandle: TDeviceHandle; AOverall: TEventOverall); stdcall;
  procedure SECUpdateActnMenusProc;

  procedure TimerCallBack(uTimer, uMessage: UINT; dwUser, dw1, dw2: DWORD); stdcall;

  function CreateWarningDialog(AText: String): TfrmWarningDialog;

implementation

uses UnitStartSplash, UnitEditEvent, UnitCheckStart;

{$R *.dfm}

procedure ServerSysRead(ABindingIP: PAnsiChar; AData: PAnsiChar; ADataSize: Integer);
var
  SaveLen: Integer;
  ByteCount: Integer;
  CMD1, CMD2: Byte;
  RecvData: AnsiString;
  R: Integer;
  SendBuffer: AnsiString;

  IsAlive: Boolean;
  IsMain: Boolean;

//label
//  GotoProcess;
begin
  if (HasMainControl) then exit;

  ServerSysCritSec.Enter;
  try
    if (ADataSize <= 0) then exit;

    SaveLen := Length(ServerSysRecvBuffer);
    SetLength(ServerSysRecvBuffer, SaveLen + ADataSize);
    Move(AData^, ServerSysRecvBuffer[SaveLen + 1], ADataSize);

    if Length(ServerSysRecvBuffer) < 1 then exit;

  //  GotoProcess:

    case Ord(ServerSysRecvBuffer[1]) of
      $02:
      begin
        if (Length(ServerSysRecvBuffer) < 3) then exit;

        ByteCount := PAnsiCharToWord(@ServerSysRecvBuffer[2]);
        if (Length(ServerSysRecvBuffer) >= ByteCount + 4) then
        begin
          if (CheckSum(ServerSysRecvBuffer)) then
          begin
            CMD1 := Ord(ServerSysRecvBuffer[4]);
            CMD2 := Ord(ServerSysRecvBuffer[5]);

            RecvData := System.Copy(ServerSysRecvBuffer, 6, ByteCount - 2);

            case CMD1 of
              $00: // System Control (0X00)
              begin
                case CMD2 of
                  $00: // Is Alive
                  begin
                    R := ServerSysIsAlive(IsAlive);
                    if (R = D_OK) then
                      SendBuffer := BoolToAnsiString(IsAlive);
                  end;
                  $01: // Is Main
                  begin
                    R := ServerSysIsMain(IsMain);
                    if (R = D_OK) then
                      SendBuffer := BoolToAnsiString(IsMain);
                  end;
                  $10: // Set Main
                  begin
                    R := ServerSysSetMain(RecvData);
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
                    $01: SECSysTransmitResponse(ABindingIP, GV_SettingSEC.SysInPort, CMD1, CMD2, PAnsiChar(SendBuffer), Length(SendBuffer));
                    $10: SECSysTransmitAck(ABindingIP, GV_SettingSEC.SysInPort);
                  end;
                end;
              end;
            end
            else
            begin
              // Sending error code
  //            SendBuffer := AnsiChar(D_ERR) + IntToAnsiString(R);
  //            FUDPOut.Send(ABindingIP, FUDPOut.Port, SendBuffer);
              SECSysTransmitError(ABindingIP, GV_SettingSEC.SysInPort, R);
            end;
          end
          else
          begin
  //          FUDPOut.Send(ABindingIP, AnsiChar(D_NAK) + IntToAnsiString(E_NAK_CHECKSUM));
            SECSysTransmitNak(ABindingIP, GV_SettingSEC.SysInPort, E_NAK_CHECKSUM);
          end;

  {        if (Length(ServerRecvBuffer) > ByteCount + 4) then
          begin
            ServerRecvBuffer := Copy(ServerRecvBuffer, ByteCount + 5, Length(ServerRecvBuffer));
            Goto GotoProcess;
          end
          else }
            ServerSysRecvBuffer := '';
        end;
      end;
      else
      begin
  //      FUDPOut.Send(ABindingIP, AnsiChar(D_ERR) + IntToAnsiString(E_UNDEFIND_COMMAND));
        SECSysTransmitError(ABindingIP, GV_SettingSEC.SysInPort, E_UNDEFIND_COMMAND);
        ServerSysRecvBuffer := '';
      end;
    end;
  finally
    ServerSysCritSec.Leave;
  end;
end;

procedure ServerRead(ABindingIP: PAnsiChar; AData: PAnsiChar; ADataSize: Integer);
var
  SaveLen: Integer;
  ByteCount: Integer;
  CMD1, CMD2: Byte;
  RecvData: AnsiString;
  R: Integer;
  SendBuffer: AnsiString;

  IsMain: Boolean;

//label
//  GotoProcess;
begin
{    Assert(False, GetMainLogStr(lsError, Format('ABindingIP = %s', [ABindingIP])));
    Assert(False, GetMainLogStr(lsError, Format('AData = %s', [AData])));
    Assert(False, GetMainLogStr(lsError, Format('ADataSize = %d', [ADataSize])));

    exit;  }

  if (HasMainControl) then exit;

  ServerCritSec.Enter;
  try
    if (ADataSize <= 0) then exit;

    SaveLen := Length(ServerRecvBuffer);
    SetLength(ServerRecvBuffer, SaveLen + ADataSize);
    Move(AData^, ServerRecvBuffer[SaveLen + 1], ADataSize);

    if Length(ServerRecvBuffer) < 1 then exit;

  //  GotoProcess:

    case Ord(ServerRecvBuffer[1]) of
      $02:
      begin
        if (Length(ServerRecvBuffer) < 2) then exit;

        ByteCount := PAnsiCharToWord(@ServerRecvBuffer[2]);
        if (Length(ServerRecvBuffer) >= ByteCount + 4) then
        begin
          if CheckSum(ServerRecvBuffer) then
          begin
            CMD1 := Ord(ServerRecvBuffer[4]);
            CMD2 := Ord(ServerRecvBuffer[5]);

            RecvData := System.Copy(ServerRecvBuffer, 6, ByteCount - 2);

            case CMD1 of
              $00: // 0X00 System Control
              begin
              end;
              $10: // 0X10 Immediate Control
              begin
              end;
              $20: // 0X20 Preset/Select Commands
              begin
              end;
              $30: // 0X30 Sense Queries
              begin
              end;
              $40: // 0X40 Cuesheet Control
              begin
                case CMD2 of
                  $00: // Begin Update
                  begin
                    R := ServerBeginUpdate(RecvData);
                  end;
                  $01: // End Update
                  begin
                    R := ServerEndUpdate(RecvData);
                  end;
                  $02: // Set DeviceCommError
                  begin
                    R := ServerSetDeviceCommError(RecvData);
                  end;
                  $03: // Set DeviceStatus
                  begin
                    R := ServerSetDeviceStatus(RecvData);
                  end;
                  $10: // Set OnAir
                  begin
                    R := ServerSetOnAir(RecvData);
                  end;
                  $11: // Set EventStatus
                  begin
                    R := ServerSetEventStatus(RecvData);
                  end;
                  $12: // Set MediaStatus
                  begin
                    R := ServerSetMediaStatus(RecvData);
                  end;
                  $13: // Set Timeline Range
                  begin
                    R := ServerSetTimelineRange(RecvData);
                  end;
                  $20: // Input CueSheet
                  begin
                    R := ServerInputCueSheet(RecvData);
  //                  if R = D_OK then SendBuffer := AnsiChar($04);
                  end;
                  $21: // Delete CueSheet
                  begin
                    R := ServerDeleteCueSheet(RecvData);
                  end;
                  $22: // Clear CueSheet
                  begin
                    R := ServerClearCueSheet(RecvData);
                  end;
                  $30: // Set Current CueSheet
                  begin
                    R := ServerSetCueSheetCurr(RecvData);
                  end;
                  $31: // Set Next CueSheet
                  begin
                    R := ServerSetCueSheetNext(RecvData);
                  end;
                  $32: // Set Target CueSheet
                  begin
                    R := ServerSetCueSheetTarget(RecvData);
                  end;
                end;
              end;
            end;

            if (R = D_OK) then
            begin
  //            TransmitResponse(ABindingIP, FUDPOut.Port, CMD1, CMD2, SendBuffer, Length(SendBuffer));
              case CMD1 of
                $00: // System Control (0X00)
                begin
                end;
                $10: // Immediate Control (0X10)
                begin
                end;
                $20: // Preset/Select Commands (0X20)
                begin
                end;
                $30: // Sense Queries (0X30)
                begin
                end;
                $40: // Event Control (0X40)
                begin
                  case CMD2 of
                    $00,
                    $01,
                    $02,
                    $03,
                    $10,
                    $11,
                    $12: SECTransmitAck(ABindingIP, GV_SettingSEC.InPort);
                    $20,
                    $21,
                    $22: SECTransmitAck(ABindingIP, GV_SettingSEC.InPort);
                    $30,
                    $31,
                    $32,
                    $34: SECTransmitAck(ABindingIP, GV_SettingSEC.InPort);
                  end;
                end;
              end;
            end
            else
            begin
              // Sending error code
  //            SendBuffer := AnsiChar(D_ERR) + IntToAnsiString(R);
  //            FUDPOut.Send(ABindingIP, FUDPOut.Port, SendBuffer);
              SECTransmitError(ABindingIP, GV_SettingSEC.InPort, R);
            end;
          end
          else
          begin
  //          FUDPOut.Send(ABindingIP, AnsiChar(D_NAK) + IntToAnsiString(E_NAK_CHECKSUM));
            SECTransmitNak(ABindingIP, GV_SettingSEC.InPort, E_NAK_CHECKSUM);
          end;

  {        if (Length(ServerRecvBuffer) > ByteCount + 4) then
          begin
            ServerRecvBuffer := Copy(ServerRecvBuffer, ByteCount + 5, Length(ServerRecvBuffer));
            Goto GotoProcess;
          end
          else }
            ServerRecvBuffer := '';
        end;
      end;
      else
      begin
  //      FUDPOut.Send(ABindingIP, AnsiChar(D_ERR) + IntToAnsiString(E_UNDEFIND_COMMAND));
        SECTransmitError(ABindingIP, GV_SettingSEC.InPort, E_UNDEFIND_COMMAND);
        ServerRecvBuffer := '';
      end;
    end;
  finally
    ServerCritSec.Leave;
  end;
end;

// 0X00 System Control
function ServerSysIsAlive(var AIsAlive: Boolean): Integer;
begin
  Result := D_FALSE;

  AIsAlive := True;

  Result := D_OK;
end;

function ServerSysIsMain(var AIsMain: Boolean): Integer;
begin
  Result := D_FALSE;

  if (GV_SECMine <> nil) then
    AIsMain := GV_SECMine^.Main
  else
    AIsMain := False;

  Result := D_OK;
end;

function ServerSysSetMain(ABuffer: AnsiString): Integer;
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
  frmSEC.FCrossCheckThread.FNumCrossCheck := 0;
end;

function ServerOpen(ABuffer: AnsiString): Integer;
var
  AID: Word;
  AName: String;
begin

  Result := E_INVALID_COMPONENT_ID;

  AID := PAnsiCharToWord(@ABuffer[1]);
  AName := Copy(ABuffer, 3, Length(ABuffer));
  if (GV_SettingGeneral.ID = AID) and
     (GV_SettingGeneral.Name = AName) then
  begin
    Result := D_OK;
  end;
end;

function ServerClose(ABuffer: AnsiString): Integer;
var
  AID: Word;
begin
  Result := E_INVALID_COMPONENT_ID;

  AID := PAnsiCharToWord(@ABuffer[1]);
  if (GV_SettingGeneral.ID = AID) then
  begin
    Result := D_OK;
  end;
end;

function ServerIsAlive(ABuffer: AnsiString): Integer;
var
  AID: Word;
begin
  Result := E_INVALID_COMPONENT_ID;

  AID := PAnsiCharToWord(@ABuffer[1]);
  if (GV_SettingGeneral.ID = AID) then
  begin
    Result := D_OK;
  end;
end;

function ServerIsMain(ABuffer: AnsiString; var AIsMain: Boolean): Integer;
var
  AMainID: Word;
  SEC: PSEC;
begin
  AIsMain := (GV_SECMine = GV_SECMain);

  Result := D_OK;
end;

function ServerSetMain(ABuffer: AnsiString): Integer;
var
  AMainID: Word;
  SEC: PSEC;
begin
  Result := E_INVALID_COMPONENT_ID;

  AMainID := PAnsiCharToWord(@ABuffer[1]);

  SEC := GetSECByID(AMainID);
  if (SEC <> nil) then
  begin
    GV_SECMain := SEC;
    Result := D_OK;
  end;
end;

// 0X20 Preset/Select Commands

// 0X30 Sense Queries

// 0X40 CueSheet Control
function ServerBeginUpdate(ABuffer: AnsiString): Integer;
var
  ChannelID: Word;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  ChannelID := PAnsiCharToWord(@ABuffer[1]);

  ChannelForm := frmSEC.GetChannelFormByID(ChannelID);
  if (ChannelForm <> nil) then
  begin
    Assert(False, 'SECBeginUpdate');
    Result := ChannelForm.SECBeginUpdateW;
  end;
end;

function ServerEndUpdate(ABuffer: AnsiString): Integer;
var
  ChannelID: Word;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  ChannelID := PAnsiCharToWord(@ABuffer[1]);

  ChannelForm := frmSEC.GetChannelFormByID(ChannelID);
  if (ChannelForm <> nil) then
  begin
    Assert(False, 'SECEndUpdate');
    Result := ChannelForm.SECEndUpdateW;
  end;
end;

function ServerSetDeviceCommError(ABuffer: AnsiString): Integer;
var
  DeviceStatus: TDeviceStatus;
  DeviceName: String;
  DeviceForm: TfrmDevice;
begin
  Result := D_FALSE;

  Move(ABuffer[1], DeviceStatus, SizeOf(TDeviceStatus));
  DeviceName := Copy(ABuffer, SizeOf(TDeviceStatus) + 1, Length(ABuffer));

  DeviceForm := frmSEC.GetDeviceForm;
  if (DeviceForm <> nil) then
    Result := DeviceForm.SECSetDeviceCommErrorW(DeviceStatus, DeviceName);
end;

function ServerSetDeviceStatus(ABuffer: AnsiString): Integer;
var
  DCSID: Word;

  DeviceHandle: TDeviceHandle;
  DeviceStatus: TDeviceStatus;
  DeviceForm: TfrmDevice;
begin
  Result := D_FALSE;

  DCSID := PAnsiCharToWord(@ABuffer[1]);
  DeviceHandle := PAnsiCharToInt(@ABuffer[3]);
  Move(ABuffer[7], DeviceStatus, SizeOf(TDeviceStatus));

  DeviceForm := frmSEC.GetDeviceForm;
  if (DeviceForm <> nil) then
    Result := DeviceForm.SECSetDeviceStatusW(DCSID, DeviceHandle, DeviceStatus);
end;

function ServerSetOnAir(ABuffer: AnsiString): Integer;
var
  ChannelID: Word;
  IsOnAir: Boolean;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  ChannelID := PAnsiCharToWord(@ABuffer[1]);
  IsOnAir   := PAnsiCharToBool(@ABuffer[3]);

  ChannelForm := frmSEC.GetChannelFormByID(ChannelID);

  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECSetOnAirW(IsOnAir);
  end;
end;

function ServerSetEventStatus(ABuffer: AnsiString): Integer;
var
  EventID: TEventID;
  EventStatus: TEventStatus;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  Move(ABuffer[1], EventID, SizeOf(TEventID));
  Move(ABuffer[SizeOf(TEventID) + 1], EventStatus, SizeOf(TEventStatus));

  ChannelForm := frmSEC.GetChannelFormByID(EventID.ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECSetEventStatusW(EventID, EventStatus);
  end;
end;

function ServerSetMediaStatus(ABuffer: AnsiString): Integer;
var
  EventID: TEventID;
  MediaStatus: TMediaStatus;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  Move(ABuffer[1], EventID, SizeOf(TEventID));
  Move(ABuffer[SizeOf(TEventID) + 1], MediaStatus, SizeOf(TMediaStatus));

  ChannelForm := frmSEC.GetChannelFormByID(EventID.ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECSetMediaStatusW(EventID, MediaStatus);
  end;
end;

function ServerSetTimelineRange(ABuffer: AnsiString): Integer;
var
  ChannelID: Word;
  StartDate, EndDate: TDateTime;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  ChannelID := PAnsiCharToWord(@ABuffer[1]);

  StartDate := PAnsiCharToDouble(@ABuffer[3]);
  EndDate   := PAnsiCharToDouble(@ABuffer[11]);

  ChannelForm := frmSEC.GetChannelFormByID(ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECSetTimelineRangeW(StartDate, EndDate);
  end;
end;

function ServerInputCueSheet(ABuffer: AnsiString): Integer;
var
  Index: Integer;
  CueSheetItem: TCueSheetItem;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  Index := PAnsiCharToInt(@ABuffer[1]);
  Move(ABuffer[5], CueSheetItem, SizeOf(TCueSheetItem));

  ChannelForm := frmSEC.GetChannelFormByID(CueSheetItem.EventID.ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECInputCueSheetW(Index, CueSheetItem);
  end;
end;

function ServerDeleteCueSheet(ABuffer: AnsiString): Integer;
var
  EventID: TEventID;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  Move(ABuffer[1], EventID, SizeOf(TEventID));

  ChannelForm := frmSEC.GetChannelFormByID(EventID.ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECDeleteCueSheetW(EventID);
  end;
end;

function ServerClearCueSheet(ABuffer: AnsiString): Integer;
var
  ChannelID: Word;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  ChannelID := PAnsiCharToWord(@ABuffer[1]);

  ChannelForm := frmSEC.GetChannelFormByID(ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECClearCueSheetW;
  end;
end;

function ServerSetCueSheetCurr(ABuffer: AnsiString): Integer;
var
  EventID: TEventID;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  Move(ABuffer[1], EventID, SizeOf(TEventID));

  ChannelForm := frmSEC.GetChannelFormByID(EventID.ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECSetCueSheetCurrW(EventID);
  end;
end;

function ServerSetCueSheetNext(ABuffer: AnsiString): Integer;
var
  EventID: TEventID;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  Move(ABuffer[1], EventID, SizeOf(TEventID));

  ChannelForm := frmSEC.GetChannelFormByID(EventID.ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECSetCueSheetNextW(EventID);
  end;
end;

function ServerSetCueSheetTarget(ABuffer: AnsiString): Integer;
var
  EventID: TEventID;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  Move(ABuffer[1], EventID, SizeOf(TEventID));

  ChannelForm := frmSEC.GetChannelFormByID(EventID.ChannelID);
  if (ChannelForm <> nil) then
  begin
    Result := ChannelForm.SECSetCueSheetTargetW(EventID);
  end;
end;

procedure DeviceStatusNotify(ADCSIP: PChar; AHandle: TDeviceHandle; AStatus: TDeviceStatus);
var
  I, J: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;

  DCSID: Word;
  DeviceForm: TfrmDevice;
begin
  if (not HasMainControl) then exit;

  if (not GV_SettingOption.OnAirCheckDeviceNotify) then exit;

  DeviceForm := frmSEC.GetDeviceForm;
  if (DeviceForm <> nil) then
  begin
    DeviceForm.SetDeviceStatus(String(ADCSIP), AHandle, AStatus);
  end;

exit;

{  Source := nil;
  for I := 0 to GV_SourceList.Count - 1 do
  begin
    SourceHandles := GV_SourceList[I]^.Handles;
    if (SourceHandles <> nil) then
    begin
      for J := 0 to SourceHandles.Count - 1 do
        if (SourceHandles[J].Handle = AHandle) then
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
           (SourceHandles[J].Handle = AHandle) then
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
  if (not HasMainControl) then exit;

  if (frmSEC <> nil) and (frmSEC.HandleAllocated) then
    with frmSEC do
    begin
      ChannelForm := GetChannelFormByID(AEventID.ChannelID);
      if (ChannelForm <> nil) then
        ChannelForm.SetEventStatus(AEventID, AStatus);
    end;
end;

procedure EventOverallNotify(ADCSIP: PChar; AHandle: TDeviceHandle; AOverall: TEventOverall);
var
  I: Integer;
  ChannelForm: TfrmChannel;
begin
{  if (frmSEC <> nil) and (frmSEC.HandleAllocated) then
    with frmSEC do
    begin
      ChannelForm := GetChannelFormByID(AOverall.ChannelID);
      if (ChannelForm <> nil) then
        ChannelForm.SetEventOverall(String(ADCSIP), AHandle, AOverall);
    end; }
end;

procedure TfrmSEC.ApplicationMessage(var Msg: TMsg; var Handled: Boolean);
var
  KeyMsg: TWMKey;
begin
{   with Msg do
      case message of
        WM_KEYDOWN:
        begin
//          KeyMsg := TWMKey(TMessage(Msg));
          if (GetKeyState(VK_CONTROL) < 0) then
            if (wParam in [67, 99]) then
            begin
{              if (Screen.ActiveControl is TAdvStringGrid) then
              begin
                (Screen.ActiveControl as TAdvStringGrid).CellEditor.
              end;
            CopyToClipboard; CopyToClipboardEvent
            Handled := True;// }{ShowMessage('1');

            IsShortCut()
            end;
        end;
      end; }

   inherited;
end;

procedure TfrmSEC.ScreenOnActiveControlChange(Sender: TObject);
begin
end;

procedure TfrmSEC.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);

  Params.Style := Params.Style or WS_MINIMIZEBOX;
end;

{procedure TfrmSEC.WndProc(var Message: TMessage);
begin
  inherited;

  case Message.Msg of
    WM_UPDATE_CURRENT_TIME: WMUpdateCurrentTime(Message);
    WM_UPDATE_ACTIVATE: WMUpdateActive(Message);
  end;
end; }

procedure TfrmSEC.WMSettingChange(var Message: TWMSettingChange);
begin
  inherited;

  actMainMenuSEC.Font.Assign(FSaveMenuFont);
  actMainMenuSEC.ColorMap.Assign(FSaveColorMap);
end;

procedure TfrmSEC.WMUpdateCurrentTime(var Message: TMessage);
begin
  with GV_TimeCurrent do
//    lblCurrentTime.Caption := Format('%0.4d-%0.2d-%0.2d %0.2d:%0.2d:%0.2d.%0.3d', [wYear, wMonth, wDay, wHour, wMinute, wSecond, wMilliseconds]);
    lblCurrentTime.Caption := Format('%0.4d-%0.2d-%0.2d %0.2d:%0.2d:%0.2d', [wYear, wMonth, wDay, wHour, wMinute, wSecond]);

//  lblCurrentTime.Caption := FormatDateTime('YYYY-MM-DD hh:nn:ss:zzz', FCurrentTime);
end;

procedure TfrmSEC.WMUpdateActive(var Message: TMessage);
begin
  DisplayActivate;
end;

procedure TfrmSEC.WMWarningDisplayDeviceCheck(var Message: TMessage);
var
  WarningText: String;
begin
  if (FWarningDialogDeviceCheck = nil) or (not FWarningDialogDeviceCheck.HandleAllocated) then
  begin
    WarningText := String(PChar(Message.LParam));
    FWarningDialogDeviceCheck := CreateWarningDialog(WarningText);

    with GV_SettingOption do
      if (OnAirCheckDeviceAlarm) then
        TAlarmThread.Create(OnAirCheckDeviceAlarmFileName, OnAirCheckDeviceAlarmDuration, OnAirCheckDeviceAlarmInterval).Start;
  end;

exit;
{  if (FWarningDialogDeviceCheck = nil) or (not FWarningDialogDeviceCheck.HandleAllocated) then
  begin
    WarningText := PString(Message.LParam);
    try
      FWarningDialogDeviceCheck := CreateWarningDialog(WarningText^);
    finally
      Dispose(WarningText);
    end;
  //  WarningText := String(PChar(Message.LParam));
  //    FWarningDialogDeviceCheck := CreateWarningDialog(WarningText);

    with GV_SettingOption do
      if (OnAirCheckDeviceAlarm) then
        TAlarmThread.Create(OnAirCheckDeviceAlarmFileName, OnAirCheckDeviceAlarmDuration, OnAirCheckDeviceAlarmInterval).Start;
  end; }
end;

function TfrmSEC.GetChannelFormByID(AChannelID: Word): TfrmChannel;
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

function TfrmSEC.GetDeviceForm: TfrmDevice;
begin
  Result := nil;

  if (HandleAllocated) then
  begin
    if (aopDevice.ControlCount > 0) then
    begin
      Result := TfrmDevice(aopDevice.Controls[0]);
    end;
  end;
end;

procedure TfrmSEC.EnableShortCutAction(AEnabled: Boolean);
var
  I: Integer;
begin
  for I := 0 to actManager.ActionCount - 1 do
  begin
    if (actManager.Actions[I].ShortCut > 0) or
       (actManager.Actions[I].SecondaryShortCuts.Count > 0) then
      actManager.Actions[I].Enabled := AEnabled;
  end;

  exit;


  actEditCutEvent.Enabled   := AEnabled;
  actEditPasteEvent.Enabled := AEnabled;

  actControlGotoCurrentEvent.Enabled := AEnabled;

  actEventInsertMainEvent.Enabled := AEnabled;
  actEventInsertJoinEvent.Enabled := AEnabled;
  actEventInsertSubEvent.Enabled := AEnabled;
  actEventInsertCommentEvent.Enabled := AEnabled;
  actEventInsertProgram.Enabled := AEnabled;

  actEventUpdateEvent.Enabled := AEnabled;
  actEventDeleteEvent.Enabled := AEnabled;
  actEventDeleteEvent.Enabled := AEnabled;

  exit;

  if (AEnabled) then
  begin
    actEditCopyEvent.ShortCut := TextToShortCut('ctrl+c');
  end
  else
  begin
    actEditCopyEvent.ShortCut := TextToShortCut('');
  end;
end;

// 0X00 System Control
function TfrmSEC.SECIsAliveW(var AIsAlive: Boolean): Integer;
begin
  Result := D_FALSE;

  AIsAlive := True;

  Result := D_OK;
end;

function TfrmSEC.SECIsMainW(var AIsMain: Boolean): Integer;
begin
  Result := D_FALSE;

  if (GV_SECMine <> nil) then
    AIsMain := GV_SECMine^.Main
  else
    AIsMain := False;

  Result := D_OK;
end;

function TfrmSEC.SECSetMainW(ABuffer: AnsiString): Integer;
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

procedure TfrmSEC.actAllChannelsTimelineGotoCurrentExecute(Sender: TObject);
begin
  inherited;

  if (frmAllChannels <> nil) then
    frmAllChannels.TimelineGotoCurrent;
end;

procedure TfrmSEC.actAllChannelsTimelineMoveLeftExecute(Sender: TObject);
begin
  inherited;

  if (frmAllChannels <> nil) then
    frmAllChannels.TimelineMoveLeft;
end;

procedure TfrmSEC.actAllChannelsTimelineMoveRightExecute(Sender: TObject);
begin
  inherited;

  if (frmAllChannels <> nil) then
    frmAllChannels.TimelineMoveRight;
end;

procedure TfrmSEC.actAllChannelsTimelineZoomInExecute(Sender: TObject);
begin
  inherited;

  if (frmAllChannels <> nil) then
    frmAllChannels.TimelineZoomIn;
end;

procedure TfrmSEC.actAllChannelsTimelineZoomOutExecute(Sender: TObject);
begin
  inherited;

  if (frmAllChannels <> nil) then
    frmAllChannels.TimelineZoomOut;
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

procedure TfrmSEC.actEventSourceExchangeExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.SourceExchangeCueSheet;
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

procedure TfrmSEC.actEditSelectMediaExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.SelectMediaFileInCueSheet;
end;

procedure TfrmSEC.actEditTimelineGotoCurrentExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.TimelineGotoCurrent;
end;

procedure TfrmSEC.actEditTimelineMoveLeftExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.TimelineMoveLeft;
end;

procedure TfrmSEC.actEditTimelineMoveRightExecute(Sender: TObject);
var
  ChannelForm: TfrmChannel;
begin
  inherited;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    ChannelForm.TimelineMoveRight;
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

  if (not HasMainControl) then exit;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
  begin
    if (ChannelForm.ChannelOnAir) then
    begin
      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(SENoCreateCuesheetWhileChannelOnair), PChar(Application.Title), MB_OK or MB_ICONWARNING or MB_TOPMOST);
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

  if (not HasMainControl) then exit;

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

  if (not HasMainControl) then exit;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  with TOpenDialog.Create(Self) do
    try
      Title := 'Open playlist file';
//      Filename := ExtractFileName(FFileName);
      InitialDir := GV_SettingGeneral.WorkCueSheetPath;
      Filter  := 'Playlist file|*.xml|Any file(*.*)|*.*';
      if (Execute) then
      begin
        ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
        if (ChannelForm <> nil) then
        begin
          if (ChannelForm.ChannelOnAir) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SENoOpenCuesheetWhileChannelOnair), PChar(Application.Title), MB_OK or MB_ICONWARNING or MB_TOPMOST);
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

  if (not HasMainControl) then exit;

  if (aoPagerMain.ActivePageIndex <= 0) then exit;

  ChannelForm := GetChannelFormByID(aoPagerMain.ActivePage.Tag);
  if (ChannelForm <> nil) then
    with TSaveDialog.Create(Self) do
      try
        Title := 'Save as playlist file';
        Filename := ExtractFileName(ChannelForm.PlayListFileName);
        DefaultExt := '.xml';
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

  if (not HasMainControl) then exit;

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

procedure TfrmSEC.actMainMenuSECGetControlClass(Sender: TCustomActionBar;
  AnItem: TActionClient; var ControlClass: TCustomActionControlClass);
begin
  inherited;
  actMainMenuSEC.ColorMap.Assign(FSaveColorMap);
end;

procedure TfrmSEC.actControlGotoCurrentEventExecute(Sender: TObject);
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

procedure TfrmSEC.FormCreate(Sender: TObject);
begin
  inherited;
  frmStartSplash := TfrmStartSplash.Create(Self);
  frmStartSplash.Show;
//  Application.ProcessMessages;
//  try
    Initialize;
//  finally
//    frmStartSplash.Close;
//    FreeAndNil(frmStartSplash);
//  end;
end;

procedure TfrmSEC.FormDestroy(Sender: TObject);
begin
  inherited;
  Finalize;
end;

procedure TfrmSEC.FormShortCut(var Msg: TWMKey; var Handled: Boolean);
var
  Grid: TAdvStringGrid;
  ShiftState: TShiftState;
begin
  inherited;
  exit;

  if (Screen.ActiveControl is TAdvStringGrid) then
  begin
    Grid := (Screen.ActiveControl as TAdvStringGrid);
    Handled := Grid.EditMode;

    if (Handled) then
    begin
      ShiftState := KeyDataToShiftState(Msg.KeyData);
      if (ssCtrl in ShiftState) and
         (Msg.CharCode in [Ord('C'), Ord('X'), Ord('V')]) then
      begin
        begin
          case Msg.CharCode of
            Ord('C'): SendMessage(Grid.NormalEdit.Handle, WM_COPY, 0, 0);
            Ord('X'): SendMessage(Grid.NormalEdit.Handle, WM_CUT, 0, 0);
            Ord('V'): SendMessage(Grid.NormalEdit.Handle, WM_PASTE, 0, 0);
          end;
        end;
      end
      else if (Msg.CharCode in [VK_SPACE]) then
//SendMessage(Grid.NormalEdit.Handle, WM_CHAR, VK_SPACE, 0)
         PostKeyEx32(VK_SPACE, ShiftState, False)
      else if (Msg.CharCode in [VK_ESCAPE]) then
//        SendMessage(Grid.NormalEdit.Handle, WM_KEYDOWN, VK_ESCAPE, 0)
         PostKeyEx32(Msg.CharCode, ShiftState, False)
      else
        Handled := False;

//SendMessage(Grid.NormalEdit.Handle, WM_KEYDOWN, Msg.CharCode, 0);
    end;
  end;
{

  SS := KeyDataToShiftState(Msg.KeyData);

    // Ctrl+/ 나 Alt+/ 이면
  if ((ssCtrl in SS) or (ssAlt in SS)) and
     (Msg.CharCode in [Ord('C')]) then
  begin
    // CharCode를 변경한다.
    ShowMessage('1');
//    Msg.CharCode := TextToShortcut('/');

    Handled := False;
  end; }
end;

procedure TfrmSEC.FormShow(Sender: TObject);
var
  TimeCaps: TTimeCaps;
begin
  inherited;

  timeGetDevCaps(@TimeCaps, SizeOf(TTimeCaps));

  FTimerID := timeSetEvent(1, TimeCaps.wPeriodMin, @TimerCallBack, 0, TIME_PERIODIC);

  if (FCrossCheckThread <> nil) then
    FCrossCheckThread.Start;

  if (frmStartSplash <> nil) then
  begin
    frmStartSplash.Close;
    FreeAndNil(frmStartSplash);
  end;
end;

procedure SECUpdateActnMenusProc;
begin
//  ShowMessage(frmSEC.actMainMenuSEC.ColorMap.Name);
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
  frmSEC.actMainMenuSEC.ColorMap := frmSEC.XPColorMap;
  frmSEC.actMainMenuSEC.Color := $00251C19;
  frmSEC.actMainMenuSEC.Font.Name := 'Century Gothic';
  frmSEC.actMainMenuSEC.Font.Color := $00FFBDAD;
  frmSEC.actMainMenuSEC.Font.Size := 9;
//  ShowMessage(Format('%x', [Integer(frmSEC.XPColorMap.FontColor)]));
end;

procedure TimerCallBack(uTimer, uMessage: UINT; dwUser, dw1, dw2: DWORD);
var
  T: TSystemTime;
begin
  GetLocalTime(T);

  if (GV_TimeBefore.wSecond <> T.wSecond) then
  begin
    GV_TimeCurrent := T;
    GV_TimeCurrent.wMilliseconds := 0; // Frame to set 0. Because milli seconds is not 0

    SetEvent(GV_TimerExecuteEvent);
    if (frmSEC <> nil) and (frmSEC.HandleAllocated) and (frmSEC.Showing) then
      PostMessage(frmSEC.Handle, WM_UPDATE_CURRENT_TIME, 0, 0);
    ResetEvent(GV_TimerExecuteEvent);

{    if (GV_TimeCurrent.wMilliseconds <> 0) then
    begin
      Inc(GV_NotZeroCount);
      Form18.Caption := Format('%d', [GV_NotZeroCount]);
    end; }

    GV_TimeBefore := T;
  end;
end;

function CreateWarningDialog(AText: String): TfrmWarningDialog;
var
  H: HWND;
begin
  Result := TfrmWarningDialog.Create(nil);
  Result.SetWarningText(AText);
  Result.Show;

  // If showing message then bring to top
  H := FindWindow('#32770', PChar(Application.Title));
  if (H <> 0) then
  begin
    SetWindowPos(H, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
  end;
end;

procedure TfrmSEC.DisplayStartCheck(ACheckStr: String);
begin
  if (frmStartSplash <> nil) and (frmStartSplash.HandleAllocated) then
    frmStartSplash.DisplayCheck(ACheckStr);
end;

procedure TfrmSEC.Initialize;
var
  R: Integer;
  I: Integer;
  OfficePage: TAdvOfficePage;
  Channel: PChannel;
  ChannelForm: TfrmChannel;

  TimeCaps: TTimeCaps;
begin
  DisplayStartCheck('Loading configuration...');

  FSaveMenuFont := TFont.Create;
  FSaveMenuFont.Assign(actMainMenuSEC.Font);
  FSaveColorMap := TXPColorMap.Create(Self);
  FSaveColorMap.Assign(actMainMenuSEC.ColorMap);


{  Application.OnMessage := ApplicationMessage;
  Screen.OnActiveControlChange := ScreenOnActiveControlChange;
  UpdateActnMenusProc := SECUpdateActnMenusProc; }

  SetEventDeadlineHour(30);

  FIsEditing := False;

  GV_LogCS := TCriticalSection.Create;

  GV_ChannelList := TChannelList.Create;
  GV_SECList := TSECList.Create;
  GV_MCCList := TMCCList.Create;
  GV_DCSList := TDCSList.Create;
  GV_SourceList := TSourceList.Create;
  GV_MCSList := TMCSList.Create;

  GV_ProgramTypeList := TProgramTypeList.Create;

  GV_ClipboardCueSheet := TClipboardCueSheet.Create;

  GV_TimerExecuteEvent := CreateEvent(nil, True, False, nil);
//  GV_TimerExecuteEvent := CreateEvent(nil, False, False, nil);
  GV_TimerCancelEvent  := CreateEvent(nil, True, False, nil);

  ServerSysRecvBuffer := '';
  ServerSysCritSec := TCriticalSection.Create;

  ServerRecvBuffer := '';
  ServerCritSec := TCriticalSection.Create;

  LoadConfig;

  if (GV_SECMine <> nil) then
    Application.Title := Format('Schedule Event Controller %s - %s', [GetFileVersionStr(Application.ExeName), GV_SECMine^.Name])
  else
    Application.Title := Format('Schedule Event Controller %s - Not assigned SEC mine', [GetFileVersionStr(Application.ExeName)]);

  WMTitleBar.Caption := Application.Title;

  DisplayStartCheck('Device opening...');

  with GV_SettingDCS do
  begin
    // System control
    R := DCSSysInitialize(SysInPort, SysOutPort, TimecodeToMilliSec(SysCheckTimeout));
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_DCSSysInitializeFailed, [R, SysInPort, SysOutPort]));

    if (SysLogInPortEnabled) then
    begin
      R := DCSSysLogInPortEnable(PChar(SysLogPath), PChar(SysLogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_DCSSysLogInPortEnableFailed, [R, SysLogPath, SysLogExt]));
    end;

    if (SysLogOutPortEnabled) then
    begin
      R := DCSSysLogOutPortEnable(PChar(SysLogPath), PChar(SysLogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_DCSSysLogOutPortEnableFailed, [R, SysLogPath, SysLogExt]));
    end;

    // Device control
    R := DCSInitialize(NotifyPort, InPort, OutPort, TimecodeToMilliSec(CommandTimeout));
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_DCSInitializeFailed, [R, NotifyPort, InPort, OutPort]));

    if (LogNotifyPortEnabled) then
    begin
      R := DCSLogNotifyPortEnable(PChar(LogPath), PChar(LogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_DCSLogNotifyPortEnableFailed, [R, LogPath, LogExt]));
    end;

    if (LogInPortEnabled) then
    begin
      R := DCSLogInPortEnable(PChar(LogPath), PChar(LogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_DCSLogInPortEnableFailed, [R, LogPath, LogExt]));
    end;

    if (LogOutPortEnabled) then
    begin
      R := DCSLogOutPortEnable(PChar(LogPath), PChar(LogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_DCSLogOutPortEnableFailed, [R, LogPath, LogExt]));
    end;

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

  DisplayStartCheck('MCC opening...');

  with GV_SettingMCC do
  begin
    // System control
    R := MCCSysInitialize(SysInPort, SysOutPort, TimecodeToMilliSec(SysCheckTimeout));
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_MCCSysInitializeFailed, [R, SysInPort]));

    if (SysLogInPortEnabled) then
    begin
      R := MCCSysLogInPortEnable(PChar(SysLogPath), PChar(SysLogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_MCCSysLogInPortEnableFailed, [R, SysLogPath, SysLogExt]));
    end;

    if (SysLogOutPortEnabled) then
    begin
      R := MCCSysLogOutPortEnable(PChar(SysLogPath), PChar(SysLogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_MCCSysLogOutPortEnableFailed, [R, LogPath, LogExt]));
    end;

    // Command control
    R := MCCInitialize(NotifyPort, InPort, OutPort, TimecodeToMilliSec(CommandTimeout));
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_MCCInitializeFailed, [R, NotifyPort, InPort, OutPort]));

    if (LogNotifyPortEnabled) then
    begin
      R := MCCLogNotifyPortEnable(PChar(LogPath), PChar(LogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_MCCLogNotifyPortEnableFailed, [R, LogPath, LogExt]));
    end;

    if (LogInPortEnabled) then
    begin
      R := MCCLogInPortEnable(PChar(LogPath), PChar(LogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_MCCLogInPortEnableFailed, [R, LogPath, LogExt]));
    end;

    if (LogOutPortEnabled) then
    begin
      R := MCCLogOutPortEnable(PChar(LogPath), PChar(LogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_MCCLogOutPortEnableFailed, [R, LogPath, LogExt]));
    end;
  end;

  DisplayStartCheck('SEC opening...');

  with GV_SettingSEC do
  begin
    // System control
    R := SECSysInitialize(SysInPort, TimecodeToMilliSec(SysCheckTimeout));
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_SECSysInitializeFailed, [R, SysInPort]));

    if (SysLogInPortEnabled) then
    begin
      R := SECSysLogInPortEnable(PChar(SysLogPath), PChar(SysLogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_SECSysLogInPortEnableFailed, [R, SysLogPath, SysLogExt]));
    end;

    R := SECSysSetServerReadProc(@ServerSysRead);
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_SECSysSetServerReadProcFailed, [R]));

    // Command control
    R := SECInitialize(InPort, TimecodeToMilliSec(CommandTimeout));
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_SECInitializeFailed, [R, InPort]));

    if (LogInPortEnabled) then
    begin
      R := SECLogInPortEnable(PChar(LogPath), PChar(LogExt));
      if (R <> D_OK) then
        Assert(False, GetMainLogStr(lsError, @LSE_SECLogInPortEnableFailed, [R, LogPath, LogExt]));
    end;

    R := SECSetServerReadProc(@ServerRead);
    if (R <> D_OK) then
      Assert(False, GetMainLogStr(lsError, @LSE_SECSetServerReadProcFailed, [R]));
  end;

  DisplayStartCheck('Main sec system checking...');

  FDCSEventThread := TDCSEventThread.Create(Self);
  FDCSEventThread.Resume;

  FSECEventThread := TSECEventThread.Create(Self);
  FSECEventThread.Resume;

  if (GV_SettingMCC.Use) then
  begin
    FMCCEventThread := TMCCEventThread.Create(Self);
    FMCCEventThread.Resume;
  end;

//  PostMessage(frmSEC.Handle, WM_UPDATE_CURRENT_TIME, 0, 0);

  GV_TimelineOnairIcon := TPicture.Create;
  try
    GV_TimelineOnairIcon.LoadFromFile(GV_SettingOption.TimelineOnairIconFileName);
  except
  end;

  GV_TimelineNextIcon := TPicture.Create;
  try
    GV_TimelineNextIcon.LoadFromFile(GV_SettingOption.TimelineNextIconFileName);
  except
  end;

  GV_TimelineNormalIcon := TPicture.Create;
  try
    GV_TimelineNormalIcon.LoadFromFile(GV_SettingOption.TimelineNormalIconFileName);
  except
  end;

  InitializeAllChannelPage;
  InitializeChannelPage;
  InitializeDevicePage;

//  SECMainCheck;

  FCrossCheckThread := TCrossCheckThread.Create(Self);
  FCrossCheckThread.DoMainCheck;

  FDCSCheckThread := TDCSCheckThread.Create(Self);
  FDCSCheckThread.Start;

  if (GV_SettingMCC.Use) then
  begin
    FMCCCheckThread := TMCCCheckThread.Create(Self);
    FMCCCheckThread.DoStartCheck;
    FMCCCheckThread.Start;
  end;

  FSECCheckThread := TSECCheckThread.Create(Self);
  FSECCheckThread.Start;

//  FDCSMediaThread := TDCSMediaThread.Create(Self);
//  FDCSMediaThread.Resume;

  FDCSDeviceThread := TDCSDeviceThread.Create(Self);
  FDCSDeviceThread.Start;

{  FTimerThread := TTimerThread.Create;
  FTimerThread.Interval := 1000;
  FTimerThread.OnTimerEvent := TimerThreadEvent;
  FTimerThread.Enabled := True; }

  FillChar(GV_TimeBefore, SizeOf(TSystemTime), #0);
  FillChar(GV_TimeCurrent, SizeOf(TSystemTime), #0);

  GetLocalTime(GV_TimeBefore);
  GetLocalTime(GV_TimeCurrent);

//  timeGetDevCaps(@TimeCaps, SizeOf(TTimeCaps));
//
//  FTimerID := timeSetEvent(1, TimeCaps.wPeriodMin, @TimerCallBack, 0, TIME_PERIODIC);


//  Maximize;
end;

procedure TfrmSEC.Finalize;
var
  R: Integer;
begin
  SetEvent(GV_TimerCancelEvent);

  if (FTimerID <> 0) then
  begin
    timeKillEvent(FTimerID); // 타이머를 해제한다.
    timeEndPeriod(0); // 타이머 주기를 해제한다.
  end;

{  FTimerThread.Terminate;
  FTimerThread.WaitFor;
  FreeAndNil(FTimerThread); }

  FCrossCheckThread.Terminate;
  FCrossCheckThread.WaitFor;
  FreeAndNil(FCrossCheckThread);

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded CrossCheck thread destroy.'));

  if (FMCCCheckThread <> nil) then
  begin
    FMCCCheckThread.Terminate;
    FMCCCheckThread.WaitFor;
    FreeAndNil(FMCCCheckThread);

    Assert(False, GetMainLogStr(lsNormal, 'Succeeded MCCCheck thread destroy.'));
  end;

  FSECCheckThread.Terminate;
  FSECCheckThread.WaitFor;
  FreeAndNil(FSECCheckThread);

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded SECCheck thread destroy.'));

  FDCSCheckThread.Terminate;
  FDCSCheckThread.WaitFor;
  FreeAndNil(FDCSCheckThread);

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded DCSCheck thread destroy.'));

  FDCSDeviceThread.Terminate;
  FDCSDeviceThread.WaitFor;
  FreeAndNil(FDCSDeviceThread);

//  FDCSMediaThread.Terminate;
//  FDCSMediaThread.WaitFor;
//  FreeAndNil(FDCSMediaThread);

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded DCSDevice thread destroy.'));

  FDCSEventThread.Terminate;
  FDCSEventThread.WaitFor;
  FreeAndNil(FDCSEventThread);

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded DCSEvent thread destroy.'));

  if (FMCCEventThread <> nil) then
  begin
    FMCCEventThread.Terminate;
    FMCCEventThread.WaitFor;
    FreeAndNil(FMCCEventThread);

    Assert(False, GetMainLogStr(lsNormal, 'Succeeded MCCEvent thread destroy.'));
  end;

  FSECEventThread.Terminate;
  FSECEventThread.WaitFor;
  FreeAndNil(FSECEventThread);

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded SECEvent thread destroy.'));

  CloseHandle(GV_TimerExecuteEvent);
  CloseHandle(GV_TimerCancelEvent);

  DeviceClose;

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded DeviceClose.'));

  FinalizeDevicePage;
  Assert(False, GetMainLogStr(lsNormal, 'Succeeded FinalizeDevicePage.'));

  FinalizeChannelPage;
  Assert(False, GetMainLogStr(lsNormal, 'Succeeded FinalizeChannelPage.'));

  FinalizeAllChannelPage;
  Assert(False, GetMainLogStr(lsNormal, 'Succeeded FinalizeAllChannelPage.'));

  R := DCSFinalize;
  if (R <> D_OK) then
    Assert(False, GetMainLogStr(lsError, @LSE_DCSFinalizeFailed, [R]));

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded DCSFinalize.'));

  R := DCSSysFinalize;
  if (R <> D_OK) then
    Assert(False, GetMainLogStr(lsError, @LSE_DCSSysFinalizeFailed, [R]));

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded DCSSysFinalize.'));

  R := SECFinalize;
  if (R <> D_OK) then
    Assert(False, GetMainLogStr(lsError, @LSE_SECFinalizeFailed, [R]));

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded SECFinalize.'));

  R := SECSysFinalize;
  if (R <> D_OK) then
    Assert(False, GetMainLogStr(lsError, @LSE_SECSysFinalizeFailed, [R]));

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded SECSysFinalize.'));

  R := MCCFinalize;
  if (R <> D_OK) then
    Assert(False, GetMainLogStr(lsError, @LSE_MCCFinalizeFailed, [R]));

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded MCCFinalize.'));

  R := MCCSysFinalize;
  if (R <> D_OK) then
    Assert(False, GetMainLogStr(lsError, @LSE_MCCSysFinalizeFailed, [R]));

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded MCCSysFinalize.'));

  SaveConfig;

  ClearProgramTypeList;

  ClearMCSList;
  ClearSourceList;
  ClearDCSList;
  ClearMCCList;
  ClearSECList;
  ClearChannelList;

  FreeAndNil(ServerCritSec);

  FreeAndNil(ServerSysCritSec);

  FreeAndNil(GV_ClipboardCueSheet);

  FreeAndNil(GV_ProgramTypeList);

  FreeAndNil(GV_MCSList);
  FreeAndNil(GV_SourceList);
  FreeAndNil(GV_DCSList);
  FreeAndNil(GV_MCCList);
  FreeAndNil(GV_SECList);
  FreeAndNil(GV_ChannelList);

  FreeAndNil(GV_TimelineOnairIcon);
  FreeAndNil(GV_TimelineNextIcon);
  FreeAndNil(GV_TimelineNormalIcon);

  FreeAndNil(FSaveMenuFont);
  FreeAndNil(FSaveColorMap);

  Assert(False, GetMainLogStr(lsNormal, '111'));
  if (FWarningDialogDeviceCheck <> nil) and (FWarningDialogDeviceCheck.HandleAllocated) then
    FreeAndNil(FWarningDialogDeviceCheck);

  Assert(False, GetMainLogStr(lsNormal, 'Succeeded Finalize.'));

  FreeAndNil(GV_LogCS);
end;

procedure TfrmSEC.InitializeAllChannelPage;
begin
  frmAllChannels := TfrmAllChannels.Create(aopAllChannel, True, 0, 0, aopAllChannel.ClientWidth, aopAllChannel.ClientHeight);
  frmAllChannels.Parent := aopAllChannel;
  frmAllChannels.Align := alClient;
  frmAllChannels.Show;
end;

procedure TfrmSEC.FinalizeAllChannelPage;
begin
  if (frmAllChannels <> nil) then
  begin
    FreeAndNil(frmAllChannels);
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

      ChannelForm := TfrmChannel.Create(OfficePage, Channel^.ID, True, 0, 0, OfficePage.ClientWidth, OfficePage.ClientHeight);
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
  Assert(False, GetMainLogStr(lsNormal, 'ChannelForm destroy.'));

          end;
        end;
        FreeAndNil(OfficePage);
  Assert(False, GetMainLogStr(lsNormal, 'OfficePage destroy.'));

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
  DeviceForm := GetDeviceForm;
  if (DeviceForm <> nil) then
  begin
    FreeAndNil(DeviceForm);
  end;
end;

procedure TfrmSEC.DisplayActivate;
begin
  if (GV_SECMine <> nil) and (GV_SECMine = GV_SECMain) then
  begin
    imgSECMain.Visible := True;
    imgSECSub.Visible := False;
  end
  else
  begin
    imgSECSub.Visible := True;
    imgSECMain.Visible := False;
  end;

  if (GV_SECMine <> nil) then
    lblSECName.Caption := String(GV_SECMine^.Name)
  else
    lblSECName.Caption := '';
end;

procedure TfrmSEC.DeviceOpen;
begin
{  frmCheckStart := TfrmCheckStart.Create(Self);
  try
    frmCheckStart.ShowModal;
  finally
    FreeAndNil(frmCheckStart);
  end; }

  frmStartSplash.DeviceOpen;
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
        if (SourceHandle = nil) then continue;
        if (SourceHandle^.DCS = nil) then continue;

        if (SourceHandle^.DCS^.Alive) and
           (SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
        begin
          R := DCSClose(SourceHandle^.DCS^.ID, SourceHandle^.Handle);
          if (R <> D_OK) then
//            ShowMessage(Format('Close Failed ID=%d, Handle=%d, Name=%s', [SourceHandle^.DCSID, SourceHandle^.Handle, GV_SourceList[I]^.Name]))
            Assert(False, GetMainLogStr(lsError, @LSE_DCSCloseDeviceFailed, [R, SourceHandle^.DCS^.ID, String(GV_SourceList[I]^.Name), SourceHandle^.Handle]))
          else
            Assert(False, GetMainLogStr(lsNormal, @LS_DCSCloseDeviceSuccess, [SourceHandle^.DCS^.ID, String(GV_SourceList[I]^.Name), SourceHandle^.Handle]));
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
var
  SystemTime: TSystemTime;
  D: TDateTime;
begin
  FCurrentTime := Now;
//  GetLocalTime(SystemTime);
//  D := Now;
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
  if (WMTitleBar.Active) and (WMTitleBar.MoveEnabled) then
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
  MCC: PMCC;
  SaveAlive, IsAlive: Boolean;
  ChannelForm: TfrmChannel;
begin
  if (not HasMainControl) then exit;

  // MCC alive check
  for I := 0 to GV_MCCList.Count - 1 do
  begin
    MCC := GV_MCCList[I];
    if (MCC <> nil) then
    begin
      SaveAlive := MCC^.Alive;

      R := MCCSysIsAlive(MCC^.HostIP, IsAlive);
      if (R = D_OK) then
      begin
        MCC^.Alive := IsAlive;
        Assert(False, GetMainLogStr(lsNormal, @LS_MCCAliveCheckSuccess, [MCC^.ID, MCC^.Name, BoolToStr(IsAlive, True)]));
      end
      else
      begin
        MCC^.Alive := False;
        Assert(False, GetMainLogStr(lsError, @LSE_MCCAliveCheckFailed, [R, MCC^.ID, MCC^.Name]));
      end;

      if ((R = D_OK) and (IsAlive) and (not SaveAlive)) then
      begin
        Assert(False, GetMainLogStr(lsNormal, Format('Start update channel status. MCC=%d, Name=%s', [MCC^.ID, MCC^.Name])));
        for J := 0 to GV_ChannelList.Count - 1 do
        begin
          ChannelForm := GetChannelFormByID(GV_ChannelList[J]^.ID);
          if (ChannelForm <> nil) then
          begin
            ChannelForm.UpdateMCCCheck(MCC, MCC^.Alive);
            Assert(False, GetMainLogStr(lsNormal, Format('Update channel status. MCC=%d, Alive=%s, Channel=%d', [MCC^.ID, BoolToStr(MCC^.Alive, True), GV_ChannelList[J]^.ID])));
          end;
        end;
        Assert(False, GetMainLogStr(lsNormal, Format('Finish update channel status. SEC=%d, Name=%s', [MCC^.ID, MCC^.Name])));
      end;
    end;
  end;
end;

procedure TfrmSEC.SECCheck;
var
  I, J: Integer;
  SEC: PSEC;

  R: Integer;
  IsMain, IsAlive, SaveAlive: Boolean;

  ChannelForm: TfrmChannel;

  NextIndex: Integer;
  NextSEC: PSEC;
begin
  if (not HasMainControl) then exit;

  // Other SEC alive check
  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (GV_SECMine <> SEC) then
    begin
      SaveAlive := SEC^.Alive;

      R := SECSysIsAlive(SEC^.HostIP, IsAlive);
      if (R = D_OK) then
      begin
        SEC^.Alive := IsAlive;
        Assert(False, GetMainLogStr(lsNormal, @LS_SECAliveCheckSuccess, [SEC^.ID, SEC^.Name, BoolToStr(IsAlive, True)]));
      end
      else
      begin
        SEC^.Alive := False;
        Assert(False, GetMainLogStr(lsError, @LSE_SECAliveCheckFailed, [R, SEC^.ID, SEC^.Name]));
      end;

      if ((R = D_OK) and (IsAlive) and (not SaveAlive)) then
      begin
        Assert(False, GetMainLogStr(lsNormal, Format('Start update channel status. SEC=%d, Name=%s', [SEC^.ID, SEC^.Name])));
        for J := 0 to GV_ChannelList.Count - 1 do
        begin
          ChannelForm := GetChannelFormByID(GV_ChannelList[J]^.ID);
          if (ChannelForm <> nil) then
          begin
            ChannelForm.UpdateSECCheck(SEC, SEC^.Alive);
            Assert(False, GetMainLogStr(lsNormal, Format('Update channel status. SEC=%d, Alive=%s, Channel=%d', [SEC^.ID, BoolToStr(SEC^.Alive, True), GV_ChannelList[J]^.ID])));
          end;
        end;
        Assert(False, GetMainLogStr(lsNormal, Format('Finish update channel status. SEC=%d, Name=%s', [SEC^.ID, SEC^.Name])));
      end;
    end;
  end;
end;

{procedure TfrmSEC.SECMainCheck;
var
  I, R: Integer;
  SEC: PSEC;

  IsMain: Boolean;
begin
  // Other SEC open
  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (SEC <> GV_SECMine) then
    begin
      R := SECOpen(SEC^.ID, SEC^.HostIP, SEC^.Name);
      if (R = D_OK) then
      begin
        Assert(False, GetMainLogStr(lsNormal, @LS_SECOpenSuccess, [SEC^.ID, SEC^.Name]));
        SEC^.Opened := True;
      end
      else
      begin
        Assert(False, GetMainLogStr(lsError, @LSE_SECOpenFailed, [R, SEC^.ID, SEC^.Name]));
        SEC^.Opened := False;
      end;
    end
    else
      SEC^.Opened := True;
  end;

  // Exist other main SEC check
  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (SEC <> GV_SECMine) and (SEC^.Opened) then
    begin
      R := SECIsMain(SEC^.ID, IsMain);
      if (R = D_OK) and (IsMain) then
      begin
        GV_SECMain := SEC;
        GV_SECMain^.Main := True;
        Assert(False, GetMainLogStr(lsNormal, @LS_SECFindMainSuccess, [SEC^.ID, SEC^.Name]));
        break;
      end;
    end;
  end;

  if (GV_SECMain = nil) then
  begin
    // If not exist main SEC then set main SEC change
    Assert(False, GetMainLogStr(lsNormal, @LS_SECNotFoundMain));
    SECMainChange;
    exit;
  end
  else if (GV_SECMain <> GV_SECMine) then
  begin
    // If config is main
    if (GV_SECMine <> nil) then
      GV_SECMine^.Main := False;

    // Save main
    SaveSECConfig;
  end
  else if (GV_SECMain = GV_SECMine) then
  begin
    SECMainDeviceControlBy;
  end;

  // Channel active
  PostMessage(Handle, WM_UPDATE_ACTIVATE, 0, 0);
end;
}
{procedure TfrmSEC.SECMainChange;
var
  I, R: Integer;
  SEC: PSEC;
begin
  // If Set self SEC to main SEC
  // Transmit main SEC to other SEC
  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (SEC <> GV_SECMine) and (SEC <> GV_SECMain) and (SEC^.Opened) then
    begin
      R := SECSetMain(SEC^.ID, GV_SECMine^.ID);
    end;
  end;

  if (GV_SECMain <> nil) then GV_SECMain^.Main := False;

  if (GV_SECMine <> nil) then GV_SECMine^.Main := True;

  GV_SECMain := GV_SECMine;

  // Set device control by
  SECMainDeviceControlBy;

  // Save main
  SaveSECConfig;

  if (GV_SECMine <> nil) then
    Assert(False, GetMainLogStr(lsNormal, @LS_SECMainChange, [GV_SECMine^.ID, GV_SECMine^.Name]));

  // Channel re active
  PostMessage(Handle, WM_UPDATE_ACTIVATE, 0, 0);
end;}

{procedure TfrmSEC.SECMainDeviceControlBy;
var
  I, J, R: Integer;
  SourceHandle: PSourceHandle;
begin
  if (not HasMainControl) then exit;

  for I := 0 to GV_SourceList.Count - 1 do
  begin
    if (GV_SourceList[I]^.Handles <> nil) then
    begin
      for J := 0 to GV_SourceList[I]^.Handles.Count - 1 do
      begin
        SourceHandle := GV_SourceList[I]^.Handles[J];

        if (SourceHandle^.DCS <> nil) and (SourceHandle^.DCS^.Opened) and
           (SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
        begin
          R := DCSSetControlBy(SourceHandle^.DCS^.ID, SourceHandle^.Handle);
          if (R = D_OK) then
          begin
            Assert(False, GetMainLogStr(lsNormal, @LS_DCSSetControlBySuccess, [SourceHandle^.DCS^.ID, GV_SourceList[I]^.Name, SourceHandle^.Handle]));
          end
          else
          begin
            Assert(False, GetMainLogStr(lsError, @LSE_DCSSetControlByFailed, [R, SourceHandle^.DCS^.ID, GV_SourceList[I]^.Name, SourceHandle^.Handle]));
          end;
        end;
      end;
    end;
  end;
end; }

{procedure TfrmSEC.SECMainAliveCheck;
var
  I, R: Integer;
  ChannelForm: TfrmChannel;
begin
//  if (HasMainControl) then exit;
//
//  // Main SEC alive check
//  if (GV_SECMain <> nil) then
//  begin
//    R := SECIsAlive(GV_SECMain^.ID);
//    if (R <> D_OK) then
//    begin
//      GV_SECMain^.Opened := False;
//
//      Inc(FNumSECMainCheck);
//
//      Assert(False, GetMainLogStr(lsError, @LSE_SECMainCheckFailed, [FNumSECMainCheck, R, GV_SECMain^.ID, GV_SECMain^.Name]));
//
//{      // If main SEC is deadlock then all channel execute end update
//      for I := 0 to GV_ChannelList.Count - 1 do
//      begin
//        ChannelForm := GetChannelFormByID(GV_ChannelList[I]^.ID);
//        if (ChannelForm <> nil) then
//          PostMessage(ChannelForm.Handle, WM_END_UPDATE, 0, 0);
//      end; }
//
//    end
//    else
//    begin
//      GV_SECMain^.Opened := True;
//
//      FNumSECMainCheck := 0;
//
//      Assert(False, GetMainLogStr(lsNormal, @LS_SECMainCheckSuccess, [GV_SECMain^.ID, GV_SECMain^.Name]));
//    end;
//  end;
//
//  if (FNumSECMainCheck >= GV_SettingSEC.NumCrossCheck) then
//  begin
//    SECMainChange;
//
//    FNumSECMainCheck := 0;
//  end;
{end; }

{function TfrmSEC.SECIsAlive(var AIsAlive: Boolean): Integer;
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

function TfrmSEC.SECGetChannelOnAir(ABuffer: AnsiString; AChannelOnAir: Boolean): Integer;
var
  ChannelID: Word;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  AChannelOnAir := False;

  ChannelID := PAnsiCharToWord(@ABuffer[1]);
  ChannelForm := GetChannelFormByID(ChannelID);
  if (ChannelForm <> nil) then
  begin
    AChannelOnAir := ChannelForm.ChannelOnAir;
    Result := D_OK;
  end;
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
    SEC^.Opened := True;

    Result := D_OK;
  end;
//  FCrossCheckThread.FNumCrossCheck := 0;
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
//  FCrossCheckThread.FNumCrossCheck := 0;
end;

function TfrmSEC.SECSetChannelOnAir(ABuffer: AnsiString): Integer;
var
  ChannelID: Word;
  ChannelOnAir: Boolean;
  ChannelForm: TfrmChannel;
begin
  Result := D_FALSE;

  ChannelID := PAnsiCharToWord(@ABuffer[1]);
  ChannelOnAir := PAnsiCharToBool(@ABuffer[3]);

  ChannelForm := GetChannelFormByID(ChannelID);
  if (ChannelForm <> nil) then
  begin
    PostMessage(ChannelForm.Handle, WM_UPDATE_CHANNEL_ONAIR, NativeInt(ChannelOnAir), 0);
    Result := D_OK;
  end;
end; }

{ TTimerThread }

//constructor TTimerThread.Create;
//begin
//  FTimerEnabledFlag := CreateEvent(nil, True, False, nil);
//  FCancelFlag := CreateEvent(nil, True, False, nil);
//  FTimerProc := nil;
//  FInterval := 1000;
//  FreeOnTerminate := False; // Main thread controls for thread destruction
//  inherited Create(False);
//end;
//
//destructor TTimerThread.Destroy; // Call TTimerThread.Free to cancel the thread
//begin
//  Terminate;
//  if GetCurrentThreadID = MainThreadID then
//  begin
//    OutputDebugString('TTimerThread.Destroy :: MainThreadID (Waitfor)');
//    Waitfor; // Synchronize
//  end;
//  CloseHandle(FCancelFlag);
//  CloseHandle(FTimerEnabledFlag);
//  OutputDebugString('TTimerThread.Destroy');
//  inherited;
//end;
//
//procedure TTimerThread.Terminate;
//begin
//  inherited Terminate;
//  ResetEvent(FTimerEnabledFlag); // Stop timer event
//  SetEvent(FCancelFlag); // Set cancel flag
//end;
//
//procedure TTimerThread.SetEnabled(AEnable: Boolean);
//begin
//  if AEnable then
//    SetEvent(FTimerEnabledFlag)
//  else
//    ResetEvent(FTimerEnabledFlag);
//end;
//
//function TTimerThread.GetEnabled: Boolean;
//begin
//  Result := WaitForSingleObject(FTimerEnabledFlag, 0) = WAIT_OBJECT_0 // Signaled
//end;
//
//procedure TTimerThread.SetInterval(AInterval: Cardinal);
//begin
//  FInterval := AInterval;
//end;
//
//procedure TTimerThread.Execute;
//var
//  WaitList: array[0..1] of THandle;
//  WaitInterval, LastProcTime: Int64;
//  Frequency, StartCount, StopCount: Int64; // minimal stop watch
//
//  SystemTime: TSystemTime;
//begin
//  QueryPerformanceFrequency(Frequency); // this will never return 0 on Windows XP or later
//  if (Frequency = 0) then Frequency := 1;
//
//  WaitList[0] := FTimerEnabledFlag;
//  WaitList[1] := FCancelFlag;
//  LastProcTime := 0;
//
//  SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_TIME_CRITICAL);
//
//{  repeat
//    GetLocalTime(SystemTime);
//  until ((SystemTime.wMilliseconds mod 1000) = 0); }
//
//  if (Enabled) and (Assigned(FTimerProc)) then
//    FTimerProc(Self);
//
//{  GetSystemTime(SystemTime);
//  WaitInterval := 1000 - SystemTime.wMilliseconds;
//  WaitForSingleObject(FCancelFlag, WaitInterval); }
//
//  while not Terminated do
//  begin
//    ResetEvent(GV_TimerExecuteEvent);
//    if (WaitForMultipleObjects(2, @WaitList[0], False, INFINITE) <> WAIT_OBJECT_0) then
//      break; // Terminate thread when FCancelFlag is signaled
//
//
////    GetLocalTime(SystemTime);
////    LastProcTime := SystemTime.wMilliseconds;
//
//    WaitInterval := FInterval - LastProcTime;
//    if (WaitInterval < 0) then
//    begin
//      WaitInterval := 0;
//    end;
////    if (WaitForSingleObject(FCancelFlag, WaitInterval) <> WAIT_TIMEOUT) then
////      break;
//
//    QueryPerformanceCounter(StartCount);
//    repeat
//      QueryPerformanceCounter(StopCount);
////      Sleep(0);
//    until ((1000 * (StopCount - StartCount) div Frequency) >= WaitInterval);
//
//    if (Enabled) then
//    begin
//      QueryPerformanceCounter(StartCount);
//      if not Terminated then
//      begin
//        if (Assigned(FTimerProc)) then
//          FTimerProc(Self);
//        SetEvent(GV_TimerExecuteEvent);
//      end;
//      QueryPerformanceCounter(StopCount);
//      // Interval adjusted for FTimerProc execution time
//      LastProcTime := 1000 * (StopCount - StartCount) div Frequency; // ElapsedMilliSeconds
//    end;
//  end;
//
//  SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_NORMAL);
//end;

procedure TfrmSEC.DCSCheck;
var
  I, J, K, N, R: Integer;
  DCS: PDCS;
  SaveMain, SaveAlive: Boolean;
  IsMain: Boolean;
  SaveConfigFlag: Boolean;

  SourceHandle: PSourceHandle;

  ChannelForm: TfrmChannel;
begin
  if (not HasMainControl) then exit;

  SaveConfigFlag := False;

  for I := 0 to GV_DCSList.Count - 1 do
  begin
    DCS := GV_DCSList[I];
    if (DCS <> nil) then
    begin
      SaveMain := DCS^.Main;
      SaveAlive := DCS^.Alive;

      R := DCSSysIsMain(DCS^.HostIP, IsMain);
      if (R = D_OK) then
      begin
        Assert(False, GetMainLogStr(lsNormal, @LS_DCSMainCheckSuccess, [DCS^.ID, DCS^.Name, BoolToStr(IsMain, True)]));
        DCS^.Main := IsMain;
        DCS^.Alive := True;
        if (DCS^.Alive) and (SaveAlive <> DCS^.Alive) then
        begin
          for J := 0 to GV_SourceList.Count - 1 do
          begin
            if (GV_SourceList[J]^.Handles <> nil) then
            begin
              for K := 0 to GV_SourceList[J]^.Handles.Count - 1 do
              begin
                SourceHandle := GV_SourceList[J]^.Handles[K];
                if (SourceHandle = nil) then continue;
                if (SourceHandle^.DCS = nil) then continue;

                if (SourceHandle^.DCS = DCS) and
                   (SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
                begin
                  R := frmSEC.DCSEventThread.SetControlBy(SourceHandle);

{                  R := DCSSetControlBy(SourceHandle^.DCS^.ID, SourceHandle^.Handle);
                  if (R = D_OK) then
                  begin
                    Assert(False, GetMainLogStr(lsNormal, @LS_DCSSetControlBySuccess, [SourceHandle^.DCS^.ID, GV_SourceList[I]^.Name, SourceHandle^.Handle]));
                  end
                  else
                  begin
                    Assert(False, GetMainLogStr(lsNormal, @LSE_DCSSetControlByFailed, [R, SourceHandle^.DCS^.ID, GV_SourceList[I]^.Name, SourceHandle^.Handle]));
                  end; }
                end;
              end;
            end;
          end;
        end;

        if (DCS^.Main) and (SaveMain <> DCS^.Main) then
        begin
          for J := 0 to GV_SourceList.Count - 1 do
          begin
            if (GV_SourceList[J]^.Handles <> nil) then
            begin
              for K := 0 to GV_SourceList[J]^.Handles.Count - 1 do
              begin
                SourceHandle := GV_SourceList[J]^.Handles[K];
                if (SourceHandle = nil) then continue;
                if (SourceHandle^.DCS = nil) then continue;

                if (SourceHandle^.DCS <> DCS) then
                begin
                  frmSEC.DCSEventThread.ClearMediaCheckQueueBySourcdeHandle(SourceHandle);
                  continue;
                end;

                if (SourceHandle^.DCS = DCS) and
                   (SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
                begin
                  for N := 0 to GV_ChannelList.Count - 1 do
                  begin
                    if (GV_ChannelList[N]^.OnAir) then
                    begin
                      ChannelForm := GetChannelFormByID(GV_ChannelList[N]^.ID);
                      if (ChannelForm <> nil) then
                        ChannelForm.MediaCheck(SourceHandle);
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
      end
      else
      begin
        Assert(False, GetMainLogStr(lsError, @LSE_DCSMainCheckFailed, [R, DCS^.ID, DCS^.Name]));
//        DCS^.Main := False;
        DCS^.Alive := False;
      end;

      if (SaveMain <> DCS^.Main) then
        SaveConfigFlag := True;
    end;
  end;

  if (SaveConfigFlag) then
    SaveDCSConfig;
end;

{ TCrossCheckThread }

constructor TCrossCheckThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FCommandCritSec := TCriticalSection.Create;
  FInCritSec := TCriticalSection.Create;

  with GV_SettingSEC do
  begin
    FUDPIn  := TUDPIn.Create(CrossPort);
    FUDPIn.LogEnabled     := True;
    FUDPIn.LogPath        := LogPath;
    FUDPIn.LogExt         := Format('%d_%s', [FUDPIn.Port, LogExt]);
  end;

  FUDPIn.OnUDPRead := UDPInRead;
  FUDPIn.Start;
  while not FUDPIn.Started do
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

  FreeAndNil(FCommandCritSec);

  FreeAndNil(FInCritSec);

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
  CurrEventID, NextEventID: TEventID;
  SendBuffer: AnsiString;
begin
  FInCritSec.Enter;
  try
    if (ADataSize <= 0) then exit;
    FRecvBuffer := FRecvBuffer + AData;

    if (Length(FRecvBuffer) < 1) then exit;

  //  if (GV_DCSMine <> nil) and (GV_DCSMine^.Main) then
    if (not FIsCommand) then
    begin
      case Ord(FRecvBuffer[1]) of
        $02:
        begin
          if (Length(FRecvBuffer) < 3) then exit;

          ByteCount := PAnsiCharToWord(@FRecvBuffer[2]);
          if (Length(FRecvBuffer) >= ByteCount + 4) then
          begin
            if (CheckSum(FRecvBuffer)) then
            begin
              CMD1 := Ord(FRecvBuffer[4]);
              CMD2 := Ord(FRecvBuffer[5]);

              FRecvData := System.Copy(FRecvBuffer, 6, ByteCount - 2);

              case CMD1 of
                $00: // System Control (0X00)
                begin
                  case CMD2 of
                    $00: // Is Alive
                    begin
                      R := FSEC.SECIsAliveW(IsAlive);
                      if (R = D_OK) then
                        SendBuffer := BoolToAnsiString(IsAlive);
                    end;
                    $01: // Is Main
                    begin
                      R := FSEC.SECIsMainW(IsMain);
                      if (R = D_OK) then
                        SendBuffer := BoolToAnsiString(IsMain);
                    end;
                    $10: // Set Main
                    begin
                      R := FSEC.SECSetMainW(FRecvData);
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
                      $01: TransmitResponse(ABindingIP, FUDPIn.Port, CMD1, CMD2, SendBuffer, Length(SendBuffer));
                      $10: TransmitAck(ABindingIP, FUDPIn.Port);
                    end;
                  end;
                end;
              end
              else
              begin
                // Sending error code
    //            SendBuffer := AnsiChar(D_ERR) + IntToAnsiString(R);
    //            FUDPOut.Send(ABindingIP, FUDPOut.Port, SendBuffer);
                TransmitError(ABindingIP, FUDPIn.Port, R);
              end;
            end
            else
            begin
    //          FUDPOut.Send(ABindingIP, AnsiChar(D_NAK) + IntToAnsiString(E_NAK_CHECKSUM));
              TransmitNak(ABindingIP, FUDPIn.Port, E_NAK_CHECKSUM);
            end;
          end;

          if (Length(FRecvBuffer) > ByteCount + 4) then
            FRecvBuffer := Copy(FRecvBuffer, ByteCount + 5, Length(FRecvBuffer))
          else
            FRecvBuffer := '';
        end;
      else
  //      FUDPOut.Send(ABindingIP, AnsiChar(D_ERR) + IntToAnsiString(E_UNDEFIND_COMMAND));
        TransmitError(ABindingIP, FUDPIn.Port, E_UNDEFIND_COMMAND);
        FRecvBuffer := '';
      end;
    end
    else
    begin
      case Ord(FRecvBuffer[1]) of
        $02:
        begin
          if (Length(FRecvBuffer) < 2) then exit;

          ByteCount := PAnsiCharToWord(@FRecvBuffer[2]);
          if (Length(FRecvBuffer) = ByteCount + 4) then
          begin
            if (CheckSum(FRecvBuffer)) then
            begin
              CMD1 := Ord(FRecvBuffer[4]);
              CMD2 := Ord(FRecvBuffer[5]);

              FRecvData := System.Copy(FRecvBuffer, 6, ByteCount - 2);

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
          else if ((ByteCount <= 0) or (Length(FRecvBuffer) > ByteCount + 4)) then
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
        FLastResult := D_FALSE;
        SetEvent(FSyncMsgEvent);
        FRecvBuffer := '';
      end;
    end;
  finally
    FInCritSec.Leave;
  end;
end;

function TCrossCheckThread.SendCommand(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
var
  Buffer: AnsiString;
  CheckSum: Byte;
  I: integer;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_STX) + WordToAnsiString($02 + ADataSize) + AnsiChar(ACMD1) + AnsiChar(ACMD2) + AData;

  CheckSum := ACMD1 + ACMD2;
  for I := 1 to ADataSize do
    CheckSum := CheckSum + Ord(AData[I]);

  CheckSum := 0 - CheckSum;
  Buffer := Buffer + AnsiChar(CheckSum);

  FUDPIn.Send(AHostIP, APort, Buffer);

  Result := D_OK;
end;

function TCrossCheckThread.SendResponse(AHostIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; AData: AnsiString; ADataSize: Integer): Integer;
var
  Buffer: AnsiString;
  CheckSum: Byte;
  I: integer;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_STX) + WordToAnsiString($02 + ADataSize) + AnsiChar(ACMD1) + AnsiChar(ACMD2) + AData;

  CheckSum := ACMD1 + ACMD2;
  for I := 1 to ADataSize do
    CheckSum := CheckSum + Ord(AData[I]);

  CheckSum := 0 - CheckSum;
  Buffer := Buffer + AnsiChar(CheckSum);

  FUDPIn.Send(AHostIP, APort, Buffer);

  Result := D_OK;
end;

function TCrossCheckThread.SendAck(AHostIP: AnsiString; APort: Word): Integer;
begin
  Result := D_FALSE;

  FUDPIn.Send(AHostIP, APort, AnsiChar(D_ACK));

  Result := D_OK;
end;

function TCrossCheckThread.SendNak(AHostIP: AnsiString; APort: Word; ANakError: Byte): Integer;
var
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_NAK) + AnsiChar(ANakError);

  FUDPIn.Send(AHostIP, APort, Buffer);

  Result := D_OK;
end;

function TCrossCheckThread.SendError(AHostIP: AnsiString; APort: Word; AErrorCode: Integer): Integer;
var
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_ERR) + IntToAnsiString(AErrorCode);

  FUDPIn.Send(AHostIP, APort, Buffer);

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

  FCommandCritSec.Enter;
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
    FCommandCritSec.Leave;

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

function TCrossCheckThread.SECIsAlive(AHostIP: AnsiString; var AIsAlive: Boolean): Integer;
var
  Buffer: AnsiString;
begin
  AIsAlive := False;

  Buffer := '';
  Result := TransmitCommand(AHostIP, FUDPIn.Port, $00, $00, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    AIsAlive := PAnsiCharToBool(@FRecvData[1]);
  end;
end;

function TCrossCheckThread.SECIsMain(AHostIP: AnsiString; var AIsMain: Boolean): Integer;
var
  Buffer: AnsiString;
begin
  AIsMain := False;

  Buffer := '';
  Result := TransmitCommand(AHostIP, FUDPIn.Port, $00, $01, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    AIsMain := PAnsiCharToBool(@FRecvData[1]);
  end;
end;

function TCrossCheckThread.SECSetMain(AHostIP: AnsiString; AMainSECID: Word): Integer;
var
  Buffer: AnsiString;
begin
  Buffer := WordToAnsiString(AMainSECID);
  Result := TransmitCommand(AHostIP, FUDPIn.Port, $00, $10, Buffer, Length(Buffer));
end;

procedure TCrossCheckThread.Execute;
begin
  { Place thread code here }

//  DoMainCheck;

  while not Terminated do
  begin
    DoCrossCheck;
    Sleep(TimecodeToMilliSec(GV_SettingSEC.CrossCheckInterval));
  end;
end;

procedure TCrossCheckThread.MainDeviceControlBy;
var
  I, J: Integer;
  R: Integer;
  SourceHandle: PSourceHandle;
begin
  // Set device control by
  for I := 0 to GV_SourceList.Count - 1 do
  begin
    if (GV_SourceList[I]^.Handles <> nil) then
    begin
      for J := 0 to GV_SourceList[I]^.Handles.Count - 1 do
      begin
        SourceHandle := GV_SourceList[I]^.Handles[J];

        if (SourceHandle^.DCS <> nil) and (SourceHandle^.DCS^.Alive) and
           (SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
        begin
          R := frmSEC.DCSEventThread.SetControlBy(SourceHandle);

{          R := DCSSetControlBy(SourceHandle^.DCS^.ID, SourceHandle^.Handle);
          if (R = D_OK) then
          begin
            Assert(False, GetMainLogStr(lsNormal, @LS_DCSSetControlBySuccess, [SourceHandle^.DCS^.ID, GV_SourceList[I]^.Name, SourceHandle^.Handle]));
          end
          else
          begin
            Assert(False, GetMainLogStr(lsError, @LSE_DCSSetControlByFailed, [R, SourceHandle^.DCS^.ID, GV_SourceList[I]^.Name, SourceHandle^.Handle]));
          end; }
        end;
      end;
    end;
  end;
end;

procedure TCrossCheckThread.MainChange;
begin
  Assert(False, GetMainLogStr(lsNormal, Format('Switch start main SEC, Name = %s, ID = %d',
                                               [String(GV_SECMine^.Name), GV_SECMine^.ID])));

  GV_SECMine^.Main := True;

  // Main/Sub 전환
  if (GV_SECMain <> nil) then
  begin
    GV_SECMain^.Main := False;
  end;

  GV_SECMain := GV_SECMine;

  // Set device control by
  MainDeviceControlBy;

  // Save main
  SaveSECConfig;

  // SEC re active
  PostMessage(FSEC.Handle, WM_UPDATE_ACTIVATE, 0, 0);

  Assert(False, GetMainLogStr(lsNormal, Format('Switch main SEC succeded, Name = %s, ID = %d',
                                               [String(GV_SECMine^.Name), GV_SECMine^.ID])));

end;

procedure TCrossCheckThread.DoMainCheck;
var
  I, R: Integer;
  SEC: PSEC;

  IsMain: Boolean;
begin
  // SEC 실행 시 현재 운행중인 Main SEC가 있으면
  // 현재 SEC를 Sub로 전환하고 SEC DCS로 부터 이벤트를 수신한다.

  if (GV_SECMine = nil) then
  begin
    Assert(False, GetMainLogStr(lsError, Format('DoMainCheck GV_SECMine is nil', [])));
    exit;
  end;

  Assert(False, GetMainLogStr(lsNormal, Format('SEC Main/sub state, name = %s, main = %s',
                                               [String(GV_SECMine^.Name), BoolToStr(GV_SECMine^.Main, True)])));

  // Another SEC main check
  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (GV_SECMine <> SEC) then
    begin
      R := SECIsMain(SEC^.HostIP, IsMain);

      Assert(False, GetMainLogStr(lsNormal, Format('Check main SEC, name = %s, ismain = %s',
                                                   [String(SEC^.Name), BoolToStr(IsMain, True)])));

      if (R = D_OK) and (IsMain) then
      begin
        GV_SECMain := SEC;
        GV_SECMain^.Main := IsMain;

        GV_SECMine^.Main := False;

        SaveSECConfig;

        break;
      end;
    end;
  end;

  if (GV_SECMain = nil) then
  begin
    MainChange;
  end
  // Set device control by
  else if (GV_SECMine <> nil) and (GV_SECMine^.Main) then
  begin
    MainDeviceControlBy;
  end;



  // SEC re active
  PostMessage(FSEC.Handle, WM_UPDATE_ACTIVATE, 0, 0);
end;

procedure TCrossCheckThread.DoCrossCheck;
var
  I, J: Integer;
  SEC: PSEC;

  R: Integer;
  IsMain, IsAlive, SaveAlive: Boolean;

  ChannelForm: TfrmChannel;

  NextIndex: Integer;
  NextSEC: PSEC;
begin
  if (HasMainControl) then exit;

  // Main SEC alive check
  for I := 0 to GV_SECList.Count - 1 do
  begin
    SEC := GV_SECList[I];
    if (GV_SECMine <> SEC) and (SEC^.Main) then
    begin
      R := SECIsMain(SEC^.HostIP, IsMain);

      if (R = D_OK) then
      begin
        SEC^.Main  := IsMain;
        SEC^.Alive := True;

        Assert(False, GetMainLogStr(lsNormal, @LS_SECMainCheckSuccess, [SEC^.ID, SEC^.Name, BoolToStr(IsMain, True)]));
      end
      else
      begin
        Assert(False, GetMainLogStr(lsError, @LSE_SECMainCheckFailed, [R, SEC^.ID, SEC^.Name]));
        SEC^.Alive := False;
      end;

      if (R <> D_OK) or (not IsMain) then
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

    // 다음 메인 SEC를 설정
    NextSEC := GV_SECList[NextIndex];

    // Send not main another DCS
    for I := 0 to GV_SECList.Count - 1 do
    begin
      SEC := GV_SECList[I];
      if (GV_SECMain <> SEC) and
         (GV_SECMine <> SEC) then
      begin
        R := SECSetMain(SEC^.HostIP, NextSEC^.ID);

        Assert(False, GetMainLogStr(lsNormal, Format('Set main SEC, name = %s, id = %d',
                                                     [String(SEC^.Name), SEC^.ID])));
      end;
    end;

    if (GV_SECMine = NextSEC) then
    begin
      MainChange;
{      GV_SECMine^.Main := True;

      // Main/Sub 전환
      if (GV_SECMain <> nil) then
      begin
        GV_SECMain^.Main := False;
        GV_SECMain := GV_SECMine;
      end;

      // Set device control by
      MainDeviceControlBy;

      // Save main
      SaveSECConfig;

      // SEC re active
      PostMessage(FSEC.Handle, WM_UPDATE_ACTIVATE, 0, 0);

      Assert(False, GetMainLogStr(lsNormal, Format('Switch main SEC succeded, name = %s, id = %d',
                                                   [String(SEC^.Name), SEC^.ID])));
                                                                                     }
    end;

    FNumCrossCheck := 0;
  end;
end;

{ TMCCCheckThread }

constructor TMCCCheckThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FExecuteEvent := CreateEvent(nil, True, True, nil);
  FCompleteEvent := CreateEvent(nil, True, True, nil);

  FEventClose := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TMCCCheckThread.Destroy;
begin
  Terminate;

  CloseHandle(FExecuteEvent);
  CloseHandle(FCompleteEvent);

  CloseHandle(FEventClose);

  inherited Destroy;
end;

procedure TMCCCheckThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FExecuteEvent);
  SetEvent(FEventClose);
end;

procedure TMCCCheckThread.SetExecute;
begin
  SetEvent(FExecuteEvent);
end;

procedure TMCCCheckThread.SetComplete;
begin
  SetEvent(FCompleteEvent);
end;

procedure TMCCCheckThread.ResetExecute;
begin
  ResetEvent(FExecuteEvent);
end;

procedure TMCCCheckThread.ResetComplete;
begin
  ResetEvent(FCompleteEvent);
end;

procedure TMCCCheckThread.WaitExecute;
begin
  WaitForSingleObject(FExecuteEvent, INFINITE);
end;

procedure TMCCCheckThread.WaitComplete;
begin
  WaitForSingleObject(FCompleteEvent, INFINITE);
end;

procedure TMCCCheckThread.DoStartCheck;
var
  I, J, R: Integer;
  MCC: PMCC;
  IsAlive: Boolean;
begin
  if (not HasMainControl) then exit;

  // MCC alive check
  for I := 0 to GV_MCCList.Count - 1 do
  begin
    MCC := GV_MCCList[I];
    if (MCC <> nil) then
    begin
      R := MCCSysIsAlive(MCC^.HostIP, IsAlive);
      if (R = D_OK) then
      begin
        MCC^.Alive := IsAlive;
        Assert(False, GetMainLogStr(lsNormal, @LS_MCCAliveCheckSuccess, [MCC^.ID, MCC^.Name, BoolToStr(IsAlive, True)]));
      end
      else
      begin
        MCC^.Alive := False;
        Assert(False, GetMainLogStr(lsError, @LSE_MCCAliveCheckFailed, [R, MCC^.ID, MCC^.Name]));
      end;
    end;
  end;
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
    R := WaitForSingleObject(FEventClose, TimecodeToMilliSec(GV_SettingMCC.SysCheckInterval));
    if (R = WAIT_OBJECT_0) then break;

    WaitExecute;
    ResetComplete;
    try
      DoMCCCheck;
    finally
      SetComplete;
    end;
  end;
end;

{ TSECCheckThread }

constructor TSECCheckThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FEventClose := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TSECCheckThread.Destroy;
begin
  Terminate;

  CloseHandle(FEventClose);

  inherited Destroy;
end;

procedure TSECCheckThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FEventClose);
end;

procedure TSECCheckThread.DoSECCheck;
var
  I, J, R: Integer;
  SaveOpened: Boolean;
  ChannelForm: TfrmChannel;
begin
  FSEC.SECCheck;
end;

procedure TSECCheckThread.Execute;
var
  R: Cardinal;
begin
  while not Terminated do
  begin
    R := WaitForSingleObject(FEventClose, TimecodeToMilliSec(GV_SettingSEC.SysCheckInterval));
    if (R = WAIT_OBJECT_0) then break;

    DoSECCheck;
  end;
end;

{ TDCSCheckThread }

constructor TDCSCheckThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FExecuteEvent := CreateEvent(nil, True, True, nil);
  FCompleteEvent := CreateEvent(nil, True, True, nil);

  FEventClose := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TDCSCheckThread.Destroy;
begin
  Terminate;

  CloseHandle(FExecuteEvent);
  CloseHandle(FCompleteEvent);

  CloseHandle(FEventClose);

  inherited Destroy;
end;

procedure TDCSCheckThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FExecuteEvent);
  SetEvent(FEventClose);
end;

procedure TDCSCheckThread.SetExecute;
begin
  SetEvent(FExecuteEvent);
end;

procedure TDCSCheckThread.SetComplete;
begin
  SetEvent(FCompleteEvent);
end;

procedure TDCSCheckThread.ResetExecute;
begin
  ResetEvent(FExecuteEvent);
end;

procedure TDCSCheckThread.ResetComplete;
begin
  ResetEvent(FCompleteEvent);
end;

procedure TDCSCheckThread.WaitExecute;
begin
  WaitForSingleObject(FExecuteEvent, INFINITE);
end;

procedure TDCSCheckThread.WaitComplete;
begin
  WaitForSingleObject(FCompleteEvent, INFINITE);
end;

procedure TDCSCheckThread.DoDCSCheck;
begin
  FSEC.DCSCheck;
end;

procedure TDCSCheckThread.Execute;
var
  R: Cardinal;
begin
  while not Terminated do
  begin
    R := WaitForSingleObject(FEventClose, TimecodeToMilliSec(GV_SettingDCS.SysCheckInterval));
    if (R = WAIT_OBJECT_0) then break;

    WaitExecute;
    ResetComplete;
    try
      DoDCSCheck;
    finally
      SetComplete;
    end;
  end;
end;

{ TDCSEventThread }

constructor TDCSEventThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FEventLock   := TCriticalSection.Create;
  FEventQueue  := TDCSEventList.Create;

  FMediaCheckLock   := TCriticalSection.Create;
  FMediaCheckQueue  := TDCSMediaCheckList.Create;
  FMediaCheckIndex  := 0;

  FInputIndex := -1;
  FInputCount := 0;

  FCMD_Result      := D_FALSE;
  FCMD_StartTime.D := 0;
  FCMD_StartTime.T := 0;
  FCMD_DurationTC  := 0;
  FCMD_Exist       := False;

  FEventCommand := CreateEvent(nil, True, False, nil);
  FEventProc := CreateEvent(nil, True, False, nil);
  FEventMediaCheck := CreateEvent(nil, True, False, nil);

  FEventCompleted := CreateEvent(nil, True, True, nil);

  FDeleteEvent := CreateEvent(nil, True, False, nil);
  FClearEvent := CreateEvent(nil, True, False, nil);
  FTakeEvent := CreateEvent(nil, True, False, nil);
  FHoldEvent := CreateEvent(nil, True, False, nil);
  FChangeDurationEvent := CreateEvent(nil, True, False, nil);
  FEventClose := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TDCSEventThread.Destroy;
begin
  CloseHandle(FEventMediaCheck);
  CloseHandle(FEventProc);
  CloseHandle(FEventCommand);
  CloseHandle(FEventCompleted);


  CloseHandle(FDeleteEvent);
  CloseHandle(FClearEvent);
  CloseHandle(FTakeEvent);
  CloseHandle(FHoldEvent);
  CloseHandle(FChangeDurationEvent);
  CloseHandle(FEventClose);

  ClearMediaCheckQueue;
  FreeAndNil(FMediaCheckQueue);

  FreeAndNil(FMediaCheckLock);

  ClearEventQueue;
  FreeAndNil(FEventQueue);

  FreeAndNil(FEventLock);

  inherited Destroy;
end;

procedure TDCSEventThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FEventCommand);
  SetEvent(FEventClose);
end;

procedure TDCSEventThread.ClearEventQueue;
var
  I: Integer;
  D: PDCSEvent;
begin
  FEventLock.Enter;
  try
    for I := FEventQueue.Count - 1 downto 0 do
    begin
      D := FEventQueue[I];
      if (D <> nil) then
      begin
        Dispose(D);
      end;
    end;
    FEventQueue.Clear;
  finally
    FEventLock.Leave;
  end;
end;

function TDCSEventThread.GetEventByID(AID:Word; AHandle: TDeviceHandle; AEventID: TEventID): PDCSEvent;
var
  I: Integer;
  E: PDCSEvent;
begin
  Result := nil;

  FEventLock.Enter;
  try
    for I := 0 to FEventQueue.Count - 1 do
    begin
      E := FEventQueue[I];
      if (E = nil) then continue;
      if (E^.SourceHandle = nil) then continue;
      if (E^.SourceHandle^.DCS = nil) then continue;

      if (E^.SourceHandle^.DCS^.ID = AID) and (E^.SourceHandle^.Handle = AHandle) and
         (IsEqualEventID(E^.Event.EventID, AEventID)) then
      begin
        Result := E;
        exit;
      end;
    end;
  finally
    FEventLock.Leave;
  end;
end;

procedure TDCSEventThread.ClearEventQueueByChannelID(ASourceHandle: PSourceHandle; AChannelID: Word);
var
  I: Integer;
  E: PDCSEvent;
begin
  FEventLock.Enter;
  try
    for I := FEventQueue.Count - 1 downto 0 do
    begin
      E := FEventQueue[I];
      if (E = nil) then continue;
      if (E^.SourceHandle = nil) then continue;

      if (E^.SourceHandle = ASourceHandle) and
         (E^.Event.EventID.ChannelID = AChannelID) then
      begin
        FEventQueue.Remove(E);
        Dispose(E);
      end;
    end;
  finally
    FEventLock.Leave;
  end;
end;

procedure TDCSEventThread.ClearMediaCheckQueue;
var
  I: Integer;
  M: PDCSMediaCheck;
begin
  FMediaCheckLock.Enter;
  try
    for I := FMediaCheckQueue.Count - 1 downto 0 do
    begin
      M := FMediaCheckQueue[I];
      if (M <> nil) then
      begin
        Dispose(M);
      end;
    end;
    FMediaCheckQueue.Clear;

    FMediaCheckIndex := 0;
  finally
    FMediaCheckLock.Leave;
  end;
end;

procedure TDCSEventThread.ClearMediaCheckQueueBySourcdeHandle(ASourceHandle: PSourceHandle);
var
  I: Integer;
  M: PDCSMediaCheck;
begin
  FMediaCheckLock.Enter;
  try
    for I := FMediaCheckQueue.Count - 1 downto 0 do
    begin
      M := FMediaCheckQueue[I];
      if (M = nil) then continue;

      if (M^.SourceHandle = ASourceHandle) then
      begin
        FMediaCheckQueue.Remove(M);
        Dispose(M);

        Dec(FMediaCheckIndex);
      end;
    end;
  finally
    FMediaCheckLock.Leave;
  end;
end;

procedure TDCSEventThread.ClearMediaCheckQueueByChannelID(ASourceHandle: PSourceHandle; AChannelID: Word);
var
  I: Integer;
  M: PDCSMediaCheck;
begin
  FMediaCheckLock.Enter;
  try
    for I := FMediaCheckQueue.Count - 1 downto 0 do
    begin
      M := FMediaCheckQueue[I];
      if (M = nil) then continue;

      if (M^.SourceHandle = ASourceHandle) and (M^.EventID.ChannelID = AChannelID) then
      begin
        FMediaCheckQueue.Remove(M);
        Dispose(M);

        Dec(FMediaCheckIndex);
      end;
    end;
  finally
    FMediaCheckLock.Leave;
  end;
end;

procedure TDCSEventThread.EventQueueCheck;
begin
  if (FEventQueue.Count > 0) then SetEvent(FEventProc);
end;

procedure TDCSEventThread.MediaCheckQueueCheck;
begin
  if (FMediaCheckQueue.Count > 0) then SetEvent(FEventMediaCheck);
end;

procedure TDCSEventThread.ResetCommand;
begin
  FillChar(FCommand, SizeOf(TDCSCommand), #0);
end;

procedure TDCSEventThread.SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType);
begin
//  WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout);
//  WaitForSingleObject(FEventCompleted, INFINITE);
  ResetCommand;

  FCommand.SourceHandle := ASourceHandle;
  FCommand.CommandType  := ACommandType;
end;

procedure TDCSEventThread.SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType; AEventID: TEventID; ATimecode: TTimecode);
begin
  SetCommand(ASourceHandle, ACommandType);

  Move(AEventID, FCommand.EventID, SizeOf(TEventID));
  FCommand.Timecode := ATimecode;
end;

procedure TDCSEventThread.SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType; AChannelID: Word);
begin
  SetCommand(ASourceHandle, ACommandType);

  FCommand.ChannelID := AChannelID;
end;

procedure TDCSEventThread.SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType; AMediaID: String);
begin
  SetCommand(ASourceHandle, ACommandType);

  StrPCopy(FCommand.MediaID, AMediaID);
end;

procedure TDCSEventThread.SetCommand(ASourceHandle: PSourceHandle; ACommandType: TDCSCommandType; AXptInput: Integer);
begin
  SetCommand(ASourceHandle, ACommandType);

  FCommand.XptInput := AXptInput;
end;

function TDCSEventThread.InputEvent(ASourceHandle: PSourceHandle; AEvent: TEvent): Integer;
var
  E: PDCSEvent;
begin
  FEventLock.Enter;
  try
    E := New(PDCSEvent);
    E^.SourceHandle := ASourceHandle;
    E^.EventType := etInput;
    Move(AEvent, E^.Event, SizeOf(TEvent));

    FEventQueue.Add(E);
  finally
    FEventLock.Leave;
  end;

  SetEvent(FEventProc);

  Result := D_OK;
end;

function TDCSEventThread.DeleteEvent(ASourceHandle: PSourceHandle; AEventID: TEventID): Integer;
var
  E: PDCSEvent;
begin
  FEventLock.Enter;
  try
    E := New(PDCSEvent);
    E^.SourceHandle := ASourceHandle;
    E^.EventType := etDelete;
    Move(AEventID, E^.EventID, SizeOf(TEventID));

    FEventQueue.Add(E);
  finally
    FEventLock.Leave;
  end;

  SetEvent(FEventProc);

  Result := D_OK;

{  Result := D_FALSE;

  SetCommand(AID, AHandle, ctDelete, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventCheck; }
end;

function TDCSEventThread.ClearEvent(ASourceHandle: PSourceHandle; AChannelID: Word): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctClear, AChannelID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TDCSEventThread.MediaCheck(ASourceHandle: PSourceHandle; AEventID: TEventID; AMediaID: String; AInTC, ADurationTC: TTimecode): Integer;
var
  M: PDCSMediaCheck;
begin
  FMediaCheckLock.Enter;
  try
    M := New(PDCSMediaCheck);
    M^.SourceHandle := ASourceHandle;
    Move(AEventID, M^.EventID, SizeOf(TEventID));
    StrPCopy(M^.MediaID, AMediaID);
    M^.InTC := AInTC;
    M^.DurationTC := ADurationTC;

    FMediaCheckQueue.Add(M);
  finally
    FMediaCheckLock.Leave;
  end;

  SetEvent(FEventMediaCheck);

  Result := D_OK;
end;

function TDCSEventThread.TakeEvent(ASourceHandle: PSourceHandle; AEventID: TEventID; ADelayTime: TTimecode): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctTake, AEventID, ADelayTime);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TDCSEventThread.HoldEvent(ASourceHandle: PSourceHandle; AEventID: TEventID): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctHold, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TDCSEventThread.ChangeDurationEvent(ASourceHandle: PSourceHandle; AEventID: TEventID; ADurTimecode: TTimecode): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctChangeDuration, AEventID, ADurTimecode);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TDCSEventThread.GetOnAirEventID(ASourceHandle: PSourceHandle; var AOnAirEventID: TEventID; var ANextEventID: TEventID): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctOnAirEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    if (FCMD_Result = D_OK) then
    begin
      AOnAirEventID := FCMD_OnAirEventID;
      ANextEventID  := FCMD_NextEventID;
    end;
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TDCSEventThread.GetEventInfo(ASourceHandle: PSourceHandle; AEventID: TEventID; var AStartTime: TEventTime; var ADurationTC: TTimecode): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctEventInfo, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    if (FCMD_Result = D_OK) then
    begin
      AStartTime := FCMD_StartTime;
      ADurationTC := FCMD_DurationTC;
    end;
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

procedure TDCSEventThread.ResetMediaCheckIndex;
begin
  FMediaCheckLock.Enter;
  try
    FMediaCheckIndex := 0;
  finally
    FMediaCheckLock.Leave;
  end;
end;

function TDCSEventThread.SetControlBy(ASourceHandle: PSourceHandle): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctControlBy);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TDCSEventThread.GetExist(ASourceHandle: PSourceHandle; AMediaID: String; var AExist: Boolean): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctExist, AMediaID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    if (FCMD_Result = D_OK) then
    begin
      AExist := FCMD_Exist;
    end;
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TDCSEventThread.GetSize(ASourceHandle: PSourceHandle; AMediaID: String; var ADuration: TTimecode): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctSize, AMediaID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    if (FCMD_Result = D_OK) then
    begin
      ADuration := FCMD_DurationTC;
    end;
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TDCSEventThread.SetXpt(ASourceHandle: PSourceHandle; AInput: Integer): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASourceHandle, ctSetXpt, AInput);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  if (WaitForSingleObject(FEventCompleted, GV_SettingDCS.CommandTimeout) = WAIT_OBJECT_0) then
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

procedure TDCSEventThread.DoEventCommand;
var
  R: Integer;
  StartTime: TEventTime;
begin
  ResetEvent(FEventCommand);

  FCMD_Result := D_FALSE;

  try
    if (not HasMainControl) then exit;

    with FCommand do
    begin
      if (SourceHandle = nil) then exit;
      if (SourceHandle^.DCS = nil) then exit;

      if (SourceHandle^.DCS^.Alive) and
         (SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
      begin
        case CommandType of
          ctControlBy:
          begin
            R := DCSSetControlBy(SourceHandle^.DCS^.ID, SourceHandle^.Handle);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSSetControlBy succeeded, dcs = %d, handle = %d', [SourceHandle^.DCS^.ID, SourceHandle^.Handle])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSSetControlBy, error = %d, dcs = %d, handle = %d', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle])));
          end;
      {    ctDelete:
            R := DCSDeleteEvent(FCommand.DCSID, FCommand.DeviceHandle, FCommand.EventID);
      }
          ctClear:
          begin
            ClearEventQueueByChannelID(SourceHandle, ChannelID);
            ClearMediaCheckQueueByChannelID(SourceHandle, ChannelID);

            R := DCSClearEvent(SourceHandle^.DCS^.ID, SourceHandle^.Handle, ChannelID);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSClearEvent succeeded, dcs = %d, handle = %d, channel = %d', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, ChannelID])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSClearEvent failed, error = %d, dcs = %d, handle = %d, channel = %d', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle, ChannelID])));
          end;

          ctTake:
          begin
            ClearEventQueueByChannelID(SourceHandle, ChannelID);

            StartTime := GetPlusEventTime(DateTimeToEventTime(Now), TimecodeToEventTime(Timecode));
            R := DCSTakeEvent(SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventID, StartTime);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSTakeEvent succeeded, dcs = %d, handle = %d, id = %s, start = %s', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(EventID), EventTimeToString(StartTime)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSTakeEvent failed, error = %d, dcs = %d, handle = %d, id = %s, start = %s', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(EventID), EventTimeToString(StartTime)])));
          end;

          ctHold:
          begin
            R := DCSHoldEvent(SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventID);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSHoldEvent succeeded, dcs = %d, handle = %d, id = %s', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSHoldEvent failed, error = %d, dcs = %d, handle = %d, id = %s', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(EventID)])));
          end;

          ctChangeDuration:
          begin
            R := DCSChangetDurationEvent(SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventID, Timecode);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSHoldEvent succeeded, dcs = %d, handle = %d, id = %s, duration = %s', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(EventID), TimecodeToString(Timecode)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSHoldEvent failed, error = %d, dcs = %d, handle = %d, id = %s, duration = %s', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(EventID), TimecodeToString(Timecode)])));
          end;

          ctOnAirEventID:
          begin
            R := DCSGetOnAirEventID(SourceHandle^.DCS^.ID, SourceHandle^.Handle, FCMD_OnAirEventID, FCMD_NextEventID);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSGetOnAirEventID succeeded, dcs = %d, handle = %d, onairid = %s, nextid = %s', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(FCMD_OnAirEventID), EventIDToString(FCMD_NextEventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSGetOnAirEventID failed, error = %d, dcs = %d, handle = %d', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle])));
          end;

          ctEventInfo:
          begin
            R := DCSGetEventInfo(SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventID, FCMD_StartTime, FCMD_DurationTC);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSGetEventInfo succeeded, dcs = %d, handle = %d, id = %s, start = %s, duration = %s', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(EventID), EventTimeToString(FCMD_StartTime), TimecodeToString(FCMD_DurationTC)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSGetEventInfo failed, error = %d, dcs = %d, handle = %d, id = %s', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle, EventIDToString(EventID)])));
          end;

          ctExist:
          begin
            R := DCSGetExist(SourceHandle^.DCS^.ID, SourceHandle^.Handle, MediaID, FCMD_Exist);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSGetExist succeeded, dcs = %d, handle = %d, media = %s, exist = %s', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, MediaID, BoolToStr(FCMD_Exist, True)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSGetExist failed, error = %d, dcs = %d, handle = %d, media = %s', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle, MediaID])));
          end;

          ctSize:
          begin
            R := DCSGetSize(SourceHandle^.DCS^.ID, SourceHandle^.Handle, MediaID, FCMD_DurationTC);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSGetSize succeeded, dcs = %d, handle = %d, media = %s, exist = %s', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, MediaID, TimecodeToString(FCMD_DurationTC)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSGetSize failed, error = %d, dcs = %d, handle = %d, media = %s', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle, MediaID])));
          end;

          ctSetXpt:
          begin
            R := DCSSetXpt(SourceHandle^.DCS^.ID, SourceHandle^.Handle, XptInput);

            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('DoEventCommand DCSSetXpt succeeded, dcs = %d, handle = %d, input = %d', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, XptInput])))
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventCommand DCSSetXpt failed, error = %d, dcs = %d, handle = %d, input = %d', [R, SourceHandle^.DCS^.ID, SourceHandle^.Handle, XptInput])));
          end;

          else
          begin
            R := D_FALSE;

            Assert(False, GetMainLogStr(lsError, Format('DoEventCommand undefined command, dcs = %d, handle = %d, command = %d', [SourceHandle^.DCS^.ID, SourceHandle^.Handle, Integer(CommandType)])));
          end;
        end;
      end;
      
      FCMD_Result := R;
    end;
  finally
    ResetCommand;

    SetEvent(FEventCompleted);
  end;
end;

procedure TDCSEventThread.DoEventProc;
var
  E: PDCSEvent;
  R: Integer;
  ChannelForm: TfrmChannel;

  MediaExist: Boolean;
  MediaDuration: TTimecode;
  EventDuration: TTimecode;
begin
//  ResetEvent(FEventProc);

  if (FEventQueue.Count <= 0) then
  begin
    ResetEvent(FEventProc);
    exit;
  end;

  try
  FEventLock.Enter;
  try
    E := FEventQueue.First;

    if (not HasMainControl) then exit;

  finally
    FEventLock.Leave;
  end;

  if (E = nil) then exit;
  if (E^.SourceHandle = nil) then exit;
  if (E^.SourceHandle^.DCS = nil) then exit;

  if (E^.SourceHandle^.DCS^.Alive) and
     (E^.SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
  begin
    case E^.EventType of
      etInput:
      begin
        R := DCSInputEvent(E^.SourceHandle^.DCS^.ID, E^.SourceHandle^.Handle, E^.Event);
        if (R <> D_OK) then
        begin
          E^.Event.Status.State := esError;
          E^.Event.Status.ErrorCode := R;

          if (E^.Event.EventType = ET_PLAYER) and (E^.SourceHandle^.DCS^.Main) then
          begin
            ChannelForm := FSEC.GetChannelFormByID(E^.Event.EventID.ChannelID);
            if (ChannelForm <> nil) then
            begin
              ChannelForm.SetEventStatus(E^.Event.EventID, E^.Event.Status);
            end;
          end;
        end;

        if (R = D_OK) then
          Assert(False, GetMainLogStr(lsNormal, Format('DoEventProc DCSInputEvent succeeded, dcs = %d, handle = %d, id = %s, start = %s', [E^.SourceHandle^.DCS^.ID, Integer(E^.SourceHandle^.Handle), EventIDToString(E^.Event.EventID), EventTimeToString(E^.Event.StartTime)])))
        else
          Assert(False, GetMainLogStr(lsError, Format('DoEventProc DCSInputEvent failed, error = %d, dcs = %d, handle = %d, id = %s, start = %s', [R, E^.SourceHandle^.DCS^.ID, Integer(E^.SourceHandle^.Handle), EventIDToString(E^.Event.EventID), EventTimeToString(E^.Event.StartTime)])));
      end;

      etDelete:
      begin
        R := DCSDeleteEvent(E^.SourceHandle^.DCS^.ID, E^.SourceHandle^.Handle, E^.EventID);
        if (R = D_OK) then
          Assert(False, GetMainLogStr(lsNormal, Format('DoEventProc DCSDeleteEvent succeeded, dcs = %d, handle = %d, id = %s', [E^.SourceHandle^.DCS^.ID, Integer(E^.SourceHandle^.Handle), EventIDToString(E^.Event.EventID)])))
        else
          Assert(False, GetMainLogStr(lsError, Format('DoEventProc DCSDeleteEvent failed, error = %d, dcs = %d, handle = %d, id = %s', [R, E^.SourceHandle^.DCS^.ID, Integer(E^.SourceHandle^.Handle), EventIDToString(E^.Event.EventID)])));
      end;
    end;
  end;

  finally
  FEventLock.Enter;
  try
  finally
    FEventQueue.Remove(E);

  Dispose(E);
    FEventLock.Leave;
  end;
  end;
end;

procedure TDCSEventThread.DoEventMediaCheck;
var
  M: PDCSMediaCheck;
  CheckTime: TEventTime;

  R: Integer;
  ChannelForm: TfrmChannel;

  MediaExist: Boolean;
  MediaDuration: TTimecode;
  EventDuration: TTimecode;

  MediaStatus: TMediaStatus;
begin
  ResetEvent(FEventMediaCheck);

  if (FMediaCheckQueue.Count <= 0) then
  begin
//    ResetEvent(FEventMediaCheck);
    exit;
  end;

//  FMediaCheckLock.Enter;

  if (FMediaCheckIndex < 0) or (FMediaCheckIndex >= FMediaCheckQueue.Count) then
  begin
    FMediaCheckLock.Enter;
    try
      FMediaCheckIndex := 0;
    finally
      FMediaCheckLock.Leave;
    end;
  end;

  try
  FMediaCheckLock.Enter;
  try
//    M := FMediaCheckQueue.First;

    M := FMediaCheckQueue[FMediaCheckIndex];
    if (not HasMainControl) then exit;
  finally
    FMediaCheckLock.Leave;
  end;

    try
      if (M = nil) then exit;
      if (M^.SourceHandle = nil) then exit;
      if (M^.SourceHandle^.DCS = nil) then exit;

      if (M^.SourceHandle^.DCS^.Alive) and
         (M^.SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
      begin
        ChannelForm := FSEC.GetChannelFormByID(M^.EventID.ChannelID);
        if (ChannelForm <> nil) then
        begin
          // 현재시각이 다음 이벤트 준비시간보다 작으면 Media check를 보류하고 다음 호출에 재검사
          if (ChannelForm.CueSheetNext <> nil) then
//             (String(ChannelForm.CueSheetNext^.Source) = String(M^.Item^.Source)) then
          begin
            CheckTime := GetPlusEventTime(DateTimeToEventTime(Now), TimecodeToEventTime(GV_SettingThresholdTime.OnAirLockTime));
            if (CompareEventTime(ChannelForm.CueSheetNext^.StartTime, CheckTime) <= 0) then
            begin
//              FMediaCheckLock.Enter;
//              try
                Inc(FMediaCheckIndex);
//              finally
//                FMediaCheckLock.Leave;
//              end;
              exit;
            end;
          end;
        end;

        R := DCSGetExist(M^.SourceHandle^.DCS^.ID, M^.SourceHandle^.Handle, M^.MediaId, MediaExist);
        if (R = D_OK) then
        begin
          Assert(False, GetMainLogStr(lsNormal, Format('DoEventMediaCheck DCSGetExist succeeded, dcs = %d, handle = %d, media = %s, exist = %s', [M^.SourceHandle^.DCS^.ID, M^.SourceHandle^.Handle, String(M^.MediaID), BoolToStr(MediaExist, True)])));

          if (MediaExist) then
          begin
            Assert(False, GetMainLogStr(lsNormal, Format('DoEventMediaCheck DCSGetExist succeeded, ''111''', [])));
            MediaStatus := msEqual;
            Assert(False, GetMainLogStr(lsNormal, Format('DoEventMediaCheck DCSGetExist succeeded, ''222''', [])));

            R := DCSGetSize(M^.SourceHandle^.DCS^.ID, M^.SourceHandle^.Handle, M^.MediaId, MediaDuration);
            Assert(False, GetMainLogStr(lsNormal, Format('DoEventMediaCheck DCSGetSize succeeded, ''111''', [])));
            if (R = D_OK) then
            begin
              EventDuration := GetPlusTimecode(M^.InTC, M^.DurationTC);
              if (EventDuration < MediaDuration) then
                MediaStatus := msLong
              else if (EventDuration > MediaDuration) then
                MediaStatus := msShort;

              Assert(False, GetMainLogStr(lsNormal, Format('DoEventMediaCheck DCSGetSize succeeded, dcs = %d, handle = %d, media = %s, duration = %s', [M^.SourceHandle^.DCS^.ID, M^.SourceHandle^.Handle, String(M^.MediaID), TimecodeToString(MediaDuration)])));
            end
            else
              Assert(False, GetMainLogStr(lsError, Format('DoEventMediaCheck DCSGetSize failed, error = %d, dcs = %d, handle = %d, media = %s', [R, M^.SourceHandle^.DCS^.ID, M^.SourceHandle^.Handle, String(M^.MediaID)])));
          end
          else
          begin
            MediaStatus := msNotExist;
          end;
        end
        else
        begin
          MediaStatus := msNotExist;

          Assert(False, GetMainLogStr(lsError, Format('DoEventMediaCheck DCSGetExist failed, error = %d, dcs = %d, handle = %d, media = %s', [R, M^.SourceHandle^.DCS^.ID, M^.SourceHandle^.Handle, String(M^.MediaID)])));
        end;

    //    ChannelForm := FSEC.GetChannelFormByID(M^.Item.EventID.ChannelID);
        if (ChannelForm <> nil) then
        begin
    //      PostMessage(ChannelForm.Handle, WM_UPDATE_MEDIA_CHECK, 0, NativeInt(M^.Item));
          ChannelForm.SetMediaStatus(M^.EventID, MediaStatus);
        end;
      end;

      FMediaCheckQueue.Remove(M);
      Dispose(M);
    finally
    end;
  finally
//    FMediaCheckLock.Leave;
  end;
end;

procedure TDCSEventThread.DoEventDeviceCheck;
var
  I, J: Integer;
  Source: PSource;
  SourceHandle: PSourceHandle;
  R: Integer;

  SaveStatus: TDeviceStatus;
  Status: TDeviceStatus;
  SourceOK: Boolean;

  DeviceForm: TfrmDevice;

  WarningCommText: String;
  WarningConnectText: String;
  WarningText: String;
begin
  if (not HasMainControl) then exit;

  if (GV_SettingOption.OnAirCheckDeviceNotify) then exit;

  WarningCommText := '';
  WarningConnectText := '';

  // Device check
  for I := 0 to GV_SourceList.Count - 1 do
  begin
    Source := GV_SourceList[I];
    SaveStatus := Source.Status;

    SourceOK := False;
    if (Source <> nil) and (Source^.Handles <> nil) then
    begin
      for J := 0 to Source^.Handles.Count - 1 do
      begin
        SourceHandle := Source^.Handles[J];
        if (SourceHandle = nil) then continue;
        if (SourceHandle^.DCS = nil) then continue;

        if (SourceHandle^.DCS^.Main) and (SourceHandle^.DCS^.Alive) then
        begin
          R := DCSGetDeviceStatus(SourceHandle^.DCS^.ID, SourceHandle^.Handle, Status);
          if (R = D_OK) then
          begin
            SourceOK := True;
            break;
          end;
        end;
      end;

      Source.CommSuccess := SourceOK;
      if (SourceOK) then Source.CommTimeout := 0;

      DeviceForm := FSEC.GetDeviceForm;
      if (DeviceForm <> nil) then
      begin
        if (SourceOK) then
        begin
          DeviceForm.SetDeviceStatus(Source, SourceHandle, Status);
          if (not Status.Connected) then
          begin
            if (WarningConnectText = '') then
              WarningConnectText := SWDeviceStatusNotConnect + #13#10 + Format('Device=%s, DCS=%s', [String(Source^.Name), SourceHandle^.DCS^.Name])
            else
              WarningConnectText := WarningConnectText + #13#10 + Format('Device=%s, DCS=%s', [String(Source^.Name), SourceHandle^.DCS^.Name]);
          end;
        end
        else
        begin
          DeviceForm.SetDeviceCommError(String(Source^.Name), Status);
          if (WarningCommText = '') then
            WarningCommText := SWDeviceCommError + #13#10 + Format('%s', [String(Source^.Name)])
          else
            WarningCommText := WarningCommText + #13#10 + Format('%s', [String(Source^.Name)]);
        end;
      end;
    end;
  end;

  if (WarningCommText <> '') or (WarningConnectText <> '') then
  begin
//    if (frmSEC.FWarningDialogDeviceCheck = nil) or (not frmSEC.FWarningDialogDeviceCheck.Showing) then
    begin
      if (WarningCommText = '') then WarningText := WarningConnectText
      else if (WarningConnectText = '') then WarningText := WarningCommText
      else WarningText := WarningCommText + #13#10 + #13#10 + WarningConnectText;

      if (WarningText <> '') then
      begin
        System.Classes.TThread.CreateAnonymousThread(
          procedure
          begin
            PostMessage(FSEC.Handle, WM_WARNING_DISPLAY_DEVICE_CHECK, 0, LParam(PChar(WarningText)));
          end).Start;
      end;
    end;
  end;
end;

procedure TDCSEventThread.Execute;
var
  R: Cardinal;
  WaitList: array[0..3] of THandle;
begin
  WaitList[0] := FEventClose;
  WaitList[1] := FEventCommand;
  WaitList[2] := FEventProc;
  WaitList[3] := FEventMediaCheck;
//  WaitList[4] := GV_TimerExecuteEvent;

{  FDeleteEvent;
  WaitList[2] := FClearEvent;
  WaitList[3] := FTakeEvent;
  WaitList[4] := FHoldEvent;
  WaitList[5] := FChangeDurationEvent;
  WaitList[6] := FEventClose; }
  while not Terminated do
  begin
    R := WaitForMultipleObjects(4, @WaitList, False, TimecodeToMilliSec(GV_SettingTimeParameter.OnAirCheckDeviceStatusInterval));// INFINITE);
    case R of
      WAIT_OBJECT_0: break;
      WAIT_OBJECT_0 + 1: DoEventCommand;
      WAIT_OBJECT_0 + 2: DoEventProc;
      WAIT_OBJECT_0 + 3: DoEventMediaCheck;
      else
      begin
        DoEventDeviceCheck;
        MediaCheckQueueCheck;
      end;
//      WAIT_OBJECT_0 + 4: DoEventDeviceCheck;
//      WAIT_OBJECT_0 + 4: DoEventHold;
//      WAIT_OBJECT_0 + 5: DoEventChangeDuration;
//      else
//        DoEventProc;
    end;
  end;
end;

{ TSECEventThread }

constructor TSECEventThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FEventLock   := TCriticalSection.Create;
  FEventQueue  := TSECEventList.Create;

  FInputIndex := -1;
  FInputCount := 0;

  FCMD_Result      := D_FALSE;
  FCMD_StartTime.D := 0;
  FCMD_StartTime.T := 0;
  FCMD_DurationTC  := 0;
  FCMD_Exist       := False;

  FEventCommand := CreateEvent(nil, True, False, nil);
  FEventProc := CreateEvent(nil, True, False, nil);

  FEventCompleted := CreateEvent(nil, True, True, nil);

  FEventClose := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TSECEventThread.Destroy;
begin
  CloseHandle(FEventProc);
  CloseHandle(FEventCommand);
  CloseHandle(FEventCompleted);

  CloseHandle(FEventClose);

  ClearEventQueue;
  FreeAndNil(FEventQueue);

  FreeAndNil(FEventLock);

  inherited Destroy;
end;

procedure TSECEventThread.ClearEventQueue;
var
  I: Integer;
  S: PSECEvent;
begin
  FEventLock.Enter;
  try
    for I := FEventQueue.Count - 1 downto 0 do
    begin
      S := FEventQueue[I];
      if (S <> nil) then
      begin
        Dispose(S);
      end;
    end;
    FEventQueue.Clear;
  finally
    FEventLock.Leave;
  end;
end;

procedure TSECEventThread.ClearEventQueueByChannelID(ASEC: PSEC; AChannelID: Word);
var
  I: Integer;
  S: PSECEvent;
begin
  FEventLock.Enter;
  try
    for I := FEventQueue.Count - 1 downto 0 do
    begin
      S := FEventQueue[I];
      if (S = nil) then continue;
      if (S^.SEC = nil) then continue;

      if (S^.SEC = ASEC) and
         (S^.Item.EventID.ChannelID = AChannelID) then
      begin
        FEventQueue.Remove(S);
        Dispose(S);
      end;
    end;
  finally
    FEventLock.Leave;
  end;
end;

procedure TSECEventThread.EventQueueCheck;
begin
  if (FEventQueue.Count > 0) then SetEvent(FEventProc);
end;

procedure TSECEventThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FEventCommand);
  SetEvent(FEventClose);
end;

procedure TSECEventThread.ResetCommand;
begin
  FillChar(FCommand, SizeOf(TSECCommand), #0);
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType);
begin
  WaitForSingleObject(FEventCompleted, INFINITE);
  ResetCommand;

  FCommand.SEC := ASEC;
  FCommand.CommandType  := ACommandType;
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; ADeviceStatus: TDeviceStatus; ADeviceName: PChar);
begin
  SetCommand(ASEC, ACommandType);

  Move(ADeviceStatus, FCommand.DeviceStatus, SizeOf(TDeviceStatus));
  StrPCopy(FCommand.DeviceName, ADeviceName);
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; ADeviceStatus: TDeviceStatus; ADCSID: Word; ADeviceHandle: TDeviceHandle);
begin
  SetCommand(ASEC, ACommandType);

  Move(ADeviceStatus, FCommand.DeviceStatus, SizeOf(TDeviceStatus));
  FCommand.DCSID := ADCSID;
  FCommand.DeviceHandle := ADeviceHandle;
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AEventID: TEventID);
begin
  SetCommand(ASEC, ACommandType);

  Move(AEventID, FCommand.EventID, SizeOf(TEventID));
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AEventID: TEventID; AEventStatus: TEventStatus);
begin
  SetCommand(ASEC, ACommandType);

  Move(AEventID, FCommand.EventID, SizeOf(TEventID));
  Move(AEventStatus, FCommand.EventStatus, SizeOf(TEventStatus));
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AEventID: TEventID; AMediaStatus: TMediaStatus);
begin
  SetCommand(ASEC, ACommandType);

  Move(AEventID, FCommand.EventID, SizeOf(TEventID));
  Move(AMediaStatus, FCommand.MediaStatus, SizeOf(TMediaStatus));
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AChannelID: Word);
begin
  SetCommand(ASEC, ACommandType);

  FCommand.ChannelID := AChannelID;
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AChannelID: Word; AIsOnAir: Boolean);
begin
  SetCommand(ASEC, ACommandType);

  FCommand.ChannelID := AChannelID;
  FCommand.IsOnAir := AIsOnAir;
end;

procedure TSECEventThread.SetCommand(ASEC: PSEC; ACommandType: TSECCommandType; AChannelID: Word; AIndex: Integer; ACount: Integer = 1);
begin
  SetCommand(ASEC, ACommandType);

  FCommand.ChannelID := AChannelID;
  FCommand.Index := AIndex;
  FCommand.Count := ACount;
end;

function TSECEventThread.SetDeviceCommError(ASEC: PSEC; ADeviceStatus: TDeviceStatus; ADeviceName: String): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASEC, sctDeviceError, ADeviceStatus, PChar(ADeviceName));

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TSECEventThread.SetDeviceStatus(ASEC: PSEC; ADeviceStatus: TDeviceStatus; ADCSID: Word; ADeviceHandle: TDeviceHandle): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASEC, sctDeviceStatus, ADeviceStatus, ADCSID, ADeviceHandle);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TSECEventThread.SetEventStatus(ASEC: PSEC; AEventID: TEventID; AEventStatus: TEventStatus): Integer;
begin
  Result := D_FALSE;

//  Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.SetEventStatus, 111 SEC=%d, ID=%s', [ASEC^.ID, EventIDToString(AEventID)])));
  SetCommand(ASEC, sctEventStatus, AEventID, AEventStatus);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.SetEventStatus, 222 SEC=%d, ID=%s', [ASEC^.ID, EventIDToString(AEventID)])));
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;
//  Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.SetEventStatus, 333 SEC=%d, ID=%s', [ASEC^.ID, EventIDToString(AEventID)])));

  EventQueueCheck;
end;

function TSECEventThread.SetMediaStatus(ASEC: PSEC; AEventID: TEventID; AMediaStatus: TMediaStatus): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASEC, sctMediaStatus, AEventID, AMediaStatus);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TSECEventThread.SetOnAir(ASEC: PSEC; AChannelID: Word; AIsOnAir: Boolean): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASEC, sctOnAir, AChannelID, AIsOnAir);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TSECEventThread.SetCueSheetCurr(ASEC: PSEC; AEventID: TEventID): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASEC, sctCurr, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TSECEventThread.SetCueSheetNext(ASEC: PSEC; AEventID: TEventID): Integer;
begin
  Result := D_FALSE;

//  Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.SetCueSheetNext, 111 SEC=%d, ID=%s', [ASEC^.ID, EventIDToString(AEventID)])));
  SetCommand(ASEC, sctNext, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.SetCueSheetNext, 222 SEC=%d, ID=%s', [ASEC^.ID, EventIDToString(AEventID)])));
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;
//  Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.SetCueSheetNext, 333 SEC=%d, ID=%s', [ASEC^.ID, EventIDToString(AEventID)])));

  EventQueueCheck;
end;

function TSECEventThread.SetCueSheetTarget(ASEC: PSEC; AEventID: TEventID): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASEC, sctTarget, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TSECEventThread.InputCueSheet(ASEC: PSEC; AIndex: Integer; AItem: TCueSheetItem): Integer;
var
  E: PSECEvent;
begin
  FEventLock.Enter;
  try
    E := New(PSECEvent);
    E^.SEC := ASEC;
    E^.EventType := setInput;
    E^.Index := AIndex;
    E^.Item := AItem;
    Move(AItem, E^.Item, SizeOf(TCueSheetItem));

    FEventQueue.Add(E);
  finally
    FEventLock.Leave;
  end;

  SetEvent(FEventProc);

  Result := D_OK;
end;

function TSECEventThread.DeleteCueSheet(ASEC: PSEC; AEventID: TEventID): Integer;
var
  E: PSECEvent;
begin
  FEventLock.Enter;
  try
    E := New(PSECEvent);
    E^.SEC := ASEC;
    E^.EventType := setDelete;
    Move(AEventID, E^.EventID, SizeOf(TEventID));

    FEventQueue.Add(E);
  finally
    FEventLock.Leave;
  end;

  SetEvent(FEventProc);

  Result := D_OK;
end;

function TSECEventThread.ClearCueSheet(ASEC: PSEC; AChannelID: Word): Integer;
begin
  Result := D_FALSE;

  SetCommand(ASEC, sctClear, AChannelID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

procedure TSECEventThread.DoEventCommand;
var
  R: Integer;

  ChannelForm: TfrmChannel;
  I: Integer;
begin
  ResetEvent(FEventCommand);

  FCMD_Result := D_FALSE;

  try
    if (not HasMainControl) then exit;

    with FCommand do
    begin
      if (SEC = nil) then exit;
//      Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand, SEC=nil', [])));

//      Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand, IP=%s, Type=%d', [SEC^.HostIP, Integer(CommandType)])));
      if (SEC^.Alive) then
      begin
        case CommandType of
          sctDeviceError:
          begin
            R := SECSetDeviceCommError(SEC^.HostIP, DeviceStatus, DeviceName);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECSetDeviceCommError succeeded, sec = %d, name = %s', [SEC^.ID, DeviceName])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECSetDeviceCommError, error = %d, sec = %d, name = %s', [R, SEC^.ID, DeviceName])));
          end;
          sctDeviceStatus:
          begin
            R := SECSetDeviceStatus(SEC^.HostIP, DCSID, DeviceHandle, DeviceStatus);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECSetDeviceStatus succeeded, sec = %d, handle = %d', [SEC^.ID, DeviceHandle])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECSetDeviceStatus, error = %d, sec = %d, handle = %d', [R, SEC^.ID, DeviceHandle])));
          end;
          sctEventStatus:
          begin
            R := SECSetEventStatus(SEC^.HostIP, EventID, EventStatus);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECSetEventStatus succeeded, sec = %d, id = %s', [SEC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECSetEventStatus, error = %d, sec = %d, id = %s', [R, SEC^.ID, EventIDToString(EventID)])));
          end;
          sctMediaStatus:
          begin
            R := SECSetMediaStatus(SEC^.HostIP, EventID, MediaStatus);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECSetMediaStatus succeeded, sec = %d, id = %s', [SEC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECSetMediaStatus, error = %d, sec = %d, id = %s', [R, SEC^.ID, EventIDToString(EventID)])));
          end;
          sctOnAir:
          begin
            R := SECSetOnAir(SEC^.HostIP, ChannelID, IsOnAir);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECSetOnAir succeeded, sec = %d, channel = %d, onair = %s', [SEC^.ID, ChannelID, BoolToStr(IsOnAir, True)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECSetOnAir, error = %d, sec = %d, channel = %d, onair = %s', [R, SEC^.ID, ChannelID, BoolToStr(IsOnAir, True)])));
          end;
          sctCurr:
          begin
            R := SECSetCueSheetCurr(SEC^.HostIP, EventID);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECSetCueSheetCurr succeeded, sec = %d, id = %s', [SEC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECSetCueSheetCurr, error = %d, sec = %d, id = %s', [R, SEC^.ID, EventIDToString(EventID)])))
          end;
          sctNext:
          begin
            R := SECSetCueSheetNext(SEC^.HostIP, EventID);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECSetCueSheetNext succeeded, sec = %d, id = %s', [SEC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECSetCueSheetNext, error = %d, sec = %d, id = %s', [R, SEC^.ID, EventIDToString(EventID)])))
          end;
          sctTarget:
          begin
            R := SECSetCueSheetTarget(SEC^.HostIP, EventID);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECSetCueSheetTarget succeeded, sec = %d, id = %s', [SEC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECSetCueSheetTarget, error = %d, sec = %d, id = %s', [R, SEC^.ID, EventIDToString(EventID)])))
          end;
          sctClear:
          begin
            ClearEventQueueByChannelID(SEC, ChannelID);
            R := SECClearCueSheet(SEC^.HostIP, ChannelID);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventCommand SECClearCueSheet succeeded, sec = %d, channel = %d', [SEC^.ID, ChannelID])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventCommand SECClearCueSheet, error = %d, sec = %d, channel = %d', [R, SEC^.ID, ChannelID])))
          end;
          else
            R := D_FALSE;
        end;
      end;

      FCMD_Result := R;
    end;
  finally
    ResetCommand;

    SetEvent(FEventCompleted);
  end;
end;

procedure TSECEventThread.DoEventProc;
var
  E: PSECEvent;
  R: Integer;
  ChannelForm: TfrmChannel;
begin
//  ResetEvent(FEventProc);

  if (FEventQueue.Count <= 0) then
  begin
    ResetEvent(FEventProc);
    exit;
  end;

  try
  FEventLock.Enter;
  try
    E := FEventQueue.First;

    if (not HasMainControl) then exit;

  finally
    FEventLock.Leave;
  end;

  if (E = nil) then exit;
  if (E^.SEC = nil) then exit;

  if (E^.SEC^.Alive) then
  begin
    case E^.EventType of
      setInput:
      begin
        R := SECInputCueSheet(E^.SEC^.HostIP, E^.Index, E^.Item);
        if (R = D_OK) then
          Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventProc SECInputCueSheet succeeded, sec = %d, id = %s', [E^.SEC^.ID, EventIDToString(E^.Item.EventID)])))
        else
          Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventProc SECInputCueSheet, error = %d, sec = %d, id = %s', [R, E^.SEC^.ID, EventIDToString(E^.Item.EventID)])))
      end;
      setDelete:
      begin
        R := SECDeleteCueSheet(E^.SEC^.HostIP, E^.EventID);
        if (R = D_OK) then
          Assert(False, GetMainLogStr(lsNormal, Format('TSECEventThread.DoEventProc SECDeleteCueSheet succeeded, sec = %d, id = %s', [E^.SEC^.ID, EventIDToString(E^.EventID)])))
        else
          Assert(False, GetMainLogStr(lsError, Format('TSECEventThread.DoEventProc SECDeleteCueSheet, error = %d, sec = %d, id = %s', [R, E^.SEC^.ID, EventIDToString(E^.EventID)])))
      end;
    end;
  end;


  finally
  FEventLock.Enter;
  try
  finally
    FEventQueue.Remove(E);

  Dispose(E);
    FEventLock.Leave;
  end;
  end;
end;

procedure TSECEventThread.Execute;
var
  R: Cardinal;
  WaitList: array[0..2] of THandle;
begin
  WaitList[0] := FEventClose;
  WaitList[1] := FEventCommand;
  WaitList[2] := FEventProc;

  while not Terminated do
  begin
    R := WaitForMultipleObjects(3, @WaitList, False, INFINITE);
    case R of
      WAIT_OBJECT_0: break;
      WAIT_OBJECT_0 + 1: DoEventCommand;
      WAIT_OBJECT_0 + 2: DoEventProc;
    end;
  end;
end;

{ TMCCEventThread }

constructor TMCCEventThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FEventLock   := TCriticalSection.Create;
  FEventQueue  := TMCCEventList.Create;

  FInputIndex := -1;
  FInputCount := 0;

  FCMD_Result      := D_FALSE;
  FCMD_StartTime.D := 0;
  FCMD_StartTime.T := 0;
  FCMD_DurationTC  := 0;
  FCMD_Exist       := False;

  FEventCommand := CreateEvent(nil, True, False, nil);
  FEventProc := CreateEvent(nil, True, False, nil);

  FEventCompleted := CreateEvent(nil, True, True, nil);

  FEventClose := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TMCCEventThread.Destroy;
begin
  CloseHandle(FEventProc);
  CloseHandle(FEventCommand);
  CloseHandle(FEventCompleted);

  CloseHandle(FEventClose);

  ClearEventQueue;
  FreeAndNil(FEventQueue);

  FreeAndNil(FEventLock);

  inherited Destroy;
end;

procedure TMCCEventThread.ClearEventQueue;
var
  I: Integer;
  M: PMCCEvent;
begin
  FEventLock.Enter;
  try
    for I := FEventQueue.Count - 1 downto 0 do
    begin
      M := FEventQueue[I];
      if (M <> nil) then
      begin
        Dispose(M);
      end;
    end;
    FEventQueue.Clear;
  finally
    FEventLock.Leave;
  end;
end;

procedure TMCCEventThread.ClearEventQueueByChannelID(AMCC: PMCC; AChannelID: Word);
var
  I: Integer;
  M: PMCCEvent;
begin
  FEventLock.Enter;
  try
    for I := FEventQueue.Count - 1 downto 0 do
    begin
      M := FEventQueue[I];
      if (M = nil) then continue;
      if (M^.MCC = nil) then continue;

      if (M^.MCC = AMCC) and
         (M^.Item.EventID.ChannelID = AChannelID) then
      begin
        FEventQueue.Remove(M);
        Dispose(M);
      end;
    end;
  finally
    FEventLock.Leave;
  end;
end;

procedure TMCCEventThread.EventQueueCheck;
begin
  if (FEventQueue.Count > 0) then SetEvent(FEventProc);
end;

procedure TMCCEventThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FEventCommand);
  SetEvent(FEventClose);
end;

procedure TMCCEventThread.ResetCommand;
begin
  FillChar(FCommand, SizeOf(TSECCommand), #0);
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType);
begin
  WaitForSingleObject(FEventCompleted, INFINITE);
  ResetCommand;

  FCommand.MCC := AMCC;
  FCommand.CommandType  := ACommandType;
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; ADeviceStatus: TDeviceStatus; ADeviceName: PChar);
begin
  SetCommand(AMCC, ACommandType);

  Move(ADeviceStatus, FCommand.DeviceStatus, SizeOf(TDeviceStatus));
  StrPCopy(FCommand.DeviceName, ADeviceName);
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; ADeviceStatus: TDeviceStatus; ADCSID: Word; ADeviceHandle: TDeviceHandle);
begin
  SetCommand(AMCC, ACommandType);

  Move(ADeviceStatus, FCommand.DeviceStatus, SizeOf(TDeviceStatus));
  FCommand.DCSID := ADCSID;
  FCommand.DeviceHandle := ADeviceHandle;
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AEventID: TEventID);
begin
  SetCommand(AMCC, ACommandType);

  Move(AEventID, FCommand.EventID, SizeOf(TEventID));
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AEventID: TEventID; AEventStatus: TEventStatus);
begin
  SetCommand(AMCC, ACommandType);

  Move(AEventID, FCommand.EventID, SizeOf(TEventID));
  Move(AEventStatus, FCommand.EventStatus, SizeOf(TEventStatus));
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AEventID: TEventID; AMediaStatus: TMediaStatus);
begin
  SetCommand(AMCC, ACommandType);

  Move(AEventID, FCommand.EventID, SizeOf(TEventID));
  Move(AMediaStatus, FCommand.MediaStatus, SizeOf(TMediaStatus));
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AChannelID: Word);
begin
  SetCommand(AMCC, ACommandType);

  FCommand.ChannelID := AChannelID;
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AChannelID: Word; AIsOnAir: Boolean);
begin
  SetCommand(AMCC, ACommandType);

  FCommand.ChannelID := AChannelID;
  FCommand.IsOnAir := AIsOnAir;
end;

procedure TMCCEventThread.SetCommand(AMCC: PMCC; ACommandType: TMCCCommandType; AChannelID: Word; AIndex: Integer; ACount: Integer = 1);
begin
  SetCommand(AMCC, ACommandType);

  FCommand.ChannelID := AChannelID;
  FCommand.Index := AIndex;
  FCommand.Count := ACount;
end;

function TMCCEventThread.SetDeviceCommError(AMCC: PMCC; ADeviceStatus: TDeviceStatus; ADeviceName: String): Integer;
begin
  Result := D_FALSE;

  SetCommand(AMCC, mctDeviceError, ADeviceStatus, PChar(ADeviceName));

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TMCCEventThread.SetDeviceStatus(AMCC: PMCC; ADeviceStatus: TDeviceStatus; ADCSID: Word; ADeviceHandle: TDeviceHandle): Integer;
begin
  Result := D_FALSE;

  SetCommand(AMCC, mctDeviceStatus, ADeviceStatus, ADCSID, ADeviceHandle);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TMCCEventThread.SetEventStatus(AMCC: PMCC; AEventID: TEventID; AEventStatus: TEventStatus): Integer;
begin
  Result := D_FALSE;

//  Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.SetEventStatus, 111 SEC=%d, ID=%s', [AMCC^.ID, EventIDToString(AEventID)])));
  SetCommand(AMCC, mctEventStatus, AEventID, AEventStatus);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.SetEventStatus, 222 SEC=%d, ID=%s', [AMCC^.ID, EventIDToString(AEventID)])));
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;
//  Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.SetEventStatus, 333 SEC=%d, ID=%s', [AMCC^.ID, EventIDToString(AEventID)])));

  EventQueueCheck;
end;

function TMCCEventThread.SetMediaStatus(AMCC: PMCC; AEventID: TEventID; AMediaStatus: TMediaStatus): Integer;
begin
  Result := D_FALSE;

  SetCommand(AMCC, mctMediaStatus, AEventID, AMediaStatus);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TMCCEventThread.SetOnAir(AMCC: PMCC; AChannelID: Word; AIsOnAir: Boolean): Integer;
begin
  Result := D_FALSE;

  SetCommand(AMCC, mctOnAir, AChannelID, AIsOnAir);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TMCCEventThread.SetCueSheetCurr(AMCC: PMCC; AEventID: TEventID): Integer;
begin
  Result := D_FALSE;

  SetCommand(AMCC, mctCurr, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TMCCEventThread.SetCueSheetNext(AMCC: PMCC; AEventID: TEventID): Integer;
begin
  Result := D_FALSE;

//  Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.SetCueSheetNext, 111 SEC=%d, ID=%s', [AMCC^.ID, EventIDToString(AEventID)])));
  SetCommand(AMCC, mctNext, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

//  Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.SetCueSheetNext, 222 SEC=%d, ID=%s', [AMCC^.ID, EventIDToString(AEventID)])));
  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;
//  Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.SetCueSheetNext, 333 SEC=%d, ID=%s', [AMCC^.ID, EventIDToString(AEventID)])));

  EventQueueCheck;
end;

function TMCCEventThread.SetCueSheetTarget(AMCC: PMCC; AEventID: TEventID): Integer;
begin
  Result := D_FALSE;

  SetCommand(AMCC, mctTarget, AEventID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

function TMCCEventThread.InputCueSheet(AMCC: PMCC; AIndex: Integer; AItem: TCueSheetItem): Integer;
var
  E: PMCCEvent;
begin
  FEventLock.Enter;
  try
    E := New(PMCCEvent);
    E^.MCC := AMCC;
    E^.EventType := metInput;
    E^.Index := AIndex;
    E^.Item := AItem;
    Move(AItem, E^.Item, SizeOf(TCueSheetItem));

    FEventQueue.Add(E);
  finally
    FEventLock.Leave;
  end;

  SetEvent(FEventProc);

  Result := D_OK;
end;

function TMCCEventThread.DeleteCueSheet(AMCC: PMCC; AEventID: TEventID): Integer;
var
  E: PMCCEvent;
begin
  FEventLock.Enter;
  try
    E := New(PMCCEvent);
    E^.MCC := AMCC;
    E^.EventType := metDelete;
    Move(AEventID, E^.EventID, SizeOf(TEventID));

    FEventQueue.Add(E);
  finally
    FEventLock.Leave;
  end;

  SetEvent(FEventProc);

  Result := D_OK;
end;

function TMCCEventThread.ClearCueSheet(AMCC: PMCC; AChannelID: Word): Integer;
begin
  Result := D_FALSE;

  SetCommand(AMCC, mctClear, AChannelID);

  ResetEvent(FEventCompleted);
  SetEvent(FEventCommand);

  if (WaitForSingleObject(FEventCompleted, INFINITE) = WAIT_OBJECT_0) then
  begin
    Result := FCMD_Result;
  end;

  EventQueueCheck;
end;

procedure TMCCEventThread.DoEventCommand;
var
  R: Integer;

  ChannelForm: TfrmChannel;
  I: Integer;
begin
  ResetEvent(FEventCommand);

  FCMD_Result := D_FALSE;

  try
    if (not HasMainControl) then exit;

    with FCommand do
    begin
      if (MCC = nil) then exit;
//      Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand, SEC=nil', [])));

//      Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand, IP=%s, Type=%d', [SEC^.HostIP, Integer(CommandType)])));
      if (MCC^.Alive) then
      begin
        case CommandType of
          mctDeviceError:
          begin
            R := MCCSetDeviceCommError(MCC^.HostIP, DeviceStatus, DeviceName);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCSetDeviceCommError succeeded, mcc = %d, name = %s', [MCC^.ID, DeviceName])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCSetDeviceCommError, error = %d, mcc = %d, name = %s', [R, MCC^.ID, DeviceName])));
          end;
          mctDeviceStatus:
          begin
            R := MCCSetDeviceStatus(MCC^.HostIP, DCSID, DeviceHandle, DeviceStatus);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCSetDeviceStatus succeeded, mcc = %d, handle = %d', [MCC^.ID, DeviceHandle])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCSetDeviceStatus, error = %d, mcc = %d, handle = %d', [R, MCC^.ID, DeviceHandle])));
          end;
          mctEventStatus:
          begin
            R := MCCSetEventStatus(MCC^.HostIP, EventID, EventStatus);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCSetEventStatus succeeded, mcc = %d, id = %s', [MCC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCSetEventStatus, error = %d, mcc = %d, id = %s', [R, MCC^.ID, EventIDToString(EventID)])));
          end;
          mctMediaStatus:
          begin
            R := MCCSetMediaStatus(MCC^.HostIP, EventID, MediaStatus);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCSetMediaStatus succeeded, mcc = %d, id = %s', [MCC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCSetMediaStatus, error = %d, mcc = %d, id = %s', [R, MCC^.ID, EventIDToString(EventID)])));
          end;
          mctOnAir:
          begin
            R := MCCSetOnAir(MCC^.HostIP, ChannelID, IsOnAir);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCSetOnAir succeeded, mcc = %d, channel = %d', [MCC^.ID, ChannelID])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCSetOnAir, error = %d, mcc = %d, channel = %d', [R, MCC^.ID, ChannelID])));
          end;
          mctCurr:
          begin
            R := MCCSetCueSheetCurr(MCC^.HostIP, EventID);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCSetCueSheetCurr succeeded, mcc = %d, id = %s', [MCC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCSetCueSheetCurr, error = %d, mcc = %d, id = %s', [R, MCC^.ID, EventIDToString(EventID)])))
          end;
          mctNext:
          begin
            R := MCCSetCueSheetNext(MCC^.HostIP, EventID);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCSetCueSheetNext succeeded, mcc = %d, id = %s', [MCC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCSetCueSheetNext, error = %d, mcc = %d, id = %s', [R, MCC^.ID, EventIDToString(EventID)])))
          end;
          mctTarget:
          begin
            R := MCCSetCueSheetTarget(MCC^.HostIP, EventID);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCSetCueSheetTarget succeeded, mcc = %d, id = %s', [MCC^.ID, EventIDToString(EventID)])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCSetCueSheetTarget, error = %d, mcc = %d, id = %s', [R, MCC^.ID, EventIDToString(EventID)])))
          end;
          mctClear:
          begin
            ClearEventQueueByChannelID(MCC, ChannelID);
            R := MCCClearCueSheet(MCC^.HostIP, ChannelID);
            if (R = D_OK) then
              Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventCommand MCCClearCueSheet succeeded, mcc = %d, channel = %d', [MCC^.ID, ChannelID])))
            else
              Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventCommand MCCClearCueSheet, error = %d, mcc = %d, channel = %d', [R, MCC^.ID, ChannelID])))
          end;
          else
            R := D_FALSE;
        end;
      end;

      FCMD_Result := R;
    end;
  finally
    ResetCommand;

    SetEvent(FEventCompleted);
  end;
end;

procedure TMCCEventThread.DoEventProc;
var
  E: PMCCEvent;
  R: Integer;
  ChannelForm: TfrmChannel;
begin
//  ResetEvent(FEventProc);

  if (FEventQueue.Count <= 0) then
  begin
    ResetEvent(FEventProc);
    exit;
  end;

  try
  FEventLock.Enter;
  try
    E := FEventQueue.First;

    if (not HasMainControl) then exit;

  finally
    FEventLock.Leave;
  end;

  if (E = nil) then exit;
  if (E^.MCC = nil) then exit;

  if (E^.MCC^.Alive) then
  begin
    case E^.EventType of
      metInput:
      begin
        R := MCCInputCueSheet(E^.MCC^.HostIP, E^.Index, E^.Item);
        if (R = D_OK) then
          Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventProc MCCInputCueSheet succeeded, mcc = %d, id = %s', [E^.MCC^.ID, EventIDToString(E^.Item.EventID)])))
        else
          Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventProc MCCInputCueSheet, error = %d, mcc = %d, id = %s', [R, E^.MCC^.ID, EventIDToString(E^.Item.EventID)])))
      end;
      metDelete:
      begin
        R := MCCDeleteCueSheet(E^.MCC^.HostIP, E^.EventID);
        if (R = D_OK) then
          Assert(False, GetMainLogStr(lsNormal, Format('TMCCEventThread.DoEventProc MCCDeleteCueSheet succeeded, mcc = %d, id = %s', [E^.MCC^.ID, EventIDToString(E^.EventID)])))
        else
          Assert(False, GetMainLogStr(lsError, Format('TMCCEventThread.DoEventProc MCCDeleteCueSheet, error = %d, mcc = %d, id = %s', [R, E^.MCC^.ID, EventIDToString(E^.EventID)])))
      end;
    end;
  end;


  finally
  FEventLock.Enter;
  try
  finally
    FEventQueue.Remove(E);

  Dispose(E);
    FEventLock.Leave;
  end;
  end;
end;

procedure TMCCEventThread.Execute;
var
  R: Cardinal;
  WaitList: array[0..2] of THandle;
begin
  WaitList[0] := FEventClose;
  WaitList[1] := FEventCommand;
  WaitList[2] := FEventProc;

  while not Terminated do
  begin
    R := WaitForMultipleObjects(3, @WaitList, False, INFINITE);
    case R of
      WAIT_OBJECT_0: break;
      WAIT_OBJECT_0 + 1: DoEventCommand;
      WAIT_OBJECT_0 + 2: DoEventProc;
    end;
  end;
end;

{ TDCSDeviceThread }

constructor TDCSDeviceThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FDeviceCheckEvent := CreateEvent(nil, True, False, nil);
  FCloseEvent := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TDCSDeviceThread.Destroy;
begin
  Terminate;

  CloseHandle(FDeviceCheckEvent);
  CloseHandle(FCloseEvent);

  inherited Destroy;
end;

procedure TDCSDeviceThread.Terminate;
begin
  inherited Terminate;

  ResetEvent(FDeviceCheckEvent);
  SetEvent(FCloseEvent);
end;

procedure TDCSDeviceThread.DeviceCheck;
begin
//  FChannelForm.FMediaCheckIntervalTime := 0;

  SetEvent(FDeviceCheckEvent);
end;

procedure TDCSDeviceThread.DoDeviceCheck;
var
  I, J: Integer;
  Source: PSource;
  SourceHandle: PSourceHandle;
  R: Integer;

  CheckTime: Integer;
  DeviceForm: TfrmDevice;

  WarningCommText: String;
  WarningConnectText: String;
  WarningText: String;
begin
  if (not HasMainControl) then exit;

  if (not GV_SettingOption.OnAirCheckDeviceNotify) then exit;

  WarningCommText := '';

  for I := 0 to GV_SourceList.Count - 1 do
  begin
    Source := GV_SourceList[I];

    if (Source <> nil) then
    begin
      if (Source^.CommSuccess) then
      begin
        Inc(Source^.CommTimeout);

        CheckTime := TimecodeToMilliSec(GV_SettingThresholdTime.OnAirCheckDeviceTimeout) div 1000;
        if (Source^.CommTimeout >= CheckTime) then
        begin
          Source^.CommSuccess := False;
          Source^.CommTimeout := CheckTime;
        end;
      end;

      if (not Source^.CommSuccess) then
      begin
        DeviceForm := FSEC.GetDeviceForm;
        if (DeviceForm <> nil) then
        begin
          DeviceForm.SetDeviceCommError(String(Source^.Name), Source^.Status);
          if (WarningCommText = '') then
            WarningCommText := SWDeviceCommError + #13#10 + Format('%s', [String(Source^.Name)])
          else
            WarningCommText := WarningCommText + #13#10 + Format('%s', [String(Source^.Name)]);
        end;
      end
      else
      begin
        if (not Source^.Status.Connected) then
        begin
          if (Source^.Handles <> nil) then
          begin
            for J := 0 to Source^.Handles.Count - 1 do
            begin
              SourceHandle := Source^.Handles[J];
              if (SourceHandle^.Handle > 0) then
                break;
            end;

            if (WarningConnectText = '') then
              WarningConnectText := SWDeviceStatusNotConnect + #13#10 + Format('Device=%s, DCS=%s', [String(Source^.Name), SourceHandle^.DCS^.Name])
            else
              WarningConnectText := WarningConnectText + #13#10 + Format('Device=%s, DCS=%s', [String(Source^.Name), SourceHandle^.DCS^.Name]);
          end;
        end;
      end;
    end;
  end;
  
  if (WarningCommText <> '') or (WarningConnectText <> '') then
  begin
//    if (frmSEC.FWarningDialogDeviceCheck = nil) or (not frmSEC.FWarningDialogDeviceCheck.Showing) then
    begin
      if (WarningCommText = '') then WarningText := WarningConnectText
      else if (WarningConnectText = '') then WarningText := WarningCommText
      else WarningText := WarningCommText + #13#10 + #13#10 + WarningConnectText;

      if (WarningText <> '') then
      begin
        System.Classes.TThread.CreateAnonymousThread(
          procedure
          begin
            PostMessage(FSEC.Handle, WM_WARNING_DISPLAY_DEVICE_CHECK, 0, LParam(PChar(WarningText)));
          end).Start;
      end;
    end;
  end;
end;

procedure TDCSDeviceThread.Execute;
var
  R: Cardinal;
  WaitList: array[0..1] of THandle;
begin
  WaitList[0] := FDeviceCheckEvent;
  WaitList[1] := FCloseEvent;
  while not Terminated do
  begin
    R := WaitForMultipleObjects(2, @WaitList, False, TimecodeToMilliSec(GV_SettingTimeParameter.OnAirCheckDeviceStatusInterval));
    if (R = (WAIT_OBJECT_0 + 1)) then break;

    DoDeviceCheck;

{    case R of
      WAIT_OBJECT_0: DoMediaCheck;
      WAIT_OBJECT_0 + 1: break;
    end;  }
  end;
end;

{ TDCSMediaThread }

constructor TDCSMediaThread.Create(ASEC: TfrmSEC);
begin
  FSEC := ASEC;

  FMediaCheckEvent := CreateEvent(nil, True, False, nil);
  FCloseEvent := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TDCSMediaThread.Destroy;
begin
  Terminate;

  CloseHandle(FMediaCheckEvent);
  CloseHandle(FCloseEvent);

  inherited Destroy;
end;

procedure TDCSMediaThread.Terminate;
begin
  inherited Terminate;

  ResetEvent(FMediaCheckEvent);
  SetEvent(FCloseEvent);
end;

procedure TDCSMediaThread.MediaCheck;
begin
//  FChannelForm.FMediaCheckIntervalTime := 0;

  SetEvent(FMediaCheckEvent);
end;

procedure TDCSMediaThread.DoMediaCheck;
var
  Channel: PChannel;
  ChannelForm: TfrmChannel;

  Index: Integer;


  I, J: Integer;
  Item, ParentItem: PCueSheetItem;

  StartTime: TDateTime;

  Source: PSource;
  SourceHandles: TSourceHandleList;

  R: Integer;
  MediaExist: Boolean;
  MediaDuration: TTimecode;
  EventDuration: TTimecode;
begin
{  with FChannelForm do
  begin
    MediaCheck;
    FMediaCheckIntervalTime := 0;
  end; }

  ResetEvent(FMediaCheckEvent);

  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    Channel :=  GV_ChannelList[I];
    if (Channel = nil) then continue;

//    if (not GetChannelOnAirByID(Channel^.ID)) then continue;

    ChannelForm := FSEC.GetChannelFormByID(Channel^.ID);
    if (ChannelForm = nil) then continue;

    ChannelForm.MediaCheck;
  end;
end;

procedure TDCSMediaThread.Execute;
var
  R: Cardinal;
  WaitList: array[0..1] of THandle;
begin
  WaitList[0] := FMediaCheckEvent;
  WaitList[1] := FCloseEvent;
  while not Terminated do
  begin
    R := WaitForMultipleObjects(2, @WaitList, False, TimecodeToMilliSec(GV_SettingTimeParameter.MediaCheckInterval));
    if (R = (WAIT_OBJECT_0 + 1)) then break;

    DoMediaCheck;

{    case R of
      WAIT_OBJECT_0: DoMediaCheck;
      WAIT_OBJECT_0 + 1: break;
    end;  }
  end;
end;

end.
