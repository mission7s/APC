unit UnitConsts;

interface

uses System.SysUtils, System.Classes, System.Generics.Collections, System.IniFiles,
     Winapi.Windows, Winapi.Messages, System.SyncObjs,
     Vcl.Forms, Vcl.Graphics,
     UnitCommons, UnitDCSDLL;

resourcestring
  // Device
  SLookingForDevice = 'Looking for device.';
  SNotFoundDeviceAndContinue = 'No device found. [Device=%s, DCS=%s]'#13#10'Do you still want to find the rest of your devices?';

  // MCC
  SLookingForMCC = 'Looking for MCC.';
  SNotFoundMCCAndContinue = 'No MCC found. [MCC=%s]'#13#10'Do you still want to find the rest of your MCC or device?';

  // Edit Event
  SEnterEventMode = '''Event mode'' is a required field.';
  SEnterStartMode = '''Start mode'' is a required field.';
  SEnterInput = '''Input'' is a required field.';
  SEnterOutput = '''Output'' is a required field.';
  SEnterTitle = '''Title'' is a required field.';
  SEnterSource = '''Source'' is a required field.';
  SEnterTransitionType = '''Transition type'' is a required field.';
  SEnterTransitionRate = '''Transition rate'' is a required field.';
  SEnterProgramType = '''Program type'' is a required field.';

  SEnterCompletedFileFolder = '캡춰 완료파일 폴더를 입력하십시오.';
  SEnterCompletedFileStoredDays = '캡춰 완료파일 보관 일 수를 입력하십시오.';

  // CheSheet
  SInvalidTimeocde = '''%s'' timecode is invalidate.';
  SStartTimeGreaterThenBeforeEndTime = 'The start time must be greater then the end time of the previous event.';
  SSubStartTimeGreaterThenParentStartTime = 'The sub event start time must be greater then the start time of the parent event.';
  SSubStartTimeLessThenParentEndTime = 'The sub event start time + duration must be less then the end time of the parent event.';
  SDurationTCGreaterThenMinDuration = '''%s'' must be greater then ''%s''.';
  SDurationTCLessThenMaxDuration = '''%s'' must be less then ''%s''.';
  SInTCLessThenDurationTC = 'The in timecode must be less then the duration timecode.';
  SInTCLessThenOutTC = 'The in timecode must be less then the out timecode.';
  SOutTCLessThenDurationTC = 'The out timecode must be less then the duration timecode.';
  SOutTCGreaterThenInTC = 'The out timecode must be grater then the in timecode.';

  SQDeleteEvent = 'Selected events and sub-events will be deleted.'#13#10'Events in the onair or cue state are not deleted.'#13#10'Are you sure you want to delete the events below?';

  SEPasteLocationIncorrect = 'Please specify only one paste location and paste again.';
  SEPasteHasProgramLocationIncorrect = 'There is a program event include the events on the clipboard.'#13#10'If clipboard has a program event, you can paste it in the program or independant main location.';
  SEPasteHasMainLocationIncorrect = 'There is a main event include the events on the clipboard.'#13#10'If clipboard has a main event, you can paste it in the main location.';
  SEPasteJoinSubLocationIncorrect = 'There is a join or sub event include the events on the clipboard.'#13#10'If clipboard not has a main event, you can paste it in the main location.';

  SENoCreateCuesheetWhileChannelOnair = 'You will not be able to create a new cuesheet while this channel is onair.';
  SENoOpenCuesheetWhileChannelOnair = 'You will not be able to open other cuesheet while this channel is onair.';

  // OnAir
  SChannelAlreadyRunning = 'This channel is already running. [%s]';
  SFreezeOnAirTimeout = 'The time remaining until the next event is less than the freeze onair timeout.';

  SQStartOnAirByManual = 'Not enough time to prepare for start onair.'#13#10'Do you want to manually start the first event after the current time?';
  SQFinishtOnAirAndPreserveEvent = 'Do you want to preserve the event queue so that you can resume onair after shutdown?';
  SQFinishtOnAirAndClearEvent = 'Are you sure you want to clear the event queue and finish the current channel completely?';

  SEChangedOnAirEvent = 'The onair event changed. Please retry start onair.';
  SENotPossibleEdit = 'You can not enter event befor onair or cued event locations.';
  SENotInsertProgramLocationIncorrect = 'Program event can only be entered on the program or independent main event.';
  SENotInsertMainLocationIncorrect = 'Main event can only be entered on the program or main or independent comment event.';
  SENotInsertJoinSubLocationIncorrect = 'Join or sub event can only be entered below the main event.';

  // Log string
  LS_DCSOpenDeviceSuccess = 'Device successfully opened. [DCS=%d, Name=%s]';
  LS_DCSCloseDeviceSuccess = 'Device successfully closed. [DCS=%d, Name=%s, Handle=%d]';

  LS_MCCOpenSuccess = 'MCC successfully opened. [MCC=%d, Name=%s]';
  LS_MCCCloseSuccess = 'MCC successfully closed. [MCC=%d, Name=%s]';
  LS_MCCCheckSuccess = 'MCC successfully alive check. [MCC=%d, Name=%s]';

  LSE_DCSInitializeFailed = 'DCS connection initialization failed. [Errorcode=%d, Notify port=%d, In port=%d, Out port=%d]';
  LSE_DCSFinalizeFailed = 'DCS connection finalization failed. [Errorcode=%d]';
  LSE_DCSSetDeviceStatusNotifyFailed = 'Setting dcs device status notify function failed. [Errorcode=%d]';
  LSE_DCSSetEventStatusNotifyFailed = 'Setting dcs event status notify function failed. [Errorcode=%d]';
  LSE_DCSSetEventOverallNotifyFailed = 'Setting dcs event overall notify function failed. [Errorcode=%d]';

  LSE_DCSCloseDeviceFailed = 'Failed to close device. [Errorcode=%d, DCS=%d, Name=%s, Handle=%d]';

  LSE_MCCInitializeFailed = 'MCC connection initialization failed. [Errorcode=%d, Notify port=%d, In port=%d, Out port=%d]';
  LSE_MCCFinalizeFailed = 'MCC connection finalization failed. [Errorcode=%d]';

  LSE_MCCCloseFailed = 'Failed to close MCC. [Errorcode=%d, MCC=%d, Name=%s]';
  LSE_MCCCheckFailed = 'Failed to alive check MCC. [Errorcode=%d, MCC=%d, Name=%s]';

  LSW_DCSOpenDeviceFailed = 'Failed to open device. [Errorcode=%d, DCS=%d, Name=%s]';

  LSW_MCCOpenFailed = 'Failed to open MCC. [Errorcode=%d, MCC=%d, Name=%s]';

const
  WM_UPDATE_CURRENT_TIME  = WM_USER + 1001;
  WM_UPDATE_ACTIVATE      = WM_USER + 1002;

  WM_UPDATE_CHANNEL_TIME  = WM_USER + 2001;
  WM_UPDATE_CURR_EVENT    = WM_USER + 3001;
  WM_UPDATE_NEXT_EVENT    = WM_USER + 3002;
  WM_UPDATE_TARGET_EVENT  = WM_USER + 3003;
  WM_UPDATE_DEVICE_STATUS = WM_USER + 4001;
  WM_UPDATE_EVENT_STATUS  = WM_USER + 4002;

  WM_EXECUTE_AUTO_LOAD_PLAYLIST  = WM_USER + 5001;
  WM_EXECUTE_AUTO_EJECT_PLAYLIST = WM_USER + 5002;

  WM_EXECUTE_MEDIA_CHECK  = WM_USER + 6001;
  WM_UPDATE_MEDIA_CHECK   = WM_USER + 6002;

  WM_EXECUTE_MCC_CHECK    = WM_USER + 7001;
  WM_UPDATE_MCC_CHECK     = WM_USER + 7002;

  WM_INVALID_EDIT         = WM_USER + 9001;

  INIT_TIME = '00:00:00';
  IDLE_TIME = '--:--:--';

  INIT_TIMECODE = '00:00:00:00';
  IDLE_TIMECODE = '--:--:--:--';

  NEW_CUESHEET_NAME = 'New cuesheet';

  // CueSheet List Grid Count
  CNT_CUESHEET_HEADER = 1;
  CNT_CUESHEET_FOOTER = 1;
  CNT_CUESHEET_COLUMNS = 25;

  // CueSheet List XML Node & Attribute Name'
  XML_CHANNEL_ID      = 'channelid';
  XML_ONAIR_DATE      = 'onairdate';
  XML_ONAIR_FLAG      = 'onairflag';
  XML_ONAIR_NO        = 'onairno';

  XML_EVENT_COUNT     = 'eventcount';
  XML_EVENTS          = 'events';
  XML_EVENT           = 'event';

  XML_ATTR_STATUS          = 'status';
  XML_ATTR_SERIAL_NO       = 'serialno';
  XML_ATTR_PROGRAM_NO      = 'programno';
  XML_ATTR_GROUP_NO        = 'groupno';
  XML_ATTR_EVENT_MODE      = 'eventmode';
  XML_ATTR_START_MODE      = 'startmode';
  XML_ATTR_START_TIME      = 'starttime';
  XML_ATTR_INPUT           = 'input';
  XML_ATTR_OUTPUT          = 'output';
  XML_ATTR_TITLE           = 'title';
  XML_ATTR_SUB_TITLE       = 'subtitle';
  XML_ATTR_SOURCE          = 'source';
  XML_ATTR_MEDIA_ID        = 'mediaid';
  XML_ATTR_DURATION_TC     = 'durationtc';
  XML_ATTR_IN_TC           = 'intc';
  XML_ATTR_OUT_TC          = 'outtc';
  XML_ATTR_TRANSITION_TYPE = 'transitiontype';
  XML_ATTR_TRANSITION_RATE = 'transitionrate';
  XML_ATTR_VIDEO_TYPE      = 'videotype';
  XML_ATTR_AUDIO_TYPE      = 'audiotype';
  XML_ATTR_CLOSED_CAPTION  = 'closedcaption';
  XML_ATTR_VOICE_ADD       = 'voiceadd';
  XML_ATTR_PROGRAM_TYPE    = 'programtype';
  XML_ATTR_NOTES           = 'notes';

  // Device List Grid Count
//  CNT_DEVICE_HEADER   = 2;
//  CNT_DEVICE_COLUMNS  = 2;
  CNT_DEVICE_HEADER   = 1;
  CNT_DEVICE_COLUMNS  = 5;//6;

  OnAirFlagNames: array[False..True] of String = ('Off air', 'On air');


type
  TSettingGeneral = packed record
    WorkCueSheetPath: String;
    LoadCueSheetPath: String;
    SaveCueSheetPath: String;
    LogPath: String;
    LogExt: String;
  end;

  TSettingSEC = packed record
    CrossPort: Word;
    NumCrossCheck: Word;
    CrossCheckInterval: TTimecode;
  end;

  TSettingMCC = packed record
    Use: Boolean;
    NotifyPort: Word;
    InPort: Word;
    OutPort: Word;
  end;

  TSettingDCS = packed record
    NotifyPort: Word;
    InPort: Word;
    OutPort: Word;
  end;

  TSettingEventColumnDefault = packed record
    StartMode: TStartMode;
    Input: TInputType;
    OutputBkgnd: TOutputBkgndType;
    OutputKeyer: TOutputKeyerType;
    Source: PSource;
    DurationTC: TTimecode;
    InTC: TTimecode;
    OutTC: TTimecode;
    TransitionType: TTRType;
    TransitionRate: TTRRate;
    ProgramType: Byte;
  end;

  TSettingTimeParameter = packed record
    MCCCheckInterval: TTimecode;

    OnAirCheckDeviceStatusInterval: TTimecode;
    OnAirLoopSleepTime: TTimecode;
    OnAirEnqueueInterval: TTimecode;
    OnAirErrorRecoveryInterval: TTimecode;
    OnAirAutoSaveCuesheetInterval: TTimecode;
    OnAirRefreshGridInterval: TTimecode;
    OnAirChannelDataUpdateInterval: TTimecode;
    MonitorCheckChannelDataInterval: TTimecode;
    MonitorLoopSleepTime: TTimecode;
    AutoIncreaseDurationAmount: TTimecode;
    AutoIncreaseDurationBefore: TTimecode;
    FindReplaceUpdateInterval: TTimecode;
    SourceExchangeUpdateInterval: TTimecode;
    StandardTimeCorrection: TTimecode;
    ForceRefreshEventQueueInterval: TTimecode;
    ForceRefreshEventQueueMinDur: TTimecode;
    RequestUpdateInterval: TTimecode;
    PerformUpdateInterval: TTimecode;
    ProbeManagerInterval: TTimecode;
    NextEventInputBefore: TTimecode;
  end;

  TSettingTresholdTime = packed record
    OnAirLockTime: TTimecode;
    EditLockTime: TTimecode;
    DeviceControlLockTime: TTimecode;
    CueAllLockTime: TTimecode;
    SetNextLockTime: TTimecode;
    HoldLockTime: TTimecode;
    EnqueueLockTime: TTimecode;
    MinDuration: TTimecode;
    MaxDuration: TTimecode;
    MinTransitionInterval: TTimecode;
    OnAirEventTransitionThreshold: TTimecode;
  end;

  TSettingOption = packed record
    AutoLoadCuesheet: Boolean;
    AutoLoadCuesheetInterval: TTimecode;
    AutoEjectCuesheet: Boolean;
    AutoEjectCuesheetInterval: TTimecode;
    MaxInputEventCount: Word;
    MediaCheckInterval: TTimecode;
    MediaCheckTime: TTimecode;
    OnAirEventHighlight: Boolean;
    OnAirEventFixedRow: Word;
  end;

  TProgramType = packed record
    Code: Byte;
    Name: array[0..PROGRAMTYPENAME_LEN] of Char;
    Color: TColor;
  end;
  PProgramType = ^TProgramType;
  TProgramTypeList = TList<PProgramType>;

{  TChannelCueSheet = packed record
    ChannelId: Word;
    OnairDate: array[0..DATE_LEN] of Char;
    OnairFlag: TOnAirFlagType;
    OnairNo: Integer;
    CueSheetList: TCueSheetList;
  end;
  PChannelCueSheet = ^TChannelCueSheet;
  TChannelCueSheetList = TList<PChannelCueSheet>; }

  TChannelCueSheet = packed record
    FileName: array[0..MAX_PATH] of Char;
    ChannelID: Word;
    OnairDate: array[0..DATE_LEN] of Char;
    OnairFlag: TOnAirFlagType;
    OnairNo: Integer;
    EventCount: Integer;
    LastSerialNo: Integer;
    LastProgramNo: Integer;
    LastGroupNo: Integer;
  end;
  PChannelCueSheet = ^TChannelCueSheet;
  TChannelCueSheetList = TList<PChannelCueSheet>;

  TPasteMode = (
    pmCut,
    pmCopy
    );

  TPasteIncluded = set of TEventMode;

  TClipboardCueSheet = class(TObject)
  private
    FPasteMode: TPasteMode;
    FPasteIncluded: TPasteIncluded;

    FChannelID: Word;
    FCueSheetList: TCueSheetList;

    function GetCount: Integer;
    function GetItem(Index: Integer): PCueSheetItem;

    procedure SetPasteMode(AValue: TPasteMode);
    procedure SetPasteIncluded(AValue: TPasteIncluded);
    procedure SetChannelID(AChannelID: Word);
    procedure SetItem(Index: Integer; Value: PCueSheetItem);
  public
    constructor Create;
    destructor Destroy; override;

    function Add(AItem: PCueSheetItem): Integer;
    function IndexOf(AItem: PCueSheetItem): Integer;

    function First: PCueSheetItem;
    function Last: PCueSheetItem;

    procedure Sort;
    procedure Clear;

    property PasteMode: TPasteMode read FPasteMode write SetPasteMode;
    property PasteIncluded: TPasteIncluded read FPasteIncluded write SetPasteIncluded;
    property ChannelID: Word read FChannelID write SetChannelID;
    property CueSheetList: TCueSheetList read FCueSheetList;
    property Count: Integer read GetCount;

    property Items[Index: Integer]: PCueSheetItem read GetItem write SetItem; Default;
  end;

  TTimelineZoomType = (ztNone,
                       zt1Second, zt2Seconds, zt5Seconds, zt10Seconds, zt15Seconds, zt30Seconds,
                       zt1Minute, zt2Minutes, zt5Minutes, zt10Minutes, zt15Minutes, zt30Minutes,
                       zt1Hour, zt2Hours, zt6Hours, zt12Hours,
                       zt1Day,
                       ztFit);

var
  // CueSheet Column Index
  IDX_COL_CUESHEET_GROUP: Word          = 0;
  IDX_COL_CUESHEET_NO: Word             = 1;
  IDX_COL_CUESHEET_EVENT_MODE: Word     = 2;
  IDX_COL_CUESHEET_EVENT_STATUS: Word   = 3;
  IDX_COL_CUESHEET_START_MODE: Word     = 4;
  IDX_COL_CUESHEET_START_DATE: Word     = 5;
  IDX_COL_CUESHEET_START_TIME: Word     = 6;
  IDX_COL_CUESHEET_INPUT: Word          = 7;
  IDX_COL_CUESHEET_OUTPUT: Word         = 8;
  IDX_COL_CUESHEET_TITLE: Word          = 9;
  IDX_COL_CUESHEET_SUB_TITLE: Word      = 10;
  IDX_COL_CUESHEET_SOURCE: Word         = 11;
  IDX_COL_CUESHEET_MEDIA_ID: Word       = 12;
  IDX_COL_CUESHEET_MEDIA_STATUS: Word   = 13;
  IDX_COL_CUESHEET_DURATON: Word        = 14;
  IDX_COL_CUESHEET_IN_TC: Word          = 15;
  IDX_COL_CUESHEET_OUT_TC: Word         = 16;
  IDX_COL_CUESHEET_TR_TYPE: Word        = 17;
  IDX_COL_CUESHEET_TR_RATE: Word        = 18;
  IDX_COL_CUESHEET_VIDEO_TYPE: Word     = 19;
  IDX_COL_CUESHEET_AUDIO_TYPE: Word     = 20;
  IDX_COL_CUESHEET_CLOSED_CAPTION: Word = 21;
  IDX_COL_CUESHEET_VOICE_ADD: Word      = 22;
  IDX_COL_CUESHEET_PROGRAM_TYPE: Word   = 23;
  IDX_COL_CUESHEET_NOTES: Word          = 24;

  // CueSheet Column Name
  NAM_COL_CUESHEET_PROGRAM: String        = '';//'Program';
  NAM_COL_CUESHEET_GROUP: String          = '';//'Group';
  NAM_COL_CUESHEET_NO: String             = 'No';
  NAM_COL_CUESHEET_EVENT_MODE: String     = 'EM';//'Event Mode';
  NAM_COL_CUESHEET_EVENT_STATUS: String   = 'Status';//'Event Status';
  NAM_COL_CUESHEET_START_MODE: String     = 'Start Mode';
  NAM_COL_CUESHEET_START_DATE: String     = 'Start Date';
  NAM_COL_CUESHEET_START_TIME: String     = 'Start Time';
  NAM_COL_CUESHEET_INPUT: String          = 'Input';
  NAM_COL_CUESHEET_OUTPUT: String         = 'Output';
  NAM_COL_CUESHEET_TITLE: String          = 'Title';
  NAM_COL_CUESHEET_SUB_TITLE: String      = 'Sub Title';
  NAM_COL_CUESHEET_SOURCE: String         = 'Source';
  NAM_COL_CUESHEET_MEDIA_ID: String       = 'Media ID';
  NAM_COL_CUESHEET_MEDIA_STATUS: String   = 'Media Status';
  NAM_COL_CUESHEET_DURATON: String        = 'Duration';
  NAM_COL_CUESHEET_IN_TC: String          = 'In TC';
  NAM_COL_CUESHEET_OUT_TC: String         = 'Out TC';
  NAM_COL_CUESHEET_TR_TYPE: String        = 'TT';//'Transition Type';
  NAM_COL_CUESHEET_TR_RATE: String        = 'TR';//'Transition Rate';
  NAM_COL_CUESHEET_VIDEO_TYPE: String     = 'VT';//'Video Type';
  NAM_COL_CUESHEET_AUDIO_TYPE: String     = 'AT';//'Audio Type';
  NAM_COL_CUESHEET_CLOSED_CAPTION: String = 'CC';//'Closed Caption';
  NAM_COL_CUESHEET_VOICE_ADD: String      = 'VA';//'Voice Add';
  NAM_COL_CUESHEET_PROGRAM_TYPE: String   = 'Program';//'Program Type';
  NAM_COL_CUESHEET_NOTES: String          = 'Notes';

  // CueSheet Column Width
  WIDTH_COL_CUESHEET_GROUP: Integer           = 40;
  WIDTH_COL_CUESHEET_NO: Integer              = 30;
  WIDTH_COL_CUESHEET_EVENT_MODE: Integer      = 20;
  WIDTH_COL_CUESHEET_EVENT_STATUS: Integer    = 60;
  WIDTH_COL_CUESHEET_START_MODE: Integer      = 76;
  WIDTH_COL_CUESHEET_START_DATE: Integer      = 76;
  WIDTH_COL_CUESHEET_START_TIME: Integer      = 76;
  WIDTH_COL_CUESHEET_INPUT: Integer           = 60;
  WIDTH_COL_CUESHEET_OUTPUT: Integer          = 60;
  WIDTH_COL_CUESHEET_TITLE: Integer           = 200;
  WIDTH_COL_CUESHEET_SUB_TITLE: Integer       = 200;
  WIDTH_COL_CUESHEET_SOURCE: Integer          = 60;
  WIDTH_COL_CUESHEET_MEDIA_ID: Integer        = 80;
  WIDTH_COL_CUESHEET_MEDIA_STATUS: Integer    = 90;
  WIDTH_COL_CUESHEET_DURATON: Integer         = 76;
  WIDTH_COL_CUESHEET_IN_TC: Integer           = 76;
  WIDTH_COL_CUESHEET_OUT_TC: Integer          = 76;
  WIDTH_COL_CUESHEET_TR_TYPE: Integer         = 40;
  WIDTH_COL_CUESHEET_TR_RATE: Integer         = 40;
  WIDTH_COL_CUESHEET_VIDEO_TYPE: Integer      = 40;
  WIDTH_COL_CUESHEET_AUDIO_TYPE: Integer      = 40;
  WIDTH_COL_CUESHEET_CLOSED_CAPTION: Integer  = 40;
  WIDTH_COL_CUESHEET_VOICE_ADD: Integer       = 40;
  WIDTH_COL_CUESHEET_PROGRAM_TYPE: Integer    = 60;
  WIDTH_COL_CUESHEET_NOTES: Integer           = 160;

  // CueSheet Column Visible
  VIS_COL_CUESHEET_NO: Boolean              = True;
  VIS_COL_CUESHEET_EVENT_MODE: Boolean      = True;
  VIS_COL_CUESHEET_EVENT_STATUS: Boolean    = True;
  VIS_COL_CUESHEET_START_MODE: Boolean      = True;
  VIS_COL_CUESHEET_START_DATE: Boolean      = True;
  VIS_COL_CUESHEET_START_TIME: Boolean      = True;
  VIS_COL_CUESHEET_TITLE: Boolean           = True;
  VIS_COL_CUESHEET_SUB_TITLE: Boolean       = True;
  VIS_COL_CUESHEET_SOURCE: Boolean          = True;
  VIS_COL_CUESHEET_MEDIA_ID: Boolean        = True;
  VIS_COL_CUESHEET_MEDIA_STATUS: Boolean    = True;
  VIS_COL_CUESHEET_DURATON: Boolean         = True;
  VIS_COL_CUESHEET_IN_TC: Boolean           = True;
  VIS_COL_CUESHEET_OUT_TC: Boolean          = True;
  VIS_COL_CUESHEET_TR_TYPE: Boolean         = True;
  VIS_COL_CUESHEET_TR_RATE: Boolean         = True;
  VIS_COL_CUESHEET_VIDEO_TYPE: Boolean      = True;
  VIS_COL_CUESHEET_AUDIO_TYPE: Boolean      = True;
  VIS_COL_CUESHEET_CLOSED_CAPTION: Boolean  = True;
  VIS_COL_CUESHEET_VOICE_ADD: Boolean       = True;
  VIS_COL_CUESHEET_PROGRAM_TYPE: Boolean    = True;
  VIS_COL_CUESHEET_NOTES: Boolean           = True;

  // Device Column Index
  IDX_COL_DEVICE_NO: Word                 = 0;
  IDX_COL_DEVICE_NAME: Word               = 1;
  IDX_COL_DEVICE_STATUS: Word             = 2;
  IDX_COL_DEVICE_TIMECODE: Word           = 3;
  IDX_COL_DEVICE_DCS: Word                = 4;
//  IDX_COL_DEVICE_CHANNEL: Word            = 5;

  // Device Column Name
  NAM_COL_DEVICE_NO: String               = 'No';
  NAM_COL_DEVICE_NAME: String             = 'Name';
  NAM_COL_DEVICE_STATUS: String           = 'Status';
  NAM_COL_DEVICE_TIMECODE: String         = 'Timecode';
  NAM_COL_DEVICE_DCS: String              = 'DCS';
//  NAM_COL_DEVICE_CHANNEL: String          = 'Channel';

  // Device Column Width
  WIDTH_COL_DEVICE_NO: Integer            = 30;
  WIDTH_COL_DEVICE_NAME: Integer          = 60;
  WIDTH_COL_DEVICE_STATUS: Integer        = 60;
  WIDTH_COL_DEVICE_TIMECODE: Integer      = 76;
  WIDTH_COL_DEVICE_DCS: Integer           = 60;
//  WIDTH_COL_DEVICE_CHANNEL: Integer       = 76;

//  TIMELINE_SIDE_FRAMES: Integer = 1800;

  HIDE_SUB_EVENT: Boolean = True;

  FORMAT_DATE: String;

  // Event status color
  COLOR_BK_EVENTSTATUS_NORMAL: TColor = $003A2C28;
  COLOR_BK_EVENTSTATUS_NEXT: TColor = clSkyBlue;
  COLOR_BK_EVENTSTATUS_CUED: TColor = clYellow;
  COLOR_BK_EVENTSTATUS_ONAIR: TColor = clMoneyGreen;
  COLOR_BK_EVENTSTATUS_DONE: TColor = clDkGray;
  COLOR_BK_EVENTSTATUS_ERROR: TColor = clRed;
  COLOR_BK_EVENTSTATUS_TARGET: TColor = clAqua;

  COLOR_TX_EVENTSTATUS_NORMAL: TColor = clSilver;
  COLOR_TX_EVENTSTATUS_NEXT: TColor = clBlack;
  COLOR_TX_EVENTSTATUS_CUED: TColor = clBlack;
  COLOR_TX_EVENTSTATUS_ONAIR: TColor = clBlack;
  COLOR_TX_EVENTSTATUS_DONE: TColor = clSilver;
  COLOR_TX_EVENTSTATUS_ERROR: TColor = clWhite;
  COLOR_TX_EVENTSTATUS_TARGET: TColor = clBlack;

  // Cuesheet cut & copy color
  COLOR_BK_CLIPBOARD: TColor = $00312420;

  COLOR_TX_CLIPBOARD: TColor = clGrayText;

  // Device Status Color
  COLOR_BK_DEVICESTATUS_NORMAL: TColor = $003A2C28;
  COLOR_BK_DEVICESTATUS_ERROR: TColor = clRed;

  COLOR_TX_DEVICESTATUS_NORMAL: TColor = clSilver;
  COLOR_TX_DEVICESTATUS_ERROR: TColor = clWhite;

  // Media Status Cell Color
  COLOR_BK_MEDIASTATUS_NORMAL: TColor = $003A2C28;
  COLOR_BK_MEDIASTATUS_NOT_EXIST: TColor = clRed;

  COLOR_TX_MEDIASTATUS_NORMAL: TColor = clSilver;
  COLOR_TX_MEDIASTATUS_NOT_EXIST: TColor = clWhite;

  GV_SettingGeneral: TSettingGeneral;
  GV_SettingSEC: TSettingSEC;
  GV_SettingMCC: TSettingMCC;
  GV_SettingDCS: TSettingDCS;
  GV_SettingEventColumnDefault: TSettingEventColumnDefault;
  GV_SettingTimeParameter: TSettingTimeParameter;
  GV_SettingTresholdTime: TSettingTresholdTime;
  GV_SettingOption: TSettingOption;

  GV_LogCS: TCriticalSection;

  GV_ChannelList: TChannelList;

  GV_SECList: TSECList;
  GV_SECMain: PSEC;
  GV_SECMine: PSEC;

  GV_MCCList: TMCCList;
  GV_DCSList: TDCSList;
  GV_SourceList: TSourceList;
  GV_MCSList: TMCSList;

  GV_ProgramTypeList: TProgramTypeList;

  GV_NullEventID: TEventID;

  GV_ClipboardCueSheet: TClipboardCueSheet;

  GV_MCCCheckExecuteEvent: THandle;

  GV_TimerExecuteEvent: THandle;
  GV_TimerCancelEvent: THandle;

  GV_ChannelMaxCount: Word = 4;
  GV_ChannelMinCount: Word = 1;
  GV_ChannelTimelineMinHeight: Word = 216;

  procedure LoadConfig;
  procedure SaveConfig;

  procedure AssertProc(const AMessage, AFileName: String; ALineNumber: Integer; AErrorAddr: Pointer);

  function GetMainLogStr(AState: TLogState; ALogStr: String): String; overload;
  function GetMainLogStr(AState: TLogState; APRec: PResStringRec): String; overload;
  function GetMainLogStr(AState: TLogState; APRec: PResStringRec; const Args: array of const): String; overload;

  function GetChannelLogStr(AState: TLogState; const AChannelID: Word; const ALogStr: String): String;

  function WriteAssertLog(const ALogStr: String): Integer;

  function CheckSum(AValue: AnsiString): Boolean;

  function GetSECByID(AID: Word): PSEC;

  function GetMCCByID(AID: Word): PMCC;

  function GetChannelByName(AName: String): PChannel;
  function GetChannelByID(AID: Byte): PChannel;
  function GetChannelNameByID(AID: Byte): String;

  function GetChannelOnAirByID(AID: Integer): Boolean;
  procedure SetChannelOnAirByID(AID: Integer; AOnAir: Boolean);

  function GetDCSByName(AName: String): PDCS;
  function GetDCSByIP(AIP: String): PDCS;
  function GetDCSNameByID(AID: Word): String;

  function GetSourceByName(AName: String): PSource;

  function GetSourceTypeByName(AName: String): TSourceType;

  function GetSourceXptByName(ARouter: TXptList; AName: String): Integer;

  function GetProgramTypeNameByCode(ACode: Byte): String;
  function GetProgramTypeColorByCode(ACode: Byte): TColor;

  procedure ClearChannelList;
  procedure ClearSECList;
  procedure ClearMCCList;
  procedure ClearDCSList;
  procedure ClearSourceList;
  procedure ClearMCSList;
  procedure ClearProgramTypeList;

implementation

{ TClipboardCueSheet }

constructor TClipboardCueSheet.Create;
begin
  FPasteMode := pmCut;
  FPasteIncluded := [];

  FChannelId := 0;

  FCueSheetList := TCueSheetList.Create;
end;

destructor TClipboardCueSheet.Destroy;
begin
  inherited Destroy;

  if (FCueSheetList <> nil) then
  begin
    FCueSheetList.Clear;
    FreeAndNil(FCueSheetList);
  end;
end;

function TClipboardCueSheet.GetCount: Integer;
begin
  Result := FCueSheetList.Count;
end;

function TClipboardCueSheet.GetItem(Index: Integer): PCueSheetItem;
begin
  Result := FCueSheetList[Index];
end;

procedure TClipboardCueSheet.SetPasteMode(AValue: TPasteMode);
begin
  FPasteMode := AValue;
end;

procedure TClipboardCueSheet.SetPasteIncluded(AValue: TPasteIncluded);
begin
  FPasteIncluded := AValue;
end;

procedure TClipboardCueSheet.SetChannelID(AChannelID: Word);
begin
  FChannelId := AChannelID;
end;

procedure TClipboardCueSheet.SetItem(Index: Integer; Value: PCueSheetItem);
begin
  Items[Index] := Value;
end;

function TClipboardCueSheet.Add(AItem: PCueSheetItem): Integer;
begin
  Result := FCueSheetList.Add(AItem);
end;

function TClipboardCueSheet.IndexOf(AItem: PCueSheetItem): Integer;
begin
  Result := FCueSheetList.IndexOf(AItem);
end;

function TClipboardCueSheet.First: PCueSheetItem;
begin
  Result := FCueSheetList.First;
end;

function TClipboardCueSheet.Last: PCueSheetItem;
begin
  Result := FCueSheetList.Last;
end;

procedure TClipboardCueSheet.Sort;
begin
  FCueSheetList.Sort;
end;

procedure TClipboardCueSheet.Clear;
begin
  FCueSheetList.Clear;
end;

{ Clobal functions }

procedure LoadConfig;
var
  IniFile: TIniFile;
  NumChannel, NumDCS, NumSource, NumXpt: Word;
  NumSEC, NumMCC, NumType: Word;
  DataStrings: TStrings;
  I, J, K: Integer;

  Section, Ident: String;
  Channel: PChannel;
  SEC: PSEC;
  MCC: PMCC;
  DCS: PDCS;
  Source: PSource;
  SourceHandle: PSourceHandle;
  Xpt: PXpt;
  MCS: PMCS;

  ProgramType: PProgramType;
begin
  // Channel
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'Channel.ini');
  try
    // Num Channel
    NumChannel := IniFile.ReadInteger('Channel', 'NumChannel', 0);

    for I := 0 to NumChannel - 1 do
    begin
      Section := Format('Channel%d', [I + 1]);

      Channel := New(PChannel);
      FillChar(Channel^, SizeOf(TChannel), #0);

      Channel^.ID := IniFile.ReadInteger(Section, 'ID', 0);
      StrPCopy(Channel^.Name, IniFile.ReadString(Section, 'Name', ''));

      GV_ChannelList.Add(Channel);
    end;
  finally
    FreeAndNil(IniFile);
  end;

  // Source
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'Source.ini');
  try
    // Num DCS
    NumDCS := IniFile.ReadInteger('DCS', 'NumDCS', 0);
    for I := 0 to NumDCS - 1 do
    begin
      Ident := Format('DCS%d', [I + 1]);

      DCS := New(PDCS);
      FillChar(DCS^, SizeOf(TDCS), #0);

      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('DCS', Ident, '0,DCS1,127.0.0.1,0,0')), DataStrings);
        DCS^.ID   := StrToIntDef(DataStrings[0], 0);
        StrPCopy(DCS^.Name, DataStrings[1]);
        StrPCopy(DCS^.HostIP, DataStrings[2]);
        DCS^.Main := StrToBool(DataStrings[3]);
        DCS^.Mine := StrToBool(DataStrings[4]);
      finally
        FreeAndNil(DataStrings);
      end;

      GV_DCSList.Add(DCS);
    end;

    // Num Source
    NumSource := IniFile.ReadInteger('Source', 'NumSource', 0);

    for I := 0 to NumSource - 1 do
    begin
//      for J := 0 to GV_DCSList.Count - 1 do
//      begin
        Section := Format('Source%d', [I + 1]);

        Source := New(PSource);
        FillChar(Source^, SizeOf(TSource), #0);

        StrPCopy(Source^.Name, IniFile.ReadString(Section, 'Name', ''));
//        Source^.Handle := -1;

        Source^.Handles := TSourceHandleList.Create;
        for J := 0 to GV_DCSList.Count - 1 do
        begin
          SourceHandle := New(PSourceHandle);
          SourceHandle^.DCSID  := GV_DCSList[J]^.ID;
          StrCopy(SourceHandle^.DCSIP, GV_DCSList[J]^.HostIP);
          SourceHandle^.Handle := -1;

          Source^.Handles.Add(SourceHandle);
        end;

        Source^.Channel := GetChannelByID(IniFile.ReadInteger(Section, 'Channel', -1));
//        Source^.DCS := GV_DCSList[J];//GetDCSByName(IniFile.ReadString(Section, 'DCS', ''));
        Source^.SourceType := GetSourceTypeByName(IniFile.ReadString(Section, 'Type', ''));
        Source^.MakeTransition := StrToBool(IniFile.ReadString(Section, 'MakeTransition', 'False'));
        Source^.SuccessiveUse := StrToBool(IniFile.ReadString(Section, 'SuccessiveUse', 'False'));
        Source^.LoadAlarm := StrToBool(IniFile.ReadString(Section, 'LoadAlarm', 'False'));
        Source^.EjectAlarm := StrToBool(IniFile.ReadString(Section, 'EjectAlarm', 'False'));
        Source^.QueryDuration := StrToBool(IniFile.ReadString(Section, 'QueryDuration', 'False'));
        Source^.PrerollTime := StringToTimecode(IniFile.ReadString(Section, 'PrerollTime', '00:00:00:00'));
        Source^.DefaultTimecode := StringToTimecode(IniFile.ReadString(Section, 'DefaultTimecode', '00:00:00:00'));
        Source^.UseDefaultTimecode := StrToBool(IniFile.ReadString(Section, 'UseDefaultTimecode', 'False'));

        FillChar(Source^.Status, SizeOf(TDeviceStatus), #0);
        case Source^.SourceType of
          ST_VSDEC,
          ST_VSENC,
          ST_VCR,
          ST_CART,
          ST_CG,
          ST_LINE,
          ST_LOGO: Source^.Status.EventType := ET_PLAYER;
          ST_ROUTER:  Source^.Status.EventType := ET_RSW;
          ST_MCS: Source^.Status.EventType := ET_SWITCHER;
          ST_GPI: Source^.Status.EventType := ET_GPI;
        end;

        if (Source^.SourceType in [ST_ROUTER, ST_MCS]) then
        begin
          Source^.Router := TXptList.Create;

          NumXpt := IniFile.ReadInteger(Section, 'NumXpt', 0);
          for K := 0 to NumXPT - 1 do
          begin
            Ident := Format('Xpt%d', [K + 1]);

            Xpt := New(PXpt);
            FillChar(Xpt^, SizeOf(TXPT), #0);

            DataStrings := TStringList.Create;
            try
              ExtractStrings([','], [' '], PChar(IniFile.ReadString(Section, Ident, '0,0')), DataStrings);
              StrPCopy(Xpt^.DeviceName, DataStrings[0]);
              Xpt^.XptIn := StrToInt(DataStrings[1]);
            finally
              FreeAndNil(DataStrings);
            end;

            Source^.Router.Add(Xpt);
          end;
        end;

        GV_SourceList.Add(Source);
//      end;
    end;
  finally
    FreeAndNil(IniFile);
  end;

  // SEC
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'SEC.ini');
  try
    Section := 'General';

    FillChar(GV_SettingGeneral, SizeOf(TSettingGeneral), #0);
    with GV_SettingGeneral do
    begin
//      WorkCueSheetPath := ExtractFilePath(Application.ExeName) + IniFile.ReadString(Section, 'WorkCueSheetPath', 'CueSheet\Work') + PathDelim;
//      LoadCueSheetPath := ExtractFilePath(Application.ExeName) + IniFile.ReadString(Section, 'LoadCueSheetPath', 'CueSheet\Load') + PathDelim;
//      SaveCueSheetPath := ExtractFilePath(Application.ExeName) + IniFile.ReadString(Section, 'SaveCueSheetPath', 'CueSheet\Save') + PathDelim;
//      LogPath          := ExtractFilePath(Application.ExeName) + IniFile.ReadString(Section, 'LogPath', 'Log') + PathDelim;

      WorkCueSheetPath := IniFile.ReadString(Section, 'WorkCueSheetPath', ExtractFilePath(Application.ExeName) + 'CueSheet\Work') + PathDelim;
      LoadCueSheetPath := IniFile.ReadString(Section, 'LoadCueSheetPath', ExtractFilePath(Application.ExeName) + 'CueSheet\Load') + PathDelim;
      SaveCueSheetPath := IniFile.ReadString(Section, 'SaveCueSheetPath', ExtractFilePath(Application.ExeName) + 'CueSheet\Save') + PathDelim;
      LogPath          := IniFile.ReadString(Section, 'LogPath', ExtractFilePath(Application.ExeName) + 'Log') + PathDelim;
      LogExt           := IniFile.ReadString(Section, 'LogExt', 'SEC.log');

      if (not DirectoryExists(LogPath)) then ForceDirectories(LogPath);
      if (not DirectoryExists(WorkCueSheetPath)) then ForceDirectories(WorkCueSheetPath);
      if (not DirectoryExists(LoadCueSheetPath)) then ForceDirectories(LoadCueSheetPath);
      if (not DirectoryExists(SaveCueSheetPath)) then ForceDirectories(SaveCueSheetPath);
    end;

    // SEC
    // SEC Port
    Section := 'SEC';

    FillChar(GV_SettingSEC, SizeOf(TSettingSEC), #0);
    with GV_SettingSEC do
    begin
      CrossPort           := IniFile.ReadInteger(Section, 'CrossPort', 9003);
      NumCrossCheck       := IniFile.ReadInteger(Section, 'NumCrossCheck', 3);
      CrossCheckInterval  := StringToTimecode(IniFile.ReadString(Section, 'CrossCheckInterval', '00:00:01:00'));
    end;

    // Num SEC
    NumSEC := IniFile.ReadInteger(Section, 'NumSEC', 0);

    for I := 0 to NumSEC - 1 do
    begin
      SEC := New(PSEC);
      FillChar(SEC^, SizeOf(TSEC), #0);

      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('SEC', Format('SEC%d', [I + 1]), '-')), DataStrings);
        SEC^.ID := StrToIntDef(DataStrings[0], 0);
        StrPCopy(SEC^.Name, DataStrings[1]);
        StrPCopy(SEC^.HostIP, DataStrings[2]);
        SEC^.Main := StrToBool(DataStrings[3]);
        SEC^.Mine := StrToBool(DataStrings[4]);
        SEC^.Alive := False;

        if (SEC^.Main) then
          GV_SECMain := SEC;

        if (SEC^.Mine) then
          GV_SECMine := SEC;
      finally
        FreeAndNil(DataStrings);
      end;

      GV_SECList.Add(SEC);
    end;

    // MCC
    // MCC Port
    Section := 'MCC';

    FillChar(GV_SettingMCC, SizeOf(TSettingMCC), #0);
    with GV_SettingMCC do
    begin
      NotifyPort  := IniFile.ReadInteger(Section, 'NotifyPort', 7000);
      InPort      := IniFile.ReadInteger(Section, 'InPort', 7002);
      OutPort     := IniFile.ReadInteger(Section, 'OutPort', 7001);
      Use         := StrToBool(IniFile.ReadString(Section, 'Use', 'False'));
    end;

    // Num MCC
    NumMCC := IniFile.ReadInteger(Section, 'NumMCC', 0);

    for I := 0 to NumMCC - 1 do
    begin
      MCC := New(PMCC);
      FillChar(MCC^, SizeOf(TMCC), #0);

      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('MCC', Format('MCC%d', [I + 1]), '-')), DataStrings);
        MCC^.ID := StrToIntDef(DataStrings[0], 0);
        StrPCopy(MCC^.Name, DataStrings[1]);
        StrPCopy(MCC^.HostIP, DataStrings[2]);
        MCC^.Main := StrToBool(DataStrings[3]);
        MCC^.Mine := StrToBool(DataStrings[4]);
        MCC^.Opened := False;
      finally
        FreeAndNil(DataStrings);
      end;

      GV_MCCList.Add(MCC);
    end;

    // DCS Port
    Section := 'DCS';

    FillChar(GV_SettingDCS, SizeOf(TSettingDCS), #0);
    with GV_SettingDCS do
    begin
      NotifyPort  := IniFile.ReadInteger(Section, 'NotifyPort', 8000);
      InPort      := IniFile.ReadInteger(Section, 'InPort', 8002);
      OutPort     := IniFile.ReadInteger(Section, 'OutPort', 8001);
    end;

    // MCS
    // Num MCS
    NumType := IniFile.ReadInteger('MCS', 'NumMCS', 0);

    for I := 0 to NumType - 1 do
    begin
      MCS := New(PMCS);
      FillChar(MCS^, SizeOf(TMCS), #0);

      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('MCS', Format('MCS%d', [I + 1]), '-')), DataStrings);
        MCS^.ChannelID := StrToIntDef(DataStrings[0], 0);
        StrPCopy(MCS^.Name, DataStrings[1]);
      finally
        FreeAndNil(DataStrings);
      end;

      GV_MCSList.Add(MCS);
    end;

    // Program Type
    // Num Type
    NumType := IniFile.ReadInteger('ProgramType', 'NumType', 0);
    for I := 0 to NumType - 1 do
    begin
      Ident := Format('Type%d', [I + 1]);

      ProgramType := New(PProgramType);
      FillChar(ProgramType^, SizeOf(TProgramType), #0);

      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('ProgramType', Ident, '0,MAIN,0,128,255')), DataStrings);
        ProgramType^.Code := StrToIntDef(DataStrings[0], 0);
        StrPCopy(ProgramType^.Name, DataStrings[1]);
        ProgramType^.Color := RGB(StrToIntDef(DataStrings[2], 0), StrToIntDef(DataStrings[3], 0), StrToIntDef(DataStrings[4], 0));
      finally
        FreeAndNil(DataStrings);
      end;

      GV_ProgramTypeList.Add(ProgramType);
    end;

    // Event Column Default value
    Section := 'EventColumnDefault';

    FillChar(GV_SettingEventColumnDefault, SizeOf(TSettingEventColumnDefault), #0);
    with GV_SettingEventColumnDefault do
    begin
      StartMode      := TStartMode(IniFile.ReadInteger(Section, 'StartMode', 0));
      Input          := TInputType(IniFile.ReadInteger(Section, 'Input', 0));
      OutputBkgnd    := TOutputBkgndType(IniFile.ReadInteger(Section, 'OutputBkgnd', 0));
      OutputKeyer    := TOutputKeyerType(IniFile.ReadInteger(Section, 'OutputKeyer', 0));
      Source         := GetSourceByName(IniFile.ReadString(Section, 'Source', ''));
      DurationTC     := TTimecode(IniFile.ReadInteger(Section, 'DurationTC', 0));
      InTC           := TTimecode(IniFile.ReadInteger(Section, 'InTC', 0));
      OutTC          := TTimecode(IniFile.ReadInteger(Section, 'OutTC', 0));
      TransitionType := TTRType(IniFile.ReadInteger(Section, 'TransitionType', 0));
      TransitionRate := TTRRate(IniFile.ReadInteger(Section, 'TransitionRate', 0));
      ProgramType    := IniFile.ReadInteger(Section, 'ProgramType', 0);
    end;

    // Time parameter
    Section := 'TimeParameter';

    FillChar(GV_SettingTimeParameter, SizeOf(TSettingTimeParameter), #0);
    with GV_SettingTimeParameter do
    begin
      MCCCheckInterval                := StringToTimecode(IniFile.ReadString(Section, 'MCCCheckInterval', '00:00:10:00'));

      OnAirCheckDeviceStatusInterval  := StringToTimecode(IniFile.ReadString(Section, 'OnAirCheckDeviceStatusInterval', '00:00:01:00'));
      OnAirCheckDeviceStatusInterval  := StringToTimecode(IniFile.ReadString(Section, 'OnAirCheckDeviceStatusInterval', '00:00:00:04'));
      OnAirEnqueueInterval            := StringToTimecode(IniFile.ReadString(Section, 'OnAirEnqueueInterval', '00:00:05:00'));
      OnAirErrorRecoveryInterval      := StringToTimecode(IniFile.ReadString(Section, 'OnAirErrorRecoveryInterval', '00:00:10:00'));
      OnAirRefreshGridInterval        := StringToTimecode(IniFile.ReadString(Section, 'OnAirRefreshGridInterval', '00:00:01:00'));
      OnAirChannelDataUpdateInterval  := StringToTimecode(IniFile.ReadString(Section, 'OnAirChannelDataUpdateInterval', '00:00:01:00'));
      MonitorCheckChannelDataInterval := StringToTimecode(IniFile.ReadString(Section, 'MonitorCheckChannelDataInterval', '00:00:01:00'));
      MonitorLoopSleepTime            := StringToTimecode(IniFile.ReadString(Section, 'MonitorLoopSleepTime', '00:00:00:04'));
      AutoIncreaseDurationAmount      := StringToTimecode(IniFile.ReadString(Section, 'AutoIncreaseDurationAmount', '00:02:00:00'));
      AutoIncreaseDurationBefore      := StringToTimecode(IniFile.ReadString(Section, 'AutoIncreaseDurationBefore', '00:00:20:00'));
      FindReplaceUpdateInterval       := StringToTimecode(IniFile.ReadString(Section, 'FindReplaceUpdateInterval', '00:00:10:00'));
      SourceExchangeUpdateInterval    := StringToTimecode(IniFile.ReadString(Section, 'SourceExchangeUpdateInterval', '00:00:10:00'));
      StandardTimeCorrection          := StringToTimecode(IniFile.ReadString(Section, 'StandardTimeCorrection', '00:00:00:00'));
      ForceRefreshEventQueueInterval  := StringToTimecode(IniFile.ReadString(Section, 'ForceRefreshEventQueueInterval', '00:30:00:00'));
      ForceRefreshEventQueueMinDur    := StringToTimecode(IniFile.ReadString(Section, 'ForceRefreshEventQueueMinDur', '00:01:00:00'));
      RequestUpdateInterval           := StringToTimecode(IniFile.ReadString(Section, 'RequestUpdateInterval', '00:00:10:00'));
      PerformUpdateInterval           := StringToTimecode(IniFile.ReadString(Section, 'PerformUpdateInterval', '00:00:10:00'));
      ProbeManagerInterval            := StringToTimecode(IniFile.ReadString(Section, 'ProbeManagerInterval', '00:00:10:00'));
      NextEventInputBefore            := StringToTimecode(IniFile.ReadString(Section, 'NextEventInputBefore', '00:00:20:00'));
    end;

    // Treshhold time
    Section := 'TresholdTime';

    FillChar(GV_SettingTresholdTime, SizeOf(TSettingTresholdTime), #0);
    with GV_SettingTresholdTime do
    begin
      OnAirLockTime                 := StringToTimecode(IniFile.ReadString(Section, 'OnAirLockTime', '00:00:30:00'));
      EditLockTime                  := StringToTimecode(IniFile.ReadString(Section, 'EditLockTime', '00:00:15:00'));
      DeviceControlLockTime         := StringToTimecode(IniFile.ReadString(Section, 'DeviceControlLockTime', '00:01:00:00'));
      CueAllLockTime                := StringToTimecode(IniFile.ReadString(Section, 'CueAllLockTime', '00:00:30:00'));
      SetNextLockTime               := StringToTimecode(IniFile.ReadString(Section, 'SetNextLockTime', '00:00:10:00'));
      HoldLockTime                  := StringToTimecode(IniFile.ReadString(Section, 'HoldLockTime', '00:00:10:00'));
      EnqueueLockTime               := StringToTimecode(IniFile.ReadString(Section, 'EnqueueLockTime', '00:00:05:00'));
      MinDuration                   := StringToTimecode(IniFile.ReadString(Section, 'MinDuration', '00:00:05:00'));
      MaxDuration                   := StringToTimecode(IniFile.ReadString(Section, 'MaxDuration', '23:59:59:29'));
      MinTransitionInterval         := StringToTimecode(IniFile.ReadString(Section, 'MinTransitionInterval', '00:00:05:00'));
      OnAirEventTransitionThreshold := StringToTimecode(IniFile.ReadString(Section, 'OnAirEventTransitionThreshold', '00:00:03:00'));
    end;

    // Option
    Section := 'Option';

    FillChar(GV_SettingOption, SizeOf(TSettingOption), #0);
    with GV_SettingOption do
    begin
      AutoLoadCuesheet              := StrToBool(IniFile.ReadString(Section, 'AutoLoadCuesheet', 'False'));
      AutoLoadCuesheetInterval      := StringToTimecode(IniFile.ReadString(Section, 'AutoLoadCuesheetInterval', '01:00:00:00'));
      AutoEjectCuesheet             := StrToBool(IniFile.ReadString(Section, 'AutoEjectCuesheet', 'False'));
      AutoEjectCuesheetInterval     := StringToTimecode(IniFile.ReadString(Section, 'AutoEjectCuesheetInterval', '01:00:00:00'));
      MaxInputEventCount            := IniFile.ReadInteger(Section, 'MaxInputEventCount', 50);
      MediaCheckInterval            := StringToTimecode(IniFile.ReadString(Section, 'MediaCheckInterval', '00:20:00:00'));
      MediaCheckTime                := StringToTimecode(IniFile.ReadString(Section, 'MediaCheckTime', '02:00:00:00'));
      OnAirEventHighlight           := StrToBool(IniFile.ReadString(Section, 'OnAirEventHighlight', 'True'));
      OnAirEventFixedRow            := IniFile.ReadInteger(Section, 'OnAirEventFixedRow', 0);
    end;
  finally
    FreeAndNil(IniFile);
  end;

  FORMAT_DATE := 'YYYY' + FormatSettings.DateSeparator + 'MM' + FormatSettings.DateSeparator + 'DD';
end;

procedure SaveConfig;
begin
end;

procedure AssertProc(const AMessage, AFileName: String; ALineNumber: Integer; AErrorAddr: Pointer);
var
  LogFilePath: String;
  LogFileName: String;
  LogStr: String;
//  DateStr, LogStr: String;

  FS: TFileStream;
  PBuffer: Pointer;
begin
  OutputDebugString(PChar(Format('%s(%dLine $%x) %s', [ExtractFileName(AFileName),
                                                       ALineNumber,
                                                       Integer(AErrorAddr),
                                                       AMessage])));

  LogStr := Format('%s %s(%dLine $%x)'#13#10, [AMessage, ExtractFileName(AFileName), ALineNumber, Integer(AErrorAddr)]);

  WriteAssertLog(LogStr);
end;

function GetMainLogStr(AState: TLogState; ALogStr: String): String;
var
  DateStr: String;
begin
  DateStr := FormatDateTime('YYYY-MM-DD tt', Now);
  case AState of
    lsNormal: Result := Format('[%s] %s', [DateStr, ALogStr]);
    lsError: Result := Format('[%s] [Error] %s', [DateStr, ALogStr]);
    lsWarning: Result := Format('[%s] [Warning] %s', [DateStr, ALogStr]);
    else
      Result := Format('[%s] %s', [DateStr, ALogStr]);
  end;
end;

function GetMainLogStr(AState: TLogState; APRec: PResStringRec): String;
begin
  Result := GetMainLogStr(AState, LoadResString(APRec));
end;

function GetMainLogStr(AState: TLogState; APRec: PResStringRec; const Args: array of const): String;
begin
  Result := GetMainLogStr(AState, Format(LoadResString(APRec), Args));
end;

function GetChannelLogStr(AState: TLogState; const AChannelID: Word; const ALogStr: String): String;
var
  DateStr: String;
begin
  DateStr := FormatDateTime('YYYY-MM-DD tt', Now);
  case AState of
    lsNormal : Result := Format('%d[%s] %s', [AChannelID, DateStr, ALogStr]);
    lsError: Result := Format('%d[%s] [Error] %s', [AChannelID, DateStr, ALogStr]);
    lsWarning: Result := Format('%d[%s] [Warning] %s', [AChannelID, DateStr, ALogStr]);
    else Result := Format('%d[%s] %s', [AChannelID, DateStr, ALogStr]);
  end;
end;

function WriteAssertLog(const ALogStr: String): Integer;
var
  DivPos: Integer;
  LogDivStr: String;
  FilePath: String;
  FileName: String;
  DateStr, LogStr: String;

  FS: TFileStream;
  PBuffer: Pointer;
  byteorder_marker: Word;
begin
  GV_LogCS.Enter;
  try
    FilePath := Format('%s%s%s', [GV_SettingGeneral.LogPath,
                                  FormatDateTime('YYYY-MM', Date),
                                  PathDelim]);

    DivPos := Pos('[', ALogStr);
    if (DivPos > 0) then
    begin
      LogDivStr := Trim(Copy(ALogStr, 1, DivPos - 1));
      LogStr    := Copy(ALogStr, DivPos, Length(ALogStr));
    end;

    if (LogDivStr = '') then
    begin
      LogDivStr := GV_SettingGeneral.LogExt;
      LogStr    := ALogStr;
    end
    else
      LogDivStr := Format('%s_%s', [LogDivStr, GV_SettingGeneral.LogExt]);

    FileName := Format('%s%s_%s', [FilePath,
                                   FormatDateTime('YYYY-MM-DD', Date),
                                   LogDivStr]);

    // Log directory
    if (not DirectoryExists(FilePath)) then
      if (not ForceDirectories(FilePath)) then
      begin
        MessageBeep(MB_ICONERROR);
        MessageBox(Application.Handle, PChar(Format('Cannot create log directory %s', [FilePath])), PChar(Application.Title), MB_OK or MB_ICONERROR);
        exit;
      end;

    try
      if (not FileExists(FileName)) then
      begin
        FS := TFileStream.Create(FileName, fmCreate or fmShareDenyNone);
        FS.Size := 0;
        FS.Position := 0;

        byteorder_marker := $FEFF;
        FS.Write(byteorder_marker, sizeof(byteorder_marker));
      end
      else FS := TFileStream.Create(FileName, fmOpenWrite or fmShareDenyNone);

      try
        FS.Position := FS.Size;
        FS.Write(PWideChar(LogStr)^, SizeOf(WideChar) * Length(LogStr));
//        FS.Write(PAnsiChar(AnsiString(LogStr))^, Length(AnsiString(LogStr)));
//        FS.Write(PChar(LogStr)^, Length(LogStr));
      finally
        if (FS <> nil) then FreeAndNil(FS);
      end;
    except
      MessageBeep(MB_ICONERROR);
      MessageBox(Application.Handle, PChar(Format('Cannot create or open log file %s', [FileName])), PChar(Application.Title), MB_OK or MB_ICONERROR);
      exit;
    end;
  finally
    GV_LogCS.Leave;
  end;
end;

function CheckSum(AValue: AnsiString): Boolean;
var
  I, Len: Integer;
  CRC: Byte;
begin
  Result := False;

  Len := Length(AValue);
  CRC := 0;
  for I := 3 to Len - 1 do
    CRC := CRC + Ord(AValue[I]);
  CRC := not (CRC) + 1;

  Result := (CRC = Ord(AValue[Len]));
end;

function GetSECByID(AID: Word): PSEC;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to GV_SECList.Count - 1 do
  begin
    if (GV_SECList[I]^.ID = AID) then
    begin
      Result := GV_SECList[I];
      break;
    end;
  end;
end;

function GetMCCByID(AID: Word): PMCC;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to GV_MCCList.Count - 1 do
  begin
    if (GV_MCCList[I]^.ID = AID) then
    begin
      Result := GV_MCCList[I];
      break;
    end;
  end;
end;

function GetChannelByName(AName: String): PChannel;
var
  I: Integer;
begin
  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    if String(GV_ChannelList[I]^.Name) = AName then
    begin
      Result := GV_ChannelList[I];
      break;
    end;
  end;
end;

function GetChannelByID(AID: Byte): PChannel;
var
  I: Integer;
begin
  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    if (GV_ChannelList[I]^.ID = AID) then
    begin
      Result := GV_ChannelList[I];
      break;
    end;
  end;
end;

function GetChannelNameByID(AID: Byte): String;
var
  I: Integer;
begin
  Result := '';
  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    if (GV_ChannelList[I]^.ID = AID) then
    begin
      Result := String(GV_ChannelList[I]^.Name);
      break;
    end;
  end;
end;

function GetChannelOnAirByID(AID: Integer): Boolean;
var
  I: Integer;
  C: PChannel;
begin
  Result := False;
  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    C := GV_ChannelList[I];
    if C^.ID = AID then
    begin
      Result := C^.OnAir;
      break;
    end;
  end;
end;

procedure SetChannelOnAirByID(AID: Integer; AOnAir: Boolean);
var
  I: Integer;
  C: PChannel;
begin
  for I := 0 to GV_ChannelList.Count - 1 do
  begin
    C := GV_ChannelList[I];
    if C^.ID = AID then
    begin
      C^.OnAir := AOnAir;
      break;
    end;
  end;
end;

function GetDCSByName(AName: String): PDCS;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to GV_DCSList.Count - 1 do
  begin
    if (String(GV_DCSList[I]^.Name) = AName) then
    begin
      Result := GV_DCSList[I];
      break;
    end;
  end;
end;

function GetDCSByIP(AIP: String): PDCS;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to GV_DCSList.Count - 1 do
  begin
    if (String(GV_DCSList[I]^.HostIP) = AIP) then
    begin
      Result := GV_DCSList[I];
      break;
    end;
  end;
end;

function GetDCSNameByID(AID: Word): String;
var
  I: Integer;
begin
  Result := '';
  for I := 0 to GV_DCSList.Count - 1 do
  begin
    if (GV_DCSList[I]^.ID = AID) then
    begin
      Result := String(GV_DCSList[I]^.Name);
      break;
    end;
  end;
end;

function GetSourceByName(AName: String): PSource;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to GV_SourceList.Count - 1 do
  begin
    if (String(GV_SourceList[I]^.Name) = AName) then
    begin
      Result := GV_SourceList[I];
      break;
    end;
  end;
end;

function GetSourceTypeByName(AName: String): TSourceType;
var
  I: TSourceType;
begin
  Result := ST_VSDEC;

  for I := ST_VSDEC to ST_LOGO do
    if (UpperCase(SourceTypeNames[I]) = UpperCase(AName)) then
    begin
      Result := I;
      break;
    end;
end;

function GetSourceXptByName(ARouter: TXptList; AName: String): Integer;
var
  I: Integer;
begin
  Result := -1;
  if (ARouter = nil) then exit;

  for I := 0 to ARouter.Count - 1 do
    if (UpperCase(ARouter[I]^.DeviceName) = UpperCase(AName)) then
    begin
      Result := ARouter[I]^.XptIn;
      break;
    end;
end;

function GetProgramTypeNameByCode(ACode: Byte): String;
var
  I: Integer;
begin
  // Program Type Name
  Result := '';

  for I := 0 to GV_ProgramTypeList.Count - 1 do
  begin
    if (GV_ProgramTypeList[I] <> nil) and
       (GV_ProgramTypeList[I]^.Code = ACode) then
    begin
      Result := String(GV_ProgramTypeList[I]^.Name);
      break;
    end;
  end;
end;

function GetProgramTypeColorByCode(ACode: Byte): TColor;
var
  I: Integer;
begin
  // Program Type Color
  Result := clNone;

  for I := 0 to GV_ProgramTypeList.Count - 1 do
  begin
    if (GV_ProgramTypeList[I] <> nil) and
       (GV_ProgramTypeList[I]^.Code = ACode) then
    begin
      Result := GV_ProgramTypeList[I]^.Color;
      break;
    end;
  end;
end;

procedure ClearChannelList;
var
  I: Integer;
begin
  // DCS
  for I := GV_ChannelList.Count - 1 downto 0 do
  begin
    Dispose(GV_ChannelList[I]);
  end;

  GV_ChannelList.Clear;
end;

procedure ClearSECList;
var
  I: Integer;
begin
  // SEC
  for I := GV_SECList.Count - 1 downto 0 do
  begin
    Dispose(GV_SECList[I]);
  end;

  GV_SECList.Clear;
end;

procedure ClearMCCList;
var
  I: Integer;
begin
  // MCC
  for I := GV_MCCList.Count - 1 downto 0 do
  begin
    Dispose(GV_MCCList[I]);
  end;

  GV_MCCList.Clear;
end;

procedure ClearDCSList;
var
  I: Integer;
begin
  // DCS
  for I := GV_DCSList.Count - 1 downto 0 do
  begin
    Dispose(GV_DCSList[I]);
  end;

  GV_DCSList.Clear;
end;

procedure ClearSourceList;
var
  I, J: Integer;
begin
  // Source
  for I := GV_SourceList.Count - 1 downto 0 do
  begin
    if (GV_SourceList[I]^.Handles <> nil) then
    begin
      for J := GV_SourceList[I]^.Handles.Count - 1 downto 0 do
        Dispose(GV_SourceList[I]^.Handles[J]);

      GV_SourceList[I]^.Handles.Clear;
      GV_SourceList[I]^.Handles.Free;
    end;

    case GV_SourceList[I]^.SourceType of
      ST_ROUTER:
        if (GV_SourceList[I]^.Router <> nil) then
        begin
          for J := GV_SourceList[I]^.Router.Count - 1 downto 0 do
            Dispose(GV_SourceList[I]^.Router[J]);
        end;
    end;
    Dispose(GV_SourceList[I]);
  end;

  GV_SourceList.Clear;
end;

procedure ClearMCSList;
var
  I, J: Integer;
begin
  // Source
  for I := GV_MCSList.Count - 1 downto 0 do
  begin
    Dispose(GV_MCSList[I]);
  end;

  GV_MCSList.Clear;
end;

procedure ClearProgramTypeList;
var
  I: Integer;
begin
  // Program Type
  for I := GV_ProgramTypeList.Count - 1 downto 0 do
  begin
    Dispose(GV_ProgramTypeList[I]);
  end;

  GV_ProgramTypeList.Clear;
end;

initialization
  AssertErrorProc := AssertProc;

end.
