unit UnitConsts;

interface

uses System.SysUtils, System.Classes, System.Generics.Collections, System.IniFiles,
     Winapi.Windows, Winapi.Messages, System.SyncObjs,
     Vcl.Forms, Vcl.Graphics, MediaInfoDLL,
     UnitCommons, UnitDCSDLL;

resourcestring
  // SEC
  SSECStart = 'SEC Starting...';
  SLookingForSEC = 'Looking for another SEC.';
  SLookingForAllComponentAndDevice = 'Would you like to find all the rest of the component or device?';
  SNotFoundComponentAndContinue = 'Not found component. [Name=%s]'#13#10'Do you still want to find the rest of your component or devices?';
  SEInvalidSettinSEC = 'Invalidate SEC setting information.';
  SEInvalidSettinProgramType = 'Invalidate programe type setting information.';

  // MCC
  SLookingForMCC = 'Looking for MCC.';
  SEInvalidSettinMCC = 'Invalidate MCC setting information.';

  // DCS
  SLookingForDCS = 'Looking for dcs.';
  SEInvalidSettinDCS = 'Invalidate dcs setting information.';

  // Device
  SLookingForDevice = 'Looking for device...';
  SLookingForAllDevice = 'Would you like to find all the rest of the devices?';
  SNotFoundDeviceAndContinue = 'Not found device. [Device=%s, DCS=%s]'#13#10'Do you still want to find the rest of your devices?';
  SEInvalidSettinXpt = 'Invalidate crosspoint setting information.';
  SEInvalidSettinMCS = 'Invalidate MCS setting information.';

  // Edit Event
  SEnterEventMode = '''Event mode'' is a required field.';
  SEnterStartMode = '''Start mode'' is a required field.';
  SEnterInput = '''Input'' is a required field.';
  SEnterOutput = '''Output'' is a required field.';
  SEnterTitle = '''Title'' is a required field.';
  SEnterSource = '''Source'' is a required field.';
  SEnterSourceLayer = '''Source layer'' is a required field.';
  SEnterTransitionType = '''Transition type'' is a required field.';
  SEnterTransitionRate = '''Transition rate'' is a required field.';
  SEnterFinishAction = '''Finish action'' is a required field.';
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

  SESourceExchangeLocationIncorrect = 'The source cannot be exchange before the current or next event.';

  SENoCreateCuesheetWhileChannelOnair = 'You will not be able to create a new cuesheet while this channel is onair.';
  SENoOpenCuesheetWhileChannelOnair = 'You will not be able to open other cuesheet while this channel is onair.';

  // OnAir
  SChannelAlreadyRunning = 'This channel is already running. [%s]';
  SFreezeOnAirTimeout = 'The time remaining until the next event is less than the freeze onair timeout.';
  SSetNextTimeout = 'The time remaining until the current event is less than the set next timeout.';
  SEnqueueTimeout = 'The time remaining until the next event is less than the enqueue timeout.';

  SQStartOnAirByManual = 'Not enough time to prepare for start onair.'#13#10'Do you want to manually start the first event after the current time?';
  SQFinishtOnAirAndPreserveEvent = 'Do you want to preserve the event queue so that you can resume onair after shutdown?';
  SQFinishtOnAirAndClearEvent = 'Are you sure you want to clear the event queue and finish the current channel completely?';
  SQSourceExchange = '%s'#13#10#13#10'Do you want to exchange the source of main event and backup event for all events from the currently selected event?';

  SENextEventNotReady = 'The next event is not ready. Please try again later.';
  SEChangedOnAirEvent = 'The onair event changed. Please retry start onair.';
  SENotPossibleEditLockTime = 'Editing is not possible because the current time is less than the edit lock time.';
  SENotPossibleEditBeforeLocationIncorrect = 'You can not enter event befor onair or cued event locations.';
  SENotInsertProgramLocationIncorrect = 'Program event can only be entered on the program or independent main event.';
  SENotInsertMainLocationIncorrect = 'Main event can only be entered on the program or main or independent comment event.';
  SENotInsertJoinSubLocationIncorrect = 'Join or sub event can only be entered below the main event.';

  // Warning
  SWDeviceCommError = 'The device cannot communicate with the DCS.';
  SWDeviceStatusNotConnect = 'Device status is not connected. ';
  SWNotExistNextEvent = 'There is no next event to play. [Channel=%s]';

  // Log string
  LS_DCSAliveCheckSuccess = 'DCS successfully alive check. [DCS=%d, Name=%s, Alive=%s]';
  LS_DCSMainCheckSuccess = 'DCS successfully main check. [DCS=%d, Name=%s, Main=%s]';

  LS_DCSOpenDeviceSuccess = 'Device successfully opened. [DCS=%d, Name=%s]';
  LS_DCSCloseDeviceSuccess = 'Device successfully closed. [DCS=%d, Name=%s, Handle=%d]';
  LS_DCSSetControlBySuccess = 'Device set controlby succeeded. [DCS=%d, Name=%s, Handle=%d]';

  LS_MCCOpenSuccess = 'MCC successfully opened. [MCC=%d, Name=%s]';
  LS_MCCCloseSuccess = 'MCC successfully closed. [MCC=%d, Name=%s]';
  LS_MCCAliveCheckSuccess = 'MCC successfully alive check. [MCC=%d, Name=%s]';

  LS_SECOpenSuccess = 'SEC successfully opened. [SEC=%d, Name=%s]';
  LS_SECCloseSuccess = 'SEC successfully closed. SEC=%d, Name=%s]';
  LS_SECAliveCheckSuccess = 'SEC successfully alive check. [SEC=%d, Name=%s, Alive=%s]';
  LS_SECMainCheckSuccess = 'SEC successfully main SEC alive check. [SEC=%d, Name=%s, Main=%s]';
  LS_SECFindMainSuccess = 'SEC successfully find main. [Main SEC=%d, Name=%s]';
  LS_SECNotFoundMain = 'Not found main SEC.';
  LS_SECMainChange = 'SEC main change. [Main SEC=%d, Name=%s]';

  LSE_DCSSysInitializeFailed = 'DCS system initialization failed. [Errorcode=%d, System in port=%d, System out port=%d]';
  LSE_DCSSysFinalizeFailed = 'DCS system finalization failed. [Errorcode=%d]';

  LSE_DCSSysLogInPortEnableFailed = 'Setting dcs system in port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_DCSSysLogOutPortEnableFailed = 'Setting dcs system out port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_DCSSysLogInPortDisableFailed = 'Setting dcs system in port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_DCSSysLogOutPortDisableFailed = 'Setting dcs system out port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_DCSInitializeFailed = 'DCS connection initialization failed. [Errorcode=%d, Notify port=%d, In port=%d, Out port=%d]';
  LSE_DCSFinalizeFailed = 'DCS connection finalization failed. [Errorcode=%d]';

  LSE_DCSLogNotifyPortEnableFailed = 'Setting dcs notify port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_DCSLogInPortEnableFailed = 'Setting dcs in port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_DCSLogOutPortEnableFailed = 'Setting dcs out port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_DCSLogNotifyPortDisableFailed = 'Setting dcs notify port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_DCSLogInPortDisableFailed = 'Setting dcs in port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_DCSLogOutPortDisableFailed = 'Setting dcs out port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_DCSSetDeviceStatusNotifyFailed = 'Setting dcs device status notify function failed. [Errorcode=%d]';
  LSE_DCSSetEventStatusNotifyFailed = 'Setting dcs event status notify function failed. [Errorcode=%d]';
  LSE_DCSSetEventOverallNotifyFailed = 'Setting dcs event overall notify function failed. [Errorcode=%d]';

  LSE_DCSAliveCheckFailed = 'Failed to alive check dcs. [Errorcode=%d, DCS=%d, Name=%s]';
  LSE_DCSMainCheckFailed = 'Failed to main check dcs. [Errorcode=%d, DCS=%d, Name=%s]';

  LSE_DCSOpenDeviceFailed = 'Failed to open device. [Errorcode=%d, DCS=%d, Name=%s]';
  LSE_DCSCloseDeviceFailed = 'Failed to close device. [Errorcode=%d, DCS=%d, Name=%s, Handle=%d]';
  LSE_DCSSetControlByFailed = 'Failed to set device controlby. [Errorcode=%d, DCS=%d, Name=%s, Handle=%d]';

  LSE_MCCSysInitializeFailed = 'MCC system initialization failed. [Errorcode=%d, System in port=%d]';
  LSE_MCCSysFinalizeFailed = 'MCC system finalization failed. [Errorcode=%d]';

  LSE_MCCSysLogInPortEnableFailed = 'Setting mcc system in port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_MCCSysLogOutPortEnableFailed = 'Setting mcc system out port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_MCCInitializeFailed = 'MCC connection initialization failed. [Errorcode=%d, Notify port=%d, In port=%d, Out port=%d]';
  LSE_MCCFinalizeFailed = 'MCC connection finalization failed. [Errorcode=%d]';

  LSE_MCCLogNotifyPortEnableFailed = 'Setting mcc notify port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_MCCLogInPortEnableFailed = 'Setting mcc in port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_MCCLogOutPortEnableFailed = 'Setting mcc in port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_MCCLogNotifyPortDisableFailed = 'Setting mcc notify port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_MCCLogInPortDisableFailed = 'Setting mcc in port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_MCCLogOutPortDisableFailed = 'Setting mcc in port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_MCCOpenFailed = 'Failed to open MCC. [Errorcode=%d, MCC=%d, Name=%s]';
  LSE_MCCCloseFailed = 'Failed to close MCC. [Errorcode=%d, MCC=%d, Name=%s]';
  LSE_MCCAliveCheckFailed = 'Failed to alive check MCC. [Errorcode=%d, MCC=%d, Name=%s]';

  LSE_SECSysInitializeFailed = 'SEC system initialization failed. [Errorcode=%d, System in port=%d]';
  LSE_SECSysFinalizeFailed = 'SEC system finalization failed. [Errorcode=%d]';

  LSE_SECSysLogInPortEnableFailed = 'Setting sec system in port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_SECSysLogOutPortEnableFailed = 'Setting sec system out port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_SECSysSetServerReadProcFailed = 'Setting sec server system read procedure failed. [Errorcode=%d]';

  LSE_SECInitializeFailed = 'SEC connection initialization failed. [Errorcode=%d, In port=%d]';
  LSE_SECFinalizeFailed = 'SEC connection finalization failed. [Errorcode=%d]';

  LSE_SECLogNotifyPortEnableFailed = 'Setting sec notify port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_SECLogInPortEnableFailed = 'Setting sec in port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_SECLogOutPortEnableFailed = 'Setting sec in port log enable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_SECLogNotifyPortDisableFailed = 'Setting sec notify port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_SECLogInPortDisableFailed = 'Setting sec in port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';
  LSE_SECLogOutPortDisableFailed = 'Setting sec in port log disable failed. [Errorcode=%d, LogPath=%s, LogExt=%s]';

  LSE_SECSetServerReadProcFailed = 'Setting sec server read procedure failed. [Errorcode=%d]';

  LSE_SECOpenFailed = 'Failed to open SEC. [Errorcode=%d, SEC=%d, Name=%s]';
  LSE_SECCloseFailed = 'Failed to close SEC. [Errorcode=%d, SEC=%d, Name=%s]';
  LSE_SECAliveCheckFailed = 'Failed to alive check SEC. [Errorcode=%d, SEC=%d, Name=%s]';
  LSE_SECMainCheckFailed = 'Failed to main check SEC. [Errorcode=%d, SEC=%d, Name=%s]';

const
  WM_UPDATE_CURRENT_TIME  = WM_USER + $1001;
  WM_UPDATE_ACTIVATE      = WM_USER + $1002;

  WM_UPDATE_CHANNEL_TIME  = WM_USER + $2001;
  WM_UPDATE_CHANNEL_TIMELINE  = WM_USER + $2002;
  WM_UPDATE_CHANNEL_ONAIR = WM_USER + $2101;

  WM_UPDATE_CURR_EVENT    = WM_USER + $3001;
  WM_UPDATE_NEXT_EVENT    = WM_USER + $3002;
  WM_UPDATE_TARGET_EVENT  = WM_USER + $3003;

  WM_UPDATE_DEVICE_COMM_ERROR = WM_USER + $4001;
  WM_UPDATE_DEVICE_STATUS     = WM_USER + $4002;

  WM_UPDATE_EVENT_STATUS  = WM_USER + $4101;
  WM_UPDATE_ERROR_DISPLAY = WM_USER + $4102;
  WM_UPDATE_WARNING_DISPLAY  = WM_USER + $4103;

  WM_EXECUTE_AUTO_LOAD_PLAYLIST  = WM_USER + $5001;
  WM_EXECUTE_AUTO_EJECT_PLAYLIST = WM_USER + $5002;

  WM_EXECUTE_SOURCE_EXCHANGE = WM_USER + $5003;

  WM_UPDATE_MEDIA_CHECK   = WM_USER + $6001;

  WM_BEGIN_UPDATE         = WM_USER + $7001;
  WM_END_UPDATE           = WM_USER + $7002;
  WM_SET_ONAIR            = WM_USER + $7003;
  WM_SET_TIMELINE_RANGE   = WM_USER + $7004;

  WM_INSERT_CUESHEET      = WM_USER + $7011;
  WM_UPDATE_CUESHEET      = WM_USER + $7012;
  WM_DELETE_CUESHEET      = WM_USER + $7013;
  WM_CLEAR_CUESHEET       = WM_USER + $7014;

  WM_INVALID_EDIT         = WM_USER + $8001;

  WM_WARNING_DISPLAY_DEVICE_CHECK     = WM_USER + $9001;
  WM_WARNING_DISPLAY_NOT_EXIST_EVENT  = WM_USER + $9001;

  INIT_TIME = '00:00:00';
  IDLE_TIME = '--:--:--';

  INIT_TIMECODE = '00:00:00:00';
  IDLE_TIMECODE = '--:--:--:--';

  NEW_CUESHEET_NAME = 'New cuesheet';

  // CueSheet List Grid Count
  CNT_CUESHEET_HEADER = 1;
  CNT_CUESHEET_FOOTER = 1;
  CNT_CUESHEET_COLUMNS = 26;

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
  XML_ATTR_SOURCE_LAYER    = 'layer';
  XML_ATTR_MEDIA_ID        = 'mediaid';
  XML_ATTR_DURATION_TC     = 'durationtc';
  XML_ATTR_IN_TC           = 'intc';
  XML_ATTR_OUT_TC          = 'outtc';
  XML_ATTR_TRANSITION_TYPE = 'transitiontype';
  XML_ATTR_TRANSITION_RATE = 'transitionrate';
  XML_ATTR_FINISH_ACTION   = 'finishaction';
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
    ID: Word;
    Name: String;
    HostIP: String;
    WorkCueSheetPath: String;
    LoadCueSheetPath: String;
    SaveCueSheetPath: String;
    MediaPath: String;
    LogPath: String;
    LogExt: String;
  end;

  TSettingSEC = packed record
    SysInPort: Word;
    SysCheckTimeout: TTimecode;
    SysCheckInterval: TTimecode;
    SysLogInPortEnabled: Boolean;
    SysLogPath: String;
    SysLogExt: String;

    NotifyPort: Word;
    InPort: Word;
    OutPort: Word;
    CrossPort: Word;
    CommandTimeout: TTimecode;
    NumCrossCheck: Word;
    CrossCheckInterval: TTimecode;
    LogNotifyPortEnabled: Boolean;
    LogInPortEnabled: Boolean;
    LogOutPortEnabled: Boolean;
    LogCrossPortEnabled: Boolean;
    LogPath: String;
    LogExt: String;
  end;

  TSettingMCC = packed record
    Use: Boolean;
    SysInPort: Word;
    SysOutPort: Word;
    SysCheckTimeout: TTimecode;
    SysCheckInterval: TTimecode;
    SysLogInPortEnabled: Boolean;
    SysLogOutPortEnabled: Boolean;
    SysLogPath: String;
    SysLogExt: String;

    NotifyPort: Word;
    InPort: Word;
    OutPort: Word;
    CommandTimeout: TTimecode;
    LogNotifyPortEnabled: Boolean;
    LogInPortEnabled: Boolean;
    LogOutPortEnabled: Boolean;
    LogPath: String;
    LogExt: String;
  end;

  TSettingDCS = packed record
    SysInPort: Word;
    SysOutPort: Word;
    SysCheckTimeout: TTimecode;
    SysCheckInterval: TTimecode;
    SysLogInPortEnabled: Boolean;
    SysLogOutPortEnabled: Boolean;
    SysLogPath: String;
    SysLogExt: String;

    NotifyPort: Word;
    InPort: Word;
    OutPort: Word;
    CommandTimeout: TTimecode;
    LogNotifyPortEnabled: Boolean;
    LogInPortEnabled: Boolean;
    LogOutPortEnabled: Boolean;
    LogPath: String;
    LogExt: String;
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
    FinishAction: TFinishAction;
    ProgramType: Byte;
  end;

  TSettingTimeParameter = packed record
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

    MediaCheckInterval: TTimecode;
    MediaCheckTime: TTimecode;

    BlankCheckInterval: TTimecode;
    BlankCheckTime: TTimecode;
  end;

  TSettingThresholdTime = packed record
    OnAirCheckDeviceTimeout: TTimecode;

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

    OnAirEventBreakingDuration: TTimecode;
  end;

  TSettingOption = packed record
    TimelineSpace: Integer;
    TimelineSpaceInterval: Integer;
    TimelineOnairIconFileName: String;
    TimelineNextIconFileName: String;
    TimelineNormalIconFileName: String;

    ChannelTimelineHeight: Integer;

    AutoLoadCuesheet: Boolean;
    AutoLoadCuesheetInterval: TTimecode;
    AutoEjectCuesheet: Boolean;
    AutoEjectCuesheetInterval: TTimecode;
    MaxInputEventCount: Word;
    OnAirEventHighlight: Boolean;
    OnAirEventFixedRow: Word;

    OnAirCheckDeviceNotify: Boolean;
    OnAirCheckDeviceAlarm: Boolean;
    OnAirCheckDeviceAlarmFileName: String;
    OnAirCheckDeviceAlarmDuration: TTimecode;
    OnAirCheckDeviceAlarmInterval: TTimecode;

    OnAirCheckEventAlarm: Boolean;
    OnAirCheckEventAlarmFileName: String;
    OnAirCheckEventAlarmDuration: TTimecode;
    OnAirCheckEventAlarmInterval: TTimecode;

    MediaCheckAlarm: Boolean;
    MediaCheckAlarmFileName: String;
    MediaCheckAlarmDuration: TTimecode;
    MediaCheckAlarmInterval: TTimecode;

    BlankCheckAlarm: Boolean;
    BlankCheckAlarmFileName: String;
    BlankCheckAlarmDuration: TTimecode;
    BlankCheckAlarmInterval: TTimecode;
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
  IDX_COL_CUESHEET_FINISH_ACTION: Word  = 19;
  IDX_COL_CUESHEET_VIDEO_TYPE: Word     = 20;
  IDX_COL_CUESHEET_AUDIO_TYPE: Word     = 21;
  IDX_COL_CUESHEET_CLOSED_CAPTION: Word = 22;
  IDX_COL_CUESHEET_VOICE_ADD: Word      = 23;
  IDX_COL_CUESHEET_PROGRAM_TYPE: Word   = 24;
  IDX_COL_CUESHEET_NOTES: Word          = 25;

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
  NAM_COL_CUESHEET_FINISH_ACTION: String  = 'FA';//'Finish Action';
  NAM_COL_CUESHEET_VIDEO_TYPE: String     = 'VT';//'Video Type';
  NAM_COL_CUESHEET_AUDIO_TYPE: String     = 'AT';//'Audio Type';
  NAM_COL_CUESHEET_CLOSED_CAPTION: String = 'CC';//'Closed Caption';
  NAM_COL_CUESHEET_VOICE_ADD: String      = 'VA';//'Voice Add';
  NAM_COL_CUESHEET_PROGRAM_TYPE: String   = 'Program';//'Program Type';
  NAM_COL_CUESHEET_NOTES: String          = 'Notes';

  // CueSheet Column Width
  WIDTH_COL_CUESHEET_GROUP: Integer           = 40;
  WIDTH_COL_CUESHEET_NO: Integer              = 30;
  WIDTH_COL_CUESHEET_EVENT_MODE: Integer      = 30;
  WIDTH_COL_CUESHEET_EVENT_STATUS: Integer    = 60;
  WIDTH_COL_CUESHEET_START_MODE: Integer      = 84;
  WIDTH_COL_CUESHEET_START_DATE: Integer      = 82;
  WIDTH_COL_CUESHEET_START_TIME: Integer      = 84;
  WIDTH_COL_CUESHEET_INPUT: Integer           = 60;
  WIDTH_COL_CUESHEET_OUTPUT: Integer          = 60;
  WIDTH_COL_CUESHEET_TITLE: Integer           = 240;
  WIDTH_COL_CUESHEET_SUB_TITLE: Integer       = 240;
  WIDTH_COL_CUESHEET_SOURCE: Integer          = 120;
  WIDTH_COL_CUESHEET_MEDIA_ID: Integer        = 100;
  WIDTH_COL_CUESHEET_MEDIA_STATUS: Integer    = 100;
  WIDTH_COL_CUESHEET_DURATON: Integer         = 84;
  WIDTH_COL_CUESHEET_IN_TC: Integer           = 84;
  WIDTH_COL_CUESHEET_OUT_TC: Integer          = 84;
  WIDTH_COL_CUESHEET_TR_TYPE: Integer         = 40;
  WIDTH_COL_CUESHEET_TR_RATE: Integer         = 40;
  WIDTH_COL_CUESHEET_FINISH_ACTION: Integer   = 50;
  WIDTH_COL_CUESHEET_VIDEO_TYPE: Integer      = 40;
  WIDTH_COL_CUESHEET_AUDIO_TYPE: Integer      = 52;
  WIDTH_COL_CUESHEET_CLOSED_CAPTION: Integer  = 46;
  WIDTH_COL_CUESHEET_VOICE_ADD: Integer       = 100;
  WIDTH_COL_CUESHEET_PROGRAM_TYPE: Integer    = 66;
  WIDTH_COL_CUESHEET_NOTES: Integer           = 240;

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
  VIS_COL_CUESHEET_FINISH_ACTION: Boolean   = True;
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
  NAM_COL_DEVICE_DCS: String              = 'Control';
//  NAM_COL_DEVICE_CHANNEL: String          = 'Channel';

  // Device Column Width
  WIDTH_COL_DEVICE_NO: Integer            = 30;
  WIDTH_COL_DEVICE_NAME: Integer          = 100;
  WIDTH_COL_DEVICE_STATUS: Integer        = 100;
  WIDTH_COL_DEVICE_TIMECODE: Integer      = 84;
  WIDTH_COL_DEVICE_DCS: Integer           = 80;
//  WIDTH_COL_DEVICE_CHANNEL: Integer       = 76;

//  TIMELINE_SIDE_FRAMES: Integer = 1800;

  HIDE_SUB_EVENT: Boolean = True;

  FORMAT_DATE: String;

  // All channel color
  COLOR_BK_ALL_CHANNEL_ONAIR: TColor = clRed;
  COLOR_BK_ALL_CHANNEL_OFFAIR: TColor = $006F6261;

  COLOR_TX_ALL_CHANNEL_ONAIR: TColor = clWhite;
  COLOR_TX_ALL_CHANNEL_OFFAIR: TColor = $00B8B8B8;

  // Event status color
  COLOR_BK_EVENTSTATUS_NORMAL: TColor = $00261B1B;//$003A2C28;
  COLOR_BK_EVENTSTATUS_NEXT: TColor = $00D4807A;//clSkyBlue;
  COLOR_BK_EVENTSTATUS_CUED: TColor = $000BA9B8;
  COLOR_BK_EVENTSTATUS_ONAIR: TColor = $0090BD4A;//clMoneyGreen;
  COLOR_BK_EVENTSTATUS_DONE: TColor = $00392F2D;//clDkGray;
  COLOR_BK_EVENTSTATUS_ERROR: TColor = $003817A5;//clRed;
  COLOR_BK_EVENTSTATUS_TARGET: TColor = $00AB7F53;//clAqua;

  COLOR_TX_EVENTSTATUS_NORMAL: TColor = $00B8B8B8;//clSilver;
  COLOR_TX_EVENTSTATUS_NEXT: TColor = clBlack;
  COLOR_TX_EVENTSTATUS_CUED: TColor = clBlack;
  COLOR_TX_EVENTSTATUS_ONAIR: TColor = clBlack;
  COLOR_TX_EVENTSTATUS_DONE: TColor = $006F6261;//clSilver;
  COLOR_TX_EVENTSTATUS_ERROR: TColor = clWhite;
  COLOR_TX_EVENTSTATUS_TARGET: TColor = clBlack;

  COLOR_TO_EVENTSTATUS_NORMAL: TColor = $0045332D;//$003A2C28;
  COLOR_TO_EVENTSTATUS_NEXT: TColor = $00A2625D;//clSkyBlue;
  COLOR_TO_EVENTSTATUS_CUED: TColor = $00177F89;//clYellow;
  COLOR_TO_EVENTSTATUS_ONAIR: TColor = $00738D3C;
  COLOR_TO_EVENTSTATUS_DONE: TColor = $00302725;//clDkGray;
  COLOR_TO_EVENTSTATUS_ERROR: TColor = $002F1485;//clRed;
  COLOR_TO_EVENTSTATUS_TARGET: TColor = $00836140;//clAqua;

  // Cuesheet select & cut & copy color
  COLOR_ROW_SELECT_CUESHHET: TColor = $00FFAB1D;//$00FFBDAD;

  COLOR_BK_COMMENT_CUESHEET: TColor = $006A4800;//$003A2C28;
  COLOR_BK_CLIPBOARD: TColor = $00312420;

  COLOR_TX_COMMENT_CUESHEET: TColor = $00FFEA00;//$003A2C28;
  COLOR_TX_CLIPBOARD: TColor = clGrayText;

  // Device Status Color
  COLOR_ROW_SELECT_DEVICE: TColor = $00FFAB1D;//$00FFBDAD;

  COLOR_BK_DEVICE_COMM_ERROR: TColor = $003817A5;//clRed;
  COLOR_BK_DEVICESTATUS_NORMAL: TColor = $00261B1B;//$003A2C28;
  COLOR_BK_DEVICESTATUS_DISCONNECT: TColor = $003817A5;//clRed;

  COLOR_TX_DEVICE_COMM_ERROR: TColor = clWhite;
  COLOR_TX_DEVICESTATUS_NORMAL: TColor = $00B8B8B8;//clSilver;
  COLOR_TX_DEVICESTATUS_DISCONNECT: TColor = clWhite;

  // Media Status Cell Color
  COLOR_BK_MEDIASTATUS_NORMAL: TColor = $00261B1B;//$003A2C28;
  COLOR_BK_MEDIASTATUS_NOT_EXIST: TColor = $003817A5;//clRed;

  COLOR_TX_MEDIASTATUS_NORMAL: TColor = $00B8B8B8;//clSilver;
  COLOR_TX_MEDIASTATUS_NOT_EXIST: TColor = clWhite;

  // Cuesheet select & cut & copy color
  COLOR_BAR_TIMELINE_RAIL: TColor = clRed;//$00FFBDAD;

  // Warning Color
  COLOR_TX_WARNING_NORMAL: TColor = $00FFD183;//$00FFBDAD;
  COLOR_TX_WARNING: TColor = $003817A5;//clRed;

  GV_SettingGeneral: TSettingGeneral;
  GV_SettingSEC: TSettingSEC;
  GV_SettingMCC: TSettingMCC;
  GV_SettingDCS: TSettingDCS;
  GV_SettingEventColumnDefault: TSettingEventColumnDefault;
  GV_SettingTimeParameter: TSettingTimeParameter;
  GV_SettingThresholdTime: TSettingThresholdTime;
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

  GV_ClipboardCueSheet: TClipboardCueSheet;

  GV_TimeBefore, GV_TimeCurrent: TSystemTime;
  GV_TimerExecuteEvent: THandle;
  GV_TimerCancelEvent: THandle;

  GV_ChannelMaxCount: Word = 4;
  GV_ChannelMinCount: Word = 1;
  GV_ChannelTimelineMinHeight: Word = 216;
  GV_ChannelTimelineTrackTopHeight: Integer = 6;

  GV_AllTimelineTrackTopHeight: Integer = 10;

  GV_TimelineOnairIcon: TPicture;
  GV_TimelineNextIcon: TPicture;
  GV_TimelineNormalIcon: TPicture;

  procedure LoadConfig;
  procedure SaveConfig;

  procedure SaveSECConfig;
  procedure SaveDCSConfig;

  procedure AssertProc(const AMessage, AFileName: String; ALineNumber: Integer; AErrorAddr: Pointer);

  function GetMainLogStr(AState: TLogState; ALogStr: String): String; overload;
  function GetMainLogStr(AState: TLogState; APRec: PResStringRec): String; overload;
  function GetMainLogStr(AState: TLogState; APRec: PResStringRec; const Args: array of const): String; overload;

  function GetChannelLogStr(AState: TLogState; const AChannelID: Word; const ALogStr: String): String;

  function WriteAssertLog(const ALogStr: String): Integer;

  function CheckSum(AValue: AnsiString): Boolean;

  function HasMainControl: Boolean;

  function GetSECByID(AID: Word): PSEC;

  function GetMCCByID(AID: Word): PMCC;

  function GetChannelByName(AName: String): PChannel;
  function GetChannelByID(AID: Byte): PChannel;
  function GetChannelNameByID(AID: Byte): String;

  function GetChannelOnAirByID(AID: Integer): Boolean;
  procedure SetChannelOnAirByID(AID: Integer; AOnAir: Boolean);

  function GetDCSByID(AID: Word): PDCS;
  function GetDCSByName(AName: String): PDCS;
  function GetDCSByIP(AIP: String): PDCS;
  function GetDCSNameByID(AID: Word): String;

  function GetSourceByName(AName: String): PSource;
  function GetSourceByIDAndHandle(ADCSID: Word; AHandle: TDeviceHandle): PSource;
  function GetSourceByIPAndHandle(ADCSIP: String; AHandle: TDeviceHandle): PSource;

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

  function GetMediaDuration(AFileName: String): TTimecode;

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

      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('DCS', Ident, '0,DCS1,127.0.0.1,0,0')), DataStrings);
        if (DataStrings.Count < 5) then
        begin
          MessageBeep(MB_ICONERROR);
          MessageBox(Application.Handle, PChar(SEInvalidSettinDCS), PChar(Application.Title), MB_OK or MB_ICONERROR or MB_TOPMOST);
          exit;
        end;

        DCS := New(PDCS);
        FillChar(DCS^, SizeOf(TDCS), #0);

        DCS^.ID   := StrToIntDef(DataStrings[0], 0);
        StrPCopy(DCS^.Name, DataStrings[1]);
        StrPCopy(DCS^.HostIP, DataStrings[2]);
        DCS^.Main := StrToBoolDef(DataStrings[3], False);
        DCS^.Mine := StrToBoolDef(DataStrings[4], False);
        DCS^.Alive := False;
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
//          SourceHandle^.DCSID  := GV_DCSList[J]^.ID;
//          StrCopy(SourceHandle^.DCSIP, GV_DCSList[J]^.HostIP);
          SourceHandle^.DCS := GV_DCSList[J];
          SourceHandle^.Handle := -1;

          Source^.Handles.Add(SourceHandle);
        end;

        Source^.Channel := GetChannelByID(IniFile.ReadInteger(Section, 'Channel', -1));
//        Source^.DCS := GV_DCSList[J];//GetDCSByName(IniFile.ReadString(Section, 'DCS', ''));
        Source^.SourceType := GetSourceTypeByName(IniFile.ReadString(Section, 'Type', ''));
        Source^.MakeTransition := StrToBoolDef(IniFile.ReadString(Section, 'MakeTransition', 'False'), False);
        Source^.SuccessiveUse := StrToBoolDef(IniFile.ReadString(Section, 'SuccessiveUse', 'False'), False);
        Source^.LoadAlarm := StrToBoolDef(IniFile.ReadString(Section, 'LoadAlarm', 'False'), False);
        Source^.EjectAlarm := StrToBoolDef(IniFile.ReadString(Section, 'EjectAlarm', 'False'), False);
        Source^.QueryDuration := StrToBoolDef(IniFile.ReadString(Section, 'QueryDuration', 'False'), False);
        Source^.PrerollTime := StringToTimecode(IniFile.ReadString(Section, 'PrerollTime', '00:00:00:00'));
        Source^.DefaultTimecode := StringToTimecode(IniFile.ReadString(Section, 'DefaultTimecode', '00:00:00:00'));
        Source^.UseDefaultTimecode := StrToBoolDef(IniFile.ReadString(Section, 'UseDefaultTimecode', 'False'), False);
        Source^.UsePCS := StrToBoolDef(IniFile.ReadString(Section, 'UsePCS', 'False'), False);

        Source^.CommSuccess := False;
        Source^.CommTimeout := 0;

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

        if (Source^.SourceType in [ST_CG]) then
        begin
          Source^.NumLayer := IniFile.ReadInteger(Section, 'NumLayer', 0);
        end
        else if (Source^.SourceType in [ST_ROUTER, ST_MCS]) then
        begin
          Source^.Router := TXptList.Create;

          NumXpt := IniFile.ReadInteger(Section, 'NumXpt', 0);
          for K := 0 to NumXPT - 1 do
          begin
            Ident := Format('Xpt%d', [K + 1]);

            DataStrings := TStringList.Create;
            try
              ExtractStrings([','], [' '], PChar(IniFile.ReadString(Section, Ident, '0,0')), DataStrings);
              if (DataStrings.Count < 2) then
              begin
                MessageBeep(MB_ICONERROR);
                MessageBox(Application.Handle, PChar(SEInvalidSettinXpt), PChar(Application.Title), MB_OK or MB_ICONERROR or MB_TOPMOST);
                exit;
              end;

              Xpt := New(PXpt);
              FillChar(Xpt^, SizeOf(TXPT), #0);

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

      ID                := IniFile.ReadInteger(Section, 'ID', 0);
      Name              := IniFile.ReadString(Section, 'Name', 'MCC');
      HostIP            := IniFile.ReadString(Section, 'HostIP', '127.0.0.1');

      WorkCueSheetPath := IniFile.ReadString(Section, 'WorkCueSheetPath', ExtractFilePath(Application.ExeName) + 'CueSheet\Work') + PathDelim;
      LoadCueSheetPath := IniFile.ReadString(Section, 'LoadCueSheetPath', ExtractFilePath(Application.ExeName) + 'CueSheet\Load') + PathDelim;
      SaveCueSheetPath := IniFile.ReadString(Section, 'SaveCueSheetPath', ExtractFilePath(Application.ExeName) + 'CueSheet\Save') + PathDelim;
      MediaPath        := IniFile.ReadString(Section, 'MediaPath', ExtractFilePath(Application.ExeName) + 'Media') + PathDelim;
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
      SysInPort             := IniFile.ReadInteger(Section, 'SysInPort', 9012);
      SysCheckTimeout       := StringToTimecode(IniFile.ReadString(Section, 'SysCheckTimeout', '00:00:01:00'));
      SysCheckInterval      := StringToTimecode(IniFile.ReadString(Section, 'SysCheckInterval', '00:00:01:00'));
      SysLogInPortEnabled   := StrToBoolDef(IniFile.ReadString(Section, 'SysLogInPortEnabled', 'False'), False);
      SysLogPath            := IniFile.ReadString(Section, 'SysLogPath', ExtractFilePath(Application.ExeName) + 'SECLog') + PathDelim;
      SysLogExt             := IniFile.ReadString(Section, 'SysLogExt', 'SysSEC.log');

      NotifyPort            := IniFile.ReadInteger(Section, 'NotifyPort', 9000);
      InPort                := IniFile.ReadInteger(Section, 'InPort', 9002);
      OutPort               := IniFile.ReadInteger(Section, 'OutPort', 9001);
      CrossPort             := IniFile.ReadInteger(Section, 'CrossPort', 9003);
      CommandTimeout        := StringToTimecode(IniFile.ReadString(Section, 'CommandTimeout', '00:00:01:00'));

      NumCrossCheck         := IniFile.ReadInteger(Section, 'NumCrossCheck', 3);
      CrossCheckInterval    := StringToTimecode(IniFile.ReadString(Section, 'CrossCheckInterval', '00:00:01:00'));

      LogNotifyPortEnabled  := StrToBoolDef(IniFile.ReadString(Section, 'LogNotifyPortEnabled', 'False'), False);
      LogInPortEnabled      := StrToBoolDef(IniFile.ReadString(Section, 'LogInPortEnabled', 'False'), False);
      LogOutPortEnabled     := StrToBoolDef(IniFile.ReadString(Section, 'LogOutPortEnabled', 'False'), False);
      LogCrossPortEnabled   := StrToBoolDef(IniFile.ReadString(Section, 'LogCrossPortEnabled', 'False'), False);
      LogPath               := IniFile.ReadString(Section, 'LogPath', ExtractFilePath(Application.ExeName) + 'SECLog') + PathDelim;
      LogExt                := IniFile.ReadString(Section, 'LogExt', 'SEC.log');
    end;

    // Num SEC
    NumSEC := IniFile.ReadInteger(Section, 'NumSEC', 0);

    GV_SECMain := nil;
    for I := 0 to NumSEC - 1 do
    begin
      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('SEC', Format('SEC%d', [I + 1]), '-')), DataStrings);
        if (DataStrings.Count < 4) then
        begin
          MessageBeep(MB_ICONERROR);
          MessageBox(Application.Handle, PChar(SEInvalidSettinSEC), PChar(Application.Title), MB_OK or MB_ICONERROR or MB_TOPMOST);
          exit;
        end;

        SEC := New(PSEC);
        FillChar(SEC^, SizeOf(TSEC), #0);

        SEC^.ID := StrToIntDef(DataStrings[0], 0);
        StrPCopy(SEC^.Name, DataStrings[1]);
        StrPCopy(SEC^.HostIP, DataStrings[2]);
        SEC^.Main := StrToBoolDef(DataStrings[3], False);
//        SEC^.Mine := StrToBoolDef(DataStrings[4], False);
        SEC^.Alive := False;

        if (SEC^.Main) then
          GV_SECMain := SEC;

        if (GV_SettingGeneral.ID = SEC^.ID) and
           (GV_SettingGeneral.Name = String(SEC^.Name)) and
           (GV_SettingGeneral.HostIP = String(SEC^.HostIP)) then
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
      Use                   := StrToBoolDef(IniFile.ReadString(Section, 'Use', 'False'), False);
      SysInPort             := IniFile.ReadInteger(Section, 'SysInPort', 7012);
      SysOutPort            := IniFile.ReadInteger(Section, 'SysOutPort', 7011);
      SysCheckTimeout       := StringToTimecode(IniFile.ReadString(Section, 'SysCheckTimeout', '00:00:01:00'));
      SysCheckInterval      := StringToTimecode(IniFile.ReadString(Section, 'SysCheckInterval', '00:00:01:00'));
      SysLogInPortEnabled   := StrToBoolDef(IniFile.ReadString(Section, 'SysLogInPortEnabled', 'False'), False);
      SysLogOutPortEnabled  := StrToBoolDef(IniFile.ReadString(Section, 'SysLogOutPortEnabled', 'False'), False);
      SysLogPath            := IniFile.ReadString(Section, 'SysLogPath', ExtractFilePath(Application.ExeName) + 'MCCLog') + PathDelim;
      SysLogExt             := IniFile.ReadString(Section, 'SysLogExt', 'SysMCC.log');

      NotifyPort    := IniFile.ReadInteger(Section, 'NotifyPort', 7000);
      InPort        := IniFile.ReadInteger(Section, 'InPort', 7002);
      OutPort       := IniFile.ReadInteger(Section, 'OutPort', 7001);

      CommandTimeout  := StringToTimecode(IniFile.ReadString(Section, 'CommandTimeout', '00:00:01:00'));

      LogNotifyPortEnabled  := StrToBoolDef(IniFile.ReadString(Section, 'LogNotifyPortEnabled', 'False'), False);
      LogInPortEnabled      := StrToBoolDef(IniFile.ReadString(Section, 'LogInPortEnabled', 'False'), False);
      LogOutPortEnabled     := StrToBoolDef(IniFile.ReadString(Section, 'LogOutPortEnabled', 'False'), False);
      LogPath               := IniFile.ReadString(Section, 'LogPath', ExtractFilePath(Application.ExeName) + 'MCCLog') + PathDelim;
      LogExt                := IniFile.ReadString(Section, 'LogExt', 'MCC.log');
    end;

    // Num MCC
    NumMCC := IniFile.ReadInteger(Section, 'NumMCC', 0);

    for I := 0 to NumMCC - 1 do
    begin
      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('MCC', Format('MCC%d', [I + 1]), '-')), DataStrings);
        if (DataStrings.Count < 3) then
        begin
          MessageBeep(MB_ICONERROR);
          MessageBox(Application.Handle, PChar(SEInvalidSettinMCC), PChar(Application.Title), MB_OK or MB_ICONERROR or MB_TOPMOST);
          exit;
        end;

        MCC := New(PMCC);
        FillChar(MCC^, SizeOf(TMCC), #0);

        MCC^.ID := StrToIntDef(DataStrings[0], 0);
        StrPCopy(MCC^.Name, DataStrings[1]);
        StrPCopy(MCC^.HostIP, DataStrings[2]);
//        MCC^.Main := StrToBoolDef(DataStrings[3], False);
//        MCC^.Mine := StrToBoolDef(DataStrings[4], False);
        MCC^.Alive := False;
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
      SysInPort             := IniFile.ReadInteger(Section, 'SysInPort', 8012);
      SysOutPort            := IniFile.ReadInteger(Section, 'SysOutPort', 8011);
      SysCheckTimeout       := StringToTimecode(IniFile.ReadString(Section, 'SysCheckTimeout', '00:00:01:00'));
      SysCheckInterval      := StringToTimecode(IniFile.ReadString(Section, 'SysCheckInterval', '00:00:01:00'));
      SysLogInPortEnabled   := StrToBoolDef(IniFile.ReadString(Section, 'SysLogInPortEnabled', 'False'), False);
      SysLogOutPortEnabled  := StrToBoolDef(IniFile.ReadString(Section, 'SysLogOutPortEnabled', 'False'), False);
      SysLogPath            := IniFile.ReadString(Section, 'SysLogPath', ExtractFilePath(Application.ExeName) + 'DCSLog') + PathDelim;
      SysLogExt             := IniFile.ReadString(Section, 'SysLogExt', 'SysDCS.log');

      NotifyPort      := IniFile.ReadInteger(Section, 'NotifyPort', 8000);
      InPort          := IniFile.ReadInteger(Section, 'InPort', 8002);
      OutPort         := IniFile.ReadInteger(Section, 'OutPort', 8001);

      CommandTimeout  := StringToTimecode(IniFile.ReadString(Section, 'CommandTimeout', '00:00:01:00'));

      LogNotifyPortEnabled  := StrToBoolDef(IniFile.ReadString(Section, 'LogNotifyPortEnabled', 'False'), False);
      LogInPortEnabled      := StrToBoolDef(IniFile.ReadString(Section, 'LogInPortEnabled', 'False'), False);
      LogOutPortEnabled     := StrToBoolDef(IniFile.ReadString(Section, 'LogOutPortEnabled', 'False'), False);
      LogPath               := IniFile.ReadString(Section, 'LogPath', ExtractFilePath(Application.ExeName) + 'DCSLog') + PathDelim;
      LogExt                := IniFile.ReadString(Section, 'LogExt', 'DCS.log');
    end;

    // MCS
    // Num MCS
    NumType := IniFile.ReadInteger('MCS', 'NumMCS', 0);

    for I := 0 to NumType - 1 do
    begin
      DataStrings := TStringList.Create;
      try
        ExtractStrings([','], [' '], PChar(IniFile.ReadString('MCS', Format('MCS%d', [I + 1]), '-')), DataStrings);
        if (DataStrings.Count < 2) then
        begin
          MessageBeep(MB_ICONERROR);
          MessageBox(Application.Handle, PChar(SEInvalidSettinMCS), PChar(Application.Title), MB_OK or MB_ICONERROR or MB_TOPMOST);
          exit;
        end;

        MCS := New(PMCS);
        FillChar(MCS^, SizeOf(TMCS), #0);

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
        if (DataStrings.Count < 5) then
        begin
          MessageBeep(MB_ICONERROR);
          MessageBox(Application.Handle, PChar(SEInvalidSettinProgramType), PChar(Application.Title), MB_OK or MB_ICONERROR or MB_TOPMOST);
          exit;
        end;

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
      FinishAction   := TFinishAction(IniFile.ReadInteger(Section, 'FinishAction', 0));
      ProgramType    := IniFile.ReadInteger(Section, 'ProgramType', 0);
    end;

    // Time parameter
    Section := 'TimeParameter';

    FillChar(GV_SettingTimeParameter, SizeOf(TSettingTimeParameter), #0);
    with GV_SettingTimeParameter do
    begin
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

      MediaCheckInterval              := StringToTimecode(IniFile.ReadString(Section, 'MediaCheckInterval', '00:20:00:00'));
      MediaCheckTime                  := StringToTimecode(IniFile.ReadString(Section, 'MediaCheckTime', '02:00:00:00'));

      BlankCheckInterval              := StringToTimecode(IniFile.ReadString(Section, 'BlankCheckInterval', '00:20:00:00'));
      BlankCheckTime                  := StringToTimecode(IniFile.ReadString(Section, 'BlankCheckTime', '02:00:00:00'));
    end;

    // Treshhold time
    Section := 'ThresholdTime';

    FillChar(GV_SettingThresholdTime, SizeOf(TSettingThresholdTime), #0);
    with GV_SettingThresholdTime do
    begin
      OnAirCheckDeviceTimeout       := StringToTimecode(IniFile.ReadString(Section, 'OnAirCheckDeviceTimeout', '00:00:04:00'));

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

      OnAirEventBreakingDuration    := StringToTimecode(IniFile.ReadString(Section, 'OnAirEventBreakingDuration', '00:05:00:00'));
    end;

    // Option
    Section := 'Option';

    FillChar(GV_SettingOption, SizeOf(TSettingOption), #0);
    with GV_SettingOption do
    begin
      TimelineSpace                   := IniFile.ReadInteger(Section, 'TimelineSpace', 200);
      TimelineSpaceInterval           := IniFile.ReadInteger(Section, 'TimelineSpaceInterval', 10);
      TimelineOnairIconFileName       := IniFile.ReadString(Section, 'TimelineOnairIconFileName', '');
      TimelineNextIconFileName        := IniFile.ReadString(Section, 'TimelineNextIconFileName', '');
      TimelineNormalIconFileName      := IniFile.ReadString(Section, 'TimelineNormalIconFileName', '');

      ChannelTimelineHeight           := IniFile.ReadInteger(Section, 'ChannelTimelineHeight', 77);

      AutoLoadCuesheet                := StrToBoolDef(IniFile.ReadString(Section, 'AutoLoadCuesheet', 'False'), False);
      AutoLoadCuesheetInterval        := StringToTimecode(IniFile.ReadString(Section, 'AutoLoadCuesheetInterval', '01:00:00:00'));
      AutoEjectCuesheet               := StrToBoolDef(IniFile.ReadString(Section, 'AutoEjectCuesheet', 'False'), False);
      AutoEjectCuesheetInterval       := StringToTimecode(IniFile.ReadString(Section, 'AutoEjectCuesheetInterval', '01:00:00:00'));
      MaxInputEventCount              := IniFile.ReadInteger(Section, 'MaxInputEventCount', 50);
      OnAirEventHighlight             := StrToBoolDef(IniFile.ReadString(Section, 'OnAirEventHighlight', 'True'), False);
      OnAirEventFixedRow              := IniFile.ReadInteger(Section, 'OnAirEventFixedRow', 0);

      OnAirCheckDeviceNotify          := StrToBoolDef(IniFile.ReadString(Section, 'OnAirCheckDeviceNotify', 'True'), False);
      OnAirCheckDeviceAlarm           := StrToBoolDef(IniFile.ReadString(Section, 'OnAirCheckDeviceAlarm', 'False'), False);
      OnAirCheckDeviceAlarmFileName   := IniFile.ReadString(Section, 'OnAirCheckDeviceAlarmFileName', '');
      OnAirCheckDeviceAlarmDuration   := StringToTimecode(IniFile.ReadString(Section, 'OnAirCheckDeviceAlarmDuration', '00:00:10:00'));
      OnAirCheckDeviceAlarmInterval   := StringToTimecode(IniFile.ReadString(Section, 'OnAirCheckDeviceAlarmInterval', '00:00:10:00'));

      OnAirCheckEventAlarm            := StrToBoolDef(IniFile.ReadString(Section, 'OnAirCheckEventAlarm', 'False'), False);
      OnAirCheckEventAlarmFileName    := IniFile.ReadString(Section, 'OnAirCheckEventAlarmFileName', '');
      OnAirCheckEventAlarmDuration    := StringToTimecode(IniFile.ReadString(Section, 'OnAirCheckEventAlarmDuration', '00:00:10:00'));
      OnAirCheckEventAlarmInterval    := StringToTimecode(IniFile.ReadString(Section, 'OnAirCheckEventAlarmInterval', '00:00:10:00'));

      MediaCheckAlarm                 := StrToBoolDef(IniFile.ReadString(Section, 'MediaCheckAlarm', 'False'), False);
      MediaCheckAlarmFileName         := IniFile.ReadString(Section, 'MediaCheckAlarmFileName', '');
      MediaCheckAlarmDuration         := StringToTimecode(IniFile.ReadString(Section, 'MediaCheckAlarmDuration', '00:00:10:00'));
      MediaCheckAlarmInterval         := StringToTimecode(IniFile.ReadString(Section, 'MediaCheckAlarmInterval', '00:00:10:00'));

      BlankCheckAlarm                 := StrToBoolDef(IniFile.ReadString(Section, 'BlankCheckAlarm', 'False'), False);
      BlankCheckAlarmFileName         := IniFile.ReadString(Section, 'BlankCheckAlarmFileName', '');
      BlankCheckAlarmDuration         := StringToTimecode(IniFile.ReadString(Section, 'BlankCheckAlarmDuration', '00:00:10:00'));
      BlankCheckAlarmInterval         := StringToTimecode(IniFile.ReadString(Section, 'BlankCheckAlarmInterval', '00:00:10:00'));
    end;
  finally
    FreeAndNil(IniFile);
  end;

  FORMAT_DATE := 'YYYY' + FormatSettings.DateSeparator + 'MM' + FormatSettings.DateSeparator + 'DD';
end;

procedure SaveConfig;
begin
end;

procedure SaveSECConfig;
var
  IniFile: TIniFile;
  I: Integer;

  Section: String;
  SEC: PSEC;
begin
  // SEC
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'SEC.ini');
  try
    // SEC
    Section := 'SEC';

    IniFile.WriteInteger(Section, 'NumSEC', GV_SECList.Count);

    for I := 0 to GV_SECList.Count - 1 do
    begin
      SEC := GV_SECList[I];
      if (SEC <> nil) then
      begin
        IniFile.WriteString(Section, Format('SEC%d', [I + 1]), Format('%d,%s,%s,%d', [SEC^.ID, String(SEC^.Name), String(SEC^.HostIP), Integer(SEC^.Main)]));
      end;
    end;
  finally
    FreeAndNil(IniFile);
  end;
end;

procedure SaveDCSConfig;
var
  IniFile: TIniFile;
  I: Integer;

  Section: String;
  DCS: PDCS;
begin
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'Source.ini');
  try
    // DCS
    Section := 'DCS';

    IniFile.WriteInteger(Section, 'NumDCS', GV_DCSList.Count);

    for I := 0 to GV_DCSList.Count - 1 do
    begin
      DCS := GV_DCSList[I];
      if (DCS <> nil) then
      begin
        IniFile.WriteString(Section, Format('DCS%d', [I + 1]), Format('%d,%s,%s,%d,%d', [DCS^.ID, String(DCS^.Name), String(DCS^.HostIP), Integer(DCS^.Main), Integer(DCS^.Mine)]));
      end;
    end;
  finally
    FreeAndNil(IniFile);
  end;
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
  DateStr := FormatDateTime('YYYY-MM-DD hh:nn:ss.zzz', Now);
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
  DateStr := FormatDateTime('YYYY-MM-DD hh:nn:ss.zzz', Now);
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
        MessageBox(Application.Handle, PChar(Format('Cannot create log directory %s', [FilePath])), PChar(Application.Title), MB_OK or MB_ICONERROR or MB_TOPMOST);
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
      MessageBox(Application.Handle, PChar(Format('Cannot create or open log file %s', [FileName])), PChar(Application.Title), MB_OK or MB_ICONERROR or MB_TOPMOST);
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

  if (Len >= 5) then
  begin
    for I := 4 to Len - 1 do
      CRC := CRC + Ord(AValue[I]);
    CRC := not (CRC) + 1;

    Result := (CRC = Ord(AValue[Len]));
  end;
end;

function HasMainControl: Boolean;
begin
  Result := (GV_SECMine <> nil) and (GV_SECMine = GV_SECMain);
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

function GetDCSByID(AID: Word): PDCS;
var
  I: Integer;
begin
  Result := nil;
  for I := 0 to GV_DCSList.Count - 1 do
  begin
    if (GV_DCSList[I]^.ID = AID) then
    begin
      Result := GV_DCSList[I];
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

function GetSourceByIDAndHandle(ADCSID: Word; AHandle: TDeviceHandle): PSource;
var
  I, J: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;
begin
  Result := nil;

  for I := 0 to GV_SourceList.Count - 1 do
  begin
    Source := GV_SourceList[I];
    if (Source <> nil) then
    begin
      SourceHandles := Source^.Handles;
      if (SourceHandles <> nil) then
      begin
        for J := 0 to SourceHandles.Count - 1 do
          if (SourceHandles[J]^.DCS <> nil) and (SourceHandles[J]^.DCS^.ID = ADCSID) and
             (SourceHandles[J].Handle = AHandle) then
          begin
            Result := Source;
            exit;
          end;
      end;
    end;
  end;
end;

function GetSourceByIPAndHandle(ADCSIP: String; AHandle: TDeviceHandle): PSource;
var
  I, J: Integer;
  Source: PSource;
  SourceHandles: TSourceHandleList;
begin
  Result := nil;

  for I := 0 to GV_SourceList.Count - 1 do
  begin
    Source := GV_SourceList[I];
    if (Source <> nil) then
    begin
      SourceHandles := Source^.Handles;
      if (SourceHandles <> nil) then
      begin
        for J := 0 to SourceHandles.Count - 1 do
          if (SourceHandles[J]^.DCS <> nil) and (String(SourceHandles[J]^.DCS^.HostIP) = ADCSIP) and
             (SourceHandles[J].Handle = AHandle) then
          begin
            Result := Source;
            exit;
          end;
      end;
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
  // Channel List
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
      ST_ROUTER,
      ST_MCS:
        if (GV_SourceList[I]^.Router <> nil) then
        begin
          for J := GV_SourceList[I]^.Router.Count - 1 downto 0 do
            Dispose(GV_SourceList[I]^.Router[J]);

          GV_SourceList[I]^.Router.Clear;
          GV_SourceList[I]^.Router.Free;
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

function GetMediaDuration(AFileName: String): TTimecode;
var
  MediaInfoHandle: NativeInt;
  DurationStr: String;
begin
  Result := 0;

  if (MediaInfoDLL_Load('MediaInfo.dll')) then
  begin
    MediaInfoHandle := MediaInfo_New;
    try
      if (MediaInfo_Open(MediaInfoHandle, PChar(AFileName)) > 0) then
      begin
        DurationStr := String(MediaInfo_Get(MediaInfoHandle, Stream_General, 0, 'Duration/String4', Info_Text, Info_Name));
        Result := StringToTimecode(DurationStr);
      end;
    finally
      MediaInfo_Close(MediaInfoHandle);
    end;
  end;
end;

initialization
  AssertErrorProc := AssertProc;

end.
