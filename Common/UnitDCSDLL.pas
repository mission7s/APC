unit UnitDCSDLL;

interface

uses Winapi.Windows, System.Types, Generics.Collections,
  UnitCommons;

//const
//  D_OK    = 0;
//  D_FALSE = -1;
//
//	{ General Error Code (10000 ~ 10099) }
//  E_INVALID_ID = 1001;
//  E_INVALID_DEVICE_NAME = 1002;
//  E_INVALID_DEVICE_HANDLE = 1003;
//  E_NOT_OPENED_DEVICE = 1004;
//  E_UNDEFIND_COMMAND = 1005;
//
//	ID_LEN = MAX_PATH;// 20;
//	HOSTIP_LEN = 20;
//	MEDIAID_LEN = MAX_PATH;//20;
//	SECNAME_LEN = 20;
//	DCSNAME_LEN = 20;
//	DEVICENAME_LEN = 20;
//	CHANNELNAME_LEN = 20;
//  DATE_LEN=8;
//	BINNO_LEN = 20;
//  TITLE_LEN = 80;
//  SUBTITLE_LEN = 80;
//  PROGRAMTYPENAME_LEN = 20;
//  NOTES_LEN = 80;
//	GPINAME_LEN = 20;
//
//  EVENTTIME_STR_LEN = 20;
//  TIMECODE_STR_LEN = 11;
//
//	MAX_NUMID = 10;
//  MAX_RETRY = 3;
//  TIME_OUT = 1000;
//
//  INVALID_DEVICE_HANDLE = -1;
//
//type
//  TErrorCode = (
//    ERR_NoError = 0,
//
//    ERR_NotExist = 1000
//    );
//
//	TDeviceHandle = Integer;
//
//  TPortMode = (pmDecoder, pmEncoder);
//
//	TTimecode = DWORD;
//	TEventTime = packed record
//		T: TTimecode;
//		D: TDate;
//	end;
//
//	TID = packed record
//		BinNo: array[0..BINNO_LEN] of Char;
//		ID: array[0..ID_LEN] of Char;
//	end;
//
//	TIDList = packed record
//		NumID: Integer;
//		ID: array[0..MAX_NUMID] of TID;
//	end;
//
//  TOnAirFlagType = (
//    FT_REGULAR = Ord('r'),
//    FT_SUBSTITUTE = Ord('s'),
//    FT_EMERGENCY = Ord('e'),
//    FT_LOG = Ord('l')
//    );
//
//  TEventID = packed record
//    ChannelID: Word;//Byte;
//    OnAirDate: array[0..DATE_LEN] of Char;// TDateTime;   // 방송일자
//    OnAirFlag: TOnAirFlagType;  // R: 정규, S: 대체, E: 긴급, L: 운행로그
//    OnAirNo: Integer;
//    SerialNo: Integer;
//  end;
//  PEventID = ^TEventID;
//
//  TFinishAction = (
//	  FA_NONE = 0,
//	  FA_STOP = 1,
//	  FA_EJECT = 2{,
//	  FA_REWIND = 3,
//	  FA_REWINDEJECT = 4 }
//    );
//
//  TEventMode = (
//    EM_PROGRAM  = $00,
//    EM_MAIN     = $01,
//    EM_JOIN     = $02,
//    EM_SUB      = $03,
//    EM_COMMENT  = $04
//    );
//
//  TStartMode = (
//    SM_ABSOLUTE   = 0,
//    SM_AUTOFOLLOW = 1,
//    SM_MANUAL     = 2,
//    SM_LOOP       = 3,
//    SM_SUBBEGIN   = 4,
//    SM_SUBEND     = 5
//    );
//
//  TInputType = (
//    IT_MAIN     = $00,
//    IT_BACKUP   = $01,
//    IT_KEYER1   = $02,
//    IT_KEYER2   = $03,
//    IT_KEYER3   = $04,
//    IT_KEYER4   = $05,
//    IT_AMIXER1  = $06,
//	  IT_AMIXER2  = $07
//    );
//
//  TOutputBkgndType = (
//    OB_NONE  = $00,
//    OB_VIDEO = $01,
//    OB_AUDIO = $02,
//    OB_BOTH  = $03
//    );
//
//  TOutputKeyerType = (
//    OK_NONE = $00,
//    OK_ON   = $01,
//    OK_OFF  = $02
//    );
//
//  TVideoType = (
////    VT_NORMAL   = $00,
////    VT_WIDE     = $01
//    VT_NONE = $00,
//    VT_UHD  = $01,
//    VT_HD   = $02,
//    VT_SD   = $03,
//    VT_3D   = $04
//    );
//
//  TAudioType = (
////	  AT_NONE     = $00,
////	  AT_LEFT     = $01,
////	  AT_RIGHT    = $02,
////	  AT_MONO     = $03,
////	  AT_STEREO   = $04
//	  AT_NONE     = $00,
//	  AT_5BY1     = $01,
//	  AT_STEREO   = $02,
//	  AT_MONO     = $03
//    );
//
//  TClosedCaption = (
//	  CC_NONE     = $00,
//	  CC_EXIST    = $01
//    );
//
//  TVoiceAdd = (
//	  VA_NONE         = $00,
//	  VA_KOREA_SCREEN = $01,
//	  VA_ENG_VOICE    = $02,
//	  VA_JPN_VOICE    = $03,
//	  VA_CHN_VOICE    = $04,
//	  VA_ETC_VOICE    = $05
//    );
//
//  TTRType = (
//    TT_CUT      = 0,
//    TT_FADE     = 1,
//    TT_CUTFADE  = 2,
//    TT_FADECUT  = 3,
//    TT_MIX      = 4
//    );
//
//  TTRRate = (
//    TR_CUT      = 0,
//    TR_FAST     = 1,
//    TR_MEDIUM   = 2,
//    TR_SLOW     = 3
//    );
//
//  TSourceType = (
//    ST_VSDEC,
//    ST_VSENC,
//    ST_VCR,
//    ST_CART,
//    ST_CG,
//    ST_ROUTER,
//    ST_MCS,
//    ST_GPI,
//    ST_LINE,
//    ST_LOGO
//    );
//
//	TPlayerEvent = packed record
//		StartTC: TTimecode;
//		PreStill: Integer;       // Unknown
//		PostStill: Integer;      // Unknown
//		FinishAction: TFinishAction;
//    ID: TID;
//	end;
//
//	TSwitcherEvent = packed record
//		MainVideo: Integer;
//		MainAudio: Integer;
//		BackupVideo: Integer;
//		BackupAudio: Integer;
//		Key1: Integer;
//		Key2: Integer;
//		Key3: Integer;
//		Key4: Integer;
//		Mix1: Integer;
//		Mix2: Integer;
//		VideoType: Integer;
//		AudioType: Integer;
//		VideoTransType: TTRType;
//		VideoTransRate: TTRRate;
//		AudioTransType: TTRType;
//		AudioTransRate: TTRRate;
//		GPI: Integer;
//	end;
//
//	TGPIEvent = packed record
//		PortNum: Integer;
//		TriggerMode: Integer;
//		AdjMillisec: Integer;
//	end;
//
//	TRSWEvent = packed record
//		XptOut: Integer;
//		XptIn: Integer;
//    XptLevel: Integer;
//		Stuffing: array[0..1] of Boolean;
//	end;
//
//  TEventType = (
//    ET_SWITCHER = $00,
//    ET_PLAYER   = $01,
//    ET_GPI      = $02,
//    ET_RSW      = $03
//    );
//
//  TEventState = (
//    esIdle,
//    esError,
//    esLoading,
//    esLoaded,
//    esCueing,
//    esCued,
//    esStandByOff,
//    esStandByOn,
//    esPreroll,
//    esOnAir,
//    esFinish,
//    esFinishing,
//    esFinished,
//    esDone,
//    esSkipped);
//
//  TEventStatus = record
//    State: TEventState;
//    ErrorCode: TErrorCode;
//  end;
//  PEventStatus = ^TEventStatus;
//
//	TEventOverall = record
//    ChannelID: Integer;
//
//		NumEventInQueue: Integer;
//		NumFreeEntryInQueue: Integer;
//
//		Time: TTIME;
//
//		LastTransitionTime: TTIME;
//		LastAiredEventID: TEVENTID;
//		OnAirEventID: TEVENTID;
//		PreparedEventID: TEVENTID;
//
//		IsHold: Boolean;
//
//{		EventID: TEVENTID;
//		ActionStartTime: TTIME;
//		nStateCode: TEventState;
//		nActionCode: BYTE;
//
//		dwErrCode: DWORD;
//		dwErrLine: DWORD;
//		dwExtraErrCode: DWORD;
//
//		EventType: BYTE;
//		case integer of
//		0: (Switcher: TSwitcherEvent);
//		1: (Player: TPlayerEvent);
//		2: (GPI: TGPIEvent);
//		3: (RSW: TRSWEvent); }
//	end;
//
//	TEvent = packed record
//		EventID: TEventID;
//
//		StartTime: TEventTime;
//		DurTime: TTimecode;
//
//		TakeEvent: Boolean;   // Aleady Onair
//		ManualEvent: Boolean;
//
//    Status: TEventStatus;
//
//		EventType: TEventType;
//		case Integer of
//		  0: (Switcher: TSwitcherEvent);
//		  1: (Player: TPlayerEvent);
//		  2: (GPI: TGPIEvent);
//		  3: (RSW: TRSWEvent);
//	end;
//  PEvent = ^TEvent;
//  TEventList = TList<PEvent>;
//
//	TXPTInfo = packed record
//		MainVideo: Integer;
//		MainAudio: Integer;
//		BackupVideo: Integer;
//		BackupAudio: Integer;
//		Key1: Integer;
//		Key2: Integer;
//		Mix1: Integer;
//		Mix2: Integer;
//	end;
//
//	TSwitcherStatus = packed record
//    Connected: Boolean;
//		PGMMainVideo: Integer;
//		PGMMainAudio: Integer;
//		PGMBackupVideo: Integer;
//		PGMBackupAudio: Integer;
//		PGMKey1: Integer;
//		PGMKey2: Integer;
//		PGMMix1: Integer;
//		PGMMix2: Integer;
//		PGMAudioType: Integer;
//
//		PSTMainVideo: Integer;
//		PSTMainAudio: Integer;
//		PSTBackupVideo: Integer;
//		PSTBackupAudio: Integer;
//		PSTKey1: Integer;
//		PSTKey2: Integer;
//		PSTMix1: Integer;
//		PSTMix2: Integer;
//		PSTAudioType: Integer;
//
//		VideoTransType: Integer;
//		VideoTransRate: Integer;
//		AudioTransType: Integer;
//		AudioTransRate: Integer;
//
//		IsTransition: Boolean;
//		IsPreroll: Boolean;
//		IsRemoteOn: Boolean;
//
////		Dummy: array[0..99] of Byte;
//
//		SMPTETime: TDateTime;
//	end;
//
//  TPlayerStatus = packed record
//    Connected: Boolean;
//    // Port status
//    Stop: Boolean;        // Video server = Idle
////    Idle: Boolean;
//    Cue: Boolean;
//    Play: Boolean;
//    Rec: Boolean;
//    Still: Boolean;
//    Jog: Boolean;
//    Shuttle: Boolean;
//    VarMode: Boolean;
//    PortBusy: Boolean;
//    CueDone: Boolean;
//    FastFoward: Boolean;
//    Rewind: Boolean;
//    Eject: Boolean;
//
//    // System status
//{    TotalTC: TTimecode;
//    AvailableTC: TTimecode;
//    NumOfIDsStored: Integer;
//
//    DiskFull: Boolean;
//    SystemDown: Boolean;
//    DiskDown: Boolean;
//    RemoteControlDisabled: Boolean; }
//
//    IDsAdded: Boolean;
//    IDsDeleted: Boolean;
//    IDsAddedToPeerArch: Boolean;
//    NoTimeCode: Boolean;
//
//    // ErrorStatus
//    SystemError: Boolean;
//    IllegalValue: Boolean;
//    InvalidPort: Boolean;
//    WrongPortType: Boolean;
//    PortLocked: Boolean;
//    NotEnoughDiskSpace: Boolean;
//    CmdWhileBusy: Boolean;
//    NotSupported: Boolean;
//    InvalidID: Boolean;
//    IDNotFound: Boolean;
//    IDAleadyExists: Boolean;
//    IDStillRecording: Boolean;
//    IDCuedOrPlaying: Boolean;
//    XFerFailed: Boolean;
//    XFerComplete: Boolean;
//    IDDeleteProtected: Boolean;
//    NotInCueState: Boolean;
//    CueNotDone: Boolean;
//    PortNotIdle: Boolean;
//    PortActive: Boolean;
//    PortIdle: Boolean;
//    OperationFailed: Boolean;
//    SystemReboot: Boolean;
//
//    // VCR Status
//    // Data 0
//    Local: Boolean;
//    HeadError: Boolean;
//    TapeTrouble: Boolean;
//    RefVdMissing: Boolean;
//    CassetteOut: Boolean;
//    // Data 1
//    StandBy: Boolean;
//
//		{
//		Added fields for KeyLogo
//		bit 7 6 5 4 3 2 1 0 (bit order)
//		bit 7 : must be zero (since distinguish between ST_UNKNOWN(0xFF) and Normal State)
//		bit6~1 : XptNum of KeyLogo (Max 64)
//		bit 0 : On/Off status of KeyLogo
//		}
//	  KeyLogos: array [0..3] of BYTE;
//
////    Dummy: array[0..167] of BYTE;
//
//		nNumID: integer;
//		CurTC: TTimecode;
//		RemainTC: TTimecode;
//		DropFrame: Boolean;
//
//		ErrorCode: DWORD;
//		ErrorLine: DWORD;
//		ExtraErrorCode: DWORD;
//  end;
//
//	TGPIStatus = packed record
// 		Connected: Boolean;
//		SystemError: BYTE;
//		ErrorCode: DWORD;
//		ErrorLine: DWORD;
//		ExtraErrorCode: DWORD;
//	end;
//
//	TRSWStatus = packed record
//		Connected: Boolean;
//		ErrorCode: DWORD;
//		ErrorLine: DWORD;
//		ExtraErrorCode: DWORD;
//	end;
//
//	TDeviceStatus = packed record
//		EventType: TEventType;
//		case Integer of
//		0: (Switcher: TSwitcherStatus);
//		1: (Player: TPlayerStatus);
//		2: (GPI: TGPIStatus);
//		3: (RSW: TRSWStatus);
//	end;
//  PDeviceStatus = ^TDeviceStatus;
//
//  TMediaStatus = (msNone, msNotExist, msEqual, msShort, msLong);

const
	DCSDLL_NAME = 'DCSDLL.DLL';

function DCSSysInitialize(AInPort, AOutPort: Word; ATimeout: Cardinal = 1000): Integer; stdcall; external DCSDLL_NAME name 'DCSSysInitialize';
function DCSSysFinalize: Integer; stdcall; external DCSDLL_NAME name 'DCSSysFinalize';

function DCSSysLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external DCSDLL_NAME name 'DCSSysLogInPortEnable';
function DCSSysLogInPortDisable: Integer; stdcall; external DCSDLL_NAME name 'DCSSysLogInPortDisable';

function DCSSysLogOutPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external DCSDLL_NAME name 'DCSSysLogOutPortEnable';
function DCSSysLogOutPortDisable: Integer; stdcall; external DCSDLL_NAME name 'DCSSysLogOutPortDisable';

function DCSInitialize(ACallbackPort, AInPort, AOutPort: Integer; ATimeout: Cardinal = 1000; ABroadcast: Boolean = False): Integer; stdcall; external DCSDLL_NAME name 'DCSInitialize';
function DCSFinalize: Integer; stdcall; external DCSDLL_NAME name 'DCSFinalize';

function DCSLogNotifyPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external DCSDLL_NAME name 'DCSLogNotifyPortEnable';
function DCSLogNotifyPortDisable: Integer; stdcall; external DCSDLL_NAME name 'DCSLogNotifyPortDisable';

function DCSLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external DCSDLL_NAME name 'DCSLogInPortEnable';
function DCSLogInPortDisable: Integer; stdcall; external DCSDLL_NAME name 'DCSLogInPortDisable';

function DCSLogOutPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external DCSDLL_NAME name 'DCSLogOutPortEnable';
function DCSLogOutPortDisable: Integer; stdcall; external DCSDLL_NAME name 'DCSLogOutPortDisable';

function DCSDeviceStatusNotify(CallbackFunc: Pointer): Integer; stdcall; external DCSDLL_NAME name 'DCSDeviceStatusNotify';
function DCSEventStatusNotify(CallbackFunc: Pointer): Integer; stdcall; external DCSDLL_NAME name 'DCSEventStatusNotify';
function DCSEventOverallNotify(CallbackFunc: Pointer): Integer; stdcall; external DCSDLL_NAME name 'DCSEventOverallNotify';

// 0X00 System Control
function DCSSysIsAlive(AIP: PChar; var AIsAlive: Boolean): Integer; stdcall; external DCSDLL_NAME name 'DCSSysIsAlive';
function DCSSysIsMain(AIP: PChar; var AIsMain: Boolean): Integer; stdcall; external DCSDLL_NAME name 'DCSSysIsMain';

// 0X10 Device Control
function DCSOpen(AID: Word; AIP: PChar; ADeviceName: PChar; var AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSOpen';
function DCSClose(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSClose';
function DCSReset(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer; stdcall; external DCSDLL_NAME name 'DCSReset';
function DCSSetControlBy(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSSetControlBy';

// 0X20 Immediate Control
function DCSStop(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSStop';
function DCSPlay(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSPlay';
function DCSPause(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSPause';
function DCSContinue(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSContinue';
function DCSFastFoward(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSFastFoward';
function DCSFastRewind(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSFastRewind';
function DCSJog(AID: Word; AHandle: TDeviceHandle; AFrameOrSpeed: Double): Integer; stdcall; external DCSDLL_NAME name 'DCSJog';
function DCSShuttle(AID: Word; AHandle: TDeviceHandle; ASpeed: Double): Integer; stdcall; external DCSDLL_NAME name 'DCSShuttle';
function DCSStandBy(AID: Word; AHandle: TDeviceHandle; AOn: Boolean): Integer; stdcall; external DCSDLL_NAME name 'DCSStandBy';
function DCSEject(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSEject';
function DCSPreroll(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSPreroll';
function DCSRecord(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSRecord';
function DCSAutoEdit(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSAutoEdit';

// 0X30 Preset/Select Commands
function DCSSetPortMode(AID: Word; AHandle: TDeviceHandle; APortMode: TPortMode): Integer; stdcall; external DCSDLL_NAME name 'DCSSetPortMode';
function DCSSetAutoStatus(AID: Word; AHandle: TDeviceHandle; AAutoStatus: Boolean): Integer; stdcall; external DCSDLL_NAME name 'DCSSetAutoStatus';
function DCSPlayCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ACueTC, ADuration: TTimecode): Integer; stdcall; external DCSDLL_NAME name 'DCSPlayCue';
function DCSRecordCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ADuration: TTimecode): Integer; stdcall; external DCSDLL_NAME name 'DCSRecordCue';
function DCSIDRename(AID: Word; AHandle: TDeviceHandle; ASourceID, ATargetID: String): Integer; stdcall; external DCSDLL_NAME name 'DCSIDRename';
function DCSIDDelete(AID: Word; AHandle: TDeviceHandle; AMediaID: String): Integer; stdcall; external DCSDLL_NAME name 'DCSIDDelete';
function DCSInEntry(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSInEntry';
function DCSOutEntry(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSOutEntry';
function DCSAInEntry(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSAInEntry';
function DCSAOutEntry(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSAOutEntry';
function DCSInReset(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSInReset';
function DCSOutReset(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSOutReset';
function DCSAInReset(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSAInReset';
function DCSAOutReset(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSAOutReset';
function DCSEditPreset(AID: Word; AHandle: TDeviceHandle; AData1, AData2: Byte): Integer; stdcall; external DCSDLL_NAME name 'DCSEditPreset';
function DCSSetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl, AInput, AInputLvl: Integer): Integer; stdcall; external DCSDLL_NAME name 'DCSSetRoute';
function DCSSetXpt(AID: Word; AHandle: TDeviceHandle; AInput: Integer): Integer; stdcall; external DCSDLL_NAME name 'DCSSetXpt';

// 0X40 Sense Queries
function DCSGetDeviceStatus(AID: Word; AHandle: TDeviceHandle; var AStatus: TDeviceStatus): Integer; stdcall; external DCSDLL_NAME name 'DCSGetDeviceStatus';
function DCSGetStorageTimeRemaining(AID: Word; AHandle: TDeviceHandle; var ATotalTC, AAvailableTC: TTimecode; AExtended: Boolean = False): Integer; stdcall; external DCSDLL_NAME name 'DCSGetStorageTimeRemaining';
function DCSGetTC(AID: Word; AHandle: TDeviceHandle; var ACurTC: TTimecode): Integer; stdcall; external DCSDLL_NAME name 'DCSGetTC';
function DCSGetRemainTC(AID: Word; AHandle: TDeviceHandle; var ARemainTC: TTimecode): Integer; stdcall; external DCSDLL_NAME name 'DCSGetRemainTC';
function DCSGetList(AID: Word; AHandle: TDeviceHandle; var AMediaIDList: TIDList; var ARemainIDCount: Integer): Integer; stdcall; external DCSDLL_NAME name 'DCSGetList';
function DCSGetNext(AID: Word; AHandle: TDeviceHandle; var AMediaIDList: TIDList; var ARemainIDCount: Integer): Integer; stdcall; external DCSDLL_NAME name 'DCSGetNext';
function DCSGetExist(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var AExist: Boolean): Integer; stdcall; external DCSDLL_NAME name 'DCSGetExist';
function DCSGetSize(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var ADuration: TTimecode): Integer; stdcall; external DCSDLL_NAME name 'DCSGetSize';
function DCSGetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl: Integer; var AInput, AInputLvl: Integer): Integer; stdcall; external DCSDLL_NAME name 'DCSGetRoute';
function DCSGetXpt(AID: Word; AHandle: TDeviceHandle; var AInput: Integer): Integer; stdcall; external DCSDLL_NAME name 'DCSGetXpt';

// 0X50 Event Control
function DCSInputEvent(AID: Word; AHandle: TDeviceHandle; AEvent: TEvent): Integer; stdcall; external DCSDLL_NAME name 'DCSInputEvent';
function DCSDeleteEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer; stdcall; external DCSDLL_NAME name 'DCSDeleteEvent';
function DCSClearEvent(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer; stdcall; external DCSDLL_NAME name 'DCSClearEvent';
function DCSTakeEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; AStartTime: TEventTime): Integer; stdcall; external DCSDLL_NAME name 'DCSTakeEvent';
function DCSHoldEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer; stdcall; external DCSDLL_NAME name 'DCSHoldEvent';
function DCSChangetDurationEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; ADuration: TTimecode): Integer; stdcall; external DCSDLL_NAME name 'DCSChangetDurationEvent';
function DCSOnAirCatchEvent(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; external DCSDLL_NAME name 'DCSOnAirCatchEvent';
function DCSGetOnAirEventID(AID: Word; AHandle: TDeviceHandle; var AOnAirEventID, ANextEventID: TEventID): Integer; stdcall; external DCSDLL_NAME name 'DCSGetOnAirEventID';
function DCSGetEventInfo(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime; var ADurationTC: TTimecode): Integer; stdcall; external DCSDLL_NAME name 'DCSGetEventInfo';
function DCSGetEventStatus(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AEventStatus: TEventStatus): Integer; stdcall; external DCSDLL_NAME name 'DCSGetEventStatus';
function DCSGetEventStartTime(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime): Integer; stdcall; external DCSDLL_NAME name 'DCSGetEventStartTime';
function DCSGetEventOverall(AID: Word; AHandle: TDeviceHandle; var AEventOverall: TEventOverall): Integer; stdcall; external DCSDLL_NAME name 'DCSGetEventOverall';

implementation

end.

