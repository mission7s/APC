unit UnitCommons;

interface

uses Winapi.Windows, System.Classes, System.SysUtils, System.Types, Vcl.Forms,
  Vcl.Dialogs, Vcl.Graphics,
  Generics.Collections, UnitTypeConvert, UnitBaseSerial;

const
  D_OK    = 0;
  D_FALSE = -1;

  D_STX   = $02;
  D_ACK   = $04;
  D_NAK   = $05;
  D_ERR   = $06;

  E_NAK_UNDEFINED         = $01;
  E_NAK_SYNTAX            = $02;
  E_NAK_CHECKSUM          = $04;
  E_NAK_PARITY            = $10;
  E_NAK_OVERRUN           = $20;
  E_NAK_FRAMING           = $40;
  E_NAK_TIMEOUT           = $80;
  E_NAK_IGNORED           = $8F;

  E_TIMEOUT               = $F0;

  E_NO_ERROR = 0;

	{ General Error Code (10000 ~ 10099) }
	E_NOT_INITED = 10000;
	E_ALREADY_INITED = 10001;
	E_NOT_OPENED = 10002;
	E_INVALID_ARG = 10003;
	E_ALLOC_RESOURCE = 10004;
	E_NO_RESOURCE = 10005;
	E_SEMAPHORE = 10006;
	E_TIMEOUT_ERROR = 10007;
	E_INVALID_RES_CMD = 10008;
	E_NOT_TERMINATED_STRING = 10009;
	E_FILE_OPEN_ERROR = 10010;
	E_FILE_SIZE_ERROR = 10011;
	E_FILE_READ_ERROR = 10012;
	E_FILE_WRITE_ERROR = 10013;
	E_ENTER_SEMAPHOR_EERROR = 10014;
	E_SERIAL_OPEN_ERROR = 10015;
	E_SERIAL_SET_CONFIG_ERROR = 10016;
	E_EXCEPTION_ERROR = 10017;
	E_NOT_LIST_ID_MODE = 10018;
	E_SYSTEM_ERROR = 10019;
	E_SERIAL_CLEAR_BUFFER = 10020;
	E_INVALID_PARAMID = 10021;
	E_SET_PARAM_ERROR = 10022;
	E_GET_PARAM_ERROR = 10023;
	E_OPEN_ERROR = 10024;
	E_CLOS_EERROR = 10025;
	E_WRITE_ERROR = 10026;
	E_READ_ERROR = 10027;
	E_SEMAPHORERE_SERVED = 10028;
	E_DRIVER_BUSY = 10029;
	E_DEVIC_EERROR = 10030;

  E_INVALID_COMPONENT_ID = 10031;
  E_UNDEFIND_COMMAND = 10099;

	{ Event Error Code(10500 ~ 10599) }
	E_INVALID_DEF_SCHEDULE = 10500;
	E_DUPLICATED_EVENT = 10501;
	E_FULL_EVENT_QUEUE = 10502;
	E_EVENT_STATE_ERROR = 10503;
	E_DELETE_EVENT_FROM_QUEUE = 10504;
	E_TIMEOUT_CANCELED = 10505;
	E_EVENT_DUPLICATED = 10506;
	E_NOT_EXIST_IN_EVENT_QUEUE = 10507;
	E_GET_EVENT_STATUS_ERROR = 10508;
	E_DELETE_EVENT_ERROR = 10509;
	E_INPUT_EVENT_ERROR = 10510;
	E_HOLD_PGM_EVENT_ERROR = 10511;
	E_CHANGE_PGM_DUR_ERROR = 10512;
	E_CUE_EVENT_ERROR = 10513;
	E_CLEAR_EVENT_QUEUE_ERROR = 10514;
	E_INVALID_DEVICE = 10515;
	E_INVALID_SCHEDULE = 10516;
	E_TAKE_MANUAL_EVENT_ERROR = 10518;
	E_PGMN_OT_FINISHED = 10519;
	E_CHANGE_CONTROLLER = 10520;
	E_TRANS_ABORTED = 10521;
	E_OTHER_EVENT_ONAIR = 10522;
	E_INVALID_SOURCE = 10523;
	E_INVALID_EVENTID = 10524;
	E_TAKE_EVENT_ERROR = 10525;
	E_PAST_EVENT_ENTERED = 10526;
	E_SOURCE_CHANGED = 10527;
	E_SCHEDULE_ERROR_OUT = 10528;
	E_EVENTID_MISMATCHED = 10529;
	E_ALREADY_ONAIRED = 10530;
	E_ALREADY_TAKEN = 10531;

	{ Device Error Code(10100 ~ 10199) }
  E_INVALID_DEVICE_NAME = 10100;
  E_INVALID_START_TIME = 10101;

  E_INVALID_DEVICE_HANDLE = 10102;
  E_NOT_OPENED_DEVICE = 10103;


	ID_LEN = MAX_PATH;// 20;
	HOSTIP_LEN = 20;
	MEDIAID_LEN = MAX_PATH;//20;
	SECNAME_LEN = 20;
	DCSNAME_LEN = 20;
	DEVICENAME_LEN = 20;
	CHANNELNAME_LEN = 20;
  DATE_LEN = 8;
	BINNO_LEN = 20;
  TITLE_LEN = 80;
  SUBTITLE_LEN = 80;
  PROGRAMTYPENAME_LEN = 20;
  NOTES_LEN = 80;
	GPINAME_LEN = 20;

  EVENTTIME_STR_LEN = 20;
  TIMECODE_STR_LEN = 11;

	MAX_NUMID = 10;
  MAX_RETRY = 3;
  TIME_OUT = 1000;

  FrameRate23_98: Double = 23.98;
  FrameRate23_976: Double = 23.976;
  FrameRate29_97: Double = 29.97;
  FrameRate30_00: Double = 30.00;
  FrameRate59_94: Double = 59.94;
  FrameRate60_00: Double = 60.00;

  INVALID_DEVICE_HANDLE = -1;

type
  TLogState = (lsNormal, lsError, lsWarning);

{  TErrorCode = (
    ERR_NOERROR = 0,

    ERR_INVALID_ID = 10001,
    ERR_INVALID_DEVICE_NAME = 10002,
    ERR_INVALID_DEVICE_HANDLE = 10003,
    ERR_NOT_OPENED_DEVICE = 10004,
    ERR_UNDEFIND_COMMAND = 10005,

    ERR_INVALID_START_TIME = 11001
    ); }

	TDeviceHandle = Integer;
  TDeviceHandleList = TList<TDeviceHandle>;

  TPortMode = (pmDecoder, pmEncoder);

	TTimecode = DWORD;
	TEventTime = packed record
		T: TTimecode;
		D: TDate;
	end;

	TID = packed record
		BinNo: array[0..BINNO_LEN] of Char;
		ID: array[0..ID_LEN] of Char;
	end;

	TIDList = packed record
		NumID: Integer;
		ID: array[0..MAX_NUMID] of TID;
	end;

  TOnAirFlagType = (
    FT_REGULAR = Ord('r'),
    FT_SUBSTITUTE = Ord('s'),
    FT_EMERGENCY = Ord('e'),
    FT_LOG = Ord('l')
    );

  TEventID = packed record
    ChannelID: Word;//Byte;
    OnAirDate: array[0..DATE_LEN] of Char;// TDateTime;   // 방송일자
    OnAirFlag: TOnAirFlagType;  // R: 정규, S: 대체, E: 긴급, L: 운행로그
    OnAirNo: Integer;
    SerialNo: Integer;
  end;
  PEventID = ^TEventID;

  TFinishAction = (
	  FA_NONE = 0,
	  FA_STOP = 1,
	  FA_EJECT = 2{,
	  FA_REWIND = 3,
	  FA_REWINDEJECT = 4 }
    );

  TEventMode = (
    EM_PROGRAM  = $00,
    EM_MAIN     = $01,
    EM_JOIN     = $02,
    EM_SUB      = $03,
    EM_COMMENT  = $04
    );

  TStartMode = (
    SM_ABSOLUTE   = 0,
    SM_AUTOFOLLOW = 1,
    SM_MANUAL     = 2,
    SM_LOOP       = 3,
    SM_SUBBEGIN   = 4,
    SM_SUBEND     = 5
    );

  TInputType = (
	  IT_NONE     = $00,
    IT_MAIN     = $01,
    IT_BACKUP   = $02,
    IT_KEYER1   = $03,
    IT_KEYER2   = $04,
    IT_KEYER3   = $05,
    IT_KEYER4   = $06,
    IT_AMIXER1  = $07,
	  IT_AMIXER2  = $08
    );

  TOutputBkgndType = (
    OB_NONE  = $00,
    OB_VIDEO = $01,
    OB_AUDIO = $02,
    OB_BOTH  = $03
    );

  TOutputKeyerType = (
    OK_NONE = $00,
    OK_ON   = $01,
    OK_OFF  = $02,
    OK_AD   = $03
    );

  TVideoType = (
//    VT_NORMAL   = $00,
//    VT_WIDE     = $01
    VT_NONE = $00,
    VT_UHD  = $01,
    VT_HD   = $02,
    VT_SD   = $03,
    VT_3D   = $04
    );

  TAudioType = (
//	  AT_NONE     = $00,
//	  AT_LEFT     = $01,
//	  AT_RIGHT    = $02,
//	  AT_MONO     = $03,
//	  AT_STEREO   = $04
	  AT_NONE     = $00,
	  AT_5BY1     = $01,
	  AT_STEREO   = $02,
	  AT_MONO     = $03
    );

  TClosedCaption = (
	  CC_NONE     = $00,
	  CC_EXIST    = $01
    );

  TVoiceAdd = (
	  VA_NONE         = $00,
	  VA_KOREA_SCREEN = $01,
	  VA_ENG_VOICE    = $02,
	  VA_JPN_VOICE    = $03,
	  VA_CHN_VOICE    = $04,
	  VA_ETC_VOICE    = $05
    );

  TTRType = (
    TT_CUT      = 0,
    TT_FADE     = 1,
    TT_CUTFADE  = 2,
    TT_FADECUT  = 3,
    TT_MIX      = 4
    );

  TTRRate = (
    TR_CUT      = 0,
    TR_FAST     = 1,
    TR_MEDIUM   = 2,
    TR_SLOW     = 3
    );

  TSourceType = (
    ST_VSDEC,
    ST_VSENC,
    ST_VCR,
    ST_CART,
    ST_CG,
    ST_ROUTER,
    ST_MCS,
    ST_GPI,
    ST_LINE,
    ST_LOGO
    );

  TPlayerAction = (
	  PA_NONE = 0,
	  PA_PLAY = 1,
	  PA_RECORD = 2,
	  PA_STOP = 3,
	  PA_AD = 4
    );

	TPlayerEvent = packed record
		StartTC: TTimecode;
		PreStill: Integer;        // Unknown
		PostStill: Integer;       // Unknown
    Layer: Word;              // CG only
		PlayerAction: TPlayerAction;
		FinishAction: TFinishAction;
    ID: TID;
	end;

	TSwitcherEvent = packed record
		MainVideo: Integer;
		MainAudio: Integer;
		BackupVideo: Integer;
		BackupAudio: Integer;
		Key1: Integer;
		Key2: Integer;
		Key3: Integer;
		Key4: Integer;
		Mix1: Integer;
		Mix2: Integer;
		VideoType: Integer;
		AudioType: Integer;
		VideoTransType: TTRType;
		VideoTransRate: TTRRate;
		AudioTransType: TTRType;
		AudioTransRate: TTRRate;
		GPI: Integer;
	end;

	TGPIEvent = packed record
		PortNum: Integer;
		TriggerMode: Integer;
		AdjMillisec: Integer;
	end;

	TRSWEvent = packed record
		XptOut: Integer;
		XptIn: Integer;
    XptLevel: Integer;
		Stuffing: array[0..1] of Boolean;
	end;

  TEventType = (
    ET_SWITCHER = $00,
    ET_PLAYER   = $01,
    ET_GPI      = $02,
    ET_RSW      = $03
    );

  TEventState = (
    esIdle,
    esError,
    esLoading,
    esLoaded,
    esCueing,
    esCued,
    esStandByOff,
    esStandByOn,
    esPreroll,
    esOnAir,
    esFinish,
    esFinishing,
    esFinished,
    esDone,
    esSkipped);

  TEventStatus = record
    State: TEventState;
//    ErrorCode: TErrorCode;
    ErrorCode: Integer;
  end;
  PEventStatus = ^TEventStatus;

	TEventOverall = record
    ChannelID: Integer;

		NumEventInQueue: Integer;
		NumFreeEntryInQueue: Integer;

		Time: TTIME;

		LastTransitionTime: TTIME;
		LastAiredEventID: TEVENTID;
		OnAirEventID: TEVENTID;
		PreparedEventID: TEVENTID;

		IsHold: Boolean;

{		EventID: TEVENTID;
		ActionStartTime: TTIME;
		nStateCode: TEventState;
		nActionCode: BYTE;

		dwErrCode: DWORD;
		dwErrLine: DWORD;
		dwExtraErrCode: DWORD;

		EventType: BYTE;
		case integer of
		0: (Switcher: TSwitcherEvent);
		1: (Player: TPlayerEvent);
		2: (GPI: TGPIEvent);
		3: (RSW: TRSWEvent); }
	end;

	TEvent = packed record
		EventID: TEventID;

		StartTime: TEventTime;
		DurTime: TTimecode;

		TakeEvent: Boolean;   // Aleady Onair
		ManualEvent: Boolean;

    Status: TEventStatus;

		EventType: TEventType;
		case Integer of
		  0: (Switcher: TSwitcherEvent);
		  1: (Player: TPlayerEvent);
		  2: (GPI: TGPIEvent);
		  3: (RSW: TRSWEvent);
	end;
  PEvent = ^TEvent;
  TEventList = TList<PEvent>;

	TXPTInfo = packed record
		MainVideo: Integer;
		MainAudio: Integer;
		BackupVideo: Integer;
		BackupAudio: Integer;
		Key1: Integer;
		Key2: Integer;
		Mix1: Integer;
		Mix2: Integer;
	end;

	TSwitcherStatus = packed record
//    Connected: Boolean;
		PGMMainVideo: Integer;
		PGMMainAudio: Integer;
		PGMBackupVideo: Integer;
		PGMBackupAudio: Integer;
		PGMKey1: Integer;
		PGMKey2: Integer;
		PGMMix1: Integer;
		PGMMix2: Integer;
		PGMAudioType: Integer;

		PSTMainVideo: Integer;
		PSTMainAudio: Integer;
		PSTBackupVideo: Integer;
		PSTBackupAudio: Integer;
		PSTKey1: Integer;
		PSTKey2: Integer;
		PSTMix1: Integer;
		PSTMix2: Integer;
		PSTAudioType: Integer;

		VideoTransType: Integer;
		VideoTransRate: Integer;
		AudioTransType: Integer;
		AudioTransRate: Integer;

		IsTransition: Boolean;
		IsPreroll: Boolean;
		IsRemoteOn: Boolean;

//		Dummy: array[0..99] of Byte;

		SMPTETime: TDateTime;
	end;

  TPlayerStatus = packed record
//    Connected: Boolean;
    // Port status
    Stop: Boolean;        // Video server = Idle
//    Idle: Boolean;
    Cue: Boolean;
    Play: Boolean;
    Rec: Boolean;
    Still: Boolean;
    Jog: Boolean;
    Shuttle: Boolean;
    VarMode: Boolean;
    PortBusy: Boolean;
    CueDone: Boolean;
    FastFoward: Boolean;
    Rewind: Boolean;
    Eject: Boolean;

    // System status
{    TotalTC: TTimecode;
    AvailableTC: TTimecode;
    NumOfIDsStored: Integer;

    DiskFull: Boolean;
    SystemDown: Boolean;
    DiskDown: Boolean;
    RemoteControlDisabled: Boolean; }

    IDsAdded: Boolean;
    IDsDeleted: Boolean;
    IDsAddedToPeerArch: Boolean;
    NoTimeCode: Boolean;

    // ErrorStatus
    SystemError: Boolean;
    IllegalValue: Boolean;
    InvalidPort: Boolean;
    WrongPortType: Boolean;
    PortLocked: Boolean;
    NotEnoughDiskSpace: Boolean;
    CmdWhileBusy: Boolean;
    NotSupported: Boolean;
    InvalidID: Boolean;
    IDNotFound: Boolean;
    IDAleadyExists: Boolean;
    IDStillRecording: Boolean;
    IDCuedOrPlaying: Boolean;
    XFerFailed: Boolean;
    XFerComplete: Boolean;
    IDDeleteProtected: Boolean;
    NotInCueState: Boolean;
    CueNotDone: Boolean;
    PortNotIdle: Boolean;
    PortActive: Boolean;
    PortIdle: Boolean;
    OperationFailed: Boolean;
    SystemReboot: Boolean;

    // VCR Status
    // Data 0
    Local: Boolean;
    HeadError: Boolean;
    TapeTrouble: Boolean;
    RefVdMissing: Boolean;
    CassetteOut: Boolean;
    // Data 1
    StandBy: Boolean;

		{
		Added fields for KeyLogo
		bit 7 6 5 4 3 2 1 0 (bit order)
		bit 7 : must be zero (since distinguish between ST_UNKNOWN(0xFF) and Normal State)
		bit6~1 : XptNum of KeyLogo (Max 64)
		bit 0 : On/Off status of KeyLogo
		}
	  KeyLogos: array [0..3] of BYTE;

//    Dummy: array[0..167] of BYTE;

		nNumID: integer;
		CurTC: TTimecode;
		RemainTC: TTimecode;
		DropFrame: Boolean;
    PortDown: Boolean;

		ErrorCode: DWORD;
		ErrorLine: DWORD;
		ExtraErrorCode: DWORD;
  end;

	TGPIStatus = packed record
// 		Connected: Boolean;
		SystemError: BYTE;
		ErrorCode: DWORD;
		ErrorLine: DWORD;
		ExtraErrorCode: DWORD;
	end;

	TRSWStatus = packed record
//		Connected: Boolean;
		ErrorCode: DWORD;
		ErrorLine: DWORD;
		ExtraErrorCode: DWORD;
	end;

	TDeviceStatus = packed record
    Connected: Boolean;
		EventType: TEventType;
		case Integer of
		0: (Switcher: TSwitcherStatus);
		1: (Player: TPlayerStatus);
		2: (GPI: TGPIStatus);
		3: (RSW: TRSWStatus);
	end;
  PDeviceStatus = ^TDeviceStatus;

  TMediaStatus = (msNone, msNotExist, msEqual, msShort, msLong);

{	TEventStatus = record
		nNumEventInQueue: integer;
		nNumFreeEntryInQueue: integer;
		Time: TEventTime;

		LastTransitionTime: TEventTime;
		LastAiredEventID: TEventID;
		OnAirEventID: TEventID;
		PreparedEventID: TEventID;

		bHold: LongBool;

		EventID: TEventID;
		ActionStartTime: TEventTime;
		nStateCode: BYTE;
		nActionCode: BYTE;

		dwErrCode: DWORD;
		dwErrLine: DWORD;
		dwExtraErrCode: DWORD;

    Status: TStatus;
{		EventType: BYTE;
		case Integer of
		0: (Switcher: TSwitcherStatus);
		1: (Player: TPlayerStatus);
		2: (GPI: TGPIStatus);
		3: (RSW: TRSWStatus); }
//	end;

  TChannel = packed record
    ID: Word;
    Name: array[0..CHANNELNAME_LEN] of Char;
    OnAir: Boolean;
    NeedUpdate: Boolean;
  end;
  PChannel = ^TChannel;
  TChannelList = TList<PChannel>;

  TSEC = packed record
    ID: Word;
    Name: array[0..SECNAME_LEN] of Char;
    HostIP: array[0..HOSTIP_LEN] of Char;
    Main: Boolean;  // Main, Sub
//    Mine: Boolean;
    Alive: Boolean;
  end;
  PSEC = ^TSEC;
  TSECList = TList<PSEC>;

  TMCC = packed record
    ID: Word;
    Name: array[0..SECNAME_LEN] of Char;
    HostIP: array[0..HOSTIP_LEN] of Char;
//    Main: Boolean;  // Main, Sub
//    Mine: Boolean;
    Alive: Boolean;
  end;
  PMCC = ^TMCC;
  TMCCList = TList<PMCC>;

  TDCS = packed record
    ID: Word;
    Name: array[0..DCSNAME_LEN] of Char;
    HostIP: array[0..HOSTIP_LEN] of Char;
    Main: Boolean;  // Main, Sub
    Mine: Boolean;
    Alive: Boolean;
  end;
  PDCS = ^TDCS;
  TDCSList = TList<PDCS>;

{  TDCS = packed record
    Name: array[0..DEVICENAME_LEN] of Char;
    IP: array[0..MAXCHAR] of Char;
    IsMain: Boolean;
    IsSelf: Boolean;
  end;
  PDCS = ^TDCS;
  TDCSList = TList<TDCS>; }

  TPortConfig = packed record
    PortNum: Word;
    PortType: TComPortType;
    BaudRate: TComPortBaudRate;
    Parity: TComPortParity;
    DataBits: TComPortDataBits;
    StopBits: TComPortStopBits;
  end;
  PPortConfig = ^TPortConfig;

  TDeviceType = (
    DT_NONE = 0, // None

    // PCS
    DT_PCS_MEDIA,     // PCS encoder/decoder
    DT_PCS_CG,        // PCS cg
    DT_PCS_SWITCHER,  // PCS master switcher

    // Video server
    DT_LOUTH,     // Louth
    DT_OMNEON,    // Omneon Video Server

    // VCR
    DT_SBC,       // Sony betacam series

    // Graphic
    DT_K3D,       // Tornado karisma3d or tornado2 TCP
    DT_TAPI,      // Tornado TCP
    DT_NSC,       // Compix news scroll

    // Line source
    DT_LINE,      // Line source

    // Router & Master switcher
    DT_VTS,       // Video Tron Master Switcher
    DT_K2E_MCS,   // K2E Master Switcher

    DT_IMG_LRC,   // Imagine Logical Router Control
    DT_GV_RCL,    // GrassValley Router Control Language
    DT_QUARTZ,    // Quartz Routing Switcher Remote Control Protocol
    DT_UTHA_RCP3, // Utha RCP-3 Remote Switcher Control Protocols


    DT_GVM,       // GVGM2100, Ross MDK keyer
    DT_GVR,       // Grass valley router
    DT_VIK        // Vikin router
    );

  // PCS media device
  TDevicePCSMedia = packed record
    HostIP: array[0..HOSTIP_LEN] of Char;
    HostPort: Word;
  end;

  // PCS cg device
  TDevicePCSCG = packed record
    HostIP: array[0..HOSTIP_LEN] of Char;
    HostPort: Word;
  end;

  // PCS switcher device
  TDevicePCSSwitcher = packed record
    HostIP: array[0..HOSTIP_LEN] of Char;
    HostPort: Word;
  end;

  // Louth
  TDeviceLOUTH = packed record
    ControlType: TControlType;
    case TControlType of
      ctSerial: (
        PortConfig: TPortConfig;
        PortNo: Integer;
        );
      ctTCP, ctUDP: (
        HostIP: array[0..HOSTIP_LEN] of Char;
        HostPort: Word;
        );
  end;

  // Line Source
  TDeviceLINE = packed record
  end;

  // OMNEON video server
  TDeviceOMNEON = packed record
    DirectorName: array[0..MAXCHAR - 1] of Char;
    PlayerName: array[0..MAXCHAR - 1] of Char;
  end;

  // Tornado karisma3d or tornado2 TCP
  TDeviceK3D = packed record
    HostIP: array[0..HOSTIP_LEN] of Char;
    HostPort: Word;
    TemplatePath: array[0..MAXCHAR - 1] of Char;
  end;

  // Tornado TCP
  TDeviceTAPI = packed record
    HostIP: array[0..HOSTIP_LEN] of Char;
    HostPort: Word;
    TemplatePath: array[0..MAXCHAR - 1] of Char;
  end;

  // Grass Valley Router
  TDeviceGVR = packed record
    ControlType: TControlType;
    case TControlType of
      ctSerial: (
        PortConfig: TPortConfig;
        );
      ctTCP, ctUDP: (
        HostIP: array[0..HOSTIP_LEN] of Char;
        HostPort: Word;
        );
  end;

  // Video Tron Master Switcher
  TDeviceVTS = packed record
    PortConfig: TPortConfig;
    PresetDelayTake: TTimecode;
    InputChannels: Word;
    PGM: Word;
    PST: Word;
    SupportKey: Boolean;
    Key1: Word;
    Key2: Word;
    Key3: Word;
  end;

  // K2E Master Switcher
  TDeviceK2EMCS = packed record
    PortConfig: TPortConfig;
    InputChannels: Word;
    PGM: Word;
    PST: Word;
    SupportKey: Boolean;
    Key1: Word;
    Key2: Word;
  end;

  // Imagine Logical Router Control
  TDeviceImagineLRC = packed record
    ControlType: TControlType;
    case TControlType of
      ctSerial: (
        PortConfig: TPortConfig;
        );
      ctTCP, ctUDP: (
        HostIP: array[0..HOSTIP_LEN] of Char;
        HostPort: Word;
        );
  end;

  // GrassValley Router Control Language
  TDeviceGvRCL = packed record
    ControlType: TControlType;
    case TControlType of
      ctSerial: (
        PortConfig: TPortConfig;
        );
      ctTCP, ctUDP: (
        HostIP: array[0..HOSTIP_LEN] of Char;
        HostPort: Word;
        );
  end;

  // Quartz Routing Switcher Remote Control Protocol
  TDeviceQuartz = packed record
    ControlType: TControlType;
    case TControlType of
      ctSerial: (
        PortConfig: TPortConfig;
        );
      ctTCP, ctUDP: (
        HostIP: array[0..HOSTIP_LEN] of Char;
        HostPort: Word;
        );
  end;

  // Utha RCP3 Remote Switcher Control Protocols
  TDeviceUthaRCP3 = packed record
    ControlType: TControlType;
    case TControlType of
      ctSerial: (
        PortConfig: TPortConfig;
        );
      ctTCP, ctUDP: (
        HostIP: array[0..HOSTIP_LEN] of Char;
        HostPort: Word;
        );
  end;

	TDevice = packed record
    Handle: TDeviceHandle;
    Name: array[0..DEVICENAME_LEN] of Char;
    FrameDelay: Word;
    PortLog: Boolean;
    CheckStatus: Boolean;
    Status: TDeviceStatus;
		DeviceType: TDeviceType;
		case TDeviceType of
      DT_PCS_MEDIA: (PCSMedia: TDevicePCSMedia);
      DT_PCS_CG: (PCSCG: TDevicePCSCG);
      DT_PCS_SWITCHER: (PCSSwitcher: TDevicePCSSwitcher);

      DT_LOUTH: (Louth: TDeviceLOUTH);
		  DT_OMNEON: (Omneon: TDeviceOMNEON);

		  DT_LINE: (Line: TDeviceLINE);

		  DT_K3D: (K3D: TDeviceK3D);
		  DT_TAPI: (Tapi: TDeviceTAPI);

		  DT_VTS: (Vts: TDeviceVTS);
		  DT_K2E_MCS: (K2eMCS: TDeviceK2EMCS);

		  DT_IMG_LRC: (ImgLRC: TDeviceImagineLRC);
		  DT_GV_RCL: (GvRCL: TDeviceGvRCL);
		  DT_QUARTZ: (Quartz: TDeviceQuartz);
		  DT_UTHA_RCP3: (UthaRCP3: TDeviceUthaRCP3);

		  DT_GVR: (Gvr: TDeviceGVR);
	end;
  PDevice = ^TDevice;
  TDeviceList = TList<PDevice>;

  TXpt = packed record
    DeviceName: array[0..DEVICENAME_LEN] of Char;
    XptIn: Integer;
  end;
  PXpt = ^TXpt;
  TXptList = TList<PXpt>;

  TMCS = packed record
    ChannelID: Word;
    Name: array[0..DEVICENAME_LEN] of Char;
  end;
  PMCS = ^TMCS;
  TMCSList = TList<PMCS>;

  TSourceHandle = packed record
//    DCSID: Word;
//    DCSIP: array[0..HOSTIP_LEN] of Char;
    DCS: PDCS;
    Handle: TDeviceHandle;
//    Status: TDeviceStatus;
  end;
  PSourceHandle = ^TSourceHandle;
  TSourceHandleList = TList<PSourceHandle>;

  TSource = packed record
    Name: array[0..DEVICENAME_LEN] of Char;
//    Handle: TDeviceHandle;
    Channel: PChannel;
//    DCS: PDCS;
    Handles: TSourceHandleList;
    SourceType: TSourceType;
    MakeTransition: Boolean;
    SuccessiveUse: Boolean;
    LoadAlarm: Boolean;
    EjectAlarm: Boolean;
    QueryDuration: Boolean;
    PrerollTime: TTimecode;
    DefaultTimecode: TTimecode;
    UseDefaultTimecode: Boolean;
    UsePCS: Boolean;
    CommSuccess: Boolean; // Communication succees flag
    CommTimeout: Word;
    Status: TDeviceStatus;
		case TSourceType of
      ST_CG: (NumLayer: Word);
		  ST_ROUTER: (Router: TXptList);
  end;
  PSource = ^TSource;
  TSourceList = TList<PSource>;

  TCueSheetItem = packed record
		EventID: TEventID;

    DisplayNo: Word;
    ProgramNo: Word;
    GroupNo: Word;

    EventMode: TEventMode;
    StartMode: TStartMode;
		StartTime: TEventTime;
    EventStatus: TEventStatus;
    DeviceStatus: TDeviceStatus;

    Title: array[0..TITLE_LEN] of Char;
    SubTitle: array[0..SUBTITLE_LEN] of Char;

    Input: TInputType;
    Output: Byte;

    Source: array[0..DEVICENAME_LEN] of Char;
    SourceLayer: Word;
    MediaId: array[0..MEDIAID_LEN] of Char;
    MediaStatus: TMediaStatus;

		DurationTC: TTimecode;
		InTC: TTimecode;
		OutTC: TTimecode;

    VideoType: TVideoType;
    AudioType: TAudioType;
    ClosedCaption: TClosedCaption;
    VoiceAdd: TVoiceAdd;

    TransitionType: TTRType;
    TransitionRate: TTRRate;

    FinishAction: TFinishAction;

    ProgramType: Byte;

    BkColor: TColor;
    TxColor: TColor;
    ToColor: TColor;

    Notes: array[0..NOTES_LEN] of Char;
  end;
  PCueSheetItem = ^TCueSheetItem;
  TCueSheetList = TList<PCueSheetItem>;

{  TCueSheet = packed record
		EventID: TEventID;

    GroupIndex: Integer;
    CueSheetNo: Integer;
    EventMode: TEventMode;
    StartMode: TStartMode;
		StartTime: TEventTime;
    EventStatus: TEventStatus;
    Status: TDeviceStatus;

    InputType: TInputType;
    Output: Byte;
    Control: Byte;

		DurTime: TTimecode;
    Title: array[0..TITLE_LEN - 1] of Char;
    SubTitle: array[0..SUBTITLE_LEN - 1] of Char;
    SourceName: array[0..DEVICENAME_LEN - 1] of Char;
    ClipID: Integer;//array[0..MEDIAID_LEN - 1] of Char;
    MaterialID: array[0..ID_LEN - 1] of Char;
    TapeID: array[0..MEDIAID_LEN - 1] of Char;
    BinNo: array[0..BINNO_LEN - 1] of Char;
    StartTC: TTimecode;
    VideoType: TVideoType;
    AudioType: TAudioType;
    TransitionType: TTransitionType;
    TransitionRate: TTransitionRate;
    ProgramType: Byte;
    OldGradeType: Byte;
    FinishAction: TFinishAction;
    CueSheetDate: TDateTime;
    CueSheetDurTime: TTimecode;
  end;
  PCueSheet = ^TCueSheet; }

  TGPIType = (gtIn, gtOut);

  TGPI = packed record
    Name: array[0..GPINAME_LEN] of Char;
    PortNo: Integer;
    GPIType: TGPIType;
    FrameDelay: Cardinal;
  end;
  PGPI = ^TGPI;

  TEditMode = (
    EM_INSERT,
    EM_UPDATE,
    EM_DELETE
    );

  TTypeConvTime = packed record
     case vType: DWord of
        0: (vtDWord: DWord);
        1: (Frame, Second, Minute, Hour: Byte);
  end;

  TTypeConvDate = packed record
     case vType: DWord of
        0: (vtDWord: DWord);
        1: (Year: Word; Month, Day: Byte);
  end;

  function OnAirDateToDate(AOnAirDate: array of Char): TDate;

  function IsEqualEventID(AID1, AID2: TEventID): Boolean;
  function EventIDToString(AEventID: TEventID): String;

  function EventTimeToStadndardEventTime(AValue: TEventTime; AFrameRate: Double = 30): TEventTime;
  function EventTimeToDateTime(AValue: TEventTime; AFrameRate: Double = 30): TDateTime;
  function EventTimeToDateTimecodeStr(AValue: TEventTime; AApplyFormat: Boolean = False): String;
  function EventTimeToDate(AValue: TEventTime; AFrameRate: Double = 30): TDate;
  function EventTimeToTime(AValue: TEventTime; AFrameRate: Double = 30): TTime;
  function EventTimeToString(AValue: TEventTime): String;
  function StringToEventTime(AValue: String; AFrameRate: Double = 30): TEventTime;

  function DateTimeToEventTime(AValue: TDateTime; AFrameRate: Double = 30): TEventTime;
  function SystemTimeToEventTime(AValue: TSystemTime; AFrameRate: Double = 30): TEventTime;

//  function DateTimecodeStrToEventTime(AValue: String; AApplyFormat: Boolean = False; AFrameRate: Double = 30): TEventTime; overload;
  function DateTimecodeStrToEventTime(AValue: String; AApplyFormat: Boolean = False; AFrameRate: Double = 30): TEventTime; overload;
  function DateTimecodeStrToEventTime(ADateStr, ATimeStr: String; AApplyFormat: Boolean = False; AFrameRate: Double = 30): TEventTime; overload;

  function EventTimeToTimecode(AValue: TEventTime): TTimecode;
  function DateTimeToTimecode(AValue: TDateTime; AFrameRate: Double = 30): TTimecode;

  function CompareEventTime(AEventTime1, AEventTime2: TEventTime): Integer;

  function IsValidTimecode(AValue: TTimecode; AFrameRate: Double = 30): Boolean;

  function StringToTimecode(AValue: String): TTimecode;
  function TimecodeToString(AValue: TTimecode; AIncludeFrame: Boolean = True): String;
  function TimecodeToEventTime(AValue: TTimecode; AFrameRate: Double = 30): TEventTime;
  function TimecodeToMilliSec(AValue: TTimecode; AFrameRate: Double = 30): DWord;
  function TimecodeToSecond(AValue: TTimecode; AFrameRate: Double = 30): Double;
  function SecondToTimeCode(AValue: Double; AFrameRate: Double = 30): TTimecode;
  function TimecodeToFrame(AValue: TTimecode; AFrameRate: Double = 30): Cardinal;
  function FrameToTimecode(AFrames: Integer; AFrameRate: Double = 30): TTimecode;
  function FrameToMilliSec(AFrames: Integer; AFrameRate: Double = 30): DWord;

  // Begin 서브이벤트의 시작 이벤트시각을 구함
  function GetEventTimeSubBegin(AParentEventTime: TEventTime; AStartTC: TTimecode): TEventTime;
  // End 서브이벤트의 시작 이벤트시각을 구함
  function GetEventTimeSubEnd(AParentEventTime: TEventTime; AParentDurTC, AStartTC: TTimecode): TEventTime;

  function GetEventEndTime(AEventTime: TEventTime; ADurTC: TTimecode): TEventTime;

  function GetDurEventTime(AEventTime1, AEventTime2: TEventTime; AFrameRate: Double = 30): TEventTime;
  function GetPlusEventTime(AEventTime1, AEventTime2: TEventTime; AFrameRate: Double = 30): TEventTime;
  function GetMinusEventTime(AEventTime1, AEventTime2: TEventTime; AFrameRate: Double = 30): TEventTime;

  function GetTimecodePerDay: TTimecode;
  function GetDurTimecode(ATimecode1, ATimecode2: TTimecode; AFrameRate: Double = 30): TTimecode;
  function GetSubTimecode(ATimecode1, ATimecode2: TTimecode; AFrameRate: Double = 30): TTimecode;
  function GetPlusTimecode(ATimecode1, ATimecode2: TTimecode; AFrameRate: Double = 30): TTimecode;
  function GetMinusTimecode(ATimecode1, ATimecode2: TTimecode; AFrameRate: Double = 30): TTimecode;

  function GetDaysByTimecode(ATimecode: TTimecode): Integer;

  function GetFileVersionStr(AFileName: String): String;

  procedure SetEventDeadlineHour(AHour: Word);
  procedure PostKeyEx32(AKey: Word; const AShift: TShiftState; ASpecialKey: Boolean);

  function MakeNetConnection(APath, AUserId, APassword: String): Boolean;

const
  EventModeNames: array[EM_PROGRAM..EM_COMMENT] of String = (
    'Program',
    'Main',
    'Join',
    'Sub',
    'Comment'
    );

  EventModeShortNames: array[EM_PROGRAM..EM_COMMENT] of String = (
    'P',
    'M',
    'J',
    'S',
    'C'
    );

  StartModeNames: array[SM_ABSOLUTE..SM_SUBEND] of String = (
    'Absolute',
    'Auto Follow',
    'Manual',
    'Loop',
    'Sub Begin',
    'Sub End'
    );

  InputTypeNames: array[IT_NONE..IT_AMIXER2] of String = (
    'None',
    'Main',
    'Backup',
    'Keyer1',
    'Keyer2',
    'Keyer3',
    'Keyer4',
    'AudioMixer1',
    'AudioMixer2'
    );

  OutputBkgndTypeNames: array[OB_NONE..OB_BOTH] of String = (
    'None',
    'Video',
    'Audio',
    'Both'
    );

  OutputKeyerTypeNames: array[OK_NONE..OK_AD] of String = (
    'None',
    'On',
    'Off',
    'AD'
    );

  SourceTypeNames: array[ST_VSDEC..ST_LOGO] of String = (
    'VSDEC',
    'VSENC',
    'VCR',
    'CART',
    'CG',
    'ROUTER',
    'MCS',
    'GPI',
    'LINE',
    'LOGO'
    );

  VideoTypeNames: array[VT_NONE..VT_3D] of String = (
    'None',
    'UHD',
    'HD',
    'SD',
    '3D'
    );

  AudioTypeNames: array[AT_NONE..AT_MONO] of String = (
    'None',
    '5.1',
    'Stereo',
    'Mono'
    );

  ClosedCaptionNames: array[CC_NONE..CC_EXIST] of String = (
    'None',
    'Exist'
    );

  VoiceAddNames: array[VA_NONE..VA_ETC_VOICE] of String = (
    'None',
    'Korea Screen',
    'English',
    'Japanese',
    'Chinese',
    'ETC'
    );

  TRTypeNames: array[TT_CUT..TT_MIX] of String = (
    'Cut',
    'Fade',
    'Cut-Fade',
    'Fade-Cut',
    'Mix'
    );

  TRRateNames: array[TR_CUT..TR_SLOW] of String = (
    'Cut',
    'Fast',
    'Medium',
    'Slow'
    );

  FinishActionNames: array[FA_NONE..FA_EJECT] of String = (
    'None',
    'Stop',
    'Eject'
    );

  EventStatusNames: array[esIdle..esSkipped] of String = (
    '', //'Idle',
    'Error',
    'Loading',
    'Loaded',
    'Cueing',
    'Cued',
    'StandBy Off',
    'StandBy On',
    'Preroll',
    'OnAir',
    'Finish',
    'Finishing',
    'Finished',
    'Done',
    'Skipped'
    );

  EditModeNames: array[EM_INSERT..EM_DELETE] of String = (
    'Insert',
    'Update',
    'Delete'
    );

  MediaStatusNames: array[msNone..msLong] of String = (
    '', //'None',
    'Not Exist',
    'Equal',
    'Short',
    'Long'
    );

var
  EventDeadlineHour: Word = 24;

implementation

uses System.DateUtils;

function OnAirDateToDate(AOnAirDate: array of Char): TDate;
var
  YY: Word;
  MM: Word;
  DD: Word;
begin
  Result := 0;
  if (SizeOf(AOnAirDate) < DATE_LEN) then exit;

  YY := StrToIntDef(Copy(String(AOnAirDate), 1, 4), 0);
  MM := StrToIntDef(Copy(String(AOnAirDate), 5, 2), 1);
  DD := StrToIntDef(Copy(String(AOnAirDate), 7, 2), 1);

  if (YY = 0) then YY := 1;
  if (MM = 0) then MM := 1;
  if (DD = 0) then DD := 1;

  Result := EncodeDate(YY, MM, DD);
end;

function IsEqualEventID(AID1, AID2: TEventID): Boolean;
begin
  Result := (AID1.ChannelID = AID2.ChannelID) and
//            (SameDate(AID1.OnAirDate, AID2.OnAirDate)) and
            (StrComp(AID1.OnAirDate, AID2.OnAirDate) = 0) and
            (AID1.OnAirFlag = AID2.OnAirFlag) and
            (AID1.OnAirNo = AID2.OnAirNo) and
            (AID1.SerialNo = AID2.SerialNo);
end;

function EventIDToString(AEventID: TEventID): String;
begin
//  Result := Format('%d-%s-%s-%d-%d', [AEventID.ChannelID, FormatDateTime('YYYYMMDD', AEventID.OnAirDate),
  Result := Format('%d-%s-%s-%d-%d', [AEventID.ChannelID, AEventID.OnAirDate,
                                      Chr(Byte(AEventID.OnAirFlag)), AEventID.OnAirNo, AEventID.SerialNo]);
end;

function EventTimeToStadndardEventTime(AValue: TEventTime; AFrameRate: Double = 30): TEventTime;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  T.vtDWord := AValue.T;
//  D.vtDWord := AValue.D;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  Result.D := AValue.D;

  if (T.Frame >= RatePerFrame) then
  begin
    Inc(T.Second, T.Frame div RatePerFrame);
    T.Frame := T.Frame mod MSecsPerSec;
  end;

  if (T.Second >= SecsPerMin) then
  begin
    Inc(T.Minute, T.Second div SecsPerMin);
    T.Second := T.Second mod SecsPerMin;
  end;

  if (T.Minute >= MinsPerHour) then
  begin
    Inc(T.Hour, T.Minute div MinsPerHour);
    T.Minute := T.Minute mod MinsPerHour;
  end;

  if (T.Hour >= HoursPerDay) then
  begin
    Result.D := IncDay(Result.D, (T.Hour div HoursPerDay));
    T.Hour := T.Hour mod HoursPerDay;
  end;

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Hour <> 0) and (T.Minute <> 0) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;
  end;
  
  Result.T := T.vtDWord;
end;

function EventTimeToDateTime(AValue: TEventTime; AFrameRate: Double = 30): TDateTime;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  RatePerFrame: Word;
  DropFrame: Boolean;
  MS: Word;
begin
  T.vtDWord := AValue.T;
//  D.vtDWord := AValue.D;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Hour <> 0) and (T.Minute <> 0) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;

    if AFrameRate <> 0 then
      MS := Round(T.Frame / AFrameRate * 1000)
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second + (T.Frame / AFrameRate))
    else
      MS := 0;
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second);
  end
  else
  begin
    if RatePerFrame <> 0 then
      MS := Round(T.Frame / RatePerFrame * 1000)
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second + (T.Frame / RatePerFrame))
    else
      MS := 0;
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second);
  end;

  try
//  Result := EncodeDate(D.Year, D.Month, D.Day);
  Result := AValue.D;

  if (MS >= MSecsPerSec) then
  begin
    Inc(T.Second, MS div MSecsPerSec);
    MS := MS mod MSecsPerSec;
  end;

  if (T.Second >= SecsPerMin) then
  begin
    Inc(T.Minute, T.Second div SecsPerMin);
    T.Second := T.Second mod SecsPerMin;
  end;

  if (T.Minute >= MinsPerHour) then
  begin
    Inc(T.Hour, T.Minute div MinsPerHour);
    T.Minute := T.Minute mod MinsPerHour;
  end;

  if (T.Hour >= HoursPerDay) then
  begin
    Result := IncDay(Result, (T.Hour div HoursPerDay));
    T.Hour := T.Hour mod HoursPerDay;
  end;

  ReplaceTime(Result, EncodeTime(T.Hour, T.Minute, T.Second, MS));
//  Result := EncodeDate(D.Year, D.Month, D.Day, T.Hour, T.Minute, T.Second, MS);
  except
  end;
end;

function EventTimeToDateTimecodeStr(AValue: TEventTime; AApplyFormat: Boolean = False): String;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  RatePerFrame: Word;
  DropFrame: Boolean;
  MS: Word;
begin
  T.vtDWord := AValue.T;
{  D.vtDWord := AValue.D;

  if (AApplyFormat) then
    Result := Format('%d%s%.2d%s%.2d %.2d:%.2d:%.2d:%.2d', [D.Year, FormatSettings.DateSeparator,
                                                            D.Month, FormatSettings.DateSeparator,
                                                            D.Day,
                                                            T.Hour, T.Minute, T.Second, T.Frame])
  else
    Result := Format('%.4d%.2d%.2d %.2d%.2d%.2d%.2d', [D.Year, D.Month, D.Day,
                                                       T.Hour, T.Minute, T.Second, T.Frame]);
}

  if (AApplyFormat) then
    Result := FormatDateTime('YYYY' + FormatSettings.DateSeparator + 'MM' + FormatSettings.DateSeparator + 'DD', AValue.D) + ' ' +
              Format('%.2d:%.2d:%.2d:%.2d', [T.Hour, T.Minute, T.Second, T.Frame])
  else
    Result := FormatDateTime('YYYYMMDD', AValue.D) + ' ' +
              Format('%.2d:%.2d:%.2d:%.2d', [T.Hour, T.Minute, T.Second, T.Frame]);
end;

function EventTimeToDate(AValue: TEventTime; AFrameRate: Double = 30): TDate;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  RatePerFrame: Word;
  DropFrame: Boolean;
  MS: Word;
begin
  T.vtDWord := AValue.T;
//  D.vtDWord := AValue.D;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Hour <> 0) and (T.Minute <> 0) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;

    if AFrameRate <> 0 then
      MS := Round(T.Frame / AFrameRate * 1000)
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second + (T.Frame / AFrameRate))
    else
      MS := 0;
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second);
  end
  else
  begin
    if RatePerFrame <> 0 then
      MS := Round(T.Frame / RatePerFrame * 1000)
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second + (T.Frame / RatePerFrame))
    else
      MS := 0;
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second);
  end;

  try
//  Result := EncodeDate(D.Year, D.Month, D.Day);
  Result := AValue.D;

  if (MS >= MSecsPerSec) then
  begin
    Inc(T.Second, MS div MSecsPerSec);
    MS := MS mod MSecsPerSec;
  end;

  if (T.Second >= SecsPerMin) then
  begin
    Inc(T.Minute, T.Second div SecsPerMin);
    T.Second := T.Second mod SecsPerMin;
  end;

  if (T.Minute >= MinsPerHour) then
  begin
    Inc(T.Hour, T.Minute div MinsPerHour);
    T.Minute := T.Minute mod MinsPerHour;
  end;

  if (T.Hour >= HoursPerDay) then
  begin
    Result := IncDay(Result, (T.Hour div HoursPerDay));
    T.Hour := T.Hour mod HoursPerDay;
  end;

  except
  end;
end;

function EventTimeToTime(AValue: TEventTime; AFrameRate: Double = 30): TTime;
var
  T: TTypeConvTime;
  RatePerFrame: Word;
  DropFrame: Boolean;
  MS: Word;
begin
  T.vtDWord := AValue.T;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Hour <> 0) and (T.Minute <> 0) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;

    if AFrameRate <> 0 then
      MS := Round(T.Frame / AFrameRate * 1000)
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second + (T.Frame / AFrameRate))
    else
      MS := 0;
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second);
  end
  else
  begin
    if RatePerFrame <> 0 then
      MS := Round(T.Frame / RatePerFrame * 1000)
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second + (T.Frame / RatePerFrame))
    else
      MS := 0;
//      Result := (T.Hour * (3600) + T.Minute * (60) + T.Second);
  end;

  try
  if (MS >= MSecsPerSec) then
  begin
    Inc(T.Second, MS div MSecsPerSec);
    MS := MS mod MSecsPerSec;
  end;

  if (T.Second >= SecsPerMin) then
  begin
    Inc(T.Minute, T.Second div SecsPerMin);
    T.Second := T.Second mod SecsPerMin;
  end;

  if (T.Minute >= MinsPerHour) then
  begin
    Inc(T.Hour, T.Minute div MinsPerHour);
    T.Minute := T.Minute mod MinsPerHour;
  end;

  if (T.Hour >= HoursPerDay) then
  begin
//    Result := IncDay(Result, T.Hour div HoursPerDay);
    T.Hour := T.Hour mod HoursPerDay;
  end;

  Result := EncodeTime(T.Hour, T.Minute, T.Second, MS);
  except
  end;
end;

function EventTimeToString(AValue: TEventTime): String;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  RatePerFrame: Word;
  DropFrame: Boolean;
  MS: Word;
begin
  T.vtDWord := AValue.T;

  Result := FormatDateTime('YYYYMMDD', AValue.D) + ' ' +
            Format('%.2d:%.2d:%.2d:%.2d', [T.Hour, T.Minute, T.Second, T.Frame]);
end;

function StringToEventTime(AValue: String; AFrameRate: Double = 30): TEventTime;
var
  T: TTypeConvTime;
  YY, MM, DD: Word;
  HH, MI, SS, FF: Word;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  if (Length(AValue) < 20) then
  begin
    Result.D := 0;
    Result.T := 0;
    exit;
  end;

  YY := StrToIntDef(Copy(AValue, 1, 4), 1899);
  MM := StrToIntDef(Copy(AValue, 5, 2), 12);
  DD := StrToIntDef(Copy(AValue, 7, 2), 30);

  HH := StrToIntDef(Copy(AValue, 10, 2), 0);
  MI := StrToIntDef(Copy(AValue, 13, 2), 0);
  SS := StrToIntDef(Copy(AValue, 16, 2), 0);
  FF := StrToIntDef(Copy(AValue, 19, 2), 0);

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  T.Hour    := HH;
  T.Minute  := MI;
  T.Second  := SS;
  T.Frame   := FF;

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;
  end;

  Result.T := T.vtDWord;
  Result.D := EncodeDate(YY, MM, DD);
end;

function DateTimeToEventTime(AValue: TDateTime; AFrameRate: Double = 30): TEventTime;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  YY, MM, DD: Word;
  HH, MI, SS, MS: Word;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  DecodeTime(AValue, HH, MI, SS, MS);

  T.Hour    := HH;
  T.Minute  := MI;
  T.Second  := SS;
  T.Frame   := Round(MS * AFrameRate / 1000);

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;
  end;

  DecodeDate(AValue, YY, MM, DD);
//  D.Year  := YY;
//  D.Month := MM;
//  D.Day   := DD;

  Result.T := T.vtDWord;
  Result.D := EncodeDate(YY, MM, DD);
end;

function SystemTimeToEventTime(AValue: TSystemTime; AFrameRate: Double = 30): TEventTime;
begin
  Result := DateTimeToEventTime(SystemTimeToDateTime(AValue), AFrameRate);
end;

function DateTimecodeStrToEventTime(AValue: String; AApplyFormat: Boolean = False; AFrameRate: Double = 30): TEventTime;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  YY, MM, DD: Word;
  HH, MI, SS, FF: Word;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  if (AApplyFormat) then
  begin
    if (Length(AValue) < 22) then
    begin
      Result.D := 0;
      Result.T := 0;
      exit;
    end;

    YY := StrToIntDef(Copy(AValue, 1, 4), 0);
    MM := StrToIntDef(Copy(AValue, 6, 2), 1);
    DD := StrToIntDef(Copy(AValue, 9, 2), 1);

    HH := StrToIntDef(Copy(AValue, 12, 2), 0);
    MI := StrToIntDef(Copy(AValue, 15, 2), 0);
    SS := StrToIntDef(Copy(AValue, 18, 2), 0);
    FF := StrToIntDef(Copy(AValue, 21, 2), 0);
  end
  else
  begin
    if (Length(AValue) < 20) then
    begin
      Result.D := 0;
      Result.T := 0;
      exit;
    end;

    YY := StrToIntDef(Copy(AValue, 1, 4), 1899);
    MM := StrToIntDef(Copy(AValue, 5, 2), 12);
    DD := StrToIntDef(Copy(AValue, 7, 2), 30);

    HH := StrToIntDef(Copy(AValue, 10, 2), 0);
    MI := StrToIntDef(Copy(AValue, 13, 2), 0);
    SS := StrToIntDef(Copy(AValue, 16, 2), 0);
    FF := StrToIntDef(Copy(AValue, 19, 2), 0);
  end;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  T.Hour    := HH;
  T.Minute  := MI;
  T.Second  := SS;
  T.Frame   := FF;

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;
  end;

{  D.Year  := YY;
  D.Month := MM;
  D.Day   := DD; }

  Result.T := T.vtDWord;
//  Result.D := D.vtDWord;
  Result.D := EncodeDate(YY, MM, DD);
end;

function DateTimecodeStrToEventTime(ADateStr, ATimeStr: String; AApplyFormat: Boolean = False; AFrameRate: Double = 30): TEventTime;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  YY, MM, DD: Word;
  HH, MI, SS, FF: Word;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  if ((AApplyFormat) and (Length(ADateStr) < 10)) or
     ((not AApplyFormat) and (Length(ADateStr) < 8)) then
  begin
    YY := 1899;
    MM := 12;
    DD := 30;
  end
  else
  begin
    YY := StrToIntDef(Copy(ADateStr, 1, 4), 1899);
    MM := StrToIntDef(Copy(ADateStr, 6, 2), 12);
    DD := StrToIntDef(Copy(ADateStr, 9, 2), 30);
  end;

  if (Length(ATimeStr) < 11) then
  begin
    HH := 0;
    MI := 0;
    SS := 0;
    FF := 0;
  end
  else
  begin
    HH := StrToIntDef(Copy(ATimeStr, 1, 2), 0);
    MI := StrToIntDef(Copy(ATimeStr, 4, 2), 0);
    SS := StrToIntDef(Copy(ATimeStr, 7, 2), 0);
    FF := StrToIntDef(Copy(ATimeStr, 10, 2), 0);
  end;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  T.Hour    := HH;
  T.Minute  := MI;
  T.Second  := SS;
  T.Frame   := FF;

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;
  end;

{  D.Year  := YY;
  D.Month := MM;
  D.Day   := DD; }

  Result.T := T.vtDWord;
//  Result.D := D.vtDWord;
  Result.D := EncodeDate(YY, MM, DD);
end;

// YYYYMMDDHHNNSSFF
{function DateTimecodeStrToEventTime(AValue: String; AApplyFormat: Boolean = False; AFrameRate: Double = 30): TEventTime;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  YY, MM, DD: Word;
  HH, MI, SS, FF: Word;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  if (AApplyFormat) then
  begin
    if (Length(AValue) < 22) then
    begin
      Result.T := 0;
      Result.D := 0;
      exit;
    end;

    YY := StrToIntDef(Copy(AValue, 1, 4), 0);
    MM := StrToIntDef(Copy(AValue, 6, 2), 1);
    DD := StrToIntDef(Copy(AValue, 9, 2), 1);

    HH := StrToIntDef(Copy(AValue, 12, 2), 12);
    MI := StrToIntDef(Copy(AValue, 15, 2), 0);
    SS := StrToIntDef(Copy(AValue, 18, 2), 0);
    FF := StrToIntDef(Copy(AValue, 21, 2), 0);
  end
  else
  begin
    if (Length(AValue) < 16) then
    begin
      Result.T := 0;
      Result.D := 0;
      exit;
    end;

    YY := StrToIntDef(Copy(AValue, 1, 4), 0);
    MM := StrToIntDef(Copy(AValue, 5, 2), 1);
    DD := StrToIntDef(Copy(AValue, 7, 2), 1);

    HH := StrToIntDef(Copy(AValue, 9, 2), 12);
    MI := StrToIntDef(Copy(AValue, 11, 2), 0);
    SS := StrToIntDef(Copy(AValue, 13, 2), 0);
    FF := StrToIntDef(Copy(AValue, 15, 2), 0);
  end;

  DropFrame := (AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  T.Hour    := HH;
  T.Minute  := MI;
  T.Second  := SS;
  T.Frame   := FF;

  if (DropFrame) then
  begin
    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;
  end;

  D.Year  := YY;
  D.Month := MM;
  D.Day   := DD;

  Result.T := T.vtDWord;
  Result.D := D.vtDWord;
end; }

function EventTimeToTimecode(AValue: TEventTime): TTimecode;
var
  T: TTypeConvTime;
  D: TTypeConvDate;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  T.vtDWord := AValue.T;
//  D.vtDWord := AValue.D;

  Result := AValue.T;
end;

function DateTimeToTimecode(AValue: TDateTime; AFrameRate: Double = 30): TTimecode;
var
  T: TTypeConvTime;
  HH, MI, SS, MS: Word;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  DecodeTime(AValue, HH, MI, SS, MS);

  T.Hour    := HH;
  T.Minute  := MI;
  T.Second  := SS;
  T.Frame   := Trunc(MS * (AFrameRate / 1000));

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;
  end;

  Result := T.vtDWord;
end;

function CompareEventTime(AEventTime1, AEventTime2: TEventTime): Integer;
var
  E1, E2: TEventTime;
begin
  E1 := EventTimeToStadndardEventTime(AEventTime1);
  E2 := EventTimeToStadndardEventTime(AEventTime2);

  Result := CompareDate(E1.D, E2.D);
  if (Result = EqualsValue) then
    if (E1.T = E2.T) then
      Result := EqualsValue
    else if (E1.T < E2.T) then
      Result := LessThanValue
    else
      Result := GreaterThanValue;
end;

function IsValidTimecode(AValue: TTimecode; AFrameRate: Double = 30): Boolean;
var
  T: TTypeConvTime;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  T.vtDWord := AValue;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  Result := False;

  if (T.Minute < 0) or (T.Minute >= MinsPerHour) then exit
  else if (T.Second < 0) or (T.Second >= SecsPerMin) then exit
  else if (T.Frame < 0) or (T.Frame >= RatePerFrame) then exit
  else if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) and (T.Frame < 1) then exit
    else if (AFrameRate = FrameRate23_976) and (T.Hour <> 0) and (T.Minute <> 0) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) and (T.Frame < 2) then exit
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) and (T.Frame < 2) then exit
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) and (T.Frame < 4) then exit;
  end;

  Result := True;
end;

function StringToTimecode(AValue: String): TTimecode;
var
  T: TTypeConvTime;
  HH, MI, SS, FF: Integer;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  Result := 0;
  if Length(AValue) < 11 then exit;

{  AValue := StringReplace(AValue, ' ', '', []);

  T.Hour    := StrToIntDef(Copy(AValue, 1, Pos(':', AValue) - 1), 0);
  AValue    := Copy(AValue, Pos(':', AValue) + 1, Length(AValue));

  T.Minute  := StrToIntDef(Copy(AValue, 1, Pos(':', AValue) - 1), 0);
  AValue    := Copy(AValue, Pos(':', AValue) + 1, Length(AValue));

  T.Second  := StrToIntDef(Copy(AValue, 1, Pos(':', AValue) - 1), 0);
  AValue    := Copy(AValue, Pos(':', AValue) + 1, Length(AValue));

  T.Frame   := StrToIntDef(AValue, 0); }


  T.Hour    := StrToIntDef(Trim(Copy(AValue, 1, 2)), 0);
  T.Minute  := StrToIntDef(Trim(Copy(AValue, 4, 2)), 0);
  T.Second  := StrToIntDef(Trim(Copy(AValue, 7, 2)), 0);
  T.Frame   := StrToIntDef(Trim(Copy(AValue, 10, 2)), 0);

  Result := T.vtDWord;
end;

function TimecodeToString(AValue: TTimecode; AIncludeFrame: Boolean = True): String;
const
  INIT_TIMECODE = '00:00:00:00';
var
  T: TTypeConvTime;
begin
  Result := INIT_TIMECODE;

  T.vtDWord := AValue;
  if (AIncludeFrame) then
    Result := Format('%.2d:%.2d:%.2d:%.2d', [T.Hour, T.Minute, T.Second, T.Frame])
  else
    Result := Format('%.2d:%.2d:%.2d', [T.Hour, T.Minute, T.Second])
end;

function TimecodeToEventTime(AValue: TTimecode; AFrameRate: Double = 30): TEventTime;
begin
  Result.D := 0;
  Result.T := AValue;
end;

function TimecodeToMilliSec(AValue: TTimecode; AFrameRate: Double = 30): DWord;
var
  T: TTypeConvTime;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  Result := Round(TimecodeToFrame(AValue, AFrameRate) * (1 / AFrameRate) * 1000);
  exit;

  T.vtDWord := AValue;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Hour <> 0) and (T.Minute <> 0) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;

    if AFrameRate <> 0 then
      Result := Trunc((T.Hour * (3600) + T.Minute * (60) + T.Second + (T.Frame / AFrameRate)) * 1000)
    else
      Result := Trunc((T.Hour * (3600) + T.Minute * (60) + T.Second) * 1000);
  end
  else
  begin
    if RatePerFrame <> 0 then
      Result := Trunc((T.Hour * (3600) + T.Minute * (60) + T.Second + (T.Frame / RatePerFrame)) * 1000)
    else
      Result := Trunc((T.Hour * (3600) + T.Minute * (60) + T.Second) * 1000);
  end;
end;

function TimecodeToSecond(AValue: TTimecode; AFrameRate: Double = 30): Double;
begin
  Result := TimecodeToFrame(AValue, AFrameRate) * (1 / AFrameRate);
end;

function SecondToTimeCode(AValue: Double; AFrameRate: Double = 30): TTimecode;
var
  T: TTypeConvTime;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  Result := FrameToTimeCode(Round(AValue * AFrameRate), AFrameRate);
  exit;

  T.Hour    := Trunc(AValue / 3600);
  T.Minute  := Trunc(AValue / 60);
  T.Second  := Trunc(AValue);
  T.Frame   := Trunc((AValue - Trunc(AValue)) * AFrameRate);

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0) then T.Frame := 1
    else if (AFrameRate = FrameRate23_976) and (T.Hour <> 0) and (T.Minute <> 0) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) = 0) then T.Frame := 2
    else }if (AFrameRate = FrameRate29_97) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 2
    else if (AFrameRate = FrameRate59_94) and (T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 10) <> 0) then T.Frame := 4;
  end;

  Result := T.vtDWord;
end;

function TimecodeToFrame(AValue: TTimecode; AFrameRate: Double = 30): Cardinal;
var
  T: TTypeConvTime;
  RatePerFrame: Word;
  DropFrame: Boolean;
begin
  Result := 0;

  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);
  if RatePerFrame <= 0 then RatePerFrame := 1;

  T.vtDWord := AValue;

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) then
    begin
      if ((T.Frame = 0) and (T.Second = 0) and ((T.Minute mod 5) <> 0)) then
        T.Frame := 1
      else if ((T.Minute <> 0) and (T.Frame in [0..1]) and (T.Second = 0) and ((T.Minute mod 5) = 0)) then
        T.Frame := 2
      else if (T.Frame >= 24) then
        T.Frame := 23;

      Result := (T.Hour * 86328) + ((T.Minute * 1440) - ((T.Minute - (T.Minute div 5)) + (T.Minute div 5) * 2)) + (T.Second * RatePerFrame) + T.Frame;
    end
    else }if (AFrameRate = FrameRate29_97) then
    begin
      if ((T.Frame in [0..1]) and (T.Second = 0) and ((T.Minute mod 10) <> 0)) then
        T.Frame := 2
      else if (T.Frame >= 30) then
        T.Frame := 29;

      Result := (T.Hour * 107892) + ((T.Minute * 1800) - ((T.Minute - (T.Minute div 10)) * 2)) + (T.Second * RatePerFrame) + T.Frame;
    end
    else if (AFrameRate = FrameRate59_94) then
    begin
      if ((T.Frame in [0..3]) and (T.Second = 0) and ((T.Minute mod 10) <> 0)) then
        T.Frame := 4
      else if (T.Frame >= 60) then
        T.Frame := 59;

      Result := (T.Hour * 215784) + ((T.Minute * 3600) - ((T.Minute - (T.Minute div 10)) * 4)) + (T.Second * RatePerFrame) + T.Frame;
    end;
  end
  else
  begin
    if (T.Frame >= RatePerFrame) then
      T.Frame := RatePerFrame - 1;

    if (AFrameRate = FrameRate23_98) or
       (AFrameRate = FrameRate23_976) then
      Result := Round(((T.Hour * 3600) + (T.Minute * 60) + T.Second) * RatePerFrame) + T.Frame
    else
      Result := Round(((T.Hour * 3600) + (T.Minute * 60) + T.Second) * AFrameRate) + T.Frame;//Format('%.2d:%.2d:%.2d:%.2d', [HH, MI, SS, FF]);
  end;
end;

function FrameToTimecode(AFrames: Integer; AFrameRate: Double): TTimecode;
var
  HH, MI, SS, FF: Integer;
  F: Integer;
  RatePerFrame: Word;
  DropFrame: Boolean;
  SampleTime: Double;
  T: TTypeConvTime;
begin
  DropFrame := //(AFrameRate = FrameRate23_976) or
               (AFrameRate = FrameRate29_97) or
               (AFrameRate = FrameRate59_94);

  RatePerFrame := Round(AFrameRate);
  if RatePerFrame <= 0 then RatePerFrame := 1;

  F := AFrames;

  if (F <= 0) then
  begin
    Result := 0;
    exit;
  end;

  if (DropFrame) then
  begin
{    if (AFrameRate = FrameRate23_976) then
    begin
      HH := (F div 86328);
      F  := (F mod 86328);
      MI := (F div 1440);
      F  := (F mod 1440) + (((MI - (MI div 10)) + (MI div 10) * 2));
      if (F >= 1440) then
      begin
        Inc(MI);
        Dec(F, 1440);
        if (MI >= 60) then
        begin
          Inc(HH);
          Dec(MI, 60);
        end;

        if ((MI mod 5) <> 0) then F := F + 1
        else if (MI <> 0) and ((MI mod 5) = 0) then F := F + 2;
      end;
    end
    else }if (AFrameRate = FrameRate29_97) then
    begin
      HH := (F div 107892);
      F  := (F mod 107892);
      MI := (F div 1800);
      F  := (F mod 1800) + (((MI - (MI div 10)) * 2));
      if (F >= 1800) then
      begin
        Inc(MI);
        Dec(F, 1800);
        if (MI >= 60) then
        begin
          Inc(HH);
          Dec(MI, 60);
        end;

        if ((MI mod 10) <> 0) then F := F + 2;
      end;
    end
    else if (AFrameRate = FrameRate59_94) then
    begin
      HH := (F div 215784);
      F  := (F mod 215784);
      MI := (F div 3600);
      F  := (F mod 3600) + (((MI - (MI div 10)) * 4));
      if (F >= 3600) then
      begin
        Inc(MI);
        Dec(F, 3600);
        if (MI >= 60) then
        begin
          Inc(HH);
          Dec(MI, 60);
        end;

        if ((MI mod 10) <> 0) then F := F + 4;
      end;
    end;

    SS := (F div RatePerFrame);
    F  := (F mod RatePerFrame);
    FF := (F);
  end
  else
  begin
    if (AFrameRate = FrameRate23_98) or
       (AFrameRate = FrameRate23_976) then
      SampleTime := F / RatePerFrame
    else
      SampleTime := F / AFrameRate;

    HH := (Trunc(SampleTime) div 3600);
    SampleTime := (SampleTime - (HH * 3600));
    MI := (Trunc(SampleTime) div 60);
    SampleTime := (SampleTime - (MI * 60));
    SS := (Trunc(SampleTime));
    SampleTime := (SampleTime - SS);
    FF := Round(SampleTime * RatePerFrame);

    if (FF >= RatePerFrame) then
    begin
      FF := 0;
      Inc(SS);

      if (SS >= 60) then
      begin
        Inc(MI);
        SS := 0;
      end;

      if (MI >= 60) then
      begin
        Inc(HH);
        MI := 0;
      end;
    end;
  end;

  T.Frame  := FF;
  T.Second := SS;
  T.Minute := MI;
  T.Hour   := HH;

  Result := T.vtDWord;
end;

function FrameToMilliSec(AFrames: Integer; AFrameRate: Double = 30): DWord;
var
  T: TTypeConvTime;
begin
  T.vtDWord := FrameToTimecode(AFrames, AFrameRate);
  Result := TimecodeToMilliSec(T.vtDWord, AFrameRate);
end;

function GetEventTimeSubBegin(AParentEventTime: TEventTime; AStartTC: TTimecode): TEventTime;
var
  Days: Integer;
  T: TTypeConvTime;
begin
  Result := AParentEventTime;
  Result.T := GetPlusTimecode(Result.T, AStartTC);//Result.T + AStartTC;

{
  // 24시 -> 0시로 처리
  Days := GetDaysByTimecode(Result.T);

  Result.D := IncDay(Result.D, Days);
  Result.T := GetMinusTimecode(Result.T, GetTimecodePerDay * Days);
}

exit;  
  T.vtDWord := Result.T;

  if (T.Hour < 0) then
  begin
    Result.D := IncDay(Result.D, -1);

    T.Hour := HoursPerDay - 1;
    Result.T := T.vtDWord;
  end
  else if (T.Hour >= HoursPerDay) then
  begin
    Result.D := IncDay(Result.D, (T.Hour div HoursPerDay));

    T.Hour := T.Hour mod HoursPerDay;
    Result.T := T.vtDWord;
  end;
end;

function GetEventTimeSubEnd(AParentEventTime: TEventTime; AParentDurTC, AStartTC: TTimecode): TEventTime;
var
  TmpTimecode: TTimecode;
  Days: Integer;
  
  T: TTypeConvTime;
  F, Frame1, Frame2: Integer;
begin
  Result := AParentEventTime;
  Result.T := GetPlusTimecode(Result.T, AParentDurTC);
  Result.T := GetMinusTimecode(Result.T, AStartTC);

{
  // 24시 -> 0시로 처리
  if (Result.T < AStartTC) then
  begin
    TmpTimecode := GetDurTimecode(Result.T, AStartTC);
    Days        := GetDaysByTimecode(TmpTimecode) + 1;

    Result.D := IncDay(Result.D, -Days);
    Result.T := GetMinusTimecode(GetTimecodePerDay, Result.T);
  end
  else
  begin
    TmpTimecode := GetMinusTimecode(Result.T, AStartTC);
    Days        := GetDaysByTimecode(TmpTimecode);

    Result.D := IncDay(Result.D, Days);
    Result.T := GetMinusTimecode(TmpTimecode, GetTimecodePerDay * Days);
  end;
}

exit;
  
  if (T.Hour < 0) then
  begin
    Result.D := IncDay(Result.D, -1);

    T.Hour := HoursPerDay - 1;
    Result.T := T.vtDWord;
  end
  else if (T.Hour >= HoursPerDay) then
  begin
    Result.D := IncDay(Result.D, (T.Hour div HoursPerDay));

    T.Hour := T.Hour mod HoursPerDay;
    Result.T := T.vtDWord;
  end;

  T.vtDWord := Result.T;
end;

function GetEventEndTime(AEventTime: TEventTime; ADurTC: TTimecode): TEventTime;
var
  Days: Integer;
  
  T: TTypeConvTime;
begin
  Result := AEventTime;
  Result.T := GetPlusTimecode(Result.T, ADurTC);

{
  // 24시 -> 0시로 처리
  Days := GetDaysByTimecode(Result.T);

  Result.D := IncDay(Result.D, Days);
  Result.T := GetMinusTimecode(Result.T, GetTimecodePerDay * Days);
}

exit;
  T.vtDWord := Result.T;

  if (T.Hour < 0) then
  begin
    Result.D := IncDay(Result.D, -1);

    T.Hour := HoursPerDay - 1;
    Result.T := T.vtDWord;
  end
  else if (T.Hour >= HoursPerDay) then
  begin
    Result.D := IncDay(Result.D, (T.Hour div HoursPerDay));

    T.Hour := T.Hour mod HoursPerDay;
    Result.T := T.vtDWord;
  end;
end;

function GetDurEventTime(AEventTime1, AEventTime2: TEventTime; AFrameRate: Double = 30): TEventTime;
var
  Days: Integer;
  F, Frame1, Frame2: Integer;
  D1, D2: TDate;
  YY, MM, DD: Word;
  DateDiff: TDate;
  DaysPerMonth: Word;
  TC: TTimecode;
  T1, T2, T: TTypeConvTime;
  HH, MI, SS, FF: Integer;
begin
//  AEventTime1 := EventTimeToStadndardEventTime(AEventTime1);
//  AEventTime2 := EventTimeToStadndardEventTime(AEventTime2);

//  ShowMessage(EventTimeToDateTimecodeStr(AEventTime1));
//  ShowMessage(EventTimeToDateTimecodeStr(AEventTime2));
  Result.D := DaySpan(AEventTime1.D, AEventTime2.D);
  Result.T := GetDurTimecode(AEventTime1.T, AEventTime2.T);


  // Deadline Hour 보다 작으면 0시 기준으로 계산
  if (T.Hour < EventDeadlineHour) then
  begin
    // 24시 -> 0시로 처리
    if ((AEventTime1.D > AEventTime2.D) and (AEventTime1.T < AEventTime2.T)) or
       ((AEventTime1.D < AEventTime2.D) and (AEventTime1.T > AEventTime2.T)) then
    begin
      if (GetTimecodePerDay > Result.T) then
        Days := GetDaysByTimecode(Result.T) + 1
      else
        Days := GetDaysByTimecode(Result.T);

      Result.D := IncDay(Result.D, -Days);
      Result.T := GetDurTimecode(GetTimecodePerDay, Result.T);
    end
    else
    begin
      Days := GetDaysByTimecode(Result.T);

      Result.D := IncDay(Result.D, Days);
      Result.T := GetMinusTimecode(Result.T, GetTimecodePerDay * Days);
    end;
  end;

exit;

  T1.vtDWord := AEventTime1.T;
  T2.vtDWord := AEventTime2.T;

  HH := T2.Hour - T1.Hour;
  MI := T2.Minute - T1.Minute;
  SS := T2.Second - T1.Second;
  FF := T2.Frame - T1.Frame;

  if (FF < 0) then
  begin
    FF := FF + Round(AFrameRate);
    Dec(SS);
  end;

  if (SS < 0) then
  begin
    SS := SS + SecsPerMin;
    Dec(MI);
  end;

  if (MI < 0) then
  begin
    MI := MI + MinsPerHour;
    Dec(HH);
  end;

  if (HH < 0) then
  begin
    HH := HH + HoursPerDay;
    Result.D := IncDay(Result.D, - 1);
  end;

  if (AFrameRate = FrameRate29_97) then
  begin
    if ((FF in [0..1]) and (SS = 0) and ((MI mod 10) <> 0)) then
      FF := 2
    else if (FF >= FrameRate29_97) then
      FF := 29;
  end;

  T1.Hour   := HH;
  T1.Minute := MI;
  T1.Second := SS;
  T1.Frame  := FF;

  Result.T := T1.vtDWord;

//  Result.D := Abs(AEventTime2.D - AEventTime1.D);

{  Frame1 := TimeCodeToFrame(AEventTime1.T, AFrameRate);
  Frame2 := TimeCodeToFrame(AEventTime2.T, AFrameRate);
  F := Frame2 - Frame1 + 1;
  if (F < 0) then F := Abs(F);

  if (F < 0) then
  begin
    D.vtDWord := Result.D;
    D.Day := D.Day - 1;
    if (D.Day = 0) then
    begin
      D.Month := D.Month - 1;
      if (D.Month = 0) then
      begin
        D.Year := D.Year - 1;
        D.Month := MonthsPerYear;
      end;

      D.Day := DaysInAMonth(D.Year, D.Month);
    end;
    Result.D := D.vtDWord;

    F := Trunc(SecsPerDay * AFrameRate) + F;
  end;
  Result.T := FrameToTimecode(F); }

//  Result.T := GetDurTimecode(AEventTime1.T, AEventTime2.T, AFrameRate);
end;

function GetPlusEventTime(AEventTime1, AEventTime2: TEventTime; AFrameRate: Double = 30): TEventTime;
var
  Days: Integer;
  
  T: TTypeConvTime;
  D1, D2: TTypeConvDate;
  DaysPerMonth: Word;
begin
//  Result := DateTimeToEventTime(EventTimeToDateTime(AEventTime1, AFrameRate) + EventTimeToDateTime(AEventTime2, AFrameRate), AFrameRate);

  Result.D := AEventTime1.D + AEventTime2.D;
  Result.T := GetPlusTimecode(AEventTime1.T, AEventTime2.T, AFrameRate);

  T.vtDWord := Result.T;

  // Deadline Hour 보다 크면 0시 기준으로 계산
  if (T.Hour >= EventDeadlineHour) then
  begin
    // 24시 -> 0시로 처리
    Days := GetDaysByTimecode(Result.T);

    Result.D := IncDay(Result.D, Days);
    Result.T := GetMinusTimecode(Result.T, GetTimecodePerDay * Days);
  end;


exit;
  T.vtDWord := Result.T;

  if (T.Hour < 0) then
  begin
    Result.D := IncDay(Result.D, -1);

    T.Hour := HoursPerDay - 1;
    Result.T := T.vtDWord;
  end
  else if (T.Hour >= HoursPerDay) then
  begin
    Result.D := IncDay(Result.D, (T.Hour div HoursPerDay));

    T.Hour := T.Hour mod HoursPerDay;
    Result.T := T.vtDWord;
  end;
end;

function GetMinusEventTime(AEventTime1, AEventTime2: TEventTime; AFrameRate: Double = 30): TEventTime;
var
  TmpTimecode: TTimecode;
  Days: Integer;
  
  T: TTypeConvTime;
  D: TTypeConvDate;
begin
//  Result := DateTimeToEventTime(EventTimeToDateTime(AEventTime1, AFrameRate) - EventTimeToDateTime(AEventTime2, AFrameRate), AFrameRate);

  Result.D := AEventTime1.D - AEventTime2.D;

  Result.T := GetMinusTimecode(AEventTime1.T, AEventTime2.T, AFrameRate);

  T.vtDWord := Result.T;

  // Deadline Hour 보다 작으면 0시 기준으로 계산
  if (T.Hour < EventDeadlineHour) then
  begin
    // 24시 -> 0시로 처리
    if ((AEventTime1.D > AEventTime2.D) and (AEventTime1.T < AEventTime2.T)) or
       ((AEventTime1.D < AEventTime2.D) and (AEventTime1.T > AEventTime2.T)) then
    begin
      TmpTimecode := GetDurTimecode(AEventTime1.T, AEventTime2.T, AFrameRate);
      if (GetTimecodePerDay > Result.T) then
        Days := GetDaysByTimecode(Result.T) + 1
      else
        Days := GetDaysByTimecode(Result.T);

      Result.D := IncDay(Result.D, -Days);
      Result.T := GetDurTimecode(GetTimecodePerDay, TmpTimecode);
    end
    else
    begin
      TmpTimecode := GetMinusTimecode(AEventTime1.T, AEventTime2.T, AFrameRate);
      Days        := GetDaysByTimecode(TmpTimecode);

      Result.D := IncDay(Result.D, Days);
      Result.T := GetMinusTimecode(TmpTimecode, GetTimecodePerDay * Days);
    end;
  end;


exit;

  T.vtDWord := Result.T;

  if (T.Hour < 0) then
  begin
    Result.D := IncDay(Result.D, -1);

    T.Hour := HoursPerDay - 1;
    Result.T := T.vtDWord;
  end
  else if (T.Hour >= HoursPerDay) then
  begin
    Result.D := IncDay(Result.D, (T.Hour div HoursPerDay));

    T.Hour := T.Hour mod HoursPerDay;
    Result.T := T.vtDWord;
  end;
end;

function GetTimecodePerDay: TTimecode;
var
  T: TTypeConvTime;
begin
  T.Hour   := 24;
  T.Minute := 0;
  T.Second := 0;
  T.Frame  := 0;

  Result := T.vtDWord;
end;

function GetDurTimecode(ATimecode1, ATimecode2: TTimecode; AFrameRate: Double): TTimecode;
var
  TC: TTimecode;
  T1, T2, TR: TTypeConvTime;
  HH, MI, SS, FF: Integer;

  F, Frame1, Frame2: Integer;
begin
  if (ATimecode1 > ATimecode2) then
  begin
    TC := ATimecode1;
    ATimecode1 := ATimecode2;
    ATimecode2 := TC;
  end;

  Frame1 := TimeCodeToFrame(ATimecode1, AFrameRate);
  Frame2 := TimeCodeToFrame(ATimecode2, AFrameRate);
  F := Frame2 - Frame1;
  if (F < 0) then F := Abs(F);

  Result := FrameToTimecode(F, AFrameRate);

  exit;
  T1.vtDWord := ATimecode1;
  T2.vtDWord := ATimecode2;

  HH := T2.Hour - T1.Hour;
  MI := T2.Minute - T1.Minute;
  SS := T2.Second - T1.Second;
  FF := T2.Frame - T1.Frame;

  if (FF < 0) then
  begin
    FF := FF + Round(AFrameRate);
    Dec(SS);
  end;

  if (SS < 0) then
  begin
    SS := SS + SecsPerMin;
    Dec(MI);
  end;

  if (MI < 0) then
  begin
    MI := MI + MinsPerHour;
    Dec(HH);
  end;

{  if (AFrameRate = FrameRate29_97) then
  begin
    if ((FF in [0..1]) and (SS = 0) and ((MI mod 10) <> 0)) then
      FF := 2
    else if (FF >= 30) then
      FF := 29;
  end; }

  TR.Hour   := HH;
  TR.Minute := MI;
  TR.Second := SS;
  TR.Frame  := FF;

  Result := TR.vtDWord;
end;

function GetSubTimecode(ATimecode1, ATimecode2: TTimecode; AFrameRate: Double = 30): TTimecode;
var
  F, Frame1, Frame2: Integer;
begin
  Frame1 := TimeCodeToFrame(ATimecode1, AFrameRate);
  Frame2 := TimeCodeToFrame(ATimecode2, AFrameRate);
  F := Frame2 - Frame1;

  Result := FrameToTimecode(F, AFrameRate);
end;

function GetPlusTimecode(ATimecode1, ATimecode2: TTimecode; AFrameRate: Double): TTimecode;
var
  T1, T2, TR: TTypeConvTime;
  HH, MI, SS, FF: Integer;

  F, Frame1, Frame2: Integer;
begin
  Frame1 := TimeCodeToFrame(ATimecode1, AFrameRate);
  Frame2 := TimeCodeToFrame(ATimecode2, AFrameRate);
  F := Frame1 + Frame2;

  Result := FrameToTimeCode(F, AFrameRate);
exit;

  T1.vtDWord := ATimecode1;
  T2.vtDWord := ATimecode2;

  HH := T1.Hour + T2.Hour;
  MI := T1.Minute + T2.Minute;
  SS := T1.Second + T2.Second;
  FF := T1.Frame + T2.Frame;

  if (FF >= Round(AFrameRate)) then
  begin
    FF := FF - Round(AFrameRate);
    Inc(SS);
  end;

  if (SS >= SecsPerMin) then
  begin
    SS := SS - SecsPerMin;
    Inc(MI);
  end;

  if (MI >= MinsPerHour) then
  begin
    MI := MI - MinsPerHour;
    Inc(HH);
  end;

  if (HH > MAXBYTE) then
  begin
    HH := MAXBYTE;
  end;

  if (AFrameRate = FrameRate29_97) then
  begin
    if ((FF in [0..1]) and (SS = 0) and ((MI mod 10) <> 0)) then
      FF := 2
    else if (FF >= FrameRate29_97) then
      FF := 29;
  end;

  TR.Hour   := HH;
  TR.Minute := MI;
  TR.Second := SS;
  TR.Frame  := FF;

  Result := TR.vtDWord;
end;

function GetMinusTimecode(ATimecode1, ATimecode2: TTimecode; AFrameRate: Double): TTimecode;
var
  T1, T2, TR: TTypeConvTime;
  HH, MI, SS, FF: Integer;

  F, Frame1, Frame2: Integer;
begin
  Frame1 := TimeCodeToFrame(ATimecode1, AFrameRate);
  Frame2 := TimeCodeToFrame(ATimecode2, AFrameRate);
  F := Frame1 - Frame2;

  Result := FrameToTimeCode(F, AFrameRate);
  exit;

  T1.vtDWord := ATimecode1;
  T2.vtDWord := ATimecode2;

  HH := T1.Hour - T2.Hour;
  MI := T1.Minute - T2.Minute;
  SS := T1.Second - T2.Second;
  FF := T1.Frame - T2.Frame;

  if (FF < 0) then
  begin
    FF := FF + Round(AFrameRate);
    Dec(SS);
  end;

  if (SS < 0) then
  begin
    SS := SS + SecsPerMin;
    Dec(MI);
  end;

  if (MI < 0) then
  begin
    MI := MI + MinsPerHour;
    Dec(HH);
  end;

  if (HH < 0) then
  begin
    HH := 0;
  end;

  if (AFrameRate = FrameRate29_97) then
  begin
    if ((FF in [0..1]) and (SS = 0) and ((MI mod 10) <> 0)) then
      FF := 2
    else if (FF >= 30) then
      FF := 29;
  end;

  TR.Hour   := HH;
  TR.Minute := MI;
  TR.Second := SS;
  TR.Frame  := FF;

  Result := TR.vtDWord;
end;

function GetDaysByTimecode(ATimecode: TTimecode): Integer;
var
  T: TTypeConvTime;
begin
  T.vtDWord := ATimecode;
  Result := T.Hour div HoursPerDay;
end;

function GetFileVersionStr(AFileName: String): String;
var
  Size, Size2: DWord;
  Pt, Pt2: Pointer;
begin
  Result := '';
  Size := GetFileVersionInfoSize(PChar(AFileName), Size2);
  if Size > 0 then begin
    GetMem(Pt, Size);
    try
      GetFileVersionInfo(PChar(AFileName), 0, Size, Pt);
      VerQueryValue (Pt, '\', Pt2, Size2);
      with TVSFixedFileInfo(Pt2^) do begin
        Result := Format('%d.%d.%d.%d', [HiWord(dwFileVersionMS),
                                         LoWord(dwFileVersionMS),
                                         HiWord(dwFileVersionLS),
                                         LoWord(dwFileVersionLS)]);
      end;
    finally
      FreeMem(Pt);
    end;
  end;
end;

procedure SetEventDeadlineHour(AHour: Word);
begin
  EventDeadlineHour := 30;
end;

procedure PostKeyEx32(AKey: Word; const AShift: TShiftState; ASpecialKey: Boolean);
type
  TShiftKeyInfo = record
    shift: Byte;
    vkey: Byte;
  end;
  ByteSet = set of 0..7;
const
  shiftkeys: array [1..3] of TShiftKeyInfo = (
    (shift: Ord(ssCtrl) ; vkey: VK_CONTROL),
    (shift: Ord(ssShift) ; vkey: VK_SHIFT),
    (shift: Ord(ssAlt) ; vkey: VK_MENU)
  );
var
  flag: DWORD;
  bShift: ByteSet absolute AShift;
  j: Integer;
begin
  for j := 1 to 3 do
  begin
    if shiftkeys[j].shift in bShift then
      keybd_event(
        shiftkeys[j].vkey, MapVirtualKey(shiftkeys[j].vkey, 0), 0, 0
    );
  end;
  if ASpecialKey then
    flag := KEYEVENTF_EXTENDEDKEY
  else
    flag := 0;

  keybd_event(AKey, MapvirtualKey(AKey, 0), flag, 0);
  flag := flag or KEYEVENTF_KEYUP;
  keybd_event(AKey, MapvirtualKey(AKey, 0), flag, 0);

  for j := 3 downto 1 do
  begin
    if shiftkeys[j].shift in bShift then
      keybd_event(
        shiftkeys[j].vkey,
        MapVirtualKey(shiftkeys[j].vkey, 0),
        KEYEVENTF_KEYUP,
        0
      );
  end;
end;

function MakeNetConnection(APath, AUserId, APassword: String): Boolean;
var
  R: DWORD;

  NetResource: TNetResource;



  I: Integer;
  DeliStr: String;
  SumPos, DeliPos: Integer;
  RemotePath: String;
begin
//    원본 : \\192.168.125.253\test\enend\adb\tntne\ssae
//    파싱 : \\192.168.125.253\test

  Result := True;

  SumPos := 0;
  DeliStr := APath;

  DeliPos := Pos('\\', DeliStr);
  if (DeliPos <= 0) then exit;

  Inc(SumPos, DeliPos + 1);
  DeliStr := Copy(DeliStr, DeliPos + 2, Length(DeliStr));

  for I := 0 to 1 do
  begin
    DeliPos := Pos('\', DeliStr);
    if (DeliPos <= 0) then break;

    Inc(SumPos, DeliPos);
    DeliStr := Copy(DeliStr, DeliPos + 1, Length(DeliStr));
  end;

  RemotePath := Copy(APath, 1, SumPos - 1);

  FillChar(NetResource, SizeOf(TNetResource), #0);
//  NetResource.dwScope := RESOURCE_GLOBALNET;
//  NetResource.dwType := RESOURCETYPE_DISK;
//  NetResource.dwDisplayType := RESOURCEDISPLAYTYPE_SHARE;
//  NetResource.dwUsage := RESOURCEUSAGE_CONNECTABLE;
//  NetResource.lpLocalName := '';
  NetResource.lpRemoteName := PChar(RemotePath);
//  NetResource.lpComment := '';
//  NetResource.lpProvider := '';

//ShowMessage(RemotePath);

//  R := WNetAddConnection2(NetResource, PChar(APassword), PChar(AUserId), CONNECT_INTERACTIVE);
  R := WNetAddConnection2(NetResource, PChar(APassword), PChar(AUserId),  0);
  if (R <> NO_ERROR) and (R <> ERROR_SESSION_CREDENTIAL_CONFLICT) then
    Result := False;

//    ShowMessage(IntToStr(R));
//  ShowMessage(Format('Remote path = %s, connected = %s', [RemotePath, BoolToStr(Result, True)]));
end;

end.
