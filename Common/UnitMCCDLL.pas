unit UnitMCCDLL;

interface

uses UnitCommons;

const
	MCCDLL_NAME = 'MCCDLL.DLL';

function MCCSysInitialize(AInPort, AOutPort: Word; ATimeout: Cardinal = 1000): Integer; stdcall; external MCCDLL_NAME name 'MCCSysInitialize';
function MCCSysFinalize: Integer; stdcall; external MCCDLL_NAME name 'MCCSysFinalize';

function MCCSysLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external MCCDLL_NAME name 'MCCSysLogInPortEnable';
function MCCSysLogInPortDisable: Integer; stdcall; external MCCDLL_NAME name 'MCCSysLogInPortDisable';

function MCCSysLogOutPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external MCCDLL_NAME name 'MCCSysLogOutPortEnable';
function MCCSysLogOutPortDisable: Integer; stdcall; external MCCDLL_NAME name 'MCCSysLogOutPortDisable';

function MCCInitialize(ANotifyPort, AInPort, AOutPort: Integer; ATimeout: Cardinal = 1000; ABroadcast: Boolean = False): Integer; stdcall; external MCCDLL_NAME name 'MCCInitialize';
function MCCFinalize: Integer; stdcall; external MCCDLL_NAME name 'MCCFinalize';

function MCCLogNotifyPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external MCCDLL_NAME name 'MCCLogNotifyPortEnable';
function MCCLogNotifyPortDisable: Integer; stdcall; external MCCDLL_NAME name 'MCCLogNotifyPortDisable';

function MCCLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external MCCDLL_NAME name 'MCCLogInPortEnable';
function MCCLogInPortDisable: Integer; stdcall; external MCCDLL_NAME name 'MCCLogInPortDisable';

function MCCLogOutPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external MCCDLL_NAME name 'MCCLogOutPortEnable';
function MCCLogOutPortDisable: Integer; stdcall; external MCCDLL_NAME name 'MCCLogOutPortDisable';

// 0X00 System Control
function MCCSysIsAlive(AIP: PChar; var AIsAlive: Boolean): Integer; stdcall; external MCCDLL_NAME name 'MCCSysIsAlive';

// 0X10 Immediate Control

// 0X20 Preset/Select Commands

// 0X30 Sense Queries

// 0X40 CueSheet Control
function MCCBeginUpdate(AIP: PChar; AChannelID: Word): Integer; stdcall; external MCCDLL_NAME name 'MCCBeginUpdate';
function MCCEndUpdate(AIP: PChar; AChannelID: Word): Integer; stdcall; external MCCDLL_NAME name 'MCCEndUpdate';

function MCCSetDeviceCommError(AIP: PChar; ADeviceStatus: TDeviceStatus; ADeviceName: PChar): Integer; stdcall; external MCCDLL_NAME name 'MCCSetDeviceCommError';
function MCCSetDeviceStatus(AIP: PChar; ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus): Integer; stdcall; external MCCDLL_NAME name 'MCCSetDeviceStatus';

function MCCSetOnAir(AIP: PChar; AChannelID: Word; AIsOnAir: Boolean): Integer; stdcall; external MCCDLL_NAME name 'MCCSetOnAir';
function MCCSetEventStatus(AIP: PChar; AEventID: TEventID; AEventStatus: TEventStatus): Integer; stdcall; external MCCDLL_NAME name 'MCCSetEventStatus';
function MCCSetMediaStatus(AIP: PChar; AEventID: TEventID; AMediaStatus: TMediaStatus): Integer; stdcall; external MCCDLL_NAME name 'MCCSetMediaStatus';

function MCCInputCueSheet(AIP: PChar; AIndex: Integer; ACueSheetItem: TCueSheetItem): Integer; stdcall; external MCCDLL_NAME name 'MCCInputCueSheet';
function MCCDeleteCueSheet(AIP: PChar; AEventID: TEventID): Integer; stdcall; external MCCDLL_NAME name 'MCCDeleteCueSheet';
function MCCClearCueSheet(AIP: PChar; AChannelID: Word): Integer; stdcall; external MCCDLL_NAME name 'MCCClearCueSheet';

function MCCSetCueSheetCurr(AIP: PChar; AEventID: TEventID): Integer; stdcall; external MCCDLL_NAME name 'MCCSetCueSheetCurr';
function MCCSetCueSheetNext(AIP: PChar; AEventID: TEventID): Integer; stdcall; external MCCDLL_NAME name 'MCCSetCueSheetNext';
function MCCSetCueSheetTarget(AIP: PChar; AEventID: TEventID): Integer; stdcall; external MCCDLL_NAME name 'MCCSetCueSheetTarget';

implementation

end.
