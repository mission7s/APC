unit UnitSECDLL;

interface

uses UnitCommons;

const
	SECDLL_NAME = 'SECDLL.DLL';

function SECSysInitialize(AInPort: Word; ATimeout: Cardinal = 1000): Integer; stdcall; external SECDLL_NAME name 'SECSysInitialize';
function SECSysFinalize: Integer; stdcall; external SECDLL_NAME name 'SECSysFinalize';

function SECSysLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external SECDLL_NAME name 'SECSysLogInPortEnable';
function SECSysLogInPortDisable: Integer; stdcall; external SECDLL_NAME name 'SECSysLogInPortDisable';

function SECSysSetServerReadProc(ReadFunc: Pointer): Integer; stdcall; external SECDLL_NAME name 'SECSysSetServerReadProc';
function SECSysTransmitResponse(AIP: PAnsiChar; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: PAnsiChar; ADataSize: Integer): Integer; stdcall; external SECDLL_NAME name 'SECSysTransmitResponse';
function SECSysTransmitAck(AIP: PAnsiChar; APort: Word): Integer; stdcall; external SECDLL_NAME name 'SECSysTransmitAck';
function SECSysTransmitNak(AIP: PAnsiChar; APort: Word; ANakError: Byte): Integer; stdcall; external SECDLL_NAME name 'SECSysTransmitNak';
function SECSysTransmitError(AIP: PAnsiChar; APort: Word; AErrorCode: Integer): Integer; stdcall; external SECDLL_NAME name 'SECSysTransmitError';

function SECInitialize(AInPort: Word; ATimeout: Cardinal = 1000; ABroadcast: Boolean = False): Integer; stdcall; external SECDLL_NAME name 'SECInitialize';
function SECFinalize: Integer; stdcall; external SECDLL_NAME name 'SECFinalize';

function SECLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall; external SECDLL_NAME name 'SECLogInPortEnable';
function SECLogInPortDisable: Integer; stdcall; external SECDLL_NAME name 'SECLogInPortDisable';

function SECSetServerReadProc(ReadFunc: Pointer): Integer; stdcall; external SECDLL_NAME name 'SECSetServerReadProc';
function SECTransmitResponse(AIP: PAnsiChar; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: PAnsiChar; ADataSize: Integer): Integer; stdcall; external SECDLL_NAME name 'SECTransmitResponse';
function SECTransmitAck(AIP: PAnsiChar; APort: Word): Integer; stdcall; external SECDLL_NAME name 'SECTransmitAck';
function SECTransmitNak(AIP: PAnsiChar; APort: Word; ANakError: Byte): Integer; stdcall; external SECDLL_NAME name 'SECTransmitNak';
function SECTransmitError(AIP: PAnsiChar; APort: Word; AErrorCode: Integer): Integer; stdcall; external SECDLL_NAME name 'SECTransmitError';

// 0X00 System Control
function SECSysIsAlive(AIP: PChar; var AIsAlive: Boolean): Integer; stdcall; external SECDLL_NAME name 'SECSysIsAlive';
function SECSysIsMain(AIP: PChar; var AIsMain: Boolean): Integer; stdcall; external SECDLL_NAME name 'SECSysIsMain';

// 0X10 Immediate Control

// 0X20 Preset/Select Commands

// 0X30 Sense Queries

// 0X40 CueSheet Control
function SECBeginUpdate(AIP: PChar; AChannelID: Word): Integer; stdcall; external SECDLL_NAME name 'SECBeginUpdate';
function SECEndUpdate(AIP: PChar; AChannelID: Word): Integer; stdcall; external SECDLL_NAME name 'SECEndUpdate';

function SECSetDeviceCommError(AIP: PChar; ADeviceStatus: TDeviceStatus; ADeviceName: PChar): Integer; stdcall; external SECDLL_NAME name 'SECSetDeviceCommError';
function SECSetDeviceStatus(AIP: PChar; ADCSID: Word; ADeviceHandle: TDeviceHandle; ADeviceStatus: TDeviceStatus): Integer; stdcall; external SECDLL_NAME name 'SECSetDeviceStatus';

function SECSetOnAir(AIP: PChar; AChannelID: Word; AIsOnAir: Boolean): Integer; stdcall; external SECDLL_NAME name 'SECSetOnAir';
function SECSetEventStatus(AIP: PChar; AEventID: TEventID; AEventStatus: TEventStatus): Integer; stdcall; external SECDLL_NAME name 'SECSetEventStatus';
function SECSetMediaStatus(AIP: PChar; AEventID: TEventID; AMediaStatus: TMediaStatus): Integer; stdcall; external SECDLL_NAME name 'SECSetMediaStatus';
function SECSetTimelineRange(AIP: PChar; AChannelID: Word; AStartDate, AEndDate: TDateTime): Integer; stdcall; external SECDLL_NAME name 'SECSetTimelineRange';

function SECInputCueSheet(AIP: PChar; AIndex: Integer; ACueSheetItem: TCueSheetItem): Integer; stdcall; external SECDLL_NAME name 'SECInputCueSheet';
function SECDeleteCueSheet(AIP: PChar; AEventID: TEventID): Integer; stdcall; external SECDLL_NAME name 'SECDeleteCueSheet';
function SECClearCueSheet(AIP: PChar; AChannelID: Word): Integer; stdcall; external SECDLL_NAME name 'SECClearCueSheet';

function SECSetCueSheetCurr(AIP: PChar; AEventID: TEventID): Integer; stdcall; external SECDLL_NAME name 'SECSetCueSheetCurr';
function SECSetCueSheetNext(AIP: PChar; AEventID: TEventID): Integer; stdcall; external SECDLL_NAME name 'SECSetCueSheetNext';
function SECSetCueSheetTarget(AIP: PChar; AEventID: TEventID): Integer; stdcall; external SECDLL_NAME name 'SECSetCueSheetTarget';

implementation

end.
