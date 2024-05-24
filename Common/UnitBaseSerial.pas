unit UnitBaseSerial;

{$MINENUMSIZE 4}

interface

uses
  Vcl.Forms, Winapi.Windows, Winapi.Messages, System.Classes, System.Win.Registry,
  System.SysUtils, System.AnsiStrings, System.SyncObjs, Winapi.MMSystem,
  Winapi.WinSock2, Dialogs;

const
  READ_TIMEOUT            = 1000;

  WM_SOCKET_MESSAGE       = WM_USER + $1001;

  DCB_Binary              = $00000001;
  DCB_ParityCheck         = $00000002;
  DCB_OutxCtsFlow         = $00000004;
  DCB_OutxDsrFlow         = $00000008;
  DCB_DtrControlMask      = $00000030;
  DCB_DtrControlDisable   = $00000000;
  DCB_DtrControlEnable    = $00000010;
  DCB_DtrControlHandshake = $00000020;
  DCB_DsrSensivity        = $00000040;
  DCB_TXContinueOnXoff    = $00000080;
  DCB_OutX                = $00000100;
  DCB_InX                 = $00000200;
  DCB_ErrorChar           = $00000400;
  DCB_NullStrip           = $00000800;
  DCB_RtsControlMask      = $00003000;
  DCB_RtsControlDisable   = $00000000;
  DCB_RtsControlEnable    = $00001000;
  DCB_RtsControlHandshake = $00002000;
  DCB_RtsControlToggle    = $00003000;
  DCB_AbortOnError        = $00004000;
  DCB_Reserveds           = $FFFF8000;

  MAX_RETRY               = 3;

  D_OK                    = 0;
  D_FALSE = -1;

  E_NAK_UNDEFINED         = $01;
  E_NAK_SYNTAX            = $02;
  E_NAK_CHECKSUM          = $04;
  E_NAK_PARITY            = $10;
  E_NAK_OVERRUN           = $20;
  E_NAK_FRAMING           = $40;
  E_NAK_TIMEOUT           = $80;
  E_NAK_IGNORED           = $8F;

  E_TIMEOUT               = $F0;

  INIT_TIMECODE = '00:00:00:00';
  IDLE_TIMECODE = '--:--:--:--';

type
  TBaseSerial = class;

  // Control Type
  TControlType = (ctSerial, ctTCP, ctUDP);

  // COM Port Type
  TComPortType = (ptRS232, ptRS422, ptRS485);

  // COM Port Baud Rates
  TComPortBaudRate = (br2400, br4800, br9600, br14400, br19200,
                      br38400, br56000, br57600, br115200);

  // COM Port Data Bits
  TComPortDataBits = (db5Bits, db6Bits, db7BITS, db8BITS);

  // COM Port Stop Bits
  TComPortStopBits = (sb1Bits, sb1HalfBits, sb2Bits );

  // COM Port Parity
  TComPortParity = (ptNone, ptOdd, ptEven, ptMark, ptSpace);

  // COM Port Hardware Handshaking
  TComPortHwHandshaking = (hhNone, hhNoneRtsOn, hhRtsCts);

  // COM Port Software Handshaing
  TComPortSwHandshaking = (shNone, shXOnXOff);

  TTypeConvWord = packed record
     case vType: Word of
        0:(vtWord: Word);
        1:(c2, c1: AnsiChar);
  end;

  TTypeConvInt64 = record
     case vType: Int64 of
        0:(c8, c7, c6, c5, c4, c3, c2, c1: AnsiChar);
        1:(vtInt64: Int64);
  end;

  TTypeConvDouble = record
     case vType: Int64 of
        0:(c8, c7, c6, c5, c4, c3, c2, c1: AnsiChar);
        1:(vtDouble: Double);
  end;

  TTypeConvInt = record
     case vType: Integer of
        0:(vtSingle: single);
        1:(c4, c3, c2, c1: AnsiChar);
        2:(vtInteger: Integer);
        3:(vtDWord: DWord);
  end;

  TTypeConvSmallInt = record
    case vType: SmallInt of
      0:(vtSmallInt: SmallInt);
      1:(c2, c1: AnsiChar);
      2:(vtWord: Word);
  end;

  TTypeConvSingle = record
  	case vType: Integer of
	    0:(c4, c3, c2, c1: AnsiChar);
     	1:(vtSingle: Single);
  end;

  TAsyncThread = class(TThread)
  private
    FBaseSerial: TBaseSerial;
    FEventMask: Cardinal;
    FRxCount: Cardinal;
    FTxEmpty: Boolean;

    FASyncComHandle: THandle;
    FCloseEvent: THandle;
    FSockReadEvent: THandle;

    procedure PortEvent;
  protected
    procedure Execute; override;
  public
    constructor Create(ABaseSerial: TBaseSerial; AASyncComHandle: THandle);
    destructor Destroy; override;

    procedure Close;
  end;

  TLogKind = (lkSend, lkReceive, lkTest);

  TOnDataEvent = procedure(Sender: TObject; Buffer: String) of object;

  TOnSocketCloseEvent = procedure(Sender: TObject) of object;

  TBaseSerial = class(TComponent)
  private
    FCriticalSection: TCriticalSection;
  protected
    FControlType: TControlType;                   // Serial or Socket
    FControlIP: AnsiString;
    FControlPort: Word;
    FSocketHandle: TSocket;
    FSocketEvent: THandle;
    FSocketConnected: Boolean;

    FComPortHandle: THANDLE;                      // COM Port Device Handle
    FComPort: Integer;
    FComPortBaudRate: TComPortBaudRate;           // COM Port speed (brXXXX)
    FComPortDataBits: TComPortDataBits;           // Data bits size (5..8)
    FComPortStopBits: TComPortStopBits;           // How many stop bits to use (1, 1.5, 2)
    FComPortParity: TComPortParity;               // Type of parity to use (none, odd, even, mark, space)
    FComPortHwHandshaking: TComPortHwHandshaking; // Type of hw handshaking to use
    FComPortSwHandshaking: TComPortSwHandshaking; // Type of sw handshaking to use
    FComPortInBufSize: Word;                      // Size of the input buffer
    FComPortOutBufSize: Word;                     // Size of the output buffer
    FEnableDTROnOpen: Boolean;                    // enable/disable DTR line on connect
    FOutputTimeout: Word;                         // output timeout - milliseconds
    FNotifyWnd: HWND;                             // This is used for the timer
    FComPortNames: TStringList;
//    FSyncMode: Boolean;
    FAsyncThread: TAsyncThread;
    FOnDataEvent: TNotifyEvent;
    FOnSocketCloseEvent: TNotifyEvent;
    FTimeOut: Word;
//    FDropFrame: Boolean;                          // True : dropframe , False : Non Drop Frame
    FSyncMsgEvent: THandle;
    FErrorString: String;
    FLastResult: Integer;
    FReceivedData: AnsiString;
    FReceivedBuffer: AnsiString;
    FLogPath: String;
    FLogEnabled: Boolean;
    FLogIsHexcode: Boolean;
    FLogExt: String;
    FOverlapMode: Boolean;

    FFlushBuf: Boolean;

    LastErrMsg: String;

    function GetRxCount: Cardinal;
    function GetTxEmpty: Boolean;
    function WriteLog(ALogKind: TLogKind; ABuffer: PAnsiChar; ADataSize: Integer; AHexcode: Boolean = True): Integer;
    function WaitForResult(ACount: Integer; AHandles: array of THandle; AWaitAll: Boolean; AMillSecond, AWakeMask: DWORD): DWORD;

//    procedure SockWndProc(var Msg: TMessage);
    function MsgWaitForHandles(ACount: Cardinal; AHandles: array of THandle; ATimeOut, AWakeMask: DWORD): DWORD;

    procedure SetControlType(AValue: TControlType);
    procedure SetComPort(AValue: Integer);
    procedure SetComPortBaudRate(AValue: TComPortBaudRate);
    procedure SetComPortDataBits(AValue: TComPortDataBits);
    procedure SetComPortStopBits(AValue: TComPortStopBits);
    procedure SetComPortParity(AValue: TComPortParity);
    procedure SetComPortHwHandshaking(AValue: TComPortHwHandshaking);
    procedure SetComPortSwHandshaking(AValue: TComPortSwHandshaking);
    procedure SetComPortInBufSize(AValue: Word);
    procedure SetComPortOutBufSize(AValue: Word);
    procedure SetTimeOut(const AValue: Word);
//    procedure SetSyncMode(AValue: Boolean);
//    procedure SetDropFrame(const AValue: Boolean);
    procedure SetOverlapMode(const AValue: Boolean);
    procedure ApplyComPortSettings;
    procedure DataEvent; virtual;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function Connect: Boolean;
    function Connected: Boolean;
    function Disconnect: Boolean;

    procedure FlushBuffers(AInBuf, AOutBuf: Boolean);
    function OutFreeSpace: Word;

    procedure ToggleDTR(AOnOff: Boolean);
    procedure ToggleRTS(AOnOff: Boolean);

    function SendData(ABuffer: Pointer; ADataSize: Integer): Integer;
    function ReadData(ABuffer: Pointer; ADataSize: Integer): Integer;
    function SendString(AStr: AnsiString): Integer;
    function ReadString(var AStr: AnsiString; ADataSize: Integer): Integer;

    { Send & Read Function }
    function SendReceiveData(ACommand: AnsiString; ALen: Integer): Integer; virtual;
    function ReceiveData: Integer; virtual; abstract;
    function Receive(ABuffer: Pointer; ADataSize: Integer): Integer;

    { RxCount <= Reading Count }
    property RxCount: Cardinal read GetRxCount;
    property LastResult: Integer read FLastResult write FLastResult;
    property ReceivedBuffer: AnsiString read FReceivedBuffer write FReceivedBuffer;
    property ReceivedData: AnsiString read FReceivedData write FReceivedData;
    property SyncMsgEvent: THandle read FSyncMsgEvent;

    property TxEmpty: Boolean read GetTxEmpty;
  published
    property ControlType: TControlType read FControlType write SetControlType default ctSerial;
    property ControlIP: AnsiString read FControlIP write FControlIP;
    property ControlPort: Word read FControlPort write FControlPort;

    // Which COM Port to use
    property ComPort: Integer read FComPort write SetComPort default 0;
    // COM Port speed (bauds)
    property ComPortBaudRate: TComPortBaudRate read FComPortBaudRate write SetComPortBaudRate default br38400;
    // Data bits to used (5..8, for the 8250 the use of 5 data bits with 2 stop bits is an invalid combination,
    // as is 6, 7, or 8 data bits with 1.5 stop bits)
    property ComPortDataBits: TComPortDataBits read FComPortDataBits write SetComPortDataBits default db8Bits;
    // Stop bits to use (1, 1.5, 2)
    property ComPortStopBits: TComPortStopBits read FComPortStopBits write SetComPortStopBits default sb1Bits;
    // Parity Type to use (none,odd,even,mark,space)
    property ComPortParity: TComPortParity read FComPortParity write SetComPortParity default ptNONE;
    // Hardware Handshaking Type to use:
    // hhNone          no handshaking
    // hhCtsRts        both cdCTS and cdRTS apply (** this is the more common method**)
    property ComPortHwHandshaking: TComPortHwHandshaking read FComPortHwHandshaking write SetComPortHwHandshaking default hhNone;
    // Software Handshaking Type to use:
    // shNone          no handshaking
    // shXOnXOff       XOn/XOff handshaking
    property ComPortSwHandshaking: TComPortSwHandshaking read FComPortSwHandshaking write SetComPortSwHandshaking default shNone;
    // Input Buffer size
    property ComPortInBufSize: Word read FComPortInBufSize write SetComPortInBufSize default 2048;
    // Output Buffer size
    property ComPortOutBufSize: Word read FComPortOutBufSize write SetComPortOutBufSize default 2048;
    // v1.02: Set to True to enable DTR line on connect and to leave it on until disconnect.
    //        Set to False to disable DTR line on connect.
    property EnableDTROnOpen: Boolean read FEnableDTROnOpen write FEnableDTROnOpen default True;
    // v1.02: Output timeout (milliseconds)
    property OutputTimeout: Word read FOutputTimeOut write FOutputTimeout default 4000;
    // Event to raise when there is data available (input buffer has data)
//    property SyncMode: Boolean read FSyncMode write SetSyncMode default True;
    property OnDataEvent: TNotifyEvent read FOnDataEvent write FOnDataEvent;
    property OnSocketCloseEvent: TNotifyEvent read FOnSocketCloseEvent write FOnSocketCloseEvent;
    property TimeOut: Word read FTimeOut write SetTimeOut default 1000;
//    property DropFrame: Boolean read FDropFrame write SetDropFrame default True;
    property ErrorString: String read FErrorString;
    property LogPath: String read FLogPath write FLogPath;
    property LogEnabled: Boolean read FLogEnabled write FLogEnabled;
    property LogIsHexcode: Boolean read FLogIsHexcode write FLogIsHexcode;
    property LogExt: String read FLogExt write FLogExt;

    property FlushBuf: Boolean read FFlushBuf write FFlushBuf;
  end;

{ Conversion Functions }
function PAnsiCharToWord(AValue: PAnsiChar): Word;
function PAnsiCharToWordR(AValue: PAnsiChar): Word;
function PAnsiCharToDWord(AValue: PAnsiChar): DWord;
function PAnsiCharToSmallInt(AValue: PAnsiChar): SmallInt;
function PAnsiCharToSmallIntR(AValue: PAnsiChar): SmallInt;
function PAnsiCharToInt(AValue: PAnsiChar): Integer;
function PAnsiCharToInt2(AValue: PAnsiChar): Integer;
function PAnsiCharToInt3(AValue: PAnsiChar): Integer;
function PAnsiCharToInt6(AValue: PAnsiChar): Integer;
function PAnsiCharToIntR(AValue: PAnsiChar): Integer;
function PAnsiCharToIntR2(AValue: PAnsiChar): Integer;
function PAnsiCharToIntR3(AValue: PAnsiChar): Integer;
function PAnsiCharToIntR6(AValue: PAnsiChar): Integer;
function PAnsiCharToInt64(AValue: PAnsiChar): Int64;
function PAnsiCharToSingle(AValue: PAnsiChar): Single;
function PAnsiCharToDouble(AValue: PAnsiChar): double;
function PAnsiCharToHexCode(AValue: PAnsiChar; ADataSize: Integer): AnsiString;
function PAnsiCharToString(AValue: PAnsiChar; ADataSize: Integer): AnsiString;
function WordToAnsiString(AValue: Word): AnsiString;
function WordToAnsiStringR(AValue: Word): AnsiString;
function WordToAnsiStringWithCRC(AValue: SmallInt; var ACrc: Integer): AnsiString;
function DWordToAnsiString(AValue: DWord): AnsiString;
function DWordToAnsiStringR(AValue: DWord): AnsiString;
function DWordToAnsiStringWithCrc(AValue: DWord; var ACrc: Integer): AnsiString;
function SmallIntToAnsiString(AValue: SmallInt): AnsiString;
function SmallIntToAnsiStringR(AValue: SmallInt): AnsiString;
function IntToAnsiString(AValue: Integer): AnsiString;
function IntToAnsiString2(AValue: Integer): AnsiString;
function IntToAnsiString3(AValue: Integer): AnsiString;
function IntToAnsiString6(AValue: Integer): AnsiString;
function IntToAnsiStringR(AValue: Integer): AnsiString;
function IntToAnsiStringR2(AValue: Integer): AnsiString;
function IntToAnsiStringR3(AValue: Integer): AnsiString;
function IntToAnsiStringR6(AValue: Integer): AnsiString;
function Int64ToAnsiString(AValue: Int64): AnsiString;
function SingleToAnsiString(AValue: Single): AnsiString;
function DoubleToAnsiString(AValue: Double): AnsiString;
function IntToBCD(AValue: Integer): Integer;
function BCDToInt(AValue: Integer): Integer;

function EncodeTimeCode(const AHour, AMinute, ASecond, AFrame: Integer; AExtended: Boolean = False): String;
procedure DecodeTimeCode(ATimeCode: String; out AHour, AMinute, ASecond, AFrame: Integer);
function FrameToTimeCode(AFrame: Integer; ADropFrame: Boolean): String;
function TimeCodeToFrame(ATC: String; ADropFrame: Boolean): Integer;
function GetTimeCodeDuration(ATC1, ATC2: String; ADropFrame: Boolean): String;
function GetPlusTimeCode(ATC1, ATC2: String; ADropFrame: Boolean): String;
function GetMinusTimeCode(ATC1, ATC2: String; ADropFrame: Boolean): String;
function IsValidTimeCode(ATC: String; ADropFrame: Boolean): Boolean;

implementation

function ValidateComPort(ComPort: PChar): Boolean;
var
  PortHandle: THandle;
begin
  // COM포트가 사용가능한지 알아낸다.

  Result := False;
  PortHandle := CreateFile(ComPort, GENERIC_READ or GENERIC_WRITE, 0, nil,
      OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
  if (PortHandle <> INVALID_HANDLE_VALUE) then
  begin
    Result := True;
    CloseHandle(PortHandle);
  end;
end;

procedure GetComPortNames(List: TStrings; OnlyAvail: Boolean = False);
var
  RegFile: TRegistry;
  DeviceList: TStrings;
  DeviceCount: Integer;
  CanAddPort: Boolean;
  CurrentPort: string;
begin
  // 현재 인스톨된 COM포트 목록을 구한다.

  if (Assigned(List)) then
  begin
    RegFile := TRegistry.Create;
    try
      RegFile.RootKey := HKEY_LOCAL_MACHINE;
      RegFile.OpenKeyReadOnly('hardware\devicemap\serialcomm');
      DeviceList := TStringList.Create;
      try
        RegFile.GetValueNames(DeviceList);
        List.Clear;
        for DeviceCount := 0 to DeviceList.Count - 1 do
        begin
            CurrentPort := RegFile.ReadString(DeviceList.Strings[DeviceCount]);
            if OnlyAvail then
                CanAddPort := ValidateComPort(PChar(CurrentPort))
            else
                CanAddPort := True;
            if CanAddPort then
                List.Add(CurrentPort);
        end;
      finally
        FreeAndNil(DeviceList);
      end;
      RegFile.CloseKey;
    finally
      FreeAndNil(RegFile);
    end;
  end;
end;

{ Conversion Functions }

function PAnsiCharToWord(AValue: PAnsiChar): Word;
var
  vtType: TTypeConvWord;
begin
  vtType.C2 := AValue[0];
  vtType.C1 := AValue[1];
  Result := vtType.vtWord;
end;

function PAnsiCharToWordR(AValue: PAnsiChar): Word;
var
  vtType : TTypeConvWord;
begin
  vtType.C1 := AValue[0];
  vtType.C2 := AValue[1];
  Result := vtType.vtWord;
end;

function PAnsiCharToDWord(AValue: PAnsiChar): DWord;
var
  vtType: TTypeConvInt;
begin
  vtType.C4 := AValue[0];
  vtType.C3 := AValue[1];
  vtType.C2 := AValue[2];
  vtType.C1 := AValue[3];
  Result := vtType.vtDWord;
end;

function PAnsiCharToDWordR(AValue: PAnsiChar): DWord;
var
  vtType: TTypeConvInt;
begin
  vtType.C1 := AValue[0];
  vtType.C2 := AValue[1];
  vtType.C3 := AValue[2];
  vtType.C4 := AValue[3];
  Result := vtType.vtDWord;
end;

function PAnsiCharToSmallInt(AValue: PAnsiChar): SmallInt;
var
  vtType: TTypeConvSmallInt;
begin
  vtType.C2 := AValue[0];
  vtType.C1 := AValue[1];
  Result := vtType.vtSmallInt;
end;

function PAnsiCharToSmallIntR(AValue: PAnsiChar): SmallInt;
var
  vtType: TTypeConvSmallInt;
begin
  vtType.C1 := AValue[0];
  vtType.C2 := AValue[1];
  Result := vtType.vtSmallInt;
end;

function PAnsiCharToInt(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C4 := AValue[0];
  vtType.C3 := AValue[1];
  vtType.C2 := AValue[2];
  vtType.C1 := AValue[3];
  Result := vtType.vtInteger;
end;

function PAnsiCharToInt2(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C4 := AValue[0];
  vtType.C3 := AValue[1];
  Result := vtType.vtInteger;
end;

function PAnsiCharToInt3(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C4 := AValue[0];
  vtType.C3 := AValue[1];
  vtType.C2 := AValue[2];
  Result := vtType.vtInteger;
end;

function PAnsiCharToInt6(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt64;
begin
  vtType.C8 := AValue[0];
  vtType.C7 := AValue[1];
  vtType.C6 := AValue[2];
  vtType.C5 := AValue[3];
  vtType.C6 := AValue[4];
  vtType.C3 := AValue[5];
  Result := vtType.vtInt64;
end;

function PAnsiCharToIntR(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C1 := AValue[0];
  vtType.C2 := AValue[1];
  vtType.C3 := AValue[2];
  vtType.C4 := AValue[3];
  Result := vtType.vtInteger;
end;

function PAnsiCharToIntR2(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C3 := AValue[0];
  vtType.C4 := AValue[1];
  Result := vtType.vtInteger;
end;

function PAnsiCharToIntR3(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C2 := AValue[0];
  vtType.C3 := AValue[1];
  vtType.C4 := AValue[2];
  Result := vtType.vtInteger;
end;

function PAnsiCharToIntR6(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt64;
begin
  vtType.C3 := AValue[0];
  vtType.C4 := AValue[1];
  vtType.C5 := AValue[2];
  vtType.C6 := AValue[3];
  vtType.C7 := AValue[4];
  vtType.C8 := AValue[5];
  Result := vtType.vtInt64;
end;

function PAnsiCharToInt64(AValue: PAnsiChar): Int64;
var
  vtType: TTypeConvInt64;
begin
  vtType.C8 := AValue[0];
  vtType.C7 := AValue[1];
  vtType.C6 := AValue[2];
  vtType.C5 := AValue[3];
  vtType.C4 := AValue[4];
  vtType.C3 := AValue[5];
  vtType.C2 := AValue[6];
  vtType.C1 := AValue[7];
  Result := vtType.vtInt64;
end;

function PAnsiCharToSingle(AValue: PAnsiChar): Single;
var
	vtType: TTypeConvSingle;
begin
	vtType.c4 := AValue[0];
	vtType.c3 := AValue[1];
  vtType.c2 := AValue[2];
  vtType.c1 := AValue[3];

  Result := vtType.vtSingle;
end;

function PAnsiCharToDouble(AValue: PAnsiChar): Double;
var
  vtType: TTypeConvDouble;
begin
  vtType.C8 := AValue[0];
  vtType.C7 := AValue[1];
  vtType.C6 := AValue[2];
  vtType.C5 := AValue[3];
  vtType.C4 := AValue[4];
  vtType.C3 := AValue[5];
  vtType.C2 := AValue[6];
  vtType.C1 := AValue[7];
  Result := vtType.vtDouble;
end;

function PAnsiCharToHexCode(AValue: PAnsiChar; ADataSize: Integer): AnsiString;
var
  I: Integer;
  V: array[0..4] of AnsiChar;
  A: PAnsiChar;
begin
//  V := AnsiStrAlloc(ADataSize);
//  StrCopy(V, AValue);
//  Move(AValue^, V^, ADataSize);
//  for I := 0 to ADataSize - 1 do
//  begin
//    Result := Result + ansi FormatFloat('00', Ord(AValue[I]));
//  end;
////  Result := '12345645y398573295732050935402hgfkjdsfjldfjlkasfnkladsf';
//  StrDispose(V);
//  exit;

  for I := 0 to ADataSize - 1 do
  begin
//    A := System.AnsiStrings.StrFmt(V, '0X%0.2x ', [Ord(AValue[I])]);
    Result := Result + {AnsiString(A);//System.AnsiStrings.}Format('0X%0.2x ', [Ord(AValue[I])]);
  end;
end;

function PAnsiCharToString(AValue: PAnsiChar; ADataSize: Integer): AnsiString;
var
  I: Integer;
begin
//    ShowMessage((AValue));

  Result := '';
  if (ADataSize <= 0) then exit;

  for I := 0 to ADataSize - 1 do
  begin
//    A := System.AnsiStrings.StrFmt(V, '0X%0.2x ', [Ord(AValue[I])]);
    Result := Result + AnsiChar(AValue[I]);
  end;

//  CopyMemory(PAnsiChar(Result), AValue, ADataSize);
//  ShowMessage(Result);
//  ShowMessage(IntToStr(Length(Result)));
end;

function WordToAnsiString(AValue: Word): AnsiString;
var
  vtType: TTypeConvWord;
begin
  vtType.vtWord := AValue;
  Result := vtType.c2 + vtType.c1;
end;

function WordToAnsiStringR(AValue: Word): AnsiString;
var
  vtType: TTypeConvWord;
begin
  vtType.vtWord := AValue;
  Result := vtType.c1 + vtType.c2;
end;

function WordToAnsiStringWithCRC(AValue: SmallInt; var ACrc: Integer): AnsiString;
var
  vtType: TTypeConvWord;
begin
  vtType.vtWord := AValue;
  ACrc := ACrc + Ord(vtType.C2) + Ord(VtType.C1);
  Result := vtType.c2 + vtType.c1;
end;

function DWordToAnsiString(AValue: DWord): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtDWord := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function DWordToAnsiStringR(AValue: DWord): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtDWord := AValue;
  Result := vtType.c1 + vtType.c2 + vtType.c3 + vtType.c4;
end;

function DWordToAnsiStringWithCrc(AValue: DWord; var ACrc: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtDWord := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
  ACrc := Ord(vtType.c4) + Ord(vtType.c3) + Ord(vtType.c2) + Ord(vtType.c1);
end;

function SmallIntToAnsiString(AValue: SmallInt): AnsiString;
var
  vtType: TTypeConvSmallInt;
begin
  vtType.vtSmallInt := AValue;
  Result := vtType.c2 + vtType.c1;
end;

function SmallIntToAnsiStringR(AValue: SmallInt): AnsiString;
var
  vtType: TTypeConvSmallInt;
begin
  vtType.vtSmallInt := AValue;
  Result := vtType.c1 + vtType.c2;
end;

function IntToAnsiString(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function IntToAnsiString2(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c4 + vtType.c3;
end;

function IntToAnsiString3(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2;
end;

function IntToAnsiString6(AValue: Integer): AnsiString;
var
	vtType: TTypeConvInt64;
begin
	vtType.vtInt64 := AValue;
  Result := vtType.c8 + vtType.c7 + vtType.c6 + vtType.c5 +
   			    vtType.c4 + vtType.c3;
end;

function IntToAnsiStringR(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c1 + vtType.c2 + vtType.c3 + vtType.c4;
end;

function IntToAnsiStringR2(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c3 + vtType.c4;
end;

function IntToAnsiStringR3(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c2 + vtType.c3 + vtType.c4;
end;

function IntToAnsiStringR6(AValue: Integer): AnsiString;
var
	vtType: TTypeConvInt64;
begin
	vtType.vtInt64 := AValue;
  Result := vtType.c3 + vtType.c4 + vtType.c5 + vtType.c6 +
   			    vtType.c7 + vtType.c8;
end;

function Int64ToAnsiString(AValue: Int64): AnsiString;
var
	vtType: TTypeConvInt64;
begin
	vtType.vtInt64 := AValue;
  Result := vtType.c8 + vtType.c7 + vtType.c6 + vtType.c5 +
   			    vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function SingleToAnsiString(AValue: Single): AnsiString;
var
	vtType: TTypeConvSingle;
begin
	vtType.vtSingle := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function DoubleToAnsiString(AValue: Double): AnsiString;
var
	vtType: TTypeConvDouble;
begin
	vtType.vtDouble := AValue;
  Result := vtType.c8 + vtType.c7 + vtType.c6 + vtType.c5 +
   			    vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function IntToBCD(AValue: Integer): Integer;
begin
  Result := AValue + 6 * (AValue div 10);
end;

function BCDToInt(AValue: Integer): Integer;
begin
  Result := AValue - 6 * (AValue div 16);
end;

function EncodeTimeCode(const AHour, AMinute, ASecond, AFrame: Integer; AExtended: Boolean = False): String;
begin
  if AExtended then
    Result := Format('%.d:%.2d:%.2d:%.2d', [AHour, AMinute, ASecond, AFrame])
  else
    Result := Format('%.2d:%.2d:%.2d:%.2d', [AHour, AMinute, ASecond, AFrame]);
end;

procedure DecodeTimeCode(ATimeCode: String; out AHour, AMinute, ASecond, AFrame: Integer);
begin
  AHour   := 0;
  AMinute := 0;
  ASecond := 0;
  AFrame  := 0;
  if Length(ATimeCode) < 11 then exit;

  if (ATimeCode[1] in ['0'..'9']) and (ATimeCode[2] in ['0'..'9']) then
    AHour   := StrToInt(Copy(ATimeCode, 1, 2));
  if (ATimeCode[4] in ['0'..'5']) and (ATimeCode[5] in ['0'..'9']) then
    AMinute := StrToInt(Copy(ATimeCode, 4, 2));
  if (ATimeCode[7] in ['0'..'5']) and (ATimeCode[8] in ['0'..'9']) then
    ASecond := StrToInt(Copy(ATimeCode, 7, 2));
  if (ATimeCode[10] in ['0'..'5']) and (ATimeCode[11] in ['0'..'9']) then
    AFrame  := StrToInt(Copy(ATimeCode, 10, 2));

{  try
    Hour   := StrToInt(Copy(TimeCode, 1, 2));
    Minute := StrToInt(Copy(TimeCode, 4, 2));
    Second := StrToInt(Copy(TimeCode, 7, 2));
  except
    Hour   := 0;
    Minute := 0;
    Second := 0;
    Frame  := 0;
  end; }
end;

function FrameToTimeCode(AFrame: Integer; ADropFrame: Boolean): String;
const
  INIT_TIMECODE = '00:00:00:00';
var
  HH, MI, SS, FF: Cardinal;
  F: Cardinal;
begin
  F := AFrame;
{  if DropFrame then
  begin
    HH := F div Round(3600 * 29.97);
    F  := F mod Round(3600 * 29.97);
    MI := F div Round(60 * 29.97);
    F  := F mod Round(60 * 29.97);
    SS := F div Round(29.97);
    F  := F mod Round(29.97);
    FF := F;
  end
  else
  begin
    HH := F div 108000;
    F  := F mod 108000;
    MI := F div 1800;
    F  := F mod 1800;
    SS := F div 30;
    F  := F mod 30;
    FF := F;
  end;
}

  if F <= 0 then
  begin
    Result := INIT_TIMECODE;
    exit;
  end;

  if ADropFrame then
  begin
    HH := (F div 107892);
    F  := (F mod 107892);
    MI := (F div 1800);
    F  := (F mod 1800) + (((MI - (MI div 10)) * 2));
    if (F >= 1800) then
    begin
      Inc(MI);
      Dec(F, 1800);
      if MI > 60 then
      begin
        Inc(HH);
        Dec(MI, 60);
      end;

      if ((MI mod 10) <> 0) then F := F + 2;
    end;
  end
  else
  begin
    HH := (F div 108000);
    F  := (F mod 108000);
    MI := (F div 1800);
    F  := (F mod 1800);
  end;
  SS := (F div 30);
  F  := (F mod 30);
  FF := (F);

  Result := Format('%.2d:%.2d:%.2d:%.2d', [HH, MI, SS, FF]);

  exit;


  HH := (F div 108000);
  F  := (F mod 108000);
  MI := (F div 1800);
  if ADropFrame then
  begin
    F  := (F mod 1800) + (((MI - (MI div 10)) * 2));
    if F >= 1800 then
    begin
      Inc(MI);
      F := (F mod 1800);
      if MI > 60 then
      begin
        Inc(HH);
        Dec(MI, 60);
      end;
    end;
  end
  else F := (F mod 1800);
  SS := (F div 30);
  F  := (F mod 30);
  FF := (F);

  Result := Format('%.2d:%.2d:%.2d:%.2d', [HH, MI, SS, FF]);
end;

function TimeCodeToFrame(ATC: String; ADropFrame: Boolean): Integer;
var
  HH, MI, SS, FF: Cardinal;
begin
  try
    HH := 0;
    MI := 0;
    SS := 0;
    FF := 0;
    if Length(ATC) < 11 then exit;

    if (ATC[1] in ['0'..'9']) and (ATC[2] in ['0'..'9']) then
      HH := StrToInt(Copy(ATC, 1, 2));
    if (ATC[4] in ['0'..'5']) and (ATC[5] in ['0'..'9']) then
      MI := StrToInt(Copy(ATC, 4, 2));
    if (ATC[7] in ['0'..'5']) and (ATC[8] in ['0'..'9']) then
      SS := StrToInt(Copy(ATC, 7, 2));
    if (ATC[10] in ['0'..'5']) and (ATC[11] in ['0'..'9']) then
      FF := StrToInt(Copy(ATC, 10, 2));

{    HH := StrToInt(Copy(TC, 1, 2));
    MI := StrToInt(Copy(TC, 4, 2));
    SS := StrToInt(Copy(TC, 7, 2));
    FF := StrToInt(Copy(TC, 10, 2)); }

//    Result := Round(HH * (3600 * 29.97) + MI * (60 * 29.97) + SS * 29.97 + FF);

//    if DropFrame then Result := (HH * 107892) + ((MI * 1800) - ((MI - (MI div 10)) * 2)) + (SS * 30) + FF
    if ADropFrame then
    begin
//      if ((MI mod 10) <> 0) and (FF <= 1) then FF := 1;
//      Result := (HH * 108000) + ((MI * 1800) - ((MI - (MI div 10)) * 2)) + (SS * 30) + FF;
      Result := (HH * 107892) + ((MI * 1800) - ((MI - (MI div 10)) * 2)) + (SS * 30) + FF;
    end
    else Result := (HH * 108000) + (MI * 1800) + (SS * 30) + FF;
  except
    Result := 0;
  end;
end;

function GetTimeCodeDuration(ATC1, ATC2: String; ADropFrame: Boolean): String;
var
  TTC: String;
{  Hour1, Minute1, Second1, Frame1: Integer;
  Hour2, Minute2, Second2, Frame2: Integer;
}
  F, Frame1, Frame2: Integer;
begin
  if ATC1 > ATC2 then
  begin
    TTC := ATC1;
    ATC1 := ATC2;
    ATC2 := TTC;
  end;

{  DecodeTimeCode(TC1, Hour1, Minute1, Second1, Frame1);
  DecodeTimeCode(TC2, Hour2, Minute2, Second2, Frame2);

  Frame1  := Frame2 - Frame1;
  Second1 := Second2 - Second1;
  Minute1 := Minute2 - Minute1;
  Hour1   := Hour2 - Hour1;

  if Frame1 < 0 then
  begin
    Dec(Second1);
    Inc(Frame1, 30);
  end;

  if Second1 < 0 then
  begin
    Dec(Minute1);
    Inc(Second1, 60);
  end;

  if Minute1 < 0 then
  begin
    Dec(Hour1);
    Inc(Minute1, 60);
  end;

//  if (DropFrame) and (Frame1 = 0) and (Second1 = 0) and ((Minute1 mod 10) <> 0) then Frame1 := 2;
  if (DropFrame) and (Frame1 = 0) and (Second1 = 0) and ((Minute1 mod 10) <> 0) then
  begin
    Dec(Frame1, 2);
    if Frame1 < 0 then
    begin
      Dec(Second1);
      Inc(Frame1, 30);
    end;
  end;

  Result := Format('%.2d:%.2d:%.2d:%.2d', [Hour1, Minute1, Second1, Frame1]); }

  Frame1 := TimeCodeToFrame(ATC1, ADropFrame);
  Frame2 := TimeCodeToFrame(ATC2, ADropFrame);
  F := Frame2 - Frame1;
  if F < 0 then F := Abs(F);
  Result := FrameToTimeCode(F, ADropFrame);
end;

function GetPlusTimeCode(ATC1, ATC2: String; ADropFrame: Boolean): String;
var
{  Hour1, Minute1, Second1, Frame1: Integer;
  Hour2, Minute2, Second2, Frame2: Integer;
}
  F, Frame1, Frame2: Integer;
begin
{  DecodeTimeCode(TC1, Hour1, Minute1, Second1, Frame1);
  DecodeTimeCode(TC2, Hour2, Minute2, Second2, Frame2);

  Hour1   := Hour1 + Hour2;
  Minute1 := Minute1 + Minute2;
  Second1 := Second1 + Second2;
  Frame1  := Frame1 + Frame2;

  if Frame1 >= 30 then
  begin
    Inc(Second1);
    Dec(Frame1, 30);
  end;

  if Second1 >= 60 then
  begin
    Inc(Minute1);
    Dec(Second1, 60);
  end;

  if Minute1 >= 60 then
  begin
    Inc(Hour1);
    Dec(Minute1, 60);
  end;

//  if (DropFrame) and (Frame1 = 0) and (Second1 = 0) and ((Minute1 mod 10) <> 0) then Frame1 := 2;
  if (DropFrame) and (Second1 = 0) and ((Minute1 mod 10) <> 0) then
  begin
    Inc(Frame1, 2);
    if Frame1 >= 30 then
    begin
      Inc(Second1);
      Dec(Frame1, 30);
    end;
  end;

  Result := Format('%.2d:%.2d:%.2d:%.2d', [Hour1, Minute1, Second1, Frame1]); }

  Frame1 := TimeCodeToFrame(ATC1, ADropFrame);
  Frame2 := TimeCodeToFrame(ATC2, ADropFrame);

  F := Frame1 + Frame2;
  if F < 0 then F := 0;
  Result := FrameToTimeCode(F, ADropFrame);
end;

function GetMinusTimeCode(ATC1, ATC2: String; ADropFrame: Boolean): String;
var
{  Hour1, Minute1, Second1, Frame1: Integer;
  Hour2, Minute2, Second2, Frame2: Integer;
}
  F, Frame1, Frame2: Integer;
begin
{  DecodeTimeCode(TC1, Hour1, Minute1, Second1, Frame1);
  DecodeTimeCode(TC2, Hour2, Minute2, Second2, Frame2);

  Hour1   := Hour1 - Hour2;
  Minute1 := Minute1 - Minute2;
  Second1 := Second1 - Second2;
  Frame1  := Frame1 - Frame2;

  if Frame1 < 0 then
  begin
    Dec(Second1);
    Inc(Frame1, 30);
  end;

  if Second1 < 0 then
  begin
    Dec(Minute1);
    Inc(Second1, 60);
  end;

  if Minute1 < 0 then
  begin
    Dec(Hour1);
    Inc(Minute1, 60);
  end;

//  if (DropFrame) and (Frame1 = 0) and (Second1 = 0) and ((Minute1 mod 10) <> 0) then Frame1 := 2;
  if (DropFrame) and (Frame1 = 0) and (Second1 = 0) and ((Minute1 mod 10) <> 0) then
  begin
    Dec(Frame1, 2);
    if Frame1 < 0 then
    begin
      Dec(Second1);
      Inc(Frame1, 30);
    end;
  end;

  Result := Format('%.2d:%.2d:%.2d:%.2d', [Hour1, Minute1, Second1, Frame1]); }

  Frame1 := TimeCodeToFrame(ATC1, ADropFrame);
  Frame2 := TimeCodeToFrame(ATC2, ADropFrame);

  F := Frame1 - Frame2;
  if F < 0 then F := 0;
  Result := FrameToTimeCode(F, ADropFrame)
end;

function IsValidTimeCode(ATC: String; ADropFrame: Boolean): Boolean;
var
  Hour, Minute, Second, Frame: Integer;
begin
  Result := False;

  Hour   := 0;
  Minute := 0;
  Second := 0;
  Frame  := 0;
  if Length(ATC) < 11 then exit;

  if (ATC[1] in ['0'..'9']) and (ATC[2] in ['0'..'9']) then
    Hour   := StrToInt(Copy(ATC, 1, 2));
  if (ATC[4] in ['0'..'5']) and (ATC[5] in ['0'..'9']) then
    Minute := StrToInt(Copy(ATC, 4, 2));
  if (ATC[7] in ['0'..'5']) and (ATC[8] in ['0'..'9']) then
    Second := StrToInt(Copy(ATC, 7, 2));
  if (ATC[10] in ['0'..'5']) and (ATC[11] in ['0'..'9']) then
    Frame  := StrToInt(Copy(ATC, 10, 2));

{  try
    Hour   := StrToInt(Copy(TC, 1, 2));
    Minute := StrToInt(Copy(TC, 4, 2));
    Second := StrToInt(Copy(TC, 7, 2));
    Frame  := StrToInt(Copy(TC, 10, 2));
  except
    Result := False;
    exit;
  end; }

  if (Hour < 0) or (Minute < 0) or (Second < 0) or (Frame < 0) then
  begin
    exit;
  end;

  if (Minute >= 60) or (Second >= 60) or (Frame >= 30) then
  begin
    exit;
  end;

  if (ADropFrame) and (((Minute mod 10) <> 0) and (Second = 0) and (Frame in [0..1])) then
  begin
    exit;
  end;

  Result := True;
end;

{ TBaseSerial }

constructor TBaseSerial.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FCriticalSection := TCriticalSection.Create;
  // Initialize to default AValues
  FComPortNames := TStringList.Create;
  GetComPortNames(FComPortNames);

  FControlType            := ctSerial;
  FControlIP              := '';
  FControlPort            := 0;
  FSocketHandle           := INVALID_SOCKET;
  FSocketConnected        := False;

  FComPortHandle          := 0;       // Not connected
  FComPortBaudRate        := br38400; // 38400 bauds
  FComPortDataBits        := db8Bits; // 8 data bits
  FComPortStopBits        := sb1Bits; // 1 stop bit
  FComPortParity          := ptNone;  // no parity
  FComPortHwHandshaking   := hhNone;  // no hardware handshaking
  FComPortSwHandshaking   := shNone;  // no software handshaking
  FComPortInBufSize       := 2048;    // input buffer of 2048 bytes
  FComPortOutBufSize      := 2048;    // output buffer of 2048 bytes
  FOutputTimeout          := 4000;    // output timeout - 4000ms
  FEnableDTROnOpen        := False;   // DTR high on connect
//  FSyncMode               := True;
  FTimeOut                := 1000;
  FErrorString            := '';
//  FDropFrame              := True;

  FFlushBuf               := True;


  FLogEnabled   := False;
  FLogIsHexcode := True;
  FLogPath      := ExtractFilePath(Application.ExeName);
  FLogExt       := '.log';

  FLastResult := D_OK;
  FReceivedData := '';
  FReceivedBuffer := '';

  FSyncMsgEvent := CreateEvent(nil, True, False, nil);
end;

destructor TBaseSerial.Destroy;
begin
  if FCriticalSection <> nil then
  begin
    FreeAndNil(FCriticalSection);
  end;

  if (Connected) then Disconnect;

  if (Assigned(FAsyncThread)) then
  begin
    FAsyncThread.Close;
    FAsyncThread.WaitFor;
    FreeAndNil(FAsyncThread);
  end;

  FreeAndNil(FComPortNames);

  CloseHandle(FSyncMsgEvent);

  inherited Destroy;
end;

function TBaseSerial.GetRxCount: Cardinal;
begin
  if FAsyncThread <> nil then Result := FAsyncThread.FRxCount
  else Result := 0;
end;

function TBaseSerial.GetTxEmpty: Boolean;
begin
  if FAsyncThread <> nil then Result := FAsyncThread.FTxEmpty
  else Result := True;
end;

function TBaseSerial.WriteLog(ALogKind: TLogKind; ABuffer: PAnsiChar; ADataSize: Integer; AHexcode: Boolean): Integer;
var
  FilePath, FileName: String;
  FileStream: TFileStream;
  S: AnsiString;
//  F: TextFile;
  WriteBuf: Pointer;
begin
  FCriticalSection.Enter;
  try
    Result := -1;

    FilePath := Format('%s%s\', [FLogPath, FormatDateTime('YYYY-MM', Date)]);
    FileName := Format('%s%s_%s', [FilePath, FormatDateTime('YYYY-MM-DD', Date), FLogExt]);

    // 디렉토리 있는지 확인 없으면 생성
    if (not DirectoryExists(FilePath)) then ForceDirectories(FilePath);

    try
      if (not FileExists(FileName)) then
        FileStream := TFileStream.Create(FileName, fmCreate or fmShareDenyNone)
      else FileStream := TFileStream.Create(FileName, fmOpenWrite or fmShareDenyNone);
      try
        FileStream.Position := FileStream.Size;

        S := FormatDateTime('[hh:nn:ss:zzz]', Now);
        case ALogKind of
          lkSend: S := S + ' [S] ';
          lkReceive: S := S + ' [R] ';
          else S := S + ' [T] ';
        end;

        if (AHexcode) then
          S := S + PAnsiCharToHexCode(ABuffer, ADataSize) + #13#10
        else
          S := S + PAnsiCharToString(ABuffer, ADataSize) + #13#10;

        GetMem(WriteBuf, Length(S) + 1);
        try
          FillChar(WriteBuf^, Length(S) + 1, #0);
          Move(S[1], WriteBuf^, Length(S));
          FileStream.Write(WriteBuf^, Length(S));
          ShowMessage(AnsiString(WriteBuf^));
        finally
          FreeMemory(WriteBuf);
        end;
      finally
        FreeAndNil(FileStream);
      end;
      Result := Length(S);
    except
  //    MessageBeep(MB_ICONERROR);
  //    Application.MessageBox(PChar(Format('Cannot create log file %s', [FileName])), PChar(Application.Title), MB_OK or MB_ICONERROR);
    end;
  finally
    FCriticalSection.Leave;
  end;
{  FCriticalSection.Enter;
  // 디렉토리 있는지 확인 없으면 생성
  if not DirectoryExists(FilePath) then ForceDirectories(FilePath);

  S := FormatDateTime('[hh:nn:ss:zzz]', Now);
  case ALogKind of
    lkSend: S := S + ' [S] ';
    lkReceive: S := S + ' [R] ';
    else S := S + ' [T] ';
  end;

  S := S + PAnsiCharToHexCode(ABuffer, ADataSize);
  try
    AssignFile(F, FileName);
    if FileExists(FileName) then Append(F)
    else Rewrite(F);
    Writeln(F, S);
  finally
    Closefile(F);
  end;
  FCriticalSection.Leave; }
end;

function TBaseSerial.WaitForResult(ACount: Integer; AHandles: array of THandle; AWaitAll: Boolean; AMillSecond, AWakeMask: DWORD): DWORD;
var
  M: DWORD;
  Msg: TMsg;
  R: Boolean;

  SystemTime: _SYSTEMTIME;
  CurrentMSec: Int64;
  StopMSec: Int64;
  A: THandle;
  I: Integer;
begin
  GetSystemTime(SystemTime);
  StopMSec := (SystemTime.wHour * 3600 + SystemTime.wMinute * 60 + SystemTime.wSecond) * 1000 +
              SystemTime.wMilliseconds + AMillSecond;

  while True do
  begin
    GetSystemTime(SystemTime);
    CurrentMSec := (SystemTime.wHour * 3600 + SystemTime.wMinute * 60 + SystemTime.wSecond) * 1000 + SystemTime.wMilliseconds;
    if (StopMSec - CurrentMSec) < 0 then
    begin
      Result := WAIT_TIMEOUT;
      break;
    end;

    M := MsgWaitForMultipleObjects(ACount, AHandles, False, StopMSec - CurrentMSec, QS_PAINT);//WakeMask);
{M := MsgWaitForMultipleObjects(1, EventArray, False, StopMSec - CurrentMSec, QS_ALLINPUT or
QS_HOTKEY	or
QS_INPUT or
QS_KEY or
QS_MOUSE or
QS_MOUSEBUTTON or
QS_MOUSEMOVE or
QS_PAINT or
QS_POSTMESSAGE or
QS_SENDMESSAGE or
QS_TIMER); }
    if M = WAIT_TIMEOUT then
    begin
      Result := M;
      break
    end
    else if M = WAIT_OBJECT_0 then
    begin
      Result := WAIT_OBJECT_0;
      break;
    end
    else if M = WAIT_OBJECT_0 + ACount then
    begin
{      while(PeekMessage(msg, 0, 0, 0, PM_REMOVE)) do
      begin
        DispatchMessage(msg);
      end; }
      R := GetMessage(Msg, 0, 0, 0);
      if (R) then
      begin
        TranslateMessage(Msg);
        DispatchMessage(Msg);
      end;
//      Application.ProcessMessages;
      Result := M;
    end;
  end;
end;

{procedure TBaseSerial.SockWndProc(var Msg: TMessage);
//var
//  PacketSize: Integer;
begin
  try
  if Msg.Msg = WM_SOCKET_MESSAGE then
  begin
    with TSocketMsg(Msg) do
    begin
      if AsyncSelectEventError(Socket, SelectEvent, SelectError) = -1 then
      begin
        case SelectEvent of
          FD_CONNECT:
            begin
              FSockConnected := False;
//              ShowMessage('error');
            end;
          FD_READ:
              ShowMessage('error');
          end;
{        case SelectEvent of
          FD_CONNECT: SetTaskFileStatus(FFileIndex, fsErrorAbort, WMFError_ConnectSelectEvent);
          FD_CLOSE: SetTaskFileStatus(FFileIndex, fsErrorAbort, WMFError_CloseSelectEvent);
          FD_ACCEPT: SetTaskFileStatus(FFileIndex, fsErrorAbort, WMFError_AcceptSelectEvent);
          FD_READ: SetTaskFileStatus(FFileIndex, fsErrorAbort, WMFError_ReadSelectEvent);
          FD_WRITE: SetTaskFileStatus(FFileIndex, fsErrorAbort, WMFError_WriteSelectEvent);
        end; }
{        exit;
      end;

      case SelectEvent of
        FD_CLOSE:
          begin
            if FSyncMsgEvent > 0 then
            begin
              try
                CloseHandle(FSyncMsgEvent);
                FSyncMsgEvent := 0;
              except
                OutputDebugString(PAnsiChar('SyncMsgEvent Handle Close Fail' + IntToStr(FSyncMsgEvent)));
              end;
            end;
          end;
        FD_CONNECT:
          begin
            FSyncMsgEvent := CreateEvent(nil, True, False, nil);
            FSockConnected := True;
//            ShowMessage('1');
          end;
        FD_READ:
          begin
            SetEvent(FAsyncThread.SockReadEvent);
          end;
        FD_WRITE:
          begin
          end;
      end;
    end;
    Msg.Result := 1;
  end;
  Msg.Result := DefWindowProc(FSockWndHandle, Msg.Msg, Msg.wParam, Msg.lParam);
  except
    if Assigned(ApplicationHandleException) then
      ApplicationHandleException(Self);
  end;
end; }

function TBaseSerial.MsgWaitForHandles(ACount: Cardinal; AHandles: array of THandle; ATimeOut, AWakeMask: DWORD): DWORD;
var
  // wait for the event to be signalled, or for the timeout (in MS) to expire.
  // allow SENT messages to be processed while we wait
  Wait, StartTime: DWord;
  // set the waiting period.
  WaitTime: Dword;
  Msg: TMsg;
  Elapsed: DWord;
begin
  WaitTime := ATimeout;

  // the timeout will eventually run down as we iterate processing messages.
  // grab the start time so that we can calculate elapsed times.
  if (WaitTime <> INFINITE) then
    StartTime := timeGetTime
  else
    StartTime := 0; // don't generate compiler hint

  repeat
    Wait := MsgWaitForMultipleObjects(ACount, AHandles, False, WaitTime, AWakeMask);
    if (Wait = WAIT_OBJECT_0 + ACount) then
    begin
      while (PeekMessage(Msg, 0, 0, 0, PM_REMOVE)) do
      begin
        TranslateMessage(Msg);
        DispatchMessage(Msg);
//        Application.HandleMessage;
      end;
      //==PeekMessage(Msg, 0, 0, 0, PM_REMOVE);

	    // If we have an explicit length of time to wait calculate the next wake up point - which might be now.
	    // If dwTimeout is INFINITE, it stays INFINITE
	    if (WaitTime <> INFINITE) then
      begin
    	  Elapsed := timeGetTime - StartTime;
        if (Elapsed >= ATimeout) then
          WaitTime := 0
        else // wake up with WAIT_TIMEOUT
    	    WaitTime := ATimeout - Elapsed;
      end;
    end;
  until (Wait <> WAIT_OBJECT_0 + ACount);

  // return True if we woke on the event handle, FALSE if we timed out.
  Result := Wait; //(Wait = WAIT_OBJECT_0);
end;

procedure TBaseSerial.SetControlType(AValue: TControlType);
begin
  if Connected then exit;
  FControlType := AValue;
end;

procedure TBaseSerial.SetComPort(AValue: Integer);
begin
  if Connected then exit;
  FComPort := AValue;
end;

procedure TBaseSerial.SetComPortBaudRate(AValue: TComPortBaudRate);
begin
  FComPortBaudRate := AValue;
  if Connected then ApplyComPortSettings;
end;

procedure TBaseSerial.SetComPortDataBits(AValue: TComPortDataBits);
begin
  FComPortDataBits := AValue;
  if Connected then ApplyComPortSettings;
end;

procedure TBaseSerial.SetComPortStopBits(AValue: TComPortStopBits);
begin
  FComPortStopBits := AValue;
  if Connected then ApplyComPortSettings;
end;

procedure TBaseSerial.SetComPortParity(AValue: TComPortParity);
begin
  FComPortParity := AValue;
  if Connected then ApplyComPortSettings;
end;

procedure TBaseSerial.SetComPortHwHandshaking(AValue: TComPortHwHandshaking);
begin
  FComPortHwHandshaking := AValue;
  if Connected then ApplyComPortSettings;
end;


procedure TBaseSerial.SetComPortSwHandshaking(AValue: TComPortSwHandshaking);
begin
  FComPortSwHandshaking := AValue;
  if Connected then ApplyComPortSettings;
end;

procedure TBaseSerial.SetComPortInBufSize(AValue: Word);
begin
  if Connected then exit;
  FComPortInBufSize := AValue;
end;

procedure TBaseSerial.SetComPortOutBufSize(AValue: Word);
begin
  if not Connected then exit;
  FComPortOutBufSize := AValue;
end;

procedure TBaseSerial.SetTimeOut(const AValue: Word);
begin
  FTimeOut := AValue;
end;

{procedure TBaseSerial.SetSyncMode(AValue: Boolean);
begin
  FSyncMode := AValue;
end; }

{procedure TBaseSerial.SetDropFrame(const AValue: Boolean);
begin
  FDropFrame := AValue;
end;}

procedure TBaseSerial.SetOverlapMode(const AValue: Boolean);
begin
  if Connected then exit;
  FOverlapMode := AValue;
end;

const
  // Win32 COM Port Baud Rates
  Win32BaudRates: array[br2400..br115200] of DWORD =
    (CBR_2400, CBR_4800, CBR_9600, CBR_14400, CBR_19200,
     CBR_38400, CBR_56000, CBR_57600, CBR_115200);

procedure TBaseSerial.ApplyComPortSettings;
var
  DCB: TDCB;
  ComStat: TComStat;
  Errors: DWORD;
begin
  if not Connected then exit;

  FillChar(DCB, SizeOf(DCB), 0 );
//  GetCommState(FComPortHandle, DCB);
//  ClearCommError(FComPortHandle, Errors, @ComStat);

  DCB.DCBLength := SizeOf(DCB);                      // dcb structure size
  DCB.BaudRate  := Win32BaudRates[FComPortBaudRate]; // baud rate to use
  DCB.Flags     := DCB_Binary;
  if EnableDTROnOpen then DCB.Flags := DCB.Flags or DCB_DtrControlEnable;

  case FComPortHwHandshaking of // Type of hardware handshaking to use
    // No hardware handshaking
    hhNone: ;
    // No hardware handshaking but set RTS high and leave it high
    hhNoneRtsOn: DCB.Flags := DCB.Flags or DCB_RtsControlEnable;
    // Rts/Cts (request-to-send/clear-to-send) hardware handshaking
    hhRtsCts: DCB.Flags := DCB.Flags or DCB_OutxCtsFlow or DCB_RtsControlHandshake;
  end;

  case FComPortSwHandshaking of // Type of software handshaking to use
    // No software handshaking
    shNone: ;
    // XOn/XOff handshaking
    shXOnXOff: DCB.Flags := DCB.Flags or DCB_OutX or DCB_InX;
  end;

  DCB.ByteSize := Byte(Ord(FComPortDataBits) + 5);// * 2;
  DCB.Parity   := Ord(FComPortParity);    // type of parity to use
  DCB.StopBits := Ord(FComPortStopbits);  // how many stop bits to use

//  if not SetCommState(FComPortHandle, DCB) then ShowMessage('1');
  SetCommState(FComPortHandle, DCB);

  SetupComm(FComPortHandle, FComPortInBufSize, FComPortOutBufSize);

  { Flush buffers }
  FlushBuffers(True, True);
end;

function TBaseSerial.Connect: Boolean;
var
  ComPortName: String;
  CommTimeouts: TCommTimeouts;

  WSAData: TWSAData;
  SockAddr: TSockAddrIn;
begin
  // Do nothing if already connected
  Result := Connected;
  if Result then exit;

  // Serial
  if (FControlType = ctSerial) then
  begin
    if FComPort = 0 then exit;

    ComPortName := 'COM' + IntToStr(FComPort);
    if Pos(ComPortName, FComPortNames.Text) <= 0 then exit;

    ComPortName := '\\.\'  + ComPortName;
    FComPortHandle := CreateFile(PChar(ComPortName),
                                 GENERIC_READ or GENERIC_WRITE,
                                 0, // Not shared
                                 nil, // No security attributes
                                 OPEN_EXISTING,
                                 FILE_FLAG_OVERLAPPED or FILE_ATTRIBUTE_NORMAL,
                                 0);

    if FComPortHandle = INVALID_HANDLE_VALUE then
    begin
//      FSyncMsgEvent  := 0;
      FComPortHandle := 0;
      exit;
    end
    else Result := True;

//{   if FSyncMode then }FSyncMsgEvent := CreateEvent(nil, True, False, nil);

    ApplyComPortSettings;

    // Setup timeouts: we disable timeouts because we are polling the com port!
    // Specifies the maximum time, in milliseconds,
    // allowed to elapse between the arrival of two
    // characters on the communications line
//    CommTimeouts.ReadIntervalTimeout := 1;
    CommTimeouts.ReadIntervalTimeout := MAXDWORD;//$ffffffff;
    // Specifies the multiplier, in milliseconds,
    // used to calculate the total time-out period
    // for read operations.
//    CommTimeouts.ReadTotalTimeoutMultiplier := $ffffffff;
    CommTimeouts.ReadTotalTimeoutMultiplier := 0;
    // Specifies the constant, in milliseconds,
    // used to calculate the total time-out period
    // for read operations.
//    CommTimeouts.ReadTotalTimeoutConstant := $14000000;
    CommTimeouts.ReadTotalTimeoutConstant := 1;
    // Specifies the multiplier, in milliseconds,
    // used to calculate the total time-out period
    // for write operations.
    CommTimeouts.WriteTotalTimeoutMultiplier := 0;
    // Specifies the constant, in milliseconds,
    // used to calculate the total time-out period
    // for write operations.
    CommTimeouts.WriteTotalTimeoutConstant := 0;

    SetCommTimeOuts(FComPortHandle, CommTimeouts);
    SetCommMask(FComPortHandle, EV_RXCHAR);

    // if FSyncmode is False, call event OnDataAvailbe
{    FAsyncThread := TAsyncThread.Create(True);
    with FAsyncThread do
    begin
      ASyncComHandle := FComPortHandle;
      FBaseSerial := Self;
      FreeOnTerminate := True;
      Start;
    end; }

    FAsyncThread := TAsyncThread.Create(Self, FComPortHandle);
    FAsyncThread.Start;
  end
  // Socket
  else
  begin
    FSocketConnected := False;

    if WSAStartup($0202, WSAData) <> 0 then
    begin
//      FSyncMsgEvent  := 0;
      FSocketHandle  := 0;
      exit;
    end;

    if (FControlType = ctTCP) then
    begin
      FSocketHandle := {SocketInit(AF_INET, SOCK_STREAM, IPPROTO_IP);//} WSASocket(AF_INET, SOCK_STREAM, IPPROTO_IP, nil, 0, WSA_FLAG_OVERLAPPED);
    end
    else
    begin
      FSocketHandle := WSASocket(AF_INET, SOCK_DGRAM, IPPROTO_IP, nil, 0, WSA_FLAG_OVERLAPPED);
    end;

    if FSocketHandle = INVALID_SOCKET then
      exit;

//    FSyncMsgEvent := CreateEvent(nil, True, False, nil);

    FSocketEvent := CreateEvent(nil, False, False, nil);
{    FSocketEvent := WSACreateEvent;
    if FSocketEvent = WSA_INVALID_EVENT then
      exit; }

    if WSAEventSelect(FSocketHandle, FSocketEvent, FD_CONNECT or FD_READ or FD_CLOSE) = SOCKET_ERROR then
      exit;

    if (FControlType = ctTCP) then
    begin
      SockAddr.sin_family := AF_INET;
      SockAddr.sin_addr.s_addr := inet_addr(PAnsiChar(FControlIP));
      SockAddr.sin_port := htons(FControlPort);

  //    if ConnectInit(FSocketHandle, SockAddr, Sizeof(SockAddr)) = SOCKET_ERROR then
      if WSAConnect(FSocketHandle, TSockAddr(SockAddr), Sizeof(SockAddr), nil, nil, nil, nil) = SOCKET_ERROR then
        if WSAGetLastError <> WSAEWOULDBLOCK then
          exit;
    end
    else
    begin
      FSocketConnected := True;
    end;

    // if FSyncmode is False, call event OnDataAvailbe
{    FAsyncThread := TAsyncThread.Create(True);
    with FAsyncThread do
    begin
      ASyncComHandle := FSocketHandle;
      FBaseSerial := Self;
      FreeOnTerminate := True;
      Start;
    end; }

    FAsyncThread := TAsyncThread.Create(Self, FSocketHandle);
    FAsyncThread.Start;
    Sleep(100);

    Result := FSocketConnected;
  end;
end;

function TBaseSerial.Disconnect: Boolean;
begin
  Result := False;

  if Connected then
  begin
//    if Assigned(FAsyncThread) then SetEvent(FAsyncThread.FCloseEvent);

{        if FSyncMsgEvent > 0 then
        begin
          try
            CloseHandle(FSyncMsgEvent);
            FSyncMsgEvent := 0;
          except
            OutputDebugString(PWideChar('SyncMsgEvent Handle Close Fail' + IntToStr(FSyncMsgEvent)));
          end;
        end; }
    if (ControlType = ctSerial) then
    begin

{//      if FSyncMode then
      begin
        if FSyncMsgEvent > 0 then
        begin
          try
            CloseHandle(FSyncMsgEvent);
            FSyncMsgEvent := 0;
          except
            OutputDebugString(PAnsiChar('SyncMsgEvent Handle Close Fail' + IntToStr(FSyncMsgEvent)));
          end;
        end;
      end; }
      FComPortHandle := 0;
    end
    else
    begin
//      WSACloseEvent(FSocketEvent);
      FSocketHandle := 0;
      FSocketConnected := False;

      WSACleanup;
    end;
    Result := True;
  end;

  if (Assigned(FAsyncThread)) then
  begin
    FAsyncThread.Close;
    FAsyncThread.WaitFor;
    FreeAndNil(FAsyncThread);
  end;
end;

function TBaseSerial.Connected: Boolean;
begin
  if (ControlType = ctSerial) then
    Result := FComPortHandle > 0
  else Result := (FSocketHandle <> INVALID_SOCKET) and FSocketConnected;
end;

procedure TBaseSerial.FlushBuffers(AInBuf, AOutBuf: Boolean);
var
  Action: DWORD;
begin
  if not Connected then exit;

  Action := 0;
  if AOutBuf then Action := Action or PURGE_TXABORT or PURGE_TXCLEAR;
  if AInBuf then Action := Action or PURGE_RXABORT or PURGE_RXCLEAR;

  PurgeComm(FComPortHandle, Action);
end;

function TBaseSerial.OutFreeSpace: Word;
var
  ComStat: TComStat;
  Errors: DWORD;
begin
  if not Connected then Result := 65535
  else
  begin
    ClearCommError(FComPortHandle, Errors, @ComStat);
    Result := FComPortOutBufSize - ComStat.cbOutQue;
  end;
end;

procedure TBaseSerial.ToggleDTR(AOnOff: Boolean);
const
  Funcs: array[Boolean] of Integer = (CLRDTR, SETDTR);
begin
  if Connected then
    EscapeCommFunction(FComPortHandle, Funcs[AOnOff]);
end;

procedure TBaseSerial.ToggleRTS(AOnOff: Boolean);
const
  Funcs: array[Boolean] of Integer = (CLRRTS, SETRTS);
begin
  if Connected then
    EscapeCommFunction(FComPortHandle, Funcs[AOnOff]);
end;

function TBaseSerial.SendData(ABuffer: Pointer; ADataSize: Integer): Integer;
var
  Sent: DWORD;
  R: DWORD;
  Overlapped: TOverlapped;
  Buf: WSABUF;

  ServerSockAddr: TSockAddrIn;

  ComStat: TComStat;
  Errors: DWORD;
begin
  Result := -1;

//    if (LogEnabled) then WriteLog(lkSend, ABuffer, Result, LogIsHexcode);
  if not Connected then exit;
  if not Assigned(FAsyncThread) then exit;

  FillChar(Overlapped, SizeOf(Overlapped), 0);
  Overlapped.hEvent := CreateEvent(nil, True, False, nil);
  if Overlapped.hEvent = 0 then exit;

  try
    if (FControlType = ctSerial) then
    begin
      if FlushBuf then FlushBuffers(True, True);
      ToggleDTR(True);
      ToggleRTS(True);

      if not WriteFile(FAsyncThread.FASyncComHandle, ABuffer^, ADataSize, Sent, @Overlapped) then
  //    if not WriteFile(FAsyncThread.ASyncComHandle, Buffer^, DataSize, Sent, nil) then
      begin
    {    Errors := GetLastError;
        while (Errors = ERROR_IO_PENDING) or (Errors = ERROR_IO_INCOMPLETE) do
        begin
          if GetOverlappedResult(FASyncThread.ASyncComHandle, Overlapped, Sent, True) then
          begin
            Result := Sent;
            break;
    //            WriteLog(lkSend, Buffer, Sent);
          end;
          Errors := GetLastError;
        end; }

        if (GetLastError <> ERROR_IO_PENDING) then exit;

    //    R := WaitForResult(1, [Overlapped.hEvent], False, FTimeOut, QS_PAINT or QS_SENDMESSAGE);
//        R := WaitForSingleObject(Overlapped.hEvent, INFINITE);
        R := WaitForSingleObject(Overlapped.hEvent, FTimeOut);
        case R of
          WAIT_OBJECT_0:
          begin
            if GetOverlappedResult(FASyncThread.FASyncComHandle, Overlapped, Sent, False) then
  //          if GetOverlappedResult(FASyncThread.ASyncComHandle, Overlapped, Sent, True) then
            begin
              Result := Sent;
//              if LogEnabled then WriteLog(lkSend, ABuffer, Sent);
            end;
          end;
        end;

        ClearCommError(FComPortHandle, Errors, @ComStat);
      end
      else
      begin
        Result := Sent;
//        if LogEnabled then WriteLog(lkSend, ABuffer, Sent);
      end;
    end
    else
    begin
      Buf.buf := ABuffer;
      Buf.len := ADataSize;

      if (ControlType = ctTCP) then
      begin
        if WSASend(FAsyncThread.FASyncComHandle, @Buf, 1, Sent, 0, @Overlapped, nil) = SOCKET_ERROR then
        begin
          if (WSAGetLastError <> WSA_IO_PENDING) then exit;
//          R := WaitForSingleObject(Overlapped.hEvent, INFINITE);
          R := WaitForSingleObject(Overlapped.hEvent, FTimeOut);
          case R of
            WAIT_OBJECT_0:
            begin
              if GetOverlappedResult(FASyncThread.FASyncComHandle, Overlapped, Sent, False) then
    //          if GetOverlappedResult(FASyncThread.ASyncComHandle, Overlapped, Sent, True) then
              begin
                Result := Sent;
  //              if LogEnabled then WriteLog(lkSend, ABuffer, Sent);
              end;
            end;
          end;
        end
        else
        begin
          Result := Sent;
  //        if LogEnabled then WriteLog(lkSend, ABuffer, Sent);
        end;
      end
      else
      begin
        // Setup address structure
        FillChar(ServerSockAddr, SizeOf(TSockAddrIn), #0);
        ServerSockAddr.sin_family := AF_INET;
        ServerSockAddr.sin_port := htons(FControlPort);
        ServerSockAddr.sin_addr.s_addr := inet_addr(PAnsiChar(FControlIP));

        if WSASendTo(FAsyncThread.FASyncComHandle, @Buf, 1, @Sent, 0, TSockAddr(ServerSockAddr), SizeOf(TSockAddr), @Overlapped, nil) = SOCKET_ERROR then
        begin
          if (WSAGetLastError <> WSA_IO_PENDING) then exit;
//          R := WaitForSingleObject(Overlapped.hEvent, INFINITE);
          R := WaitForSingleObject(Overlapped.hEvent, FTimeOut);
          case R of
            WAIT_OBJECT_0:
            begin
              if GetOverlappedResult(FASyncThread.FASyncComHandle, Overlapped, Sent, False) then
    //          if GetOverlappedResult(FASyncThread.ASyncComHandle, Overlapped, Sent, True) then
              begin
                Result := Sent;
  //              if LogEnabled then WriteLog(lkSend, ABuffer, Sent);
              end;
            end;
          end;
        end
        else
        begin
          Result := Sent;
  //        if LogEnabled then WriteLog(lkSend, ABuffer, Sent);
        end;
      end;

{      Sent := SendInit(FSocketHandle, PChar(ABuffer), ADataSize, 0);

      if Sent = SOCKET_ERROR then
        if WSAGetLastError <> WSAEWOULDBLOCK then
        begin
          exit;
        end;

      Result := Sent;
      if LogEnabled then WriteLog(lkSend, ABuffer, Sent); }
    end;
    if (LogEnabled) and (Result > 0) then WriteLog(lkSend, ABuffer, Result, LogIsHexcode);
  finally
    CloseHandle(Overlapped.hEvent);
  end;
end;

function TBaseSerial.ReadData(ABuffer: Pointer; ADataSize: Integer): Integer;
const
  RECV_BUF_LEN = 1024;
var
  OverLapped: TOverLapped;
  R: Integer;
  Read: DWORD;
  Buf: WSABUF;
  RecvFlags: Cardinal;

  ClientAddrSize: Integer;
  ClientSockAddr: TSockAddrIn;

  ComStat: TComStat;
  Errors: DWORD;
begin
  Result := -1;
  if ADataSize <= 0 then exit;

//    if (LogEnabled) then WriteLog(lkReceive, ABuffer, ADataSize, LogIsHexcode);

  if not Connected then exit;
  if not Assigned(FAsyncThread) then exit;

  FillChar(OverLapped, SizeOf(OverLapped), 0);
  OverLapped.hEvent := CreateEvent(nil, True, False, nil);
  if OverLapped.hEvent = 0 then exit;

  try
    if (FControlType = ctSerial) then
    begin
      ToggleDTR(False);
      ToggleRTS(False);

      if not ReadFile(FAsyncThread.FAsyncComHandle, ABuffer^, ADataSize, Read, @OverLapped) then
  //    if not ReadFile(FAsyncThread.AsyncComHandle, ABuffer^, ADataSize, Read, nil) then
      begin
  {      Errors := GetLastError;
        while (Errors = ERROR_IO_PENDING) or (Errors = ERROR_IO_INCOMPLETE) do
        begin
          if GetOverlappedResult(FASyncThread.ASyncComHandle, Overlapped, Read, True) then
          begin
            Result := Read;
            break;
    //            WriteLog(lkSend, Buffer, Sent);
          end;
          Errors := GetLastError;
        end; }

        if (GetLastError <> ERROR_IO_PENDING) then exit;

  //      R := WaitForResult(1, [Overlapped.hEvent], False, READ_TIMEOUT, QS_PAINT or QS_SENDMESSAGE);
        R := WaitForSingleObject(OverLapped.hEvent, READ_TIMEOUT);
        case R of
          WAIT_OBJECT_0:
          begin
            if GetOverlappedResult(FASyncThread.FASyncComHandle, OverLapped, Read, False) then
//            if GetOverlappedResult(FASyncThread.ASyncComHandle, OverLapped, Read, True) then
            begin
              Result := Read;
//              if LogEnabled then WriteLog(lkReceive, ABuffer, Read);
            end;
          end;
        end;

        ClearCommError(FComPortHandle, Errors, @ComStat);
      end
      else
      begin
        Result := Read;
//        if LogEnabled then WriteLog(lkReceive, ABuffer, Read);
      end;
    end
    else
    begin
      Buf.buf := ABuffer;
      Buf.len := ADataSize;
      RecvFlags := 0;

      if (FControlType = ctTCP) then
      begin
        if WSARecv(FAsyncThread.FASyncComHandle, @Buf, 1, Read, RecvFlags, @Overlapped, nil) = SOCKET_ERROR then
        begin
          if (WSAGetLastError <> WSA_IO_PENDING) then exit;
          R := WaitForSingleObject(OverLapped.hEvent, READ_TIMEOUT);
          case R of
            WAIT_OBJECT_0:
            begin
              if GetOverlappedResult(FASyncThread.FASyncComHandle, OverLapped, Read, False) then
  //            if GetOverlappedResult(FASyncThread.ASyncComHandle, OverLapped, Read, True) then
              begin
                Result := Read;
  //              if LogEnabled then WriteLog(lkReceive, ABuffer, Read);
              end;
            end;
          end;
        end
        else
        begin
          Result := Read;
  //        if LogEnabled then WriteLog(lkReceive, ABuffer, Read);
        end;
      end
      else
      begin
        ClientAddrSize := SizeOf(TSockAddrIn);

        if WSARecvFrom(FAsyncThread.FASyncComHandle, @Buf, 1, @Read, RecvFlags, @TSockAddr(ClientSockAddr), @ClientAddrSize, @Overlapped, nil) = SOCKET_ERROR then
        begin
          if (WSAGetLastError <> WSA_IO_PENDING) then exit;
          R := WaitForSingleObject(OverLapped.hEvent, READ_TIMEOUT);
          case R of
            WAIT_OBJECT_0:
            begin
              if GetOverlappedResult(FASyncThread.FASyncComHandle, OverLapped, Read, False) then
  //            if GetOverlappedResult(FASyncThread.ASyncComHandle, OverLapped, Read, True) then
              begin
                Result := Read;
  //              if LogEnabled then WriteLog(lkReceive, ABuffer, Read);
              end;
            end;
          end;
        end
        else
        begin
          Result := Read;
  //        if LogEnabled then WriteLog(lkReceive, ABuffer, Read);
        end;
      end;

{      Read := RecvInit(FAsyncThread.ASyncComHandle, PChar(Buffer), DataSize, 0);

      if Read = SOCKET_ERROR then
        if WSAGetLastError <> WSAEWOULDBLOCK then
        begin
          exit;
        end;

      Result := Read;
      if LogEnabled then WriteLog(lkReceive, Buffer, Read); }
    end;
    if (Result > 0) then Dec(FAsyncThread.FRxCount, Result);
    if (LogEnabled) and (Result > 0) then WriteLog(lkReceive, ABuffer, Result, LogIsHexcode);
  finally
    CloseHandle(OverLapped.hEvent);
  end;
end;

function TBaseSerial.SendString(AStr: AnsiString): Integer;
var
  Buffer: Pointer;
  BufSize: Integer;
begin
  Result := -1;

  if Length(AStr) <= 0 then exit;
  Result := SendData(PAnsiChar(AStr), Length(AStr));

{  BufSize := Length(AStr);
  if BufSize <= 0 then exit;

  Buffer := GetMemory(BufSize);
  try
    FillChar(Buffer^, BufSize, #0);
    Move(AStr[1], Buffer^, Result);
    Result := SendData(Buffer, BufSize);
//    CopyMemory(@S, Buffer, DataSize);
  finally
    FreeMemory(Buffer);
  end; }
end;

function TBaseSerial.ReadString(var AStr: AnsiString; ADataSize: Integer): Integer;
var
  Buffer: Pointer;
begin
  Result := -1;

  SetLength(AStr, 0);
  if ADataSize <= 0 then exit;

  SetLength(AStr, ADataSize);
//  FillChar(S, DataSize, #0);
  Result := ReadData(PAnsiChar(AStr), ADataSize);

{//  SetLength(S, DataSize);
  Buffer := GetMemory(ADataSize);
  try
    FillChar(Buffer^, ADataSize, #0);
    Result := ReadData(Buffer, ADataSize);
//    CopyMemory(@S, Buffer, DataSize);
    if Result > 0 then
    begin
      SetLength(AStr, Result);
      Move(Buffer^, AStr[1], Result);
    end;
  finally
    FreeMemory(Buffer);
  end; }
end;

function TBaseSerial.SendReceiveData(ACommand: AnsiString; ALen: Integer): Integer;
var
  ErrCnt: Integer;
begin
  Result := D_FALSE;
  ErrCnt := 0;
  while ErrCnt < MAX_RETRY do
  begin
    if SendData(PAnsiChar(ACommand), ALen) <= 0 then
    begin
      Inc(ErrCnt);
      Continue;
    end
    else
    begin
      Result := D_OK;
      break;
    end;
  end;
end;

function TBaseSerial.Receive(ABuffer: Pointer; ADataSize: Integer): Integer;
var
  OverLapped: TOverLapped;
  RecvSize: Cardinal;
  Buf: WSABUF;
  RecvFlags: Cardinal;
begin
  Result := 0;
  if ADataSize <= 0 then exit;

  if (FControlType = ctSerial) then
  begin
    if Cardinal(ADataSize) > FAsyncThread.FRxCount then
      ADataSize := FAsyncThread.FRxCount;

    FillChar(OverLapped, SizeOf(OverLapped), 0);
    ReadFile(FAsyncThread.FASyncComHandle, ABuffer^, ADataSize, Cardinal(Result), @OverLapped);
    Dec(FAsyncThread.FRxCount, Result);
  end
  else
  begin
    if Cardinal(ADataSize) > FAsyncThread.FRxCount then
      ADataSize := FAsyncThread.FRxCount;

    Buf.buf := ABuffer;
    Buf.len := ADataSize;
    RecvFlags := 0;

    FillChar(OverLapped, SizeOf(OverLapped), 0);
    WSARecv(FAsyncThread.FASyncComHandle, @Buf, 1, RecvSize, RecvFlags, @OverLapped, nil);
    Result := RecvSize;
    Dec(FAsyncThread.FRxCount, Result);

{    RecvSize := RecvInit(FAsyncThread.ASyncComHandle, PAnsiChar(ABuffer), ADataSize, 0);
    if RecvSize = SOCKET_ERROR then
      if WSAGetLastError <> WSAEWOULDBLOCK then
      begin
        exit;
      end;

    Result := RecvSize;
    Dec(FAsyncThread.FRxCount, Result);  }
  end;
end;

procedure TBaseSerial.DataEvent;
var
  DataSize: Integer;
  Buffer: AnsiString;
  ByteCount: Integer;

  R: Integer;
begin
  DataSize := RxCount;
  if DataSize = 0 then exit;

{  while DataSize > 0 do
  begin
    R := ReadString(Buffer, DataSize);
    if R <= 0 then exit;

    FReceivedBuffer := FReceivedBuffer + Buffer;
    Dec(DataSize, R);
  end;



exit; }

  if ReadString(Buffer, DataSize) <= 0 then exit;
  FReceivedBuffer := FReceivedBuffer + Buffer;
//  if Assigned(FOnDataEvent) then FOnDataEvent(Self);
end;

{ TAsyncThread }

constructor TAsyncThread.Create(ABaseSerial: TBaseSerial; AASyncComHandle: THandle);
begin
  FBaseSerial := ABaseSerial;
  FASyncComHandle := AASyncComHandle;

  FCloseEvent := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;

  inherited Create(True);
end;

destructor TAsyncThread.Destroy;
begin
  CloseHandle(FCloseEvent);

  inherited Destroy;
end;

procedure TAsyncThread.Close;
begin
  SetEvent(FCloseEvent);
  Terminate;
end;

procedure TAsyncThread.Execute;
var
  WaitHandles: array[0..1] of THandle;
  OverLapped: TOverLapped;
  WaitEvent: Cardinal;
  a: Boolean;
  fWaitingOnStat, bWaitingOnStatusHandle: Boolean;
  dwOvRes: Cardinal;

	NetEvent: TWSANetworkEvents;
begin
  FTxEmpty := False;

  FillChar(OverLapped, SizeOf(OverLapped), 0);
//  CloseEvent := CreateEvent(nil, True, False, nil);

  if (FBaseSerial.ControlType = ctSerial) then
  begin
    OverLapped.hEvent := CreateEvent(nil, True, True, nil);
    WaitHandles[0] := FCloseEvent;
    WaitHandles[1] := OverLapped.hEvent;

    fWaitingOnStat := False;
    while not Terminated do
    begin
  {    SetCommMask(ASyncComHandle, EV_RXCHAR);
      WaitCommEvent(ASyncComHandle, FEventMask, nil);
      if (FEventMask and EV_RXCHAR) > 0 then PortEvent;
      Continue; }

  {      if not (fWaitingOnStat) then
        begin
           if not (WaitCommEvent(ASyncComHandle, FEventMask, @OverLapped)) then
           begin
              if (GetLastError = ERROR_IO_PENDING) then
                 fWaitingOnStat := TRUE
              else
                 // error in WaitCommEvent; abort
                 break;
           end
  {         else
              // WaitCommEvent returned immediately.
              // Deal with status event as appropriate.
              ReportStatusEvent(dwCommEvent); }
  {      end;

        // Check on overlapped operation.
        if (fWaitingOnStat) then
        begin
           // Wait a little while for an event to occur.
           WaitEvent := WaitForSingleObject(OverLapped.hEvent, 1000);
           case WaitEvent of
               // Event occurred.
               WAIT_OBJECT_0:
               begin
                   if not (GetOverlappedResult(ASyncComHandle, OverLapped, dwOvRes, FALSE)) then
                      // An error occurred in the overlapped operation;
                      // call GetLastError to find out what it was
                      // and abort if it is fatal.
                   else PortEvent;
                      // Status event is stored in the event flag
                      // specified in the original WaitCommEvent call.
                      // Deal with the status event as appropriate.
                      ;//ReportStatusEvent(dwCommEvent);

                   // Set fWaitingOnStat flag to indicate that a new
                   // WaitCommEvent is to be issued.
                   fWaitingOnStat := FALSE;
  //                 break;

  //             case WAIT_TIMEOUT:
                   // Operation isn't complete yet. fWaitingOnStatusHandle flag
                   // isn't changed since I'll loop back around and I don't want
                   // to issue another WaitCommEvent until the first one finishes.
                   //
                   // This is a good time to do some background work.
  //                DoBackgroundWork();
  //                 break;
               end
               else
               begin
                   // Error in the WaitForSingleObject; abort
                   // This indicates a problem with the OVERLAPPED structure's
                   // event handle.
  //                CloseHandle(OverLapped.hEvent);
                  break;
               end;
           end;

      end; }
  {    Continue;
      if not WaitCommEvent(ASyncComHandle, FEventMask, @OverLapped) then
      begin
        if GetLastError = ERROR_IO_PENDING then Continue
        else continue;
      end;
  //    if a then showmessage('1');  }

  //    WaitEvent := FBaseSerial.WaitForResult(2, WaitHandles, False, 1000, QS_PAINT or QS_SENDMESSAGE);
  //    SetCommMask(ASyncComHandle, EV_RXCHAR);
  //    FEventMask := EV_RXCHAR;
  //    WaitCommEvent(ASyncComHandle, FEventMask, @OverLapped);

      WaitCommEvent(FASyncComHandle, FEventMask, @OverLapped);
      WaitEvent := WaitForMultipleObjects(2, @WaitHandles, False, INFINITE);
      case WaitEvent of
        WAIT_OBJECT_0:
        begin
          CloseHandle(OverLapped.hEvent);
//          CloseHandle(FCloseEvent);
          CloseHandle(FASyncComHandle);
          Terminate;
          Break;
        end;
        WAIT_OBJECT_0 + 1: PortEvent;
      end;
    end;
  end
  else
  begin
    WaitHandles[0] := FCloseEvent;
    WaitHandles[1] := FBaseSerial.FSocketEvent;

    while not Terminated do
    begin
      WaitEvent := WSAWaitForMultipleEvents(2, @WaitHandles, False, INFINITE, False);
      case WaitEvent of
        WAIT_OBJECT_0:
        begin
//          CloseHandle(SockReadEvent);
          CloseHandle(FBaseSerial.FSocketEvent);
//          CloseHandle(FCloseEvent);
          CloseSocket(FASyncComHandle);
          Terminate;
          Break;
        end;
        WAIT_OBJECT_0 + 1:
        begin
          if WSAEnumNetworkEvents(FASyncComHandle, FBaseSerial.FSocketEvent, NetEvent) = SOCKET_ERROR then Continue;

          if (NetEvent.lNetworkEvents and FD_CONNECT) > 0 then
          begin
            if NetEvent.iErrorCode[FD_CONNECT_BIT] <> 0 then
            begin
              ///Fail 2.1 connect Error
            end
            else
            begin
              ///No Error. Event Process
              FBaseSerial.FSocketConnected := True;
            end;
          end;
          if (NetEvent.lNetworkEvents and FD_READ) > 0 then
          begin
            if NetEvent.iErrorCode[FD_READ_BIT] <> 0 then
            begin
              ///Fail 2.1 read Error
            end
            else
            begin
              ///No Error. Event Process
              PortEvent;
            end;
          end;
          if (NetEvent.lNetworkEvents and FD_CLOSE) > 0 then
          begin
            if NetEvent.iErrorCode[FD_CLOSE_BIT] <> 0 then
            begin
              ///Fail 2.1 close Error
            end
            else
            begin
              ///No Error. Event Process
              FBaseSerial.FSocketConnected := False;
              if Assigned(FBaseSerial.FOnSocketCloseEvent) then FBaseSerial.FOnSocketCloseEvent(FBaseSerial);
            end;
          end;
        end;
      end;
    end;
  end;
end;

// synchronized method
procedure TAsyncThread.PortEvent;
var
  ComStat: TComStat;
  ErrorMask: Cardinal;
  TrashCan: String[255];
begin
  if (FBaseSerial.ControlType = ctSerial) then
  begin
  ClearCommError(FASyncComHandle, ErrorMask, @ComStat);

  if (FEventMask and EV_RXCHAR) > 0 then
  begin
    FRxCount := ComStat.cbInQue;
    if FRxCount > 0 then
    begin
      FBaseSerial.DataEvent;
      if Assigned(FBaseSerial.FOnDataEvent) then FBaseSerial.FOnDataEvent(FBaseSerial);

{      if not Assigned(FBaseSerial.FOnDataEvent) then
        while FRxCount > 0 do
          FBaseSerial.Receive(@TrashCan[1], High(TrashCan))
      else FBaseSerial.FOnDataEvent(FBaseSerial); }
    end;
  end;



{
    Have to add them also in the SetCommMask
   if ((FEventMask and EV_TXEMPTY) > 0) and Assigned(FBaseSerial.FOnDataSent) then
       FBaseSerial.FOnDataSent(FBaseSerial);
   if (FEventMask and EV_CTS)  > 0 then
   if (FEventMask and EV_DSR)  > 0 then
   if (FEventMask and EV_RLSD) > 0 then
   if (FEventMask and EV_RING) > 0 then
   if (FEventMask and EV_ERR)  > 0 then
}
  end
  else
  begin
    if FBaseSerial.FSocketHandle <> INVALID_SOCKET then
    begin
      FRxCount := 0;
      ioctlSocket(FBaseSerial.FSocketHandle, FIONREAD, FRxCount);
      if FRxCount > 0 then
      begin
        FBaseSerial.DataEvent;
        if Assigned(FBaseSerial.FOnDataEvent) then FBaseSerial.FOnDataEvent(FBaseSerial);
      end;
    end;
    ResetEvent(FSockReadEvent);
  end;
end;

end.
