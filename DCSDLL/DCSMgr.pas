unit DCSMgr;

interface

uses System.Classes, Winapi.Windows, System.SysUtils, Vcl.Dialogs, System.SyncObjs,
  Generics.Collections,
  UnitCommons, UnitTypeConvert,
  UnitUDPIn, UnitUDPOut;

type
  TDCSDevice = record
    Handle: TDeviceHandle;
    Name: array[0..DEVICENAME_LEN - 1] of Char;
    ID: Word;
    IP: array[0..HOSTIP_LEN - 1] of Char;
    Port: Integer;
  end;
  PDCSDevice = ^TDCSDevice;
  TDCSDeviceList = TList<PDCSDevice>;

  TDCSEventNotifyThread = class;

  TDCSMgr = class(TComponent)
  private
    FSysCMD1, FSysCMD2: Byte;
    FSysSyncMsgEvent: THandle;
    FSysRecvBuffer: AnsiString;
    FSysRecvData: AnsiString;
    FSysLastResult: Integer;

    FDeviceList: TDCSDeviceList;

    FCMD1, FCMD2: Byte;
    FSyncMsgEvent: THandle;
    FRecvBuffer: AnsiString;
    FRecvData: AnsiString;
    FLastResult: Integer;

    FNotifyBuffer: AnsiString;
    FNotifyData: AnsiString;

    FSysCommandCritSec: TCriticalSection;

    FSysInCritSec: TCriticalSection;

    FCommandCritSec: TCriticalSection;

    FNotifyCritSec: TCriticalSection;
    FInCritSec: TCriticalSection;

    FEventNotifyThread: TDCSEventNotifyThread;

    function GetDeviceByHandle(AID: Word; AHandle: TDeviceHandle): PDCSDevice;
    function GetDeviceByName(AID: Word; AName: String): PDCSDevice;
    function GetValidDeviceHandle(AID: Word; AHandle: TDeviceHandle; var Device: PDCSDevice): Integer;

    function SendSysCommand(AIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: AnsiString; ADataSize: Integer): Integer;
    function SendCommand(AIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: AnsiString; ADataSize: Integer): Integer;
  public
    FUDPSysIn: TUDPIn;
    FUDPSysOut: TUDPOut;

    FSysInPort, FSysOutPort: Word;
    FSysTimeout: Cardinal;

    FUDPNotify: TUDPIn;
    FUDPIn: TUDPIn;
    FUDPOut: TUDPOut;

    FNotifyPort, FInPort, FOutPort: Word;
    FTimeout: Cardinal;

    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    function CheckSum(AValue: AnsiString): Boolean;

    procedure UDPSysInRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);

    procedure UDPNotifyRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);
    procedure UDPInRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);

    function TransmitSysCommand(AIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: AnsiString; ADataSize: Integer): Integer;
    function TransmitCommand(AIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: AnsiString; ADataSize: Integer): Integer;
    function TransmitDevice(ACMD1, ACMD2: Byte; AID: Word; AHandle: TDeviceHandle; ADataBuf: AnsiString; ADataSize: Integer): Integer;

    // 0X00 System Control
    function SysIsAlive(AIP: PChar; var AIsAlive: Boolean): Integer;                                                 // 0X00.00
    function SysIsMain(AIP: PChar; var AIsMain: Boolean): Integer;                                                   // 0X00.01

    // 0X10 Device Control
    function Open(AID: Word; AIP: PChar; ADeviceName: PChar; var AHandle: TDeviceHandle): Integer;                   // 0X10.00
    function Close(AID: Word; AHandle: TDeviceHandle): Integer;                                                      // 0X10.01
    function Reset(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer;                                    // 0X10.02
    function SetControlBy(AID: Word; AHandle: TDeviceHandle): Integer;                                               // 0X10.03

    // 0X20 Immediate Control
    function Stop(AID: Word; AHandle: TDeviceHandle): Integer;                                                       // 0X20.00
    function Play(AID: Word; AHandle: TDeviceHandle): Integer;                                                       // 0X20.01
    function Pause(AID: Word; AHandle: TDeviceHandle): Integer;                                                      // 0X20.02
    function Continue(AID: Word; AHandle: TDeviceHandle): Integer;                                                   // 0X20.03
    function FastFoward(AID: Word; AHandle: TDeviceHandle): Integer;                                                 // 0X20.04
    function FastRewind(AID: Word; AHandle: TDeviceHandle): Integer;                                                 // 0X20.05
    function Jog(AID: Word; AHandle: TDeviceHandle; AFrameOrSpeed: Double): Integer;                                 // 0X20.06
    function Shuttle(AID: Word; AHandle: TDeviceHandle; ASpeed: Double): Integer;                                    // 0X20.07
    function StandBy(AID: Word; AHandle: TDeviceHandle; AOn: Boolean): Integer;                                      // 0X20.08
    function Eject(AID: Word; AHandle: TDeviceHandle): Integer;                                                      // 0X20.09
    function Preroll(AID: Word; AHandle: TDeviceHandle): Integer;                                                    // 0X20.10 VCR only
    function Rec(AID: Word; AHandle: TDeviceHandle): Integer;                                                        // 0X20.20
    function AutoEdit(AID: Word; AHandle: TDeviceHandle): Integer;                                                   // 0X20.21 VCR only

    // 0X30 Preset/Select Commands
    function SetPortMode(AID: Word; AHandle: TDeviceHandle; APortMode: TPortMode): Integer;                          // 0X30.00 Video server only
    function SetAutoStatus(AID: Word; AHandle: TDeviceHandle; AAutoStatus: Boolean = True): Integer;                 // 0X30.01
    function PlayCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ACueTC, ADuration: TTimecode): Integer;    // 0X30.10
    function RecordCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ADuration: TTimecode): Integer;          // 0X30.20
    function IDRename(AID: Word; AHandle: TDeviceHandle; ASourceID, ATargetID: String): Integer;                     // 0X30.30
    function IDDelete(AID: Word; AHandle: TDeviceHandle; AMediaID: String): Integer;                                 // 0X30.31
    function InEntry(AID: Word; AHandle: TDeviceHandle): Integer;                                                    // 0X30.40 VCR only
    function OutEntry(AID: Word; AHandle: TDeviceHandle): Integer;                                                   // 0X30.41 VCR only
    function AInEntry(AID: Word; AHandle: TDeviceHandle): Integer;                                                   // 0X30.42 VCR only
    function AOutEntry(AID: Word; AHandle: TDeviceHandle): Integer;                                                  // 0X30.43 VCR only
    function InReset(AID: Word; AHandle: TDeviceHandle): Integer;                                                    // 0X30.44 VCR only
    function OutReset(AID: Word; AHandle: TDeviceHandle): Integer;                                                   // 0X30.45 VCR only
    function AInReset(AID: Word; AHandle: TDeviceHandle): Integer;                                                   // 0X30.46 VCR only
    function AOutReset(AID: Word; AHandle: TDeviceHandle): Integer;                                                  // 0X30.47 VCR only
    function EditPreset(AID: Word; AHandle: TDeviceHandle; AData1, AData2: Byte): Integer;                           // 0X30.48 VCR only
    function SetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl, AInput, AInputLvl: Integer): Integer;   // 0X30.50 Router only
    function SetXpt(AID: Word; AHandle: TDeviceHandle; AInput: Integer): Integer;                                     // 0X30.60 MCS only

    // 0X40 Sense Queries
    function GetDeviceStatus(AID: Word; AHandle: TDeviceHandle; var AStatus: TDeviceStatus): Integer;                 // 0X40.00
    function GetStorageTimeRemaining(AID: Word; AHandle: TDeviceHandle; var ATotal, AAvailable: TTimecode; Extended: Boolean = False): Integer; // 0X40.01
    function GetTC(AID: Word; AHandle: TDeviceHandle; var ACurTC: TTimecode): Integer;                                // 0X40.10
    function GetRemainTC(AID: Word; AHandle: TDeviceHandle; var ARemainTC: TTimecode): Integer;                       // 0X40.11
    function GetList(AID: Word; AHandle: TDeviceHandle; var AIDList: TIDList; var ARemainIDCount: Integer): Integer;  // 0X40.20
    function GetNext(AID: Word; AHandle: TDeviceHandle; var AIDList: TIDList; var ARemainIDCount: Integer): Integer;  // 0X40.21
    function GetExist(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var AExist: Boolean): Integer;             // 0X40.22
    function GetSize(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var ADuration: TTimecode): Integer;         // 0X40.23
    function GetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl: Integer; var AInput, AInputLvl: Integer): Integer;  // 0X40.50 Router only
    function GetXpt(AID: Word; AHandle: TDeviceHandle; var AInput: Integer): Integer;                                 // 0X40.60 Router only

    // 0X50 Event Control
    function InputEvent(AID: Word; AHandle: TDeviceHandle; AEvent: TEvent): Integer;                                  // 0X50.00
    function DeleteEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer;                             // 0X50.01
    function ClearEvent(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer;                                // 0X50.02
    function TakeEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; AStartTime: TEventTime): Integer;       // 0X50.10
    function HoldEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer;                               // 0X50.11
    function ChangeDurationEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; ADuration: TTimecode): Integer; // 0X50.12
    function OnAirCatchEvent(AID: Word; AHandle: TDeviceHandle): Integer;                                             // 0X50.13
    function GetOnAirEventID(AID: Word; AHandle: TDeviceHandle; var AOnAirEventID, ANextEventID: TEventID): Integer;  // 0X50.20
    function GetEventInfo(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime; var ADurationTC: TTimecode): Integer;  // 0X50.21
    function GetEventStatus(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AEventStatus: TEventStatus): Integer;  // 0X50.22
    function GetEventStartTime(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime): Integer;   // 0X50.23
    function GetEventOverall(AID: Word; AHandle: TDeviceHandle; var AEventOverall: TEventOverall): Integer;   // 0X50.24

//    property UDPNotify: TUDPIn read FUDPNotify write FUDPNotify;
//    property UDPIn: TUDPIn read FUDPIn write FUDPIn;
//    property UDPOut: TUDPOut read FUDPOut write FUDPOut;
//
//    property Timeout: Cardinal read FTimeout write FTimeout;

//    property UDPSysIn: TUDPIn read FUDPSysIn write FUDPSysIn;
//    property UDPSysOut: TUDPOut read FUDPSysOut write FUDPSysOut;
//    property SysInPort: Word read FSysInPort write FSysInPort;
//    property SysOutPort: Word read FSysOutPort write FSysOutPort;
//    property SysTimeout: Cardinal read FSysTimeout write FSysTimeout;
  end;

  TDCSEventNotifyThread = class(TThread)
  private
    FDCSMgr: TDCSMgr;
    FExecuteEvent: THandle;
    FEventID: TEventID;
    FEventStatus: TEventStatus;

    procedure DoControl;
  protected
    procedure Execute; override;
  public
    constructor Create(ADCSMgr: TDCSMgr); virtual;
    destructor Destroy; override;

    procedure SetEventStatus(AEventID: TEventID; AEventStatus: TEventStatus);
  end;

var
  VDCSMgr: TDCSMgr;

implementation

uses DLLConsts;

constructor TDCSMgr.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  // System control
  FSysCommandCritSec := TCriticalSection.Create;

  FSysInCritSec := TCriticalSection.Create;

{  FUDPSysIn := TUDPIn.Create;
  FUDPSysIn.OnUDPRead := UDPSysInRead;

  FUDPSysOut := TUDPOut.Create; }

  FSysInPort  := 0;
  FSysOutPort := 0;

  FSysCMD1 := $00;
  FSysCMD2 := $00;

  FSysSyncMsgEvent := CreateEvent(nil, True, False, nil);

  FSysTimeout := TIME_OUT;

  // Device control
  FCommandCritSec := TCriticalSection.Create;

  FNotifyCritSec := TCriticalSection.Create;
  FInCritSec := TCriticalSection.Create;

  FDeviceList := TDCSDeviceList.Create;

{  FUDPNotify := TUDPIn.Create;
  FUDPNotify.OnUDPRead := UDPNotifyRead;

  FUDPIn := TUDPIn.Create;
  FUDPIn.OnUDPRead := UDPInRead;

  FUDPOut := TUDPOut.Create; }

  FCMD1 := $00;
  FCMD2 := $00;

  FSyncMsgEvent := CreateEvent(nil, True, False, nil);

  FTimeout := TIME_OUT;

  FEventNotifyThread := TDCSEventNotifyThread.Create(Self);
  FEventNotifyThread.Start;
end;

destructor TDCSMgr.Destroy;
var
  I: Integer;
begin
  if Assigned(FEventNotifyThread) then
  begin
    FEventNotifyThread.Terminate;
    FreeAndNil(FEventNotifyThread);
  end;

  CloseHandle(FSyncMsgEvent);

  if (FUDPNotify <> nil) then
  begin
//    if (WaitForSingleObject(FUDPNotify.Handle, 0) = WAIT_OBJECT_0) then
    begin
      FUDPNotify.Close;
      FUDPNotify.Terminate;
      FUDPNotify.WaitFor;
    end;
    FreeAndNil(FUDPNotify);
  end;

  if (FUDPIn <> nil) then
  begin
//    if (WaitForSingleObject(FUDPIn.Handle, 0) = WAIT_OBJECT_0) then
    begin
      FUDPIn.Close;
      FUDPIn.Terminate;
      FUDPIn.WaitFor;
    end;
    FreeAndNil(FUDPIn);
  end;

  if (FUDPOut <> nil) then
  begin
//    if (WaitForSingleObject(FUDPOut.Handle, 0) = WAIT_OBJECT_0) then
    begin
      FUDPOut.Close;
      FUDPOut.Terminate;
      FUDPOut.WaitFor;
    end;
    FreeAndNil(FUDPOut);
  end;

//  FreeAndNil(FUDPNotify);
//  FreeAndNil(FUDPIn);
//  FreeAndNil(FUDPOut);

  CloseHandle(FSysSyncMsgEvent);

  if (FUDPSysIn <> nil) then
  begin
//    if (WaitForSingleObject(FUDPSysIn.Handle, 0) = WAIT_OBJECT_0) then
    begin
      FUDPSysIn.Close;
      FUDPSysIn.Terminate;
      FUDPSysIn.WaitFor;
    end;
    FreeAndNil(FUDPSysIn);
  end;

  if (FUDPSysOut <> nil) then
  begin
//    if (WaitForSingleObject(FUDPSysOut.Handle, 0) = WAIT_OBJECT_0) then
    begin
      FUDPSysOut.Close;
      FUDPSysOut.Terminate;
      FUDPSysOut.WaitFor;
    end;
    FreeAndNil(FUDPSysOut);
  end;

//  FreeAndNil(FUDPSysIn);
//  FreeAndNil(FUDPSysOut);

  for I := FDeviceList.Count - 1 downto 0 do
  begin
    Dispose(FDeviceList[I]);
  end;
  FDeviceList.Clear;
  FreeAndNil(FDeviceList);

  FreeAndNil(FCommandCritSec);

  FreeAndNil(FNotifyCritSec);
  FreeAndNil(FInCritSec);

  FreeAndNil(FSysCommandCritSec);

  FreeAndNil(FSysInCritSec);

  inherited Destroy;
end;

function TDCSMgr.GetDeviceByHandle(AID: Word; AHandle: TDeviceHandle): PDCSDevice;
var
  I: Integer;
  P: PDCSDevice;
begin
  Result := nil;
  for I := 0 to FDeviceList.Count - 1 do
  begin
    P := FDeviceList[I];
    if (FDeviceList[I]^.ID = AID) and
       (FDeviceList[I]^.Handle = AHandle) then
    begin
      Result := P;
      break;
    end;
  end;
end;

function TDCSMgr.GetDeviceByName(AID: Word; AName: String): PDCSDevice;
var
  I: Integer;
  P: PDCSDevice;
begin
  Result := nil;
  for I := 0 to FDeviceList.Count - 1 do
  begin
    P := FDeviceList[I];
    if (FDeviceList[I]^.ID = AID) and
       (String(FDeviceList[I]^.Name) = AName) then
    begin
      Result := P;
      break;
    end;
  end;
end;

function TDCSMgr.GetValidDeviceHandle(AID: Word; AHandle: TDeviceHandle; var Device: PDCSDevice): Integer;
begin
  Result := D_FALSE;
  Device := nil;

  if (AHandle < 0) then
  begin
    Result := E_INVALID_DEVICE_HANDLE;
    exit;
  end;

  Device := GetDeviceByHandle(AID, AHandle);
  if (Device = nil) then
  begin
    Result := E_NOT_OPENED_DEVICE;
    exit;
  end;

  Result := D_OK;
end;

function TDCSMgr.SendSysCommand(AIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: AnsiString; ADataSize: Integer): Integer;
var
  Buffer: AnsiString;
  CheckSum: Byte;
  I: integer;
begin
  Result := D_FALSE;

  if (FUDPSysOut = nil) then exit;

  Buffer := AnsiChar(D_STX) + WordToAnsiString($02 + ADataSize) + AnsiChar(ACMD1) + AnsiChar(ACMD2) + ADataBuf;

  CheckSum := ACMD1 + ACMD2;
  for I := 1 to ADataSize do
    CheckSum := CheckSum + Ord(ADataBuf[I]);

  CheckSum := 0 - CheckSum;
  Buffer := Buffer + AnsiChar(CheckSum);

  FUDPSysOut.Send(AIP, APort, Buffer);

  Result := D_OK;
end;

function TDCSMgr.SendCommand(AIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: AnsiString; ADataSize: Integer): Integer;
var
  Buffer: AnsiString;
  CheckSum: Byte;
  I: integer;
begin
  Result := D_FALSE;

  Buffer := AnsiChar(D_STX) + WordToAnsiString($02 + ADataSize) + AnsiChar(ACMD1) + AnsiChar(ACMD2) + ADataBuf;

  CheckSum := ACMD1 + ACMD2;
  for I := 1 to ADataSize do
    CheckSum := CheckSum + Ord(ADataBuf[I]);

  CheckSum := 0 - CheckSum;
  Buffer := Buffer + AnsiChar(CheckSum);

//  FUDPOut.Send(AHost, APort, Buffer);
  FUDPOut.Send(AIP, APort, Buffer);

  Result := D_OK;
end;

function TDCSMgr.CheckSum(AValue: AnsiString): Boolean;
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

function TDCSMgr.TransmitSysCommand(AIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: AnsiString; ADataSize: Integer): Integer;
var
  R: DWORD;
begin
  Result := D_FALSE;

  ResetEvent(FSysSyncMsgEvent);
  FSysRecvBuffer := '';
  FSysRecvData := '';

  FSysCMD1 := ACMD1;
  FSysCMD2 := ACMD2;

  FSysCommandCritSec.Enter;
  try
    if (SendSysCommand(AIP, APort, ACMD1, ACMD2, ADataBuf, ADataSize) = NOERROR) then
    begin
      R := WaitForSingleObject(FSysSyncMsgEvent, FSysTimeout);
      case R of
        WAIT_OBJECT_0:
          begin
            Result := FSysLastResult;
          end;
        else Result := E_TIMEOUT;
       end;
    end;
  finally
    FSysCommandCritSec.Leave;
  end;
end;

function TDCSMgr.TransmitCommand(AIP: AnsiString; APort: Word; ACMD1, ACMD2: Byte; ADataBuf: AnsiString; ADataSize: Integer): Integer;
var
  R: DWORD;
begin
  Result := D_FALSE;

  FInCritSec.Enter;
  try
  ResetEvent(FSyncMsgEvent);
  FRecvBuffer := '';
  FRecvData := '';

  FCMD1 := ACMD1;
  FCMD2 := ACMD2;
  finally
    FInCritSec.Leave;
  end;

  FCommandCritSec.Enter;
  try
    if (SendCommand(AIP, APort, ACMD1, ACMD2, ADataBuf, ADataSize) = NOERROR) then
    begin
      R := WaitForSingleObject(FSyncMsgEvent, FTimeout);
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
  end;
end;

function TDCSMgr.TransmitDevice(ACMD1, ACMD2: Byte; AID: Word; AHandle: TDeviceHandle; ADataBuf: AnsiString; ADataSize: Integer): Integer;
var
  P: PDCSDevice;
begin
  Result := D_FALSE;

  P := GetDeviceByHandle(AID, AHandle);
  if P = nil then exit;

  Result := TransmitCommand(P^.IP, P^.Port, ACMD1, ACMD2, ADataBuf, ADataSize);
end;

function WriteLog(Log: AnsiString): Integer;
var
  FilePath, FileName: AnsiString;
  FileStream: TFileStream;
  S: AnsiString;
  F: TextFile;
begin
  Result := -1;

  FilePath := 'C:\01';
  FileName := FilePath + '\' + '01.log';

  // 디렉토리 있는지 확인 없으면 생성
  if not DirectoryExists(FilePath) then ForceDirectories(FilePath);

  S := FormatDateTime('[hh:nn:ss:zzz]', Now);
  S := S + ' ' + Log;
  try
    AssignFile(F, FileName);
    if FileExists(FileName) then Append(F)
    else Rewrite(F);
    Writeln(F, S);
  finally
    Closefile(F);
  end;
end;

procedure TDCSMgr.UDPSysInRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);
var
  ByteCount: Integer;
  CMD1, CMD2: Byte;
  Device: TDeviceHandle;
begin
  FSysInCritSec.Enter;
  try
    if (ADataSize <= 0) then exit;
    FSysRecvBuffer := FSysRecvBuffer + AData;

    if (Length(FSysRecvBuffer) < 1) then exit;

    case Ord(FSysRecvBuffer[1]) of
      $02:
        begin
          if (Length(FSysRecvBuffer) < 3) then exit;

          ByteCount := PAnsiCharToWord(@FSysRecvBuffer[2]);
          if (ByteCount > 0) and (Length(FSysRecvBuffer) = ByteCount + 4) then
          begin
            if (CheckSum(FSysRecvBuffer)) then
            begin
              CMD1 := Ord(FSysRecvBuffer[4]);
              CMD2 := Ord(FSysRecvBuffer[5]);

              FSysRecvData := System.Copy(FSysRecvBuffer, 6, ByteCount - 2);

              if (CMD1 = FSysCMD1) and (CMD2 = FSysCMD2 + $80)  then
                FSysLastResult := D_OK
              else
                FSysLastResult := D_FALSE;
            end
            else
              FSysLastResult := E_NAK_CHECKSUM;

            SetEvent(FSysSyncMsgEvent);

  //          if (Length(FSysRecvBuffer) > ByteCount + 4) then
  //            FSysRecvBuffer := Copy(FSysRecvBuffer, ByteCount + 5, Length(FSysRecvBuffer))
  //          else
              FSysRecvBuffer := '';
          end
          else if (ByteCount <= 0) or (Length(FSysRecvBuffer) > ByteCount + 4) then
          begin
            FSysLastResult := D_FALSE;
            SetEvent(FSysSyncMsgEvent);
            FSysRecvBuffer:= '';
          end;
        end;
      $04: // ACK
        begin
          FSysLastResult := D_OK;
          SetEvent(FSysSyncMsgEvent);
          FSysRecvBuffer := '';
        end;
      $05: // NAK
        begin
          if Length(FSysRecvBuffer) < 2 then exit;
          if (Ord(FSysRecvBuffer[2]) and E_NAK_UNDEFINED) > 0 then FSysLastResult := E_NAK_UNDEFINED
          else if (Ord(FSysRecvBuffer[2]) and E_NAK_SYNTAX) > 0 then FSysLastResult := E_NAK_SYNTAX
          else if (Ord(FSysRecvBuffer[2]) and E_NAK_CHECKSUM) > 0 then FSysLastResult := E_NAK_CHECKSUM
          else if (Ord(FSysRecvBuffer[2]) and E_NAK_PARITY) > 0 then FSysLastResult := E_NAK_PARITY
          else if (Ord(FSysRecvBuffer[2]) and E_NAK_OVERRUN) > 0 then FSysLastResult := E_NAK_OVERRUN
          else if (Ord(FSysRecvBuffer[2]) and E_NAK_FRAMING) > 0 then FSysLastResult := E_NAK_FRAMING
          else if (Ord(FSysRecvBuffer[2]) and E_NAK_TIMEOUT) > 0 then FSysLastResult := E_NAK_TIMEOUT
          else FSysLastResult := E_NAK_IGNORED;
          SetEvent(FSysSyncMsgEvent);
          FSysRecvBuffer := '';
        end;
      $06: // Error
        begin
          if Length(FSysRecvBuffer) < 5 then exit;
          FSysLastResult := PAnsiCharToInt(@FSysRecvBuffer[2]);
          SetEvent(FSysSyncMsgEvent);
          FSysRecvBuffer := '';
        end;
      else
        begin
          FSysLastResult := D_FALSE;
          SetEvent(FSysSyncMsgEvent);
          FSysRecvBuffer := '';
        end;
    end;
  finally
    FSysInCritSec.Leave;
  end;
end;

procedure TDCSMgr.UDPNotifyRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);
{const
  EventStatusNames: array[TEventStatus] of AnsiString =
    ('', 'Error', 'Skipped', 'Idle', 'Loading', 'Loaded', 'Cueing', 'Cued',
     'StandByOff', 'StandByOn', 'Preroll', 'OnAir',
     'Finish', 'Finishing', 'Finished', 'Done'); }
var
  ByteCount: Integer;
  CMD1, CMD2: Byte;
  NameLen: Integer;
  DeviceHandle: TDeviceHandle;
  EventID: TEventID;
  DeviceStatus: TDeviceStatus;
  EventStatus: TEventStatus;

  ChannelID: Word;
  EventOverall: TEventOverall;
begin
  FNotifyCritSec.Enter;
  try
    if (ADataSize <= 0) then exit;
    FNotifyBuffer := FNotifyBuffer + AData;

    if (Length(FNotifyBuffer) < 1) then exit;

    case Ord(FNotifyBuffer[1]) of
      $02:
      begin
        if (Length(FNotifyBuffer) < 3) then exit;

        ByteCount := PAnsiCharToWord(@FNotifyBuffer[2]);
        if (ByteCount > 0) and (Length(FNotifyBuffer) = ByteCount + 4) then
        begin
          if (CheckSum(FNotifyBuffer)) then
          begin
            CMD1 := Ord(FNotifyBuffer[4]);
            CMD2 := Ord(FNotifyBuffer[5]);
            FNotifyData := System.Copy(FNotifyBuffer, 6, ByteCount - 2);

            case CMD1 of
              $00:
                case CMD2 of
                  $80:
                    begin
                      DeviceHandle := PAnsiCharToInt(@FNotifyData[1]);
                      FNotifyData := Copy(FNotifyData, 5, Length(FNotifyData));

                      Move(FNotifyData[1], DeviceStatus, SizeOf(TDeviceStatus));
                      if (Assigned(@DeviceStatusNotifyProc)) then
                        DeviceStatusNotifyProc(PChar(String(ABindingIP)), DeviceHandle, DeviceStatus);
                    end;
                  $81:
                    begin
                      Move(FNotifyData[1], EventID, SizeOf(TEventID));
                      Move(FNotifyData[SizeOf(TEventID) + 1], EventStatus, Sizeof(TEventStatus));
//                      EventStatus.State := TEventState(Ord(FNotifyData[SizeOf(TEventID) + 1]));
//                      EventStatus.ErrorCode := TErrorCode(PAnsiCharToInt(@FNotifyData[SizeOf(TEventID) + SizeOf(TEventState) + 1]));
//                      WriteLog(EventIDToAnsiString(EventID) + ' : ' + EventStatusNames[EventStatus]);
                      if (Assigned(@EventStatusNotifyProc)) then
                        EventStatusNotifyProc(PChar(String(ABindingIP)), EventID, EventStatus);
//                      FEventNotifyThread.SetEventStatus(EventID, EventStatus);
                    end;
                  $82:
                    begin
                      DeviceHandle := PAnsiCharToWord(@FNotifyData[1]);
                      FNotifyData := Copy(FNotifyData, 5, Length(FNotifyData));

                      Move(FNotifyData[1], EventOverall, SizeOf(TEventOverall));
                      if (Assigned(@EventOverallNotifyProc)) then
                        EventOverallNotifyProc(PChar(String(ABindingIP)), DeviceHandle, EventOverall);
                    end;
                end;
            end;
          end
          else
            FLastResult := E_NAK_CHECKSUM;

          FNotifyBuffer := '';
        end
        else if ((ByteCount <= 0) or (Length(FNotifyBuffer) > ByteCount + 4)) then
        begin
          FNotifyBuffer := '';
        end;
      end;
      else
      begin
        FNotifyBuffer := '';
      end;
    end;
  finally
    FNotifyCritSec.Leave;
  end;
end;

procedure TDCSMgr.UDPInRead(const ABindingIP: AnsiString; const AData: AnsiString; const ADataSize: Integer);
var
  ByteCount: Integer;
  CMD1, CMD2: Byte;
  Device: TDeviceHandle;
begin
  FInCritSec.Enter;
  try
//      WriteLog('UDPInRead1');
    if (ADataSize <= 0) then exit;
//      WriteLog('UDPInRead2');
    FRecvBuffer := FRecvBuffer + AData;
//      WriteLog('UDPInRead3');

    if (Length(FRecvBuffer) < 1) then exit;
//      WriteLog('UDPInRead4');

    case Ord(FRecvBuffer[1]) of
      $02:
        begin
//      WriteLog('UDPInRead5');
          if (Length(FRecvBuffer) < 3) then exit;
//      WriteLog('UDPInRead6');

          ByteCount := PAnsiCharToWord(@FRecvBuffer[2]);
//      WriteLog('UDPInRead7');
          if (ByteCount > 0) and (Length(FRecvBuffer) = ByteCount + 4) then
          begin
//      WriteLog('UDPInRead8');
            if (CheckSum(FRecvBuffer)) then
            begin
//      WriteLog('UDPInRead9');
              CMD1 := Ord(FRecvBuffer[4]);
              CMD2 := Ord(FRecvBuffer[5]);

              FRecvData := System.Copy(FRecvBuffer, 6, ByteCount - 2);

              if (CMD1 = FCMD1) and (CMD2 = FCMD2 + $80)  then
                FLastResult := D_OK
              else
                FLastResult := D_FALSE;
//      WriteLog('UDPInRead10');
            end
            else
              FLastResult := E_NAK_CHECKSUM;

//      WriteLog('UDPInRead11');
            SetEvent(FSyncMsgEvent);

  //          if (Length(FRecvBuffer) > ByteCount + 4) then
  //            FRecvBuffer := Copy(FRecvBuffer, ByteCount + 5, Length(FRecvBuffer))
  //          else
//      WriteLog('UDPInRead111');
              FRecvBuffer := '';
//      WriteLog('UDPInRead12');
          end
          else if (ByteCount <= 0) or (Length(FRecvBuffer) > ByteCount + 4) then
          begin
//      WriteLog('UDPInRead13');
            FLastResult := D_FALSE;
            SetEvent(FSyncMsgEvent);
            FRecvBuffer:= '';
//      WriteLog('UDPInRead14');
          end;
        end;
      $04: // ACK
        begin
//      WriteLog('UDPInRead15');
          FLastResult := D_OK;
          SetEvent(FSyncMsgEvent);
          FRecvBuffer := '';
//      WriteLog('UDPInRead16');
        end;
      $05: // NAK
        begin
//      WriteLog('UDPInRead17');
          if Length(FRecvBuffer) < 2 then exit;
//      WriteLog('UDPInRead18');
          if (Ord(FRecvBuffer[2]) and E_NAK_UNDEFINED) > 0 then FLastResult := E_NAK_UNDEFINED
          else if (Ord(FRecvBuffer[2]) and E_NAK_SYNTAX) > 0 then FLastResult := E_NAK_SYNTAX
          else if (Ord(FRecvBuffer[2]) and E_NAK_CHECKSUM) > 0 then FLastResult := E_NAK_CHECKSUM
          else if (Ord(FRecvBuffer[2]) and E_NAK_PARITY) > 0 then FLastResult := E_NAK_PARITY
          else if (Ord(FRecvBuffer[2]) and E_NAK_OVERRUN) > 0 then FLastResult := E_NAK_OVERRUN
          else if (Ord(FRecvBuffer[2]) and E_NAK_FRAMING) > 0 then FLastResult := E_NAK_FRAMING
          else if (Ord(FRecvBuffer[2]) and E_NAK_TIMEOUT) > 0 then FLastResult := E_NAK_TIMEOUT
          else FLastResult := E_NAK_IGNORED;
          SetEvent(FSyncMsgEvent);
          FRecvBuffer := '';
//      WriteLog('UDPInRead19');
        end;
      $06: // Error
        begin
//      WriteLog('UDPInRead19-1');
          if Length(FRecvBuffer) < 5 then exit;
//      WriteLog('UDPInRead19-2');
          FLastResult := PAnsiCharToInt(@FRecvBuffer[2]);
          SetEvent(FSyncMsgEvent);
          FRecvBuffer := '';
//      WriteLog('UDPInRead19-3');
        end;
      else
        begin
//      WriteLog('UDPInRead20');
          FLastResult := D_FALSE;
          SetEvent(FSyncMsgEvent);
          FRecvBuffer := '';
//      WriteLog('UDPInRead21');
        end;
    end;
//      WriteLog('UDPInRead22');
  finally
    FInCritSec.Leave;
  end;
end;

// 0X00 System Control
function TDCSMgr.SysIsAlive(AIP: PChar; var AIsAlive: Boolean): Integer;
var
  Buffer: AnsiString;
begin
  AIsAlive := False;

  Buffer := '';
  Result := TransmitSysCommand(AIP, FSysOutPort, $00, $00, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    if (Length(FSysRecvData) < 1) then exit;
    AIsAlive := PAnsiCharToBool(@FSysRecvData[1]);
  end;
end;

function TDCSMgr.SysIsMain(AIP: PChar; var AIsMain: Boolean): Integer;
var
  Buffer: AnsiString;
begin
  AIsMain := False;

  Buffer := '';
  Result := TransmitSysCommand(AIP, FSysOutPort, $00, $01, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    if (Length(FSysRecvData) < 1) then exit;
    AIsMain := PAnsiCharToBool(@FSysRecvData[1]);
  end;
end;

// 0X00 Device Control
function TDCSMgr.Open(AID: Word; AIP: PChar; ADeviceName: PChar; var AHandle: TDeviceHandle): Integer;
var
  Buffer: AnsiString;
  P: PDCSDevice;
begin
  Result := D_FALSE;
  AHandle := 0;

  P := GetDeviceByName(AID, ADeviceName);
  if P = nil then
  begin
    P := New(PDCSDevice);
    FillChar(P^, SizeOf(TDCSDevice), #0);

    P^.Handle := -1;
    StrPCopy(P^.Name, ADeviceName);
    P^.ID := AID;
    StrPCopy(P^.IP, AIP);
    P^.Port := FOutPort;
    FDeviceList.Add(P);
  end;

  Buffer := ADeviceName;
  Result := TransmitCommand(P^.IP, P^.Port, $10, $00, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    if (Length(FRecvData) < 4) then exit;
    AHandle := PAnsiCharToInt(@FRecvData[1]);
//    P^.Handle := AHandle;
//    ShowMessage(IntToStr(PAnsiCharToInt(@FRecvData[1])));
//    ShowMessage(Format('ID=%d, IP=%s, DeviceName=%s, DeviceHandle=%d', [AID, AIP, ADeviceName, AHandle]));
  end
  else
    AHandle := INVALID_DEVICE_HANDLE;

  P^.Handle := AHandle;
end;

function TDCSMgr.Close(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(P^.Handle);
  Result := TransmitCommand(P^.IP, P^.Port, $10, $01, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    I := FDeviceList.IndexOf(P);
    FDeviceList.Delete(I);
    Dispose(P);
  end;
end;

function TDCSMgr.Reset(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + WordToAnsiString(AChannelID);
  Result := TransmitCommand(P^.IP, P^.Port, $10, $02, Buffer, Length(Buffer));
end;

function TDCSMgr.SetControlBy(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $10, $03, Buffer, Length(Buffer));
end;

// Video server only
// 0X20 Immediate Control
function TDCSMgr.Stop(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $00, Buffer, Length(Buffer));
end;

function TDCSMgr.Play(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $01, Buffer, Length(Buffer));
end;

function TDCSMgr.Pause(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $02, Buffer, Length(Buffer));
end;

function TDCSMgr.Continue(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $03, Buffer, Length(Buffer));
end;

function TDCSMgr.FastFoward(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $04, Buffer, Length(Buffer));
end;

function TDCSMgr.FastRewind(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $05, Buffer, Length(Buffer));
end;

function TDCSMgr.Jog(AID: Word; AHandle: TDeviceHandle; AFrameOrSpeed: Double): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + DoubleToAnsiString(AFrameOrSpeed);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $06, Buffer, Length(Buffer));
end;

function TDCSMgr.Shuttle(AID: Word; AHandle: TDeviceHandle; ASpeed: Double): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + DoubleToAnsiString(ASpeed);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $07, Buffer, Length(Buffer));
end;

function TDCSMgr.StandBy(AID: Word; AHandle: TDeviceHandle; AOn: Boolean): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  case AOn of
    True: Data := AnsiChar($01);
    else Data := AnsiChar($00);
  end;
  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $20, $08, Buffer, Length(Buffer));
end;

function TDCSMgr.Eject(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $09, Buffer, Length(Buffer));
end;

function TDCSMgr.Preroll(AID: Word; AHandle: TDeviceHandle): Integer; // VCR only
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $10, Buffer, Length(Buffer));
end;

function TDCSMgr.Rec(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $20, Buffer, Length(Buffer));
end;

function TDCSMgr.AutoEdit(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $20, $21, Buffer, Length(Buffer));
end;

// 0X30 Preset/Select Commands
function TDCSMgr.SetPortMode(AID: Word; AHandle: TDeviceHandle; APortMode: TPortMode): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + AnsiChar(Ord(APortMode));
  Result := TransmitCommand(P^.IP, P^.Port, $30, $00, Buffer, Length(Buffer));
end;

function TDCSMgr.SetAutoStatus(AID: Word; AHandle: TDeviceHandle; AAutoStatus: Boolean): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + AnsiChar(Ord(AAutoStatus));
  Result := TransmitCommand(P^.IP, P^.Port, $30, $01, Buffer, Length(Buffer));
end;

function TDCSMgr.PlayCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ACueTC, ADuration: TTimecode): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + WordToAnsiString(Length(AnsiString(AMediaID))) + AnsiString(AMediaID) + IntToAnsiString(ACueTC) + IntToAnsiString(ADuration);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $10, Buffer, Length(Buffer));
end;

function TDCSMgr.RecordCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ADuration: TTimecode): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + WordToAnsiString(Length(AnsiString(AMediaID))) + AnsiString(AMediaID) + IntToAnsiString(ADuration);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $20, Buffer, Length(Buffer));
end;

function TDCSMgr.IDRename(AID: Word; AHandle: TDeviceHandle; ASourceID, ATargetID: String): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + WordToAnsiString(Length(AnsiString(ASourceID))) + AnsiString(ASourceID) + WordToAnsiString(Length(AnsiString(ATargetID))) + AnsiString(ATargetID);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $30, Buffer, Length(Buffer));
end;

function TDCSMgr.IDDelete(AID: Word; AHandle: TDeviceHandle; AMediaID: String): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + WordToAnsiString(Length(AnsiString(AMediaID))) + AnsiString(AMediaID);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $31, Buffer, Length(Buffer));
end;

function TDCSMgr.InEntry(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $40, Buffer, Length(Buffer));
end;

function TDCSMgr.OutEntry(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $41, Buffer, Length(Buffer));
end;

function TDCSMgr.AInEntry(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $42, Buffer, Length(Buffer));
end;

function TDCSMgr.AOutEntry(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $43, Buffer, Length(Buffer));
end;

function TDCSMgr.InReset(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $44, Buffer, Length(Buffer));
end;

function TDCSMgr.OutReset(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $45, Buffer, Length(Buffer));
end;

function TDCSMgr.AInReset(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $46, Buffer, Length(Buffer));
end;

function TDCSMgr.AOutReset(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $47, Buffer, Length(Buffer));
end;

function TDCSMgr.EditPreset(AID: Word; AHandle: TDeviceHandle; AData1, AData2: Byte): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + AnsiChar(AData1) + AnsiChar(AData2);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $48, Buffer, Length(Buffer));
end;

function TDCSMgr.SetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl, AInput, AInputLvl: Integer): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + IntToAnsiString(AOutput) + IntToAnsiString(AOutputLvl) + IntToAnsiString(AInput) + IntToAnsiString(AInputLvl);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $50, Buffer, Length(Buffer));
end;

function TDCSMgr.SetXpt(AID: Word; AHandle: TDeviceHandle; AInput: Integer): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + IntToAnsiString(AInput);
  Result := TransmitCommand(P^.IP, P^.Port, $30, $60, Buffer, Length(Buffer));
end;

// 0X40 Sense Queries
function TDCSMgr.GetDeviceStatus(AID: Word; AHandle: TDeviceHandle; var AStatus: TDeviceStatus): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $00, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) > 1) then
      Move(FRecvData[1], AStatus, SizeOf(TDeviceStatus));
  end
  else
    case AStatus.EventType of
      ET_SWITCHER: FillChar(AStatus.Switcher, SizeOf(TSwitcherStatus), #0);
      ET_PLAYER: FillChar(AStatus.Switcher, SizeOf(TPlayerStatus), #0);
      ET_GPI: FillChar(AStatus.Switcher, SizeOf(TGPIStatus), #0);
      ET_RSW: FillChar(AStatus.Switcher, SizeOf(TRSWStatus), #0);
    end;
end;

function TDCSMgr.GetStorageTimeRemaining(AID: Word; AHandle: TDeviceHandle; var ATotal, AAvailable: TTimecode; Extended: Boolean = False): Integer;
var
  P: PDCSDevice;
  T: TTypeConvTime;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  if Extended then Buffer := Buffer + AnsiChar($41)
  else Buffer := Buffer + AnsiChar($01);

  Result := TransmitCommand(P^.IP, P^.Port, $40, $01, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < 4) then exit;
    T.Frame   := Ord(FRecvData[1]);
    T.Second  := Ord(FRecvData[2]);
    T.Minute  := Ord(FRecvData[3]);
    T.Hour    := Ord(FRecvData[4]);
    ATotal    := T.vtDWord;

    if (Length(FRecvData) < 8) then exit;
    T.Frame   := Ord(FRecvData[5]);
    T.Second  := Ord(FRecvData[6]);
    T.Minute  := Ord(FRecvData[7]);
    T.Hour    := Ord(FRecvData[8]);
    AAvailable  := T.vtDWord;
  end
  else
  begin
    ATotal      := 0;
    AAvailable  := 0;
  end;
end;

function TDCSMgr.GetTC(AID: Word; AHandle: TDeviceHandle; var ACurTC: TTimecode): Integer;
var
  P: PDCSDevice;
  T: TTypeConvTime;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $10, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < 4) then exit;
    T.Frame   := Ord(FRecvData[1]);
    T.Second  := Ord(FRecvData[2]);
    T.Minute  := Ord(FRecvData[3]);
    T.Hour    := Ord(FRecvData[4]);
    ACurTC    := T.vtDWord;
  end
  else
  begin
    ACurTC    := 0;
  end;
end;

function TDCSMgr.GetRemainTC(AID: Word; AHandle: TDeviceHandle; var ARemainTC: TTimecode): Integer;
var
  P: PDCSDevice;
  T: TTypeConvTime;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $11, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < 4) then exit;
    T.Frame   := Ord(FRecvData[1]);
    T.Second  := Ord(FRecvData[2]);
    T.Minute  := Ord(FRecvData[3]);
    T.Hour    := Ord(FRecvData[4]);
    ARemainTC := T.vtDWord;
  end
  else
  begin
    ARemainTC    := 0;
  end;
end;

function TDCSMgr.GetList(AID: Word; AHandle: TDeviceHandle; var AIDList: TIDList; var ARemainIDCount: Integer): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;

  IDs, ID: AnsiString;
  I, IDLen: Integer;
begin
  Result := D_FALSE;
  FillChar(AIDList, SizeOf(TIDList), #0);
  ARemainIDCount := 0;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $20, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < SizeOf(TIDList)) then exit;
    Move(FRecvData[1], AIDList, SizeOf(TIDList));

    if (Length(FRecvData) < SizeOf(TIDList) + 4) then exit;
    ARemainIDCount := PAnsiCharToInt(@FRecvData[SizeOf(TIDList) + 1]);
  end;
end;

function TDCSMgr.GetNext(AID: Word; AHandle: TDeviceHandle; var AIDList: TIDList; var ARemainIDCount: Integer): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;

  IDs, ID: AnsiString;
  I, IDLen: Integer;
begin
  Result := D_FALSE;
  FillChar(AIDList, SizeOf(TIDList), #0);
  ARemainIDCount := 0;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $21, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < SizeOf(TIDList)) then exit;
    Move(FRecvData[1], AIDList, SizeOf(TIDList));

    if (Length(FRecvData) < SizeOf(TIDList) + 4) then exit;
    ARemainIDCount := PAnsiCharToInt(@FRecvData[SizeOf(TIDList) + 1]);
  end;
end;

function TDCSMgr.GetExist(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var AExist: Boolean): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + WordToAnsiString(Length(AnsiString(AMediaID))) + AnsiString(AMediaID);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $22, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < 1) then exit;
    AExist := Ord(FRecvData[1]) > $00;
  end
  else
  begin
    AExist := False;
  end;
end;

function TDCSMgr.GetSize(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var ADuration: TTimecode): Integer;
var
  P: PDCSDevice;
  T: TTypeConvTime;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + WordToAnsiString(Length(AnsiString(AMediaID))) + AnsiString(AMediaID);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $23, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < 4) then exit;
    T.Frame   := Ord(FRecvData[1]);
    T.Second  := Ord(FRecvData[2]);
    T.Minute  := Ord(FRecvData[3]);
    T.Hour    := Ord(FRecvData[4]);
    ADuration := T.vtDWord;
  end
  else
  begin
    ADuration := 0;
  end;
end;

function TDCSMgr.GetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl: Integer; var AInput, AInputLvl: Integer): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + IntToAnsiString(AOutput) + IntToAnsiString(AOutputLvl);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $50, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    if (Length(FRecvData) < 8) then exit;
    AInput    := PAnsiCharToInt(@FRecvData[1]);
    AInputLvl := PAnsiCharToInt(@FRecvData[5]);
  end
  else
  begin
    AInput    := -1;
    AInputLvl := -1;
  end;
end;

function TDCSMgr.GetXpt(AID: Word; AHandle: TDeviceHandle; var AInput: Integer): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $40, $60, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    if (Length(FRecvData) < 4) then exit;
    AInput := PAnsiCharToInt(@FRecvData[1]);
  end
  else
  begin
    AInput := -1;
  end;
end;

// 0X50 Event Control
function TDCSMgr.InputEvent(AID: Word; AHandle: TDeviceHandle; AEvent: TEvent): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  SetLength(Data, SizeOf(TEvent));
  Move(AEvent, Data[1], SizeOf(TEvent));

  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $50, $00, Buffer, Length(Buffer));
end;

function TDCSMgr.DeleteEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  SetLength(Data, SizeOf(TEventID));
  Move(AEventID, Data[1], SizeOf(TEventID));

  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $50, $01, Buffer, Length(Buffer));
end;

function TDCSMgr.ClearEvent(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + WordToAnsiString(AChannelID);
  Result := TransmitCommand(P^.IP, P^.Port, $50, $02, Buffer, Length(Buffer));
end;

function TDCSMgr.TakeEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; AStartTime: TEventTime): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  SetLength(Data, SizeOf(TEventID) + SizeOf(TEventTime));
  Move(AEventID, Data[1], SizeOf(TEventID));
  Move(AStartTime, Data[1 + SizeOf(TEventID)], SizeOf(TEventTime));

  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $50, $10, Buffer, Length(Buffer));
end;

function TDCSMgr.HoldEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  SetLength(Data, SizeOf(TEventID));
  Move(AEventID, Data[1], SizeOf(TEventID));

  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $50, $11, Buffer, Length(Buffer));
end;

function TDCSMgr.ChangeDurationEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; ADuration: TTimecode): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  SetLength(Data, SizeOf(TEventID));
  Move(AEventID, Data[1], SizeOf(TEventID));

  Buffer := IntToAnsiString(AHandle) + Data + DWordToAnsiString(ADuration);
  Result := TransmitCommand(P^.IP, P^.Port, $50, $12, Buffer, Length(Buffer));
end;

function TDCSMgr.OnAirCatchEvent(AID: Word; AHandle: TDeviceHandle): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $50, $13, Buffer, Length(Buffer));
end;

function TDCSMgr.GetOnAirEventID(AID: Word; AHandle: TDeviceHandle; var AOnAirEventID, ANextEventID: TEventID): Integer;
var
  P: PDCSDevice;
  Buffer: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle);
  Result := TransmitCommand(P^.IP, P^.Port, $50, $20, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    if (Length(FRecvData) < SizeOf(TEventID)) then exit;
    Move(FRecvData[1], AOnAirEventID, SizeOf(TEventID));

    if (Length(FRecvData) < SizeOf(TEventID) * 2) then exit;
    Move(FRecvData[SizeOf(TEventID) + 1], ANextEventID, SizeOf(TEventID));
  end;
end;

function TDCSMgr.GetEventInfo(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime; var ADurationTC: TTimecode): Integer; // 0X40.21
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  SetLength(Data, SizeOf(TEventID));
  Move(AEventID, Data[1], SizeOf(TEventID));

  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $50, $21, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    if (Length(FRecvData) < 4) then exit;
    AStartTime.T := PAnsiCharToDWord(@FRecvData[1]);

    if (Length(FRecvData) < 12) then exit;
    AStartTime.D := PAnsiCharToDouble(@FRecvData[5]);

    if (Length(FRecvData) < 16) then exit;
    ADurationTC  := PAnsiCharToDWord(@FRecvData[13]);
  end;
end;

function TDCSMgr.GetEventStatus(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AEventStatus: TEventStatus): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  SetLength(Data, SizeOf(TEventID));
  Move(AEventID, Data[1], SizeOf(TEventID));

  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $50, $22, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < SizeOf(TEventStatus)) then exit;
    Move(FRecvData[1], AEventStatus, SizeOf(TEventStatus));
  end;
end;

function TDCSMgr.GetEventStartTime(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  SetLength(Data, SizeOf(TEventID));
  Move(AEventID, Data[1], SizeOf(TEventID));

  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $50, $23, Buffer, Length(Buffer));
  if Result = D_OK then
  begin
    if (Length(FRecvData) < 4) then exit;
    AStartTime.T := PAnsiCharToInt(@FRecvData[1]);

    if (Length(FRecvData) < 12) then exit;
    AStartTime.D := PAnsiCharToDouble(@FRecvData[5]);
  end;
end;

function TDCSMgr.GetEventOverall(AID: Word; AHandle: TDeviceHandle; var AEventOverall: TEventOverall): Integer;
var
  P: PDCSDevice;
  Buffer, Data: AnsiString;
  I: Integer;
begin
  Result := D_FALSE;
  FillChar(AEventOverall, SizeOf(TEventOverall), #0);

  Result := GetValidDeviceHandle(AID, AHandle, P);
  if (Result <> D_OK) then exit;

  Buffer := IntToAnsiString(AHandle) + Data;
  Result := TransmitCommand(P^.IP, P^.Port, $50, $24, Buffer, Length(Buffer));
  if (Result = D_OK) then
  begin
    if (Length(FRecvData) < SizeOf(TEventOverall)) then exit;
    Move(FRecvData[1], AEventOverall, SizeOf(TEventOverall));
  end;
end;

{ TDCSEventNotifyThread }

constructor TDCSEventNotifyThread.Create(ADCSMgr: TDCSMgr);
begin
  FDCSMgr := ADCSMgr;

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TDCSEventNotifyThread.Destroy;
begin
  FDCSMgr := nil;

  inherited Destroy;
end;

procedure TDCSEventNotifyThread.SetEventStatus(AEventID: TEventID; AEventStatus: TEventStatus);
begin
  FEventID := AEventID;
  FEventStatus := AEventStatus;
  SetEvent(FExecuteEvent);
end;

procedure TDCSEventNotifyThread.DoControl;
begin
//  if Assigned(@EventStatusNotifyProc) then EventStatusNotifyProc(FEventID, FEventStatus);
//  ResetEvent(FExecuteEvent);
end;

procedure TDCSEventNotifyThread.Execute;
var
  R: Integer;
begin
  FExecuteEvent := CreateEvent(nil, True, False, nil);

  while not Terminated do
  begin
        R := WaitForSingleObject(FExecuteEvent, INFINITE);
        if R = WAIT_OBJECT_0 then
        (DoControl);
  end;
  CloseHandle(FExecuteEvent);
end;

end.
