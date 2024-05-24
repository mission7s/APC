unit DLLConsts;

interface

uses System.Classes, Winapi.Windows, System.SysUtils, Dialogs, Forms,
  UnitCommons, UnitUDPIn, UnitUDPOut,
  DCSMgr;

type
  // Notify function of status
  TDeviceStatusNotifyProc = procedure (ADCSIP: PChar; ADeviceHandle: TDeviceHandle; AStatus: TDeviceStatus); stdcall;
  PDeviceStatusNotifyProc = ^TDeviceStatusNotifyProc;

  // Notify function of event status
  TEventStatusNotifyProc = procedure (ADCSIP: PChar; AEventID: TEventID; AEvenStatus: TEventStatus); stdcall;
  PEventStatusNotifyProc = ^TEventStatusNotifyProc;

  // Notify function of event overall
  TEventOverallNotifyProc = procedure (ADCSIP: PChar; ADeviceHandle: TDeviceHandle; AEvenOverall: TEventOverall); stdcall;
  PEventOverallNotifyProc = ^TEventOverallNotifyProc;

var
  DeviceStatusNotifyProc: TDeviceStatusNotifyProc;
  EventStatusNotifyProc: TEventStatusNotifyProc;
  EventOverallNotifyProc: TEventOverallNotifyProc;

function DCSSysInitialize(AInPort, AOutPort: Word; ATimeout: Cardinal = 1000): Integer; stdcall;
function DCSSysFinalize: Integer; stdcall;

function DCSSysLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall;
function DCSSysLogInPortDisable: Integer; stdcall;

function DCSSysLogOutPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall;
function DCSSysLogOutPortDisable: Integer; stdcall;

function DCSInitialize(ANotifyPort, AInPort, AOutPort: Word; ATimeout: Cardinal = 1000; ABroadcast: Boolean = False): Integer; stdcall;
function DCSFinalize: Integer; stdcall;

function DCSLogNotifyPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall;
function DCSLogNotifyPortDisable: Integer; stdcall;

function DCSLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall;
function DCSLogInPortDisable: Integer; stdcall;

function DCSLogOutPortEnable(ALogPath: PChar; ALogExt: PChar): Integer; stdcall;
function DCSLogOutPortDisable: Integer; stdcall;

function DCSDeviceStatusNotify(NotifyFunc: Pointer): Integer; stdcall;
function DCSEventStatusNotify(NotifyFunc: Pointer): Integer; stdcall;
function DCSEventOverallNotify(NotifyFunc: Pointer): Integer; stdcall;

// 0X00 System Control
function DCSSysIsAlive(AIP: PChar; var AIsAlive: Boolean): Integer; stdcall;
function DCSSysIsMain(AIP: PChar; var AIsMain: Boolean): Integer; stdcall;

// 0X10 Device Control
function DCSOpen(AID: Word; AIP: PChar; ADeviceName: PChar; var AHandle: TDeviceHandle): Integer; stdcall;
function DCSClose(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSReset(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer; stdcall;
function DCSSetControlBy(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;

// 0X20 Immediate Control
function DCSStop(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSPlay(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSPause(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSContinue(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSFastFoward(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSFastRewind(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSJog(AID: Word; AHandle: TDeviceHandle; AFrameOrSpeed: Double): Integer; stdcall;
function DCSShuttle(AID: Word; AHandle: TDeviceHandle; ASpeed: Double): Integer; stdcall;
function DCSStandBy(AID: Word; AHandle: TDeviceHandle; AOn: Boolean): Integer; stdcall;
function DCSEject(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSPreroll(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSRecord(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSAutoEdit(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only

// 0X30 Preset/Select Commands
function DCSSetPortMode(AID: Word; AHandle: TDeviceHandle; APortMode: TPortMode): Integer; stdcall; // Video server only
function DCSSetAutoStatus(AID: Word; AHandle: TDeviceHandle; AAutoStatus: Boolean = True): Integer; stdcall;
function DCSPlayCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ACueTC, ADuration: TTimecode): Integer; stdcall;
function DCSRecordCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ADuration: TTimecode): Integer; stdcall;
function DCSIDRename(AID: Word; AHandle: TDeviceHandle; ASourceID, ATargetID: String): Integer; stdcall;
function DCSIDDelete(AID: Word; AHandle: TDeviceHandle; AMediaID: String): Integer; stdcall;
function DCSInEntry(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
function DCSOutEntry(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
function DCSAInEntry(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
function DCSAOutEntry(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
function DCSInReset(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
function DCSOutReset(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
function DCSAInReset(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
function DCSAOutReset(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
function DCSEditPreset(AID: Word; AHandle: TDeviceHandle; AData1, AData2: Byte): Integer; stdcall; // VCR only
function DCSSetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl, AInput, AInputLvl: Integer): Integer; stdcall; // Router only
function DCSSetXpt(AID: Word; AHandle: TDeviceHandle; AInput: Integer): Integer; stdcall; // MCS only

// 0X40 Sense Queries
function DCSGetDeviceStatus(AID: Word; AHandle: TDeviceHandle; var AStatus: TDeviceStatus): Integer; stdcall;
function DCSGetStorageTimeRemaining(AID: Word; AHandle: TDeviceHandle; var ATotalTC, AAvailableTC: TTimecode; AExtended: Boolean = False): Integer; stdcall;
function DCSGetTC(AID: Word; AHandle: TDeviceHandle; var ACurTC: TTimecode): Integer; stdcall;
function DCSGetRemainTC(AID: Word; AHandle: TDeviceHandle; var ARemainTC: TTimecode): Integer; stdcall;
function DCSGetList(AID: Word; AHandle: TDeviceHandle; var AIDList: TIDList; var ARemainIDCount: Integer): Integer; stdcall;
function DCSGetNext(AID: Word; AHandle: TDeviceHandle; var AIDList: TIDList; var ARemainIDCount: Integer): Integer; stdcall;
function DCSGetExist(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var AExist: Boolean): Integer; stdcall;
function DCSGetSize(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var ADuration: TTimecode): Integer; stdcall;
function DCSGetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl: Integer; var AInput, AInputLvl: Integer): Integer; stdcall; // Router only
function DCSGetXpt(AID: Word; AHandle: TDeviceHandle; var AInput: Integer): Integer; stdcall; // MCS only

// 0X50 Event Control
function DCSInputEvent(AID: Word; AHandle: TDeviceHandle; AEvent: TEvent): Integer; stdcall;
function DCSDeleteEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer; stdcall;
function DCSClearEvent(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer; stdcall;
function DCSTakeEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; AStartTime: TEventTime): Integer; stdcall;
function DCSHoldEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer; stdcall;
function DCSChangetDurationEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; ADuration: TTimecode): Integer; stdcall;
function DCSOnAirCatchEvent(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
function DCSGetOnAirEventID(AID: Word; AHandle: TDeviceHandle; var AOnAirEventID, ANextEventID: TEventID): Integer; stdcall;
function DCSGetEventInfo(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime; var ADurationTC: TTimecode): Integer; stdcall;
function DCSGetEventStatus(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AEvenStatus: TEventStatus): Integer; stdcall;
function DCSGetEventStartTime(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime): Integer; stdcall;
function DCSGetEventOverall(AID: Word; AHandle: TDeviceHandle; var AEventOverall: TEventOverall): Integer; stdcall;

exports // exports 절이 추가된다.
  DCSSysInitialize name 'DCSSysInitialize',
  DCSSysFinalize name 'DCSSysFinalize',

  DCSSysLogInPortEnable name 'DCSSysLogInPortEnable',
  DCSSysLogInPortDisable name 'DCSSysLogInPortDisable',

  DCSSysLogOutPortEnable name 'DCSSysLogOutPortEnable',
  DCSSysLogOutPortDisable name 'DCSSysLogOutPortDisable',

  DCSInitialize name 'DCSInitialize',
  DCSFinalize name 'DCSFinalize',

  DCSLogNotifyPortEnable name 'DCSLogNotifyPortEnable',
  DCSLogNotifyPortDisable name 'DCSLogNotifyPortDisable',

  DCSLogInPortEnable name 'DCSLogInPortEnable',
  DCSLogInPortDisable name 'DCSLogInPortDisable',

  DCSLogOutPortEnable name 'DCSLogOutPortEnable',
  DCSLogOutPortDisable name 'DCSLogOutPortDisable',

  DCSDeviceStatusNotify name 'DCSDeviceStatusNotify',
  DCSEventStatusNotify name 'DCSEventStatusNotify',
  DCSEventOverallNotify name 'DCSEventOverallNotify',

  DCSSysIsAlive name 'DCSSysIsAlive',
  DCSSysIsMain name 'DCSSysIsMain',

  DCSOpen name 'DCSOpen',
  DCSClose name 'DCSClose',
  DCSReset name 'DCSReset',
  DCSSetControlBy name 'DCSSetControlBy',

  DCSStop name 'DCSStop',
  DCSPlay name 'DCSPlay',
  DCSPause name 'DCSPause',
  DCSContinue name 'DCSContinue',
  DCSFastFoward name 'DCSFastFoward',
  DCSFastRewind name 'DCSFastRewind',
  DCSJog name 'DCSJog',
  DCSShuttle name 'DCSShuttle',
  DCSStandBy name 'DCSStandBy',
  DCSEject name 'DCSEject',
  DCSPreroll name 'DCSPreroll',
  DCSRecord name 'DCSRecord',
  DCSAutoEdit name 'DCSAutoEdit',

  DCSSetPortMode name 'DCSSetPortMode',
  DCSSetAutoStatus name 'DCSSetAutoStatus',
  DCSPlayCue name 'DCSPlayCue',
  DCSRecordCue name 'DCSRecordCue',
  DCSIDRename name 'DCSIDRename',
  DCSIDDelete name 'DCSIDDelete',
  DCSInEntry name 'DCSInEntry',
  DCSOutEntry name 'DCSOutEntry',
  DCSAInEntry name 'DCSAInEntry',
  DCSAOutEntry name 'DCSAOutEntry',
  DCSInReset name 'DCSInReset',
  DCSOutReset name 'DCSOutReset',
  DCSAInReset name 'DCSAInReset',
  DCSAOutReset name 'DCSAOutReset',
  DCSEditPreset name 'DCSEditPreset',
  DCSSetRoute name 'DCSSetRoute',
  DCSSetXpt name 'DCSSetXpt',

  DCSGetDeviceStatus name 'DCSGetDeviceStatus',
  DCSGetStorageTimeRemaining name 'DCSGetStorageTimeRemaining',
  DCSGetTC name 'DCSGetTC',
  DCSGetRemainTC name 'DCSGetRemainTC',
  DCSGetList name 'DCSGetList',
  DCSGetNext name 'DCSGetNext',
  DCSGetExist name 'DCSGetExist',
  DCSGetSize name 'DCSGetSize',
  DCSGetRoute name 'DCSGetRoute',
  DCSGetXpt name 'DCSGetXpt',

  DCSInputEvent name 'DCSInputEvent',
  DCSDeleteEvent name 'DCSDeleteEvent',
  DCSClearEvent name 'DCSClearEvent',
  DCSTakeEvent name 'DCSTakeEvent',
  DCSHoldEvent name 'DCSHoldEvent',
  DCSChangetDurationEvent name 'DCSChangetDurationEvent',
  DCSOnAirCatchEvent name 'DCSOnAirCatchEvent',
  DCSGetOnAirEventID name 'DCSGetOnAirEventID',
  DCSGetEventInfo name 'DCSGetEventInfo',
  DCSGetEventStatus name 'DCSGetEventStatus',
  DCSGetEventStartTime name 'DCSGetEventStartTime',
  DCSGetEventOverall name 'DCSGetEventOverall';

implementation

function DCSSysInitialize(AInPort, AOutPort: Word; ATimeout: Cardinal = 1000): Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    if (FUDPSysIn = nil) then
    begin
      FUDPSysIn := TUDPIn.Create;
      FUDPSysIn.OnUDPRead := UDPSysInRead;
    end;

    if (FUDPSysOut = nil) then
    begin
      FUDPSysOut := TUDPOut.Create;
    end;

    if (WaitForSingleObject(FUDPSysIn.Handle, 0) = WAIT_OBJECT_0) then exit;
    if (WaitForSingleObject(FUDPSysOut.Handle, 0) = WAIT_OBJECT_0) then exit;

    FSysInPort  := AInPort;
    FSysOutPort := AOutPort;

    FSysTimeout := ATimeout;

    FUDPSysIn.Port := AInPort;
    FUDPSysIn.Broadcast := False;
    FUDPSysIn.AsyncMode := True;
    FUDPSysIn.Start;

    while not (FUDPSysIn.Started) do
      Sleep(30);

    FUDPSysOut.Port := AOutPort;
    FUDPSysOut.Broadcast := False;
    FUDPSysOut.AsyncMode := True;
    FUDPSysOut.Start;

    while not (FUDPSysOut.Started) do
      Sleep(30);
  end;

  Result := D_OK;
end;

function DCSSysFinalize: Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    if (FUDPSysIn <> nil) then
    begin
//      if (WaitForSingleObject(FUDPSysIn.Handle, 0) = WAIT_OBJECT_0) then
      begin
        FUDPSysIn.Close;
        FUDPSysIn.Terminate;
        FUDPSysIn.WaitFor;
      end;
      FreeAndNil(FUDPSysIn);
    end;

    if (FUDPSysOut <> nil) then
    begin
//      if (WaitForSingleObject(FUDPSysOut.Handle, 0) = WAIT_OBJECT_0) then
      begin
        FUDPSysOut.Close;
        FUDPSysOut.Terminate;
        FUDPSysOut.WaitFor;
      end;
      FreeAndNil(FUDPSysOut);
    end;
  end;

  Result := D_OK;
end;


function DCSSysLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPSysIn.LogEnabled := True;
    FUDPSysIn.LogPath    := String(ALogPath);
    FUDPSysIn.LogExt     := Format('%d_%s', [FUDPSysIn.Port, ALogExt]);
  end;

  Result := D_OK;
end;

function DCSSysLogInPortDisable: Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPSysIn.LogEnabled := False;
  end;

  Result := D_OK;
end;

function DCSSysLogOutPortEnable(ALogPath: PChar; ALogExt: PChar): Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPSysOut.LogEnabled := True;
    FUDPSysOut.LogPath    := String(ALogPath);
    FUDPSysOut.LogExt     := Format('%d_%s', [FUDPSysOut.Port, ALogExt]);
  end;

  Result := D_OK;
end;

function DCSSysLogOutPortDisable: Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPSysOut.LogEnabled := False;
  end;

  Result := D_OK;
end;

function DCSInitialize(ANotifyPort, AInPort, AOutPort: Word; ATimeout: Cardinal = 1000; ABroadcast: Boolean = False): Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    if (FUDPNotify = nil) then
    begin
      FUDPNotify := TUDPIn.Create;
      FUDPNotify.OnUDPRead := UDPNotifyRead;
    end;

    if (FUDPIn = nil) then
    begin
      FUDPIn := TUDPIn.Create;
      FUDPIn.OnUDPRead := UDPInRead;
    end;

    if (FUDPOut = nil) then
    begin
      FUDPOut := TUDPOut.Create;
    end;

    if (WaitForSingleObject(FUDPNotify.Handle, 0) = WAIT_OBJECT_0) then exit;
    if (WaitForSingleObject(FUDPIn.Handle, 0) = WAIT_OBJECT_0) then exit;
    if (WaitForSingleObject(FUDPOut.Handle, 0) = WAIT_OBJECT_0) then exit;

    FNotifyPort := ANotifyPort;

    FInPort  := AInPort;
    FOutPort := AOutPort;

    FTimeout := ATimeout;

{  UDPNotify.LogPath := ExtractFilePath(Application.ExeName) + 'UDPLog\';
  UDPNotify.LogExt  := Format('%d.log', [UDPNotify.Port]);
  UDPNotify.LogEnabled := True; }
    FUDPNotify.Port := ANotifyPort;
    FUDPNotify.Broadcast := ABroadcast;
    FUDPNotify.AsyncMode := True;
    FUDPNotify.Start;

    while not (FUDPNotify.Started) do
      Sleep(30);

    FUDPIn.Port := AInPort;
    FUDPIn.Broadcast := ABroadcast;
    FUDPIn.AsyncMode := True;
    FUDPIn.Start;
{  FUDPIn.LogPath := ExtractFilePath(Application.ExeName) + 'UDPLog\';
  FUDPIn.LogExt  := Format('%d.log', [FUDPIn.Port]);
  FUDPIn.LogEnabled := True; }

    while not (FUDPIn.Started) do
      Sleep(30);

//    Sleep(2000);

    FUDPOut.Port := AOutPort;
    FUDPOut.Broadcast := ABroadcast;
    FUDPOut.AsyncMode := True;
    FUDPOut.Start;
{  FUDPOut.LogPath := ExtractFilePath(Application.ExeName) + 'UDPLog\';
  FUDPOut.LogExt  := Format('%d.log', [FUDPOut.Port]);
  FUDPOut.LogEnabled := True; }

    while not (FUDPOut.Started) do
      Sleep(30);

    Sleep(1000);
  end;

  Result := D_OK;
end;

function DCSFinalize: Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    @DeviceStatusNotifyProc := nil;
    @EventStatusNotifyProc := nil;
    @EventOverallNotifyProc := nil;

    if (FUDPNotify <> nil) then
    begin
//      if (WaitForSingleObject(FUDPNotify.Handle, 0) = WAIT_OBJECT_0) then
      begin
        FUDPNotify.Close;
        FUDPNotify.Terminate;
        FUDPNotify.WaitFor;
      end;
      FreeAndNil(FUDPNotify);
    end;

    if (FUDPIn <> nil) then
    begin
//      if (WaitForSingleObject(FUDPIn.Handle, 0) = WAIT_OBJECT_0) then
      begin
        FUDPIn.Close;
        FUDPIn.Terminate;
        FUDPIn.WaitFor;
      end;
      FreeAndNil(FUDPIn);
    end;

    if (FUDPOut <> nil) then
    begin
//      if (WaitForSingleObject(FUDPOut.Handle, 0) = WAIT_OBJECT_0) then
      begin
        FUDPOut.Close;
        FUDPOut.Terminate;
        FUDPOut.WaitFor;
      end;
      FreeAndNil(FUDPOut);
    end;

{    if (WaitForSingleObject(UDPNotify.Handle, 0) = WAIT_OBJECT_0) then
    begin
      UDPNotify.Close;
      UDPNotify.Terminate;
      UDPNotify.WaitFor;
    end
    else exit;

    if (WaitForSingleObject(UDPIn.Handle, 0) = WAIT_OBJECT_0) then
    begin
      UDPIn.Close;
      UDPIn.Terminate;
      UDPIn.WaitFor;
    end
    else exit;

    if (WaitForSingleObject(UDPOut.Handle, 0) = WAIT_OBJECT_0) then
    begin
      UDPOut.Close;
      UDPOut.Terminate;
      UDPOut.WaitFor;
    end
    else exit; }
  end;

  Result := D_OK;
end;

function DCSLogNotifyPortEnable(ALogPath: PChar; ALogExt: PChar): Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPNotify.LogEnabled := True;
    FUDPNotify.LogPath    := String(ALogPath);
    FUDPNotify.LogExt     := Format('%d_%s', [FUDPNotify.Port, ALogExt]);
  end;

  Result := D_OK;
end;

function DCSLogNotifyPortDisable: Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPNotify.LogEnabled := False;
  end;

  Result := D_OK;
end;

function DCSLogInPortEnable(ALogPath: PChar; ALogExt: PChar): Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPIn.LogEnabled := True;
    FUDPIn.LogPath    := String(ALogPath);
    FUDPIn.LogExt     := Format('%d_%s', [FUDPIn.Port, ALogExt]);
  end;

  Result := D_OK;
end;

function DCSLogInPortDisable: Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPIn.LogEnabled := False;
  end;

  Result := D_OK;
end;

function DCSLogOutPortEnable(ALogPath: PChar; ALogExt: PChar): Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPOut.LogEnabled := True;
    FUDPOut.LogPath    := String(ALogPath);
    FUDPOut.LogExt     := Format('%d_%s', [FUDPOut.Port, ALogExt]);
  end;

  Result := D_OK;
end;

function DCSLogOutPortDisable: Integer;
begin
  Result := D_FALSE;

  with VDCSMgr do
  begin
    FUDPOut.LogEnabled := False;
  end;

  Result := D_OK;
end;

function DCSDeviceStatusNotify(NotifyFunc: Pointer): Integer;
begin
  Result := D_FALSE;
  if (NotifyFunc = nil) then exit;

  @DeviceStatusNotifyProc := NotifyFunc;
  Result := D_OK;
end;

function DCSEventStatusNotify(NotifyFunc: Pointer): Integer;
begin
  Result := D_FALSE;
  if (NotifyFunc = nil) then exit;

  @EventStatusNotifyProc := NotifyFunc;
  Result := D_OK;
end;

function DCSEventOverallNotify(NotifyFunc: Pointer): Integer;
begin
  Result := D_FALSE;
  if (NotifyFunc = nil) then exit;

  @EventOverallNotifyProc := NotifyFunc;
  Result := D_OK;
end;

// 0X00 System Control
function DCSSysIsAlive(AIP: PChar; var AIsAlive: Boolean): Integer;
begin
  Result := VDCSMgr.SysIsAlive(AIP, AIsAlive);
end;

function DCSSysIsMain(AIP: PChar; var AIsMain: Boolean): Integer;
begin
  Result := VDCSMgr.SysIsMain(AIP, AIsMain);
end;

// 0X10 Device Control
function DCSOpen(AID: Word; AIP: PChar; ADeviceName: PChar; var AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.Open(AID, AIP, ADeviceName, AHandle);
end;

function DCSClose(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.Close(AID, AHandle);
end;

function DCSReset(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer;
begin
  Result := VDCSMgr.Reset(AID, AHandle, AChannelID);
end;

function DCSSetControlBy(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
begin
  Result := VDCSMgr.SetControlBy(AID, AHandle);
end;

// 0X20 Immediate Control
function DCSStop(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.Stop(AID, AHandle);
end;

function DCSPlay(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.Play(AID, AHandle);
end;

function DCSPause(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.Pause(AID, AHandle);
end;

function DCSContinue(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
begin
  Result := VDCSMgr.Continue(AID, AHandle);
end;

function DCSFastFoward(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.FastFoward(AID, AHandle);
end;

function DCSFastRewind(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.FastRewind(AID, AHandle);
end;

function DCSJog(AID: Word; AHandle: TDeviceHandle; AFrameOrSpeed: Double): Integer;
begin
  Result := VDCSMgr.Jog(AID, AHandle, AFrameOrSpeed);
end;

function DCSShuttle(AID: Word; AHandle: TDeviceHandle; ASpeed: Double): Integer;
begin
  Result := VDCSMgr.Shuttle(AID, AHandle, ASpeed);
end;

function DCSStandBy(AID: Word; AHandle: TDeviceHandle; AOn: Boolean): Integer;
begin
  Result := VDCSMgr.StandBy(AID, AHandle, AOn);
end;

function DCSEject(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.Eject(AID, AHandle);
end;

function DCSPreroll(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
begin
  Result := VDCSMgr.Preroll(AID, AHandle);
end;

function DCSRecord(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.Rec(AID, AHandle);
end;

function DCSAutoEdit(AID: Word; AHandle: TDeviceHandle): Integer; stdcall; // VCR only
begin
  Result := VDCSMgr.AutoEdit(AID, AHandle);
end;

// 0X30 Preset/Select Commands
// Video server only
function DCSSetPortMode(AID: Word; AHandle: TDeviceHandle; APortMode: TPortMode): Integer;
begin
  Result := VDCSMgr.SetPortMode(AID, AHandle, APortMode);
end;

function DCSSetAutoStatus(AID: Word; AHandle: TDeviceHandle; AAutoStatus: Boolean = True): Integer; stdcall;
begin
  Result := VDCSMgr.SetAutoStatus(AID, AHandle, AAutoStatus);
end;

function DCSPlayCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ACueTC, ADuration: TTimecode): Integer;
begin
  Result := VDCSMgr.PlayCue(AID, AHandle, AMediaID, ACueTC, ADuration);
end;

function DCSRecordCue(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; ADuration: TTimecode): Integer;
begin
  Result := VDCSMgr.RecordCue(AID, AHandle, AMediaID, ADuration);
end;

function DCSIDRename(AID: Word; AHandle: TDeviceHandle; ASourceID, ATargetID: String): Integer;
begin
  Result := VDCSMgr.IDRename(AID, AHandle, ASourceID, ATargetID);
end;

function DCSIDDelete(AID: Word; AHandle: TDeviceHandle; AMediaID: String): Integer;
begin
  Result := VDCSMgr.IDDelete(AID, AHandle, AMediaID);
end;

function DCSInEntry(AID: Word; AHandle: TDeviceHandle): Integer; // VCR only
begin
  Result := VDCSMgr.InEntry(AID, AHandle);
end;

function DCSOutEntry(AID: Word; AHandle: TDeviceHandle): Integer; // VCR only
begin
  Result := VDCSMgr.OutEntry(AID, AHandle);
end;

function DCSAInEntry(AID: Word; AHandle: TDeviceHandle): Integer; // VCR only
begin
  Result := VDCSMgr.AInEntry(AID, AHandle);
end;

function DCSAOutEntry(AID: Word; AHandle: TDeviceHandle): Integer; // VCR only
begin
  Result := VDCSMgr.AOutEntry(AID, AHandle);
end;

function DCSInReset(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.InReset(AID, AHandle);
end;

function DCSOutReset(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.OutReset(AID, AHandle);
end;

function DCSAInReset(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.AInReset(AID, AHandle);
end;

function DCSAOutReset(AID: Word; AHandle: TDeviceHandle): Integer;
begin
  Result := VDCSMgr.AOutReset(AID, AHandle);
end;

function DCSEditPreset(AID: Word; AHandle: TDeviceHandle; AData1, AData2: Byte): Integer;
begin
  Result := VDCSMgr.EditPreset(AID, AHandle, AData1, AData2);
end;

function DCSSetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl, AInput, AInputLvl: Integer): Integer; // Router only
begin
  Result := VDCSMgr.SetRoute(AID, AHandle, AOutput, AOutputLvl, AInput, AInputLvl);
end;

function DCSSetXpt(AID: Word; AHandle: TDeviceHandle; AInput: Integer): Integer; stdcall; // MCS only
begin
  Result := VDCSMgr.SetXpt(AID, AHandle, AInput);
end;

// 0X40 Sense Queries
function DCSGetDeviceStatus(AID: Word; AHandle: TDeviceHandle; var AStatus: TDeviceStatus): Integer;
begin
  Result := VDCSMgr.GetDeviceStatus(AID, AHandle, AStatus);
end;

function DCSGetStorageTimeRemaining(AID: Word; AHandle: TDeviceHandle; var ATotalTC, AAvailableTC: TTimecode; AExtended: Boolean = False): Integer;
begin
  Result := VDCSMgr.GetStorageTimeRemaining(AID, AHandle, ATotalTC, AAvailableTC, AExtended);
end;

function DCSGetTC(AID: Word; AHandle: TDeviceHandle; var ACurTC: TTimecode): Integer;
begin
  Result := VDCSMgr.GetTC(AID, AHandle, ACurTC);
end;

function DCSGetRemainTC(AID: Word; AHandle: TDeviceHandle; var ARemainTC: TTimecode): Integer;
begin
  Result := VDCSMgr.GetRemainTC(AID, AHandle, ARemainTC);
end;

function DCSGetList(AID: Word; AHandle: TDeviceHandle; var AIDList: TIDList; var ARemainIDCount: Integer): Integer;
begin
  Result := VDCSMgr.GetList(AID, AHandle, AIDList, ARemainIDCount);
end;

function DCSGetNext(AID: Word; AHandle: TDeviceHandle; var AIDList: TIDList; var ARemainIDCount: Integer): Integer;
begin
  Result := VDCSMgr.GetNext(AID, AHandle, AIDList, ARemainIDCount);
end;

function DCSGetExist(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var AExist: Boolean): Integer;
begin
  Result := VDCSMgr.GetExist(AID, AHandle, AMediaID, AExist);
end;

function DCSGetSize(AID: Word; AHandle: TDeviceHandle; AMediaID: PChar; var ADuration: TTimecode): Integer;
begin
  Result := VDCSMgr.GetSize(AID, AHandle, AMediaID, ADuration);
end;

function DCSGetRoute(AID: Word; AHandle: TDeviceHandle; AOutput, AOutputLvl: Integer; var AInput, AInputLvl: Integer): Integer;
begin
  Result := VDCSMgr.GetRoute(AID, AHandle, AOutput, AOutputLvl, AInput, AInputLvl);
end;

function DCSGetXpt(AID: Word; AHandle: TDeviceHandle; var AInput: Integer): Integer; stdcall; // MCS only
begin
  Result := VDCSMgr.GetXpt(AID, AHandle, AInput);
end;

// 0X50 Event Control
function DCSInputEvent(AID: Word; AHandle: TDeviceHandle; AEvent: TEvent): Integer;
begin
  Result := VDCSMgr.InputEvent(AID, AHandle, AEvent);
end;

function DCSDeleteEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer;
begin
  Result := VDCSMgr.DeleteEvent(AID, AHandle, AEventID);
end;

function DCSClearEvent(AID: Word; AHandle: TDeviceHandle; AChannelID: Word): Integer;
begin
  Result := VDCSMgr.ClearEvent(AID, AHandle, AChannelID);
end;

function DCSTakeEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; AStartTime: TEventTime): Integer;
begin
  Result := VDCSMgr.TakeEvent(AID, AHandle, AEventID, AStartTime);
end;

function DCSHoldEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID): Integer;
begin
  Result := VDCSMgr.HoldEvent(AID, AHandle, AEventID);
end;

function DCSChangetDurationEvent(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; ADuration: TTimecode): Integer; stdcall;
begin
  Result := VDCSMgr.ChangeDurationEvent(AID, AHandle, AEventID, ADuration);
end;

function DCSOnAirCatchEvent(AID: Word; AHandle: TDeviceHandle): Integer; stdcall;
begin
  Result := VDCSMgr.OnAirCatchEvent(AID, AHandle);
end;

function DCSGetOnAirEventID(AID: Word; AHandle: TDeviceHandle; var AOnAirEventID, ANextEventID: TEventID): Integer; stdcall;
begin
  Result := VDCSMgr.GetOnAirEventID(AID, AHandle, AOnAirEventID, ANextEventID);
end;

function DCSGetEventInfo(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime; var ADurationTC: TTimecode): Integer; stdcall;
begin
  Result := VDCSMgr.GetEventInfo(AID, AHandle, AEventID, AStartTime, ADurationTC);
end;

function DCSGetEventStatus(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AEvenStatus: TEventStatus): Integer; stdcall;
begin
  Result := VDCSMgr.GetEventStatus(AID, AHandle, AEventID, AEvenStatus);
end;

function DCSGetEventStartTime(AID: Word; AHandle: TDeviceHandle; AEventID: TEventID; var AStartTime: TEventTime): Integer; stdcall;
begin
  Result := VDCSMgr.GetEventStartTime(AID, AHandle, AEventID, AStartTime);
end;

function DCSGetEventOverall(AID: Word; AHandle: TDeviceHandle; var AEventOverall: TEventOverall): Integer; stdcall;
begin
  Result := VDCSMgr.GetEventOverall(AID, AHandle, AEventOverall);
end;

end.
