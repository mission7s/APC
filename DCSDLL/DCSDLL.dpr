library DCSDLL;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  System.SysUtils,
  System.Classes,
  Winapi.Windows,
  Vcl.Forms,
  DLLConsts in 'DLLConsts.pas',
  DCSMgr in 'DCSMgr.pas',
  UnitUDPIn in '..\UDP\UnitUDPIn.pas',
  UnitUDPOut in '..\UDP\UnitUDPOut.pas',
  UnitCommons in '..\Common\UnitCommons.pas',
  UnitTypeConvert in '..\Common\UnitTypeConvert.pas',
  UnitBaseSerial in '..\Common\UnitBaseSerial.pas',
  UnitUDPCommons in '..\UDP\UnitUDPCommons.pas';

{$R *.res}

var
  SaveExit: Pointer;

procedure NewExit; far;
begin
  ExitProc := SaveExit;
end {NewExit};

procedure NewDllProc(Reason: Integer);
var
  ALockList: TList;
  I: Integer;
begin
  case Reason of
    DLL_PROCESS_ATTACH:
      begin
        VDCSMgr := TDCSMgr.Create(nil);
      end;
    DLL_PROCESS_DETACH:
      begin
        FreeAndNil(VDCSMgr);
      end;
  end;
end;

begin
  { Library initialization code }
  SaveExit := ExitProc;  { Save exit procedure chain }
  ExitProc := @NewExit;  { Install NewExit exit procedure }
  DllProc := @NewDllProc;
  NewDllProc(DLL_PROCESS_ATTACH);
end.
