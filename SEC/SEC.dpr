program SEC;

uses
//  FastMM4,
//  FastMM4Messages,
  Vcl.Forms, Winapi.Windows,
  UnitCommons in '..\Common\UnitCommons.pas',
  UnitNormalForm in '..\Common\UnitNormalForm.pas' {frmNormal},
  UnitSingleForm in '..\Common\UnitSingleForm.pas' {frmSingle},
  UnitTypeConvert in '..\Common\UnitTypeConvert.pas',
  UnitWorkForm in '..\Common\UnitWorkForm.pas' {frmWork},
  UnitSEC in 'UnitSEC.pas' {frmSEC},
  UnitChannel in 'UnitChannel.pas' {frmChannel},
  UnitConsts in 'UnitConsts.pas',
  UnitInspectPlayList in 'UnitInspectPlayList.pas' {frmInspectPlayList},
  UnitDevice in 'UnitDevice.pas' {frmDevice},
  UnitEditEvent in 'UnitEditEvent.pas' {frmEditEvent},
  Vcl.Themes,
  Vcl.Styles,
  UnitSelectStartOnAir in 'UnitSelectStartOnAir.pas' {frmSelectStartOnAir},
  UnitBaseSerial in '..\Common\UnitBaseSerial.pas',
  UnitSelectPlayList in 'UnitSelectPlayList.pas' {frmSelectPlaylist},
  UnitDCSDLL in '..\Common\UnitDCSDLL.pas',
  UnitUDPIn in '..\UDP\UnitUDPIn.pas',
  UnitUDPOut in '..\UDP\UnitUDPOut.pas',
  UnitAllChannels in 'UnitAllChannels.pas' {frmAllChannels},
  UnitAllChannelsPanel in 'UnitAllChannelsPanel.pas' {frmAllChannelsPanel},
  UnitUDPCommons in '..\UDP\UnitUDPCommons.pas',
  UnitCheckStart in 'UnitCheckStart.pas' {frmCheckStart},
  UnitWarningDialog in 'UnitWarningDialog.pas' {frmWarningDialog},
  UnitAlarmThread in 'UnitAlarmThread.pas',
  UnitMCCDLL in '..\Common\UnitMCCDLL.pas',
  UnitSECDLL in '..\Common\UnitSECDLL.pas',
  UnitOpenPlayList in 'UnitOpenPlayList.pas' {frmOpenPlayList},
  UnitStartSplash in 'UnitStartSplash.pas' {frmStartSplash};

{$R *.res}

const
  MUTEX_NAME = 'SEC_V1.0';

var
  Handle: THandle;

begin
  System.ReportMemoryLeaksOnShutdown := True;

  CreateMutex(nil, True, MUTEX_NAME);

  if (GetLastError = ERROR_ALREADY_EXISTS) then
  begin
    Handle := FindWindow('TfrmSEC', nil);
    if (Handle <> 0) then
    begin
      SetForegroundWindow(Handle);
      ShowWindow(Handle, SW_RESTORE);
    end;

    Halt;
  end;

  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Schedule event controller';
  Application.CreateForm(TfrmSEC, frmSEC);
  Application.Run;
end.
