unit UnitCheckStart;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitSingleForm, WMTools, WMControls,
  Vcl.ExtCtrls, Vcl.Imaging.pngimage, Vcl.ComCtrls, Vcl.StdCtrls,
  System.Actions, Vcl.ActnList,
  UnitCommons, UnitConsts, UnitDCSDLL, UnitMCCDLL, UnitSECDLL;

type
  TDeviceCheckThread = class;

  TfrmCheckStart = class(TfrmSingle)
    lblChecking: TLabel;
    lblDeviceName: TLabel;
    pBarStart: TProgressBar;
    wmibCancel: TWMImageButton;
    aLstStart: TActionList;
    actCancel: TAction;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
  private
    { Private declarations }
    FIsCheckCancel: Boolean;
    FIsCheckAll: Boolean;
    FDeviceCheckThread: TDeviceCheckThread;
  public
    { Public declarations }
  end;

  TDeviceCheckThread = class(TThread)
  private
    { Private declarations }
    FForm: TfrmCheckStart;
    procedure DoControl;
  protected
    procedure Execute; override;
  public
    constructor Create(AForm: TfrmCheckStart);
    destructor Destroy; override;
  end;

var
  frmCheckStart: TfrmCheckStart;

implementation

{$R *.dfm}

procedure TfrmCheckStart.actCancelExecute(Sender: TObject);
begin
  inherited;
  FIsCheckCancel := True;
end;

procedure TfrmCheckStart.FormCreate(Sender: TObject);
begin
  inherited;

  lblChecking.Caption := SSECStart;
  lblDeviceName.Caption := '';

  FIsCheckCancel := False;
  FIsCheckAll := False;

  FDeviceCheckThread := TDeviceCheckThread.Create(Self);
end;

procedure TfrmCheckStart.FormDestroy(Sender: TObject);
begin
  inherited;
  if (FDeviceCheckThread <> nil) then
  begin
    FDeviceCheckThread.Terminate;
    FDeviceCheckThread.WaitFor;
    FreeAndNil(FDeviceCheckThread);
  end;
end;

procedure TfrmCheckStart.FormShow(Sender: TObject);
begin
  inherited;

  if (FDeviceCheckThread <> nil) then
    FDeviceCheckThread.Start;
end;

{ TDeviceCheckThread }

constructor TDeviceCheckThread.Create(AForm: TfrmCheckStart);
begin
  FForm := AForm;

  FreeOnTerminate := False;

  inherited Create(True);
end;

destructor TDeviceCheckThread.Destroy;
begin
  inherited Destroy;
end;

procedure TDeviceCheckThread.Execute;
begin
  { Place thread code here }
  Synchronize(DoControl);
end;

procedure TDeviceCheckThread.DoControl;
var
  ErrorString: String;

  MCCName: String;
  MCCCount: Integer;

  SECName: String;
  SECCount: Integer;

  DeviceName: String;
  DCSName: String;
  HandleCount: Integer;

  I, J: Integer;
  R: Integer;
  IsAlive: Boolean;
  SourceHandle: PSourceHandle;
  DeviceHandle: TDeviceHandle;
begin
  with FForm do
  begin
    try
      FIsCheckCancel := False;
      FIsCheckAll := False;

{      if (GV_SettingMCC.Use) then
        MCCCount := GV_MCCList.Count
      else
        MCCCount := 0;

      SECCount := GV_SECList.Count - 1; }

      HandleCount := 0;
      for I := 0 to GV_SourceList.Count - 1 do
      begin
        if (GV_SourceList[I]^.Handles <> nil) then
        begin
          for J := 0 to GV_SourceList[I]^.Handles.Count - 1 do
            Inc(HandleCount);
        end;
      end;

      pBarStart.Min := 0;
      pBarStart.Max := {MCCCount + SECCount + }HandleCount;
      pBarStart.Position := 0;
      Application.ProcessMessages;

{      if (GV_SettingMCC.Use) then
      begin
        // MCC check
        lblChecking.Caption := SLookingForMCC;

        for I := 0 to GV_MCCList.Count - 1 do
        begin
          lblDeviceName.Caption := Format('%s', [GV_MCCList[I]^.Name]);
          pBarStart.Position := pBarStart.Position + 1;
          Application.ProcessMessages;

          R := MCCOpen(GV_MCCList[I]^.ID, GV_MCCList[I]^.HostIP, GV_MCCList[I]^.Name);
          if (R = D_OK) then
          begin
            Assert(False, GetMainLogStr(lsNormal, @LS_MCCOpenSuccess, [GV_MCCList[I]^.ID, String(GV_MCCList[I]^.Name)]));
            GV_MCCList[I]^.Alive := True;
          end
          else
          begin
            Assert(False, GetMainLogStr(lsError, @LSE_MCCOpenFailed, [R, GV_MCCList[I]^.ID, String(GV_MCCList[I]^.Name)]));
            GV_MCCList[I]^.Alive := False;

  //              ShowMessage(IntToStr(R));
            if (not FIsCheckAll) then
            begin
              ErrorString := Format(SNotFoundComponentAndContinue, [GV_MCCList[I]^.Name]);
              MessageBeep(MB_ICONWARNING);
              R := MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_YESNO or MB_ICONWARNING or MB_TOPMOST);
              if (R = ID_NO) then
              begin
                FIsCheckCancel := True;
                break;
              end
              else
              begin
                ErrorString := SLookingForAllComponentAndDevice;
                MessageBeep(MB_ICONQUESTION);
                R := MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_YESNO or MB_ICONQUESTION or MB_TOPMOST);
                if (R = ID_YES) then
                begin
                  FIsCheckAll := True;
                end
              end;
            end;
          end;
        end;
      end;

      if (FIsCheckCancel) then exit; }

{      if (GV_SECList.Count > 1) then
      begin
        // SEC check
        lblChecking.Caption := SLookingForSEC;

        for I := 0 to GV_SECList.Count - 1 do
        begin
          if (GV_SECMine = GV_SECList[I]) then continue;

          lblDeviceName.Caption := Format('%s', [GV_SECList[I]^.Name]);
          pBarStart.Position := pBarStart.Position + 1;
          Application.ProcessMessages;

          R := SECOpen(GV_SECList[I]^.ID, GV_SECList[I]^.HostIP, GV_SECList[I]^.Name);
          if (R = D_OK) then
          begin
            Assert(False, GetMainLogStr(lsNormal, @LS_SECOpenSuccess, [GV_SECList[I]^.ID, String(GV_SECList[I]^.Name)]));
            GV_SECList[I]^.Opened := True;
          end
          else
          begin
            Assert(False, GetMainLogStr(lsError, @LSE_SECOpenFailed, [R, GV_SECList[I]^.ID, String(GV_SECList[I]^.Name)]));
            GV_SECList[I]^.Opened := False;

  //              ShowMessage(IntToStr(R));
            if (not FIsCheckAll) then
            begin
              ErrorString := Format(SNotFoundComponentAndContinue, [GV_SECList[I]^.Name]);
              MessageBeep(MB_ICONWARNING);
              R := MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_YESNO or MB_ICONWARNING or MB_TOPMOST);
              if (R = ID_NO) then
              begin
                FIsCheckCancel := True;
                break;
              end
              else
              begin
                ErrorString := SLookingForAllComponentAndDevice;
                MessageBeep(MB_ICONQUESTION);
                R := MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_YESNO or MB_ICONQUESTION or MB_TOPMOST);
                if (R = ID_YES) then
                begin
                  FIsCheckAll := True;
                end
              end;
            end;
          end;
        end;
      end;  }

      // DCS check
      lblChecking.Caption := SLookingForDevice;
      for I := 0 to GV_DCSList.Count - 1 do
      begin
        DCSName := String(GV_DCSList[I]^.Name);
        lblDeviceName.Caption := Format('%s', [DCSName]);

        R := DCSSysIsAlive(GV_DCSList[I]^.HostIP, IsAlive);
        if (R = D_OK) then
        begin
          GV_DCSList[I]^.Alive := IsAlive;
          Assert(False, GetMainLogStr(lsNormal, @LS_DCSAliveCheckSuccess, [GV_DCSList[I]^.ID, String(GV_DCSList[I]^.Name), BoolToStr(IsAlive, True)]));
        end
        else
        begin
          GV_DCSList[I]^.Alive := False;
          Assert(False, GetMainLogStr(lsError, @LSE_DCSAliveCheckFailed, [R, GV_DCSList[I]^.ID, String(GV_DCSList[I]^.Name)]));
        end;
      end;

      // Device check
      lblChecking.Caption := SLookingForDevice;
      for I := 0 to GV_SourceList.Count - 1 do
      begin
        DeviceName := GV_SourceList[I]^.Name;
        if (GV_SourceList[I]^.Handles <> nil) then
        begin
          for J := 0 to GV_SourceList[I]^.Handles.Count - 1 do
          begin
            SourceHandle := GV_SourceList[I]^.Handles[J];
            if (SourceHandle^.DCS = nil) then continue;

//            DCSName := GetDCSNameByID(SourceHandle^.DCSID);
            DCSName := String(SourceHandle^.DCS^.Name);
            lblDeviceName.Caption := Format('%s, %s', [DeviceName, DCSName]);
            pBarStart.Position := pBarStart.Position + 1;
            Application.ProcessMessages;

            if (not FIsCheckCancel) then
            begin
        //    ShowMessage(DCS^.HostIP);
        //    ShowMessage(GV_SourceList[I]^.Name);
              R := DCSOpen(SourceHandle^.DCS^.ID, SourceHandle^.DCS^.HostIP, GV_SourceList[I]^.Name, DeviceHandle);
              if (R = D_OK) then
              begin
                SourceHandle^.Handle := DeviceHandle;
//                GV_SourceList[I]^.CommSuccess := True;
      //          ShowMessage(Format('Success ID=%d, IP=%s, DeviceName=%s, DeviceHandle=%d', [SourceHandle^.DCSID, SourceHandle^.DCSIP, GV_SourceList[I]^.Name, DeviceHandle]));

                Assert(False, GetMainLogStr(lsNormal, @LS_DCSOpenDeviceSuccess, [SourceHandle^.DCS^.ID, String(GV_SourceList[I]^.Name)]));
              end
              else
              begin
                SourceHandle^.Handle := DeviceHandle;

                Assert(False, GetMainLogStr(lsError, @LSE_DCSOpenDeviceFailed, [R, SourceHandle^.DCS^.ID, String(GV_SourceList[I]^.Name)]));

  //              ShowMessage(IntToStr(R));
                if (not FIsCheckAll) then
                begin
                  ErrorString := Format(SNotFoundDeviceAndContinue, [DeviceName, DCSName]);
                  MessageBeep(MB_ICONWARNING);
                  R := MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_YESNO or MB_ICONWARNING or MB_TOPMOST);
                  if (R = ID_NO) then
                  begin
                    FIsCheckCancel := True;
                  end
                  else
                  begin
                    ErrorString := SLookingForAllDevice;
                    MessageBeep(MB_ICONQUESTION);
                    R := MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_YESNO or MB_ICONQUESTION or MB_TOPMOST);
                    if (R = ID_YES) then
                    begin
                      FIsCheckAll := True;
                    end
                  end;
                end;
              end;
            end
            else
              SourceHandle^.Handle := INVALID_DEVICE_HANDLE;
          end;
        end;
      end;
    finally
      Close;
    end;
  end;
end;

end.
