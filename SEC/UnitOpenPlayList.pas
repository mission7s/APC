unit UnitOpenPlayList;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitSingleForm, WMTools, WMControls,
  Vcl.ExtCtrls, Vcl.Imaging.pngimage, Vcl.ComCtrls, Vcl.StdCtrls,
  System.Actions, Vcl.ActnList,
  UnitCommons, UnitConsts, UnitDCSDLL, UnitMCCDLL;

type
  TOpenPlaylistThread = class;

  TfrmOpenPlayList = class(TfrmSingle)
    lblPlaylistName: TLabel;
    lblWait: TLabel;
    pBarStart: TProgressBar;
    wmibCancel: TWMImageButton;
    aLstStart: TActionList;
    actCancel: TAction;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure lblPlaylistNameClick(Sender: TObject);
  private
    { Private declarations }
    FIsCheckCancel: Boolean;
    FOpenPlaylistThread: TOpenPlaylistThread;
  public
    { Public declarations }
  end;

  TOpenPlaylistThread = class(TThread)
  private
    { Private declarations }
    FForm: TfrmOpenPlaylist;
    procedure DoControl;
  protected
    procedure Execute; override;
  public
    constructor Create(AForm: TfrmOpenPlaylist);
    destructor Destroy; override;
  end;

var
  frmOpenPlayList: TfrmOpenPlayList;

implementation

{$R *.dfm}

procedure TfrmOpenPlayList.FormCreate(Sender: TObject);
begin
  inherited;

  lblPlaylistName.Caption := '';

  FIsCheckCancel := False;

  FOpenPlaylistThread := TOpenPlaylistThread.Create(Self);
end;

procedure TfrmOpenPlayList.FormDestroy(Sender: TObject);
begin
  inherited;
  if (FOpenPlaylistThread <> nil) then
  begin
    FOpenPlaylistThread.Terminate;
    FOpenPlaylistThread.WaitFor;
    FreeAndNil(FOpenPlaylistThread);
  end;
end;

procedure TfrmOpenPlayList.FormShow(Sender: TObject);
begin
  inherited;

  if (FOpenPlaylistThread <> nil) then
    FOpenPlaylistThread.Start;
end;

procedure TfrmOpenPlayList.lblPlaylistNameClick(Sender: TObject);
begin
  inherited;

end;

{ TOpenPlaylistThread }

constructor TOpenPlaylistThread.Create(AForm: TfrmOpenPlaylist);
begin
  FForm := AForm;

  FreeOnTerminate := False;

  inherited Create(True);
end;

destructor TOpenPlaylistThread.Destroy;
begin
  inherited Destroy;
end;

procedure TOpenPlaylistThread.Execute;
begin
  { Place thread code here }
  Synchronize(DoControl);
end;

procedure TOpenPlaylistThread.DoControl;
var
  ErrorString: String;

  MCCName: String;
  MCCCount: Integer;

  DeviceName: String;
  DCSName: String;
  HandleCount: Integer;

  I, J: Integer;
  R: Integer;
  IsMain: Boolean;
  SourceHandle: PSourceHandle;
  DeviceHandle: TDeviceHandle;
begin
{  with FForm do
  begin
    try
      FIsCheckCancel := False;

      if (GV_SettingMCC.Use) then
        MCCCount := GV_MCCList.Count
      else
        MCCCount := 0;

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
      pBarStart.Max := MCCCount + HandleCount;
      pBarStart.Position := 0;
      Application.ProcessMessages;

      if (GV_SettingMCC.Use) then
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
            GV_MCCList[I]^.Opened := True;
          end
          else
          begin
            Assert(False, GetMainLogStr(lsError, @LSE_MCCOpenFailed, [R, GV_MCCList[I]^.ID, String(GV_MCCList[I]^.Name)]));
            GV_MCCList[I]^.Opened := False;

  //              ShowMessage(IntToStr(R));
            ErrorString := Format(SNotFoundMCCAndContinue, [GV_MCCList[I]^.Name]);
            MessageBeep(MB_ICONWARNING);
            R := MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_YESNO or MB_ICONWARNING);
            if (R = ID_NO) then
            begin
              FIsCheckCancel := True;
              break;
            end;
          end;
        end;
      end;

      if (FIsCheckCancel) then exit;

      // DCS check
  {    for I := 0 to GV_DCSList.Count - 1 do
      begin
        R := DCSIsMain(GV_DCSList[I]^.ID, GV_DCSList[I]^.HostIP, IsMain);
        if (R = D_OK) then
        begin
          GV_DCSList[I]^.Main := IsMain;
        end
        else
          GV_DCSList[I]^.Main := False;
      end;  }

      // Device check
{      lblChecking.Caption := SLookingForDevice;
      for I := 0 to GV_SourceList.Count - 1 do
      begin
        DeviceName := GV_SourceList[I]^.Name;
        if (GV_SourceList[I]^.Handles <> nil) then
        begin
          for J := 0 to GV_SourceList[I]^.Handles.Count - 1 do
          begin
            SourceHandle := GV_SourceList[I]^.Handles[J];
            DCSName := GetDCSNameByID(SourceHandle^.DCSID);
            lblDeviceName.Caption := Format('%s, %s', [DeviceName, DCSName]);
            pBarStart.Position := pBarStart.Position + 1;
            Application.ProcessMessages;

            if (not FIsCheckCancel) then
            begin
        //    ShowMessage(DCS^.HostIP);
        //    ShowMessage(GV_SourceList[I]^.Name);
              R := DCSOpen(SourceHandle^.DCSID, SourceHandle^.DCSIP, GV_SourceList[I]^.Name, DeviceHandle);
              if (R = D_OK) then
              begin
                SourceHandle^.Handle := DeviceHandle;
      //          ShowMessage(Format('Success ID=%d, IP=%s, DeviceName=%s, DeviceHandle=%d', [SourceHandle^.DCSID, SourceHandle^.DCSIP, GV_SourceList[I]^.Name, DeviceHandle]));

                Assert(False, GetMainLogStr(lsNormal, @LS_DCSOpenDeviceSuccess, [SourceHandle^.DCSID, String(GV_SourceList[I]^.Name)]));
              end
              else
              begin
                SourceHandle^.Handle := DeviceHandle;

                Assert(False, GetMainLogStr(lsError, @LSE_DCSOpenDeviceFailed, [R, SourceHandle^.DCSID, String(GV_SourceList[I]^.Name)]));

  //              ShowMessage(IntToStr(R));
                ErrorString := Format(SNotFoundDeviceAndContinue, [DeviceName, DCSName]);
                MessageBeep(MB_ICONWARNING);
                R := MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_YESNO or MB_ICONWARNING);
                if (R = ID_NO) then
                begin
                  FIsCheckCancel := True;
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
  end; }
end;

end.
