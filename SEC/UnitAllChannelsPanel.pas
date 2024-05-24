unit UnitAllChannelsPanel;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, WMControls,
  UnitCommons, UnitConsts, Vcl.Imaging.pngimage;

type
  TTimelineTimerThread = class;

  TfrmAllChannelsPanel = class(TForm)
    pnlChannel: TWMPanel;
    lblChannelName: TLabel;
    pnlPlayedTime: TWMPanel;
    lblPlayedTime: TLabel;
    Label2: TLabel;
    pnlRemainingTime: TWMPanel;
    lblRemainingTime: TLabel;
    Label3: TLabel;
    pnlNextStart: TWMPanel;
    lblNextStart: TLabel;
    Label5: TLabel;
    pnlNextDuration: TWMPanel;
    lblNextDuration: TLabel;
    Label7: TLabel;
    imgOnair: TImage;
    imgOffair: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure pnlChannelDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    FChannelID: Word;

    FTimerThread: TTimelineTimerThread;

    procedure Initialize;
    procedure Finalize;

    procedure GotoPlaylist(AChannelID: Word);
  protected
//    procedure WndProc(var Message: TMessage); override;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent; AChannelID: Word); overload;

    procedure SetChannelOnAir(AOnAir: Boolean);
    procedure SetChannelTime(APlayedTime, ARemainingTime, ANextStart, ANextDuration: String);
  end;


  TTimelineTimerThread = class(TThread)
  private
    FChannelPanelForm: TfrmAllChannelsPanel;

    FCloseEvent: THandle;
  protected
    procedure Execute; override;
  public
    constructor Create(AChannelPanelForm: TfrmAllChannelsPanel);
    destructor Destroy; override;

    procedure Terminate;
  end;

var
  frmAllChannelsPanel: TfrmAllChannelsPanel;

implementation

uses UnitSEC, UnitAllChannels, System.DateUtils, System.Math;

{$R *.dfm}

{procedure TfrmAllChannelsPanel.WndProc(var Message: TMessage);
begin
  inherited;
  case Message.Msg of
    WM_UPDATE_CHANNEL_TIME:
    begin
    end;

    WM_UPDATE_CURR_EVENT:
    begin
    end;

    WM_UPDATE_NEXT_EVENT:
    begin
    end;

  end;
end; }

constructor TfrmAllChannelsPanel.Create(AOwner: TComponent; AChannelID: Word);
begin
  inherited Create(AOwner);

  FChannelID := AChannelID;
end;

procedure TfrmAllChannelsPanel.Initialize;
var
  Channel: PChannel;
begin
  Channel := GetChannelByID(FChannelID);
  if (Channel <> nil) then
  begin
    lblChannelName.Caption := String(Channel^.Name);
  end;

  lblPlayedTime.Caption    := IDLE_TIME;
  lblRemainingTime.Caption := IDLE_TIME;
  lblNextStart.Caption     := IDLE_TIMECODE;
  lblNextDuration.Caption  := IDLE_TIMECODE;

  SetChannelOnAir(False);

  FTimerThread := TTimelineTimerThread.Create(Self);
end;

procedure TfrmAllChannelsPanel.pnlChannelDblClick(Sender: TObject);
begin
  GotoPlaylist(FChannelID);
end;

procedure TfrmAllChannelsPanel.Finalize;
begin
  if (FTimerThread <> nil) then
  begin
    FTimerThread.Terminate;
    FTimerThread.WaitFor;
    FreeAndNil(FTimerThread);
  end;
end;

procedure TfrmAllChannelsPanel.GotoPlaylist(AChannelID: Word);
var
  I: Integer;
begin
  with frmSEC.aoPagerMain do
  begin
    for I := 1 to AdvPageCount - 1 do
    begin
      if (AdvPages[I].Tag = AChannelID) then
      begin
        ActivePageIndex := I;
        break;
      end;
    end;
  end;
end;

procedure TfrmAllChannelsPanel.FormCreate(Sender: TObject);
begin
  inherited;

  Initialize;
end;

procedure TfrmAllChannelsPanel.FormDestroy(Sender: TObject);
begin
  Finalize;

  inherited;
end;

procedure TfrmAllChannelsPanel.FormShow(Sender: TObject);
begin
  inherited;

  if (FTimerThread <> nil) then
    FTimerThread.Start;
end;

procedure TfrmAllChannelsPanel.SetChannelOnAir(AOnAir: Boolean);
begin
  if (AOnAir) then
  begin
    pnlChannel.Color          := COLOR_BK_ALL_CHANNEL_ONAIR;
    pnlChannel.ColorHighLight := COLOR_BK_ALL_CHANNEL_ONAIR;
    pnlChannel.ColorShadow    := COLOR_BK_ALL_CHANNEL_ONAIR;

    imgOnair.Visible  := True;
    imgOffair.Visible := False;

    lblChannelName.Font.Color := COLOR_TX_ALL_CHANNEL_ONAIR;
  end
  else
  begin
    pnlChannel.Color          := COLOR_BK_ALL_CHANNEL_OFFAIR;
    pnlChannel.ColorHighLight := COLOR_BK_ALL_CHANNEL_OFFAIR;
    pnlChannel.ColorShadow    := COLOR_BK_ALL_CHANNEL_OFFAIR;

    imgOffair.Visible := True;
    imgOnair.Visible  := False;

    lblChannelName.Font.Color := COLOR_TX_ALL_CHANNEL_OFFAIR;
  end;
end;

procedure TfrmAllChannelsPanel.SetChannelTime(APlayedTime, ARemainingTime, ANextStart, ANextDuration: String);
begin
  lblPlayedTime.Caption    := APlayedTime;
  lblRemainingTime.Caption := ARemainingTime;
  lblNextStart.Caption     := ANextStart;
  lblNextDuration.Caption  := ANextDuration;
end;

{ TTimelineTimerThread }

constructor TTimelineTimerThread.Create(AChannelPanelForm: TfrmAllChannelsPanel);
begin
  FChannelPanelForm := AChannelPanelForm;

  FCloseEvent := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TTimelineTimerThread.Destroy;
begin
  Terminate;

  CloseHandle(FCloseEvent);

  inherited Destroy;
end;

procedure TTimelineTimerThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FCloseEvent);
end;

procedure TTimelineTimerThread.Execute;
var
  WaitList: array[0..1] of THandle;
begin
  WaitList[0] := GV_TimerExecuteEvent;
//  WaitList[1] := GV_TimerCancelEvent;
  WaitList[1] := FCloseEvent;
  while not Terminated do
  begin
    if (WaitForMultipleObjects(2, @WaitList, False, INFINITE) <> WAIT_OBJECT_0) then
      break; // Terminate thread when GV_TimerCancelEvent is signaled

    PostMessage(FChannelPanelForm.Handle, WM_UPDATE_CHANNEL_TIME, 0, 0);
  end;
end;

end.
