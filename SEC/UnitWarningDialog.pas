unit UnitWarningDialog;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitSingleForm, Vcl.StdCtrls,
  Vcl.Imaging.pngimage, WMTools, WMControls, Vcl.ExtCtrls, System.Actions,
  Vcl.ActnList, UnitConsts;

type
  TWarningTimerThread = class;

  TfrmWarningDialog = class(TfrmSingle)
    lblWaraningText: TLabel;
    actWarningDialog: TActionList;
    actOK: TAction;
    wmibOK: TWMImageButton;
    mmWarning: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actOKExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    FTimerThread: TWarningTimerThread;
    FWarningBlink: Boolean;

    procedure Initialize;
    procedure Finalize;

    procedure WMUpdateWarningDisplay(var Message: TMessage); message WM_UPDATE_WARNING_DISPLAY;
    procedure WMCommand(var Msg: TWMCommand); message WM_COMMAND;
  protected
  public
    { Public declarations }
    procedure SetWarningText(AText: String);
  end;

  TWarningTimerThread = class(TThread)
  private
    FWarningDialog: TfrmWarningDialog;

    FCloseEvent: THandle;
  protected
    procedure Execute; override;
  public
    constructor Create(AWarningDialog: TfrmWarningDialog);
    destructor Destroy; override;

    procedure Terminate;
  end;

var
  frmWarningDialog: TfrmWarningDialog;

implementation

{$R *.dfm}

procedure SetMargins(AHandle: HWND);
var
  Rect: TRect;
begin
  SendMessage(AHandle, EM_GETRECT, 0, Longint(@Rect));
  Rect.Right := Rect.Right - GetSystemMetrics(SM_CXHSCROLL);
  SendMessage(AHandle, EM_SETRECT, 0, Longint(@Rect));
end;

procedure TfrmWarningDialog.WMCommand(var Msg: TWMCommand);
begin
  if (Msg.Ctl = mmWarning.Handle) and (Msg.NotifyCode = EN_UPDATE) then
  begin
    if (mmWarning.Lines.Count > 12) then   // maximum 12 lines
      mmWarning.ScrollBars := ssVertical
    else
    begin
      if mmWarning.ScrollBars <> ssNone then
      begin
        mmWarning.ScrollBars := ssNone;
        SetMargins(mmWarning.Handle);
      end;
    end;
  end;
  inherited;
end;

procedure TfrmWarningDialog.WMUpdateWarningDisplay(var Message: TMessage);
begin
  FWarningBlink := not FWarningBlink;

  if (FWarningBlink) then
  begin
    lblWaraningText.Font.Color := COLOR_TX_WARNING;
    mmWarning.Font.Color := COLOR_TX_WARNING;
  end
  else
  begin
    lblWaraningText.Font.Color := COLOR_TX_WARNING_NORMAL;
    mmWarning.Font.Color := COLOR_TX_WARNING_NORMAL;
  end;
end;

procedure TfrmWarningDialog.FormCreate(Sender: TObject);
begin
  inherited;

  Initialize;
end;

procedure TfrmWarningDialog.FormDestroy(Sender: TObject);
begin
  inherited;

  Finalize;
end;

procedure TfrmWarningDialog.FormShow(Sender: TObject);
begin
  inherited;
  if (FTimerThread <> nil) then
    FTimerThread.Start;
end;

procedure TfrmWarningDialog.Initialize;
begin
  lblWaraningText.Caption := '';

  mmWarning.ScrollBars := ssNone;
  mmWarning.Clear;
  SetMargins(mmWarning.Handle);

  FWarningBlink := False;

  FTimerThread := TWarningTimerThread.Create(Self);
end;

procedure TfrmWarningDialog.actOKExecute(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TfrmWarningDialog.Finalize;
begin
  if (FTimerThread <> nil) then
  begin
    FTimerThread.Terminate;
    FTimerThread.WaitFor;
    FreeAndNil(FTimerThread);
  end;
end;

procedure TfrmWarningDialog.SetWarningText(AText: String);
begin
  WMTitleBar.Caption := Format('Warning - %s', [FormatDateTime('dddddd tt', Now)]);
  lblWaraningText.Caption := AText;

  mmWarning.Text := AText;
  if (mmWarning.Lines.Count > 12) then   // maximum 12 lines
    mmWarning.ScrollBars := ssVertical
  else
  begin
    if (mmWarning.ScrollBars <> ssNone) then
    begin
      mmWarning.ScrollBars := ssNone;
      SetMargins(mmWarning.Handle);
    end;
  end;
end;

{ TWarningTimerThread }

constructor TWarningTimerThread.Create(AWarningDialog: TfrmWarningDialog);
begin
  FWarningDialog := AWarningDialog;

  FCloseEvent := CreateEvent(nil, True, False, nil);

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TWarningTimerThread.Destroy;
begin
  Terminate;

  CloseHandle(FCloseEvent);

  inherited Destroy;
end;

procedure TWarningTimerThread.Terminate;
begin
  inherited Terminate;

  SetEvent(FCloseEvent);
end;

procedure TWarningTimerThread.Execute;
var
  R: Integer;
  WaitList: array[0..1] of THandle;
begin
  WaitList[0] := FCloseEvent;
//  WaitList[1] := GV_TimerCancelEvent;
  WaitList[1] := GV_TimerExecuteEvent;
  while not Terminated do
  begin
    R := WaitForMultipleObjects(2, @WaitList, False, INFINITE);
    case R of
      WAIT_OBJECT_0: break;
      WAIT_OBJECT_0 + 1: PostMessage(FWarningDialog.Handle, WM_UPDATE_WARNING_DISPLAY, 0, 0);
    end;
  end;
end;

end.
