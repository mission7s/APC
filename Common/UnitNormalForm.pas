unit UnitNormalForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, VCl.Forms, Winapi.MultiMon,
  Vcl.Dialogs, WMTools, WMControls, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Imaging.pngimage;

type
  TfrmNormal = class(TForm)
    WMPanel: TWMPanel;
    WMTitleBar: TWMTitleBar;
    WMIBClose: TWMImageSpeedButton;
    pnlDesc: TWMPanel;
    WMIBMinimize: TWMImageSpeedButton;
    WMIBMaximize: TWMImageSpeedButton;
    WMIBRestore: TWMImageSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure WMIBCloseClick(Sender: TObject);
    procedure WMTitleBarSystemMenuClick(Sender: TObject;
      DoubleClicked: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure WMTitleBarMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure WMIBRestoreClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure WMIBMaximizeClick(Sender: TObject);
    procedure WMIBMinimizeClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    FPanelInstance: Pointer;
    FDefPanelProc: Pointer;
    FPanelHandle: HWND;

    FX, FY: Integer;
    FW, FH: Integer;

    FFirstMove: Boolean;

    function IsMaximized: Boolean;
    procedure SetWindowBounds(L, T, W, H: Integer);

    procedure PanelWndProc(var Message: TMessage);
    procedure WMNCHitTest(var Message: TWMNCHitTest); message WM_NCHITTEST;
  protected
    procedure Maximize;
    procedure Minimize;
    procedure Restore;
  public
    { Public declarations }
  end;

var
  frmNormal: TfrmNormal;

implementation

{$R *.dfm}

procedure TfrmNormal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmNormal.FormCreate(Sender: TObject);
begin
  FPanelInstance := MakeObjectInstance(PanelWndProc);

  FPanelHandle := WMPanel.Handle;
  FDefPanelProc := Pointer(GetWindowLong(FPanelHandle, GWL_WNDPROC));
  SetWindowLong(FPanelHandle, GWL_WNDPROC, LongInt(FPanelInstance));

  WMIBRestore.Visible  := IsMaximized;
  WMIBMaximize.Visible := not IsMaximized;

  FFirstMove := True;
end;

procedure TfrmNormal.FormDestroy(Sender: TObject);
begin
  if FPanelHandle <> 0 then
    SetWindowLong(FPanelHandle, GWL_WNDPROC, LongInt(FDefPanelProc));
  FreeObjectInstance(FPanelInstance);
end;

procedure TfrmNormal.FormResize(Sender: TObject);
begin
  WMIBRestore.Visible  := IsMaximized;
  WMIBMaximize.Visible := not IsMaximized;
end;

procedure TfrmNormal.FormShow(Sender: TObject);
begin
  WMPanel.SetFocus;
end;

function TfrmNormal.IsMaximized: Boolean;
var
  Monitor: TMonitor;
  L, T, W, H: Integer;
begin
  Monitor := Screen.MonitorFromWindow(Handle);
  with Monitor.WorkareaRect do
  begin
    L := Left;
    T := Top;
    W := Right - Left;
    H := Bottom - Top;
  end;

  Result := (Left <= L) and (Top <= T) and (Width >= W) and (Height >= H);
end;

procedure TfrmNormal.SetWindowBounds(L, T, W, H: Integer);
begin
  FX  := Left;
  FY  := Top;
  FW  := Width;
  FH  := Height;

  SetBounds(L, T, W, H);
end;

procedure TfrmNormal.Maximize;
var
  Monitor: TMonitor;
  W, H: Integer;
begin
  Monitor := Screen.MonitorFromWindow(Handle);
  with Monitor.WorkareaRect do
  begin
    W := Right - Left;
    H := Bottom - Top;
    SetWindowBounds(Left, Top, W, H);
  end;

  WMTitleBar.MoveEnabled := False;
end;

procedure TfrmNormal.Minimize;
begin
  Application.Minimize;
end;

procedure TfrmNormal.Restore;
begin
  SetWindowBounds(FX, FY, FW, FH);
  WMTitleBar.MoveEnabled := True;
end;

procedure TfrmNormal.WMIBCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmNormal.WMIBMaximizeClick(Sender: TObject);
begin
  Maximize;
end;

procedure TfrmNormal.WMIBMinimizeClick(Sender: TObject);
begin
  Minimize;
end;

procedure TfrmNormal.WMIBRestoreClick(Sender: TObject);
begin
  Restore;
end;

procedure TfrmNormal.WMTitleBarMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if (Button = mbLeft) and (ssDouble in Shift) then
  begin
    if (not IsMaximized) then Maximize
    else Restore;
  end;
end;

procedure TfrmNormal.WMTitleBarSystemMenuClick(Sender: TObject;
  DoubleClicked: Boolean);
begin
  if (DoubleClicked) then Close;
end;

procedure TfrmNormal.PanelWndProc(var Message: TMessage);
begin
  with Message do
  begin
    case Msg of
      WM_NCHITTEST:
        with TWMNCHitTest(Message) do
          begin
            Result := Winapi.Windows.HTTRANSPARENT;
            Exit;
          end;
    end;
    Result := CallWindowProc(FDefPanelProc, FPanelHandle, Msg, WParam, LParam);
  end;
end;

procedure TfrmNormal.WMNCHitTest(var Message: TWMNCHitTest);
const
  BorderSize = 8;
var
  P: TPoint;
begin
  inherited;
  P := Point(Message.XPos, Message.YPos);
  P := ScreenToClient(P);
  if (P.X < BorderSize) and (P.Y < BorderSize) then
    Message.Result := htTopLeft
  else if (P.X > Width - BorderSize) and (P.Y < BorderSize) then
    Message.Result := htTopRight
  else if (P.X > Width - BorderSize) and (P.Y > Height - BorderSize) then
    Message.Result := htBottomRight
  else if (P.X < BorderSize) and (P.Y > Height - BorderSize) then
    Message.Result := htBottomLeft
  else if (P.X < BorderSize) then
    Message.Result := htLeft
  else if (P.Y < BorderSize) then
    Message.Result := htTop
  else if (P.X > Width - BorderSize) then
    Message.Result := htRight
  else if (P.Y > Height - BorderSize) then
    Message.Result := htBottom;
end;

end.
