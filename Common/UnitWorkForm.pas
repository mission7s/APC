unit UnitWorkForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, WMControls, ExtCtrls, MultiMon, Math, Vcl.Themes, Vcl.StdCtrls,
  WMTools, Vcl.Menus, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnPopup,
  Vcl.Imaging.pngimage;

type
  TfrmWork = class(TForm)
    WMPanel: TWMPanel;
    WMTitleBar: TWMTitleBar;
    pnlDesc: TWMPanel;
    WMIBClose: TWMImageSpeedButton;
    WMIBMinimize: TWMImageSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure WMIBCloseClick(Sender: TObject);
    procedure FormCanResize(Sender: TObject; var NewWidth,
      NewHeight: Integer; var Resize: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure WMTitleBarSystemMenuClick(Sender: TObject;
      DoubleClicked: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure WMIBMinimizeClick(Sender: TObject);
  private
    { Private declarations }
    FScreenWorkAreaWidth: Integer;
    FScreenWorkAreaHeight: Integer;
    FChangeLeft: Integer;
    FChangeTop: Integer;
    FChangeWidth: Integer;
    FChangeHeight: Integer;
    FScreenWorkAreaChanged: Boolean;

    FSaveWidth: Integer;
    FSaveHeight: Integer;

    FCombine: Boolean;
    FCombineLeft: Integer;
    FCombineWidth: Integer;
    FCombineHeight: Integer;

    FSnapped: Boolean;
    FThresh: Integer;
    FSizing: Boolean;
    FSizingLeft: Boolean;
    FSizingRight: Boolean;
    FSizingTop: Boolean;
    FSizingBottom: Boolean;
    FSizingDockLeft: Boolean;
    FSizingDockRight: Boolean;
    FSizingDockTop: Boolean;
    FSizingDockBottom: Boolean;
    FMoving: Boolean;
    FClosing: Boolean;

    FSizingLeftMin: Integer;

    FPanelInstance: Pointer;
    FDefPanelProc: Pointer;
    FPanelHandle: HWND;

    function GetMonitorWorkArea(AMonHandle: HMONITOR): TRect;

    procedure SetCombine(Value: Boolean);
    procedure PanelWndProc(var Message: TMessage);
    procedure WMActivate(var Message: TWMActivate); message WM_ACTIVATE;
    procedure WMNCHitTest(var Message: TWMNCHitTest); message WM_NCHITTEST;
    procedure WMNCLButtonDown(var Msg: TWMNCLButtonDown); message WM_NCLBUTTONDOWN;
    procedure WMEnterSizeMove(var Msg: TMessage); message WM_ENTERSIZEMOVE;
    procedure WMExitSizeMove(var Msg: TMessage); message WM_EXITSIZEMOVE;
    procedure WMMoving(var Msg: TWMMoving); message WM_MOVING;
    procedure WMMove(var Msg: TWMMove); message WM_MOVE;
    procedure WMSizing(var Msg: TMessage); message WM_SIZING;
    procedure WMSize(var Msg: TWMSize); message WM_SIZE;
    procedure WMWindowPosChanging(var Msg: TWMWindowPosChanging); message WM_WINDOWPOSCHANGING;
    procedure WMWindowPosChanged(var Msg: TWMWindowPosChanged); message WM_WINDOWPOSCHANGED;
    procedure WMDisplayChange(var Message: TWMDisplayChange); message WM_DISPLAYCHANGE;

    procedure ScaleCtrls(Control: TWinControl; WW, PW, HH, PH: Integer);
    procedure ChangeDims(Control: TControl; WW, PW, HH, PH: Integer);
  protected
    procedure CreateParams(var Params: TCreateParams); override;
    procedure CombineForm(AToWinControl: TWinControl);
    procedure SplitForm(AFormStyle: TFormStyle = fsNormal; ABorderStyle: TFormBorderStyle = bsSizeable);
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); overload;
    constructor Create(AOwner: TComponent; ACombine: Boolean; ALeft, ATop, AWidth, AHeight: Integer); overload;
    destructor Destroy; override;

    property Closing: Boolean read FClosing;
    property Combine: Boolean read FCombine write SetCombine;
    property CombineLeft: Integer read FCombineLeft write FCombineLeft;
    property CombineWidth: Integer read FCombineWidth write FCombineWidth;
    property CombineHeight: Integer read FCombineHeight write FCombineHeight;
  end;

var
  frmWork: TfrmWork;

implementation

{$R *.dfm}

procedure TfrmWork.CombineForm(AToWinControl: TWinControl);
var
  P: TPoint;
begin
  FCombine := True;

  if not WMTitleBar.Visible then
  begin
    Show;
    exit;
  end;

  P.X := Left;
  P.Y := Top;
//  Parent := AToWinControl;
//  P := ClientToScreen(P);
//  P := ScreenToClient(P);
  if P.X < 1 then P.X := 1;
  if P.Y < 1 then P.Y := 1;


  Hide;
  try
    WMTitleBar.Visible := False;
    SetBounds(P.X, P.Y{ + WMTitleBar.Height}, Width, Height - WMTitleBar.Height);
  finally
    Show;
  end;
end;

procedure TfrmWork.SplitForm;
var
  P: TPoint;
begin
  FCombine := False;

  if WMTitleBar.Visible then
  begin
    Show;
    exit;
  end;

  P := Point(0, 0);
  P := ClientToScreen(P);

  Hide;
  try
    WMTitleBar.Visible := True;
    SetBounds(P.X, P.Y{ - WMTitleBar.Height}, Width, Height + WMTitleBar.Height);
  finally
    Show;
  end;
end;

constructor TfrmWork.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{  FSnapped := False;
  FThresh  := 30; //get threshold
  FSizing  := False;

//  WMPanel.DoubleBuffered := True;
  FPanelInstance := MakeObjectInstance(PanelWndProc);

  FPanelHandle := WMPanel.Handle;
  FDefPanelProc := Pointer(GetWindowLong(FPanelHandle, GWL_WNDPROC));
  SetWindowLong(FPanelHandle, GWL_WNDPROC, LongInt(FPanelInstance)); }
end;

constructor TfrmWork.Create(AOwner: TComponent; ACombine: Boolean; ALeft, ATop, AWidth, AHeight: Integer);
var
  P: TPoint;
begin
  inherited Create(AOwner);
  SetBounds(ALeft, ATop, AWidth, AHeight);
  FCombineLeft := ALeft;
  FCombineWidth := AWidth;
  FCombineHeight := AHeight;
  FCombine := ACombine;

  if (FCombine) then
  begin
    P.X := Left;
    P.Y := Top;

    if P.X < 1 then P.X := 1;
    if P.Y < 1 then P.Y := 1;

    WMTitleBar.Visible := False;
    SetBounds(P.X, P.Y{ + WMTitleBar.Height}, Width, Height - WMTitleBar.Height);
  end;
end;

destructor TfrmWork.Destroy;
begin
{  if FPanelHandle <> 0 then
    SetWindowLong(FPanelHandle, GWL_WNDPROC, LongInt(FDefPanelProc));
  FreeObjectInstance(FPanelInstance);  }

  inherited Destroy;
end;

procedure TfrmWork.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
//  with Params do Style := (Style or WS_POPUP) and not WS_CAPTION;
//  with Params do Style := (Style or WS_THICKFRAME or WS_CAPTION);
end;

procedure TfrmWork.FormShow(Sender: TObject);
var
  W, H: Integer;
begin
  inherited;
  {  W := Width - (GetSystemMetrics(SM_CXDLGFRAME) + GetSystemMetrics(SM_CXBORDER)) * 2;
  H := Height - GetSystemMetrics(SM_CYCAPTION) - (GetSystemMetrics(SM_CXDLGFRAME) + GetSystemMetrics(SM_CXBORDER)) * 2;
  SetBounds(Left, Top, W, H); }
//  CombineForm(Self, frmWVMain);
//  WMPanel.SetFocus;
end;

procedure TfrmWork.WMIBCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmWork.WMIBMinimizeClick(Sender: TObject);
begin
  ShowWindow(Handle, SW_MINIMIZE);
end;

function TfrmWork.GetMonitorWorkArea(AMonHandle: HMONITOR): TRect;
var
  MonInfo: TMonitorInfo;
begin
  MonInfo.cbSize := SizeOf(MonInfo);
  GetMonitorInfo(AMonHandle, @MonInfo);
  Result := MonInfo.rcWork;
end;

procedure TfrmWork.SetCombine(Value: Boolean);
begin
  if Value then
  begin
//  FCombineWidth := Width;
//  FCombineHeight := Height;
    CombineForm(TWinControl(Owner));
  end
  else SplitForm;
end;

procedure TfrmWork.PanelWndProc(var Message: TMessage);
begin
  with Message do
  begin
    case Msg of
      WM_NCHITTEST:
        with TWMNCHitTest(Message) do
          begin
            Result := Windows.HTTRANSPARENT;
            Exit;
          end;
    end;
    Result := CallWindowProc(FDefPanelProc, FPanelHandle, Msg, WParam, LParam);
  end;
end;

procedure TfrmWork.WMActivate(var Message: TWMActivate);
begin
{  case Message.Active of
    WA_INACTIVE:
    begin
      WMPanel.ColorHighLight := $00727272;
      WMPanel.ColorShadow := $00131313;
      WMPanel.Repaint;
    end
    else
    begin
      WMPanel.ColorHighLight := $0029BC68;
      WMPanel.ColorShadow := $0029BC68;
      WMPanel.Repaint;
    end;
  end; }
  inherited;
end;

procedure TfrmWork.WMNCHitTest(var Message: TWMNCHitTest);
const
  BorderSize = 8;
var
  P: TPoint;
begin
  inherited;
  if FCombine then exit;

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

procedure TfrmWork.WMNCLButtonDown(var Msg: TWMNCLButtonDown);
var
  I, J: Integer;
  SR: TRect;
  FormHandle: HWND;
  
  // Checking left adjoin form
  // Checking right adjoin form
  // Checking top adjoin form
  // Checking bottom adjoin form

  // Checking left sizing on top adjoin form
  // return adjoin top value

  // Checking left sizing on bottom adjoin form
  // return adjoin bottom value

  // Checking right sizing on top adjoin form
  // return adjoin top value

  // Checking right sizing on bottom adjoin form
  // return adjoin bottom value

  // Checking top sizing on left adjoin form
  // return adjoin left value

  // Checking top sizing on right adjoin form
  // return adjoin right value

  // Checking bottom sizing on left adjoin form
  // return adjoin left value

  // Checking bottom sizing on right adjoin form
  // return adjoin right value

  // Get adjoin left form in top from bottom
  function GetSizingLeftFormRange(ALeft, ATop, ABottom: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockRight) and
         (WorkForm.Handle <> 0) and 
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ALeft = SR.Right) and
           ((InRange(SR.Top, ATop, ABottom)) or
            (InRange(SR.Bottom, ATop, ABottom)) or
            (InRange(ATop, SR.Top, SR.Bottom)) or
            (InRange(ABottom, SR.Top, SR.Bottom))) then
        begin
          WorkForm.FSizingDockRight := True;
          Result := WorkForm;
          break;
        end;
      end;  
    end;
  end;
  
  // Get adjoin right form in top from bottom
  function GetSizingRightFormRange(ARight, ATop, ABottom: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockLeft) and
         (WorkForm.Handle <> 0) and 
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ARight = SR.Left) and
           ((InRange(SR.Top, ATop, ABottom)) or
            (InRange(SR.Bottom, ATop, ABottom)) or
            (InRange(ATop, SR.Top, SR.Bottom)) or
            (InRange(ABottom, SR.Top, SR.Bottom))) then
        begin
          WorkForm.FSizingDockLeft := True;
          Result := WorkForm;
          break;
        end;
      end;  
    end;
  end;
  
  // Get adjoin top form in left from right
  function GetSizingTopFormRange(ATop, ALeft, ARight: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockBottom) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ATop = SR.Bottom) and
           ((InRange(SR.Left, ALeft, ARight)) or
            (InRange(SR.Right, ALeft, ARight)) or
            (InRange(ALeft, SR.Left, SR.Right)) or
            (InRange(ARight, SR.Left, SR.Right))) then
        begin
          WorkForm.FSizingDockBottom := True;
          Result := WorkForm;
          break;
        end;
      end;
    end;
  end;

  // Get adjoin bottom form in left from right
  function GetSizingBottomFormRange(ABottom, ALeft, ARight: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockTop) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ABottom = SR.Top) and
           ((InRange(SR.Left, ALeft, ARight)) or
            (InRange(SR.Right, ALeft, ARight)) or
            (InRange(ALeft, SR.Left, SR.Right)) or
            (InRange(ARight, SR.Left, SR.Right))) then
        begin
          WorkForm.FSizingDockTop := True;
          Result := WorkForm;
          break;
        end;
      end;
    end;
  end;

  // Get adjoin top form on left sizing
  function GetSizingTopFormLeft(ATop, ALeft: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockLeft) and
         (WorkForm.Handle <> 0) and 
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ATop = SR.Bottom) and
           (ALeft = SR.Left) then
        begin
          WorkForm.FSizingDockLeft := True;
          Result := WorkForm;
          break;
        end;
      end;  
    end;
  end;
  
  // Get adjoin bottom form on left sizing
  function GetSizingBottomFormLeft(ABottom, ALeft: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockLeft) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ABottom = SR.Top) and
           (ALeft = SR.Left) then
        begin
          WorkForm.FSizingDockLeft := True;
          Result := WorkForm;
          break;
        end;
      end;
    end;
  end;

  // Get adjoin top form on right sizing
  function GetSizingTopFormRight(ATop, ARight: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockRight) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ATop = SR.Bottom) and
           (ARight = SR.Right) then
        begin
          WorkForm.FSizingDockRight := True;
          Result := WorkForm;
          break;
        end;
      end;  
    end;
  end;

  // Get adjoin bottom form on right sizing
  function GetSizingBottomFormRight(ABottom, ARight: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockRight) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ABottom = SR.Top) and
           (ARight = SR.Right) then
        begin
          WorkForm.FSizingDockRight := True;
          Result := WorkForm;
          break;
        end;
      end;
    end;
  end;

  // Get adjoin left form on top sizing
  function GetSizingLeftFormTop(ALeft, ATop: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockTop) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ALeft = SR.Right) and
           (ATop = SR.Top) then
        begin
          WorkForm.FSizingDockTop := True;
          Result := WorkForm;
          break;
        end;
      end;
    end;
  end;

  // Get adjoin right form on top sizing
  function GetSizingRightFormTop(ARight, ATop: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockTop) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ARight = SR.Left) and
           (ATop = SR.Top) then
        begin
          WorkForm.FSizingDockTop := True;
          Result := WorkForm;
          break;
        end;
      end;
    end;
  end;

  // Get adjoin left form on bottom sizing
  function GetSizingLeftFormBottom(ALeft, ABottom: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockBottom) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ALeft = SR.Right) and
           (ABottom = SR.Bottom) then
        begin
          WorkForm.FSizingDockBottom := True;
          Result := WorkForm;
          break;
        end;
      end;
    end;
  end;

  // Get adjoin right form on bottom sizing
  function GetSizingRightFormBottom(ARight, ABottom: Integer): TfrmWork;
  var
    I: Integer;
    WorkForm: TfrmWork;
    SR: TRect;
  begin
    Result := nil;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      WorkForm := TfrmWork(Screen.Forms[I]);
      if not (WorkForm.FSizingDockBottom) and
         (WorkForm.Handle <> 0) and
         (IsWindowVisible(WorkForm.Handle)) and
         (GetWindowRect(WorkForm.Handle, SR)) then
      begin
        if (ARight = SR.Left) and
           (ABottom = SR.Bottom) then
        begin
          WorkForm.FSizingDockBottom := True;
          Result := WorkForm;
          break;
        end;
      end;
    end;
  end;

  // Checking left sizing
  procedure CheckSizingLeft;
  var
    I: Integer;
    R, SR: TRect;
    LeftForm, TopForm, RightForm, BottomForm: TfrmWork;
  begin
    FSizingDockLeft := True;

    GetWindowRect(Handle, R);
    repeat
      LeftForm   := nil;
      RightForm  := nil;
      TopForm    := nil;
      BottomForm := nil;

      // Checking left sizing adjoin top and bottom
      LeftForm := GetSizingLeftFormRange(R.Left, R.Top, R.Bottom);
      if (LeftForm <> nil) then
      begin
        GetWindowRect(LeftForm.Handle, SR);
        if SR.Top < R.Top then
        begin
          R.Top := SR.Top;

          // Checking sizingdockright adjoin top and bottom
          RightForm := GetSizingRightFormRange(R.Left, SR.Top, SR.Bottom);
          if RightForm <> nil then
          begin
            GetWindowRect(RightForm.Handle, SR);
            if SR.Top < R.Top then
            begin
              R.Top := SR.Top;

              // Checking left sizing adjoin top
              TopForm := GetSizingTopFormLeft(R.Left, SR.Top);
            end;
          end;
        end
        else if SR.Bottom > R.Bottom then
        begin
          R.Bottom := SR.Bottom;

          // Checking sizingdockright adjoin top and bottom
          RightForm := GetSizingRightFormRange(R.Left, SR.Top, SR.Bottom);
          if RightForm <> nil then
          begin
            GetWindowRect(RightForm.Handle, SR);
            if SR.Bottom > R.Bottom then
            begin
              R.Bottom := SR.Bottom;

              // Checking left sizing adjoin bottom
              BottomForm := GetSizingBottomFormLeft(R.Left, SR.Bottom);
            end;
          end;
        end;
      end
      else
      begin
        // Checking left sizing adjoin top
        TopForm := GetSizingTopFormLeft(R.Top, R.Left);
        if (TopForm <> nil) then
        begin
          GetWindowRect(TopForm.Handle, SR);
          R.Top := SR.Top;
        end;

        // Checking left sizing adjoin bottom
        BottomForm := GetSizingBottomFormLeft(R.Bottom, R.Left);
        if (BottomForm <> nil) then
        begin
          GetWindowRect(BottomForm.Handle, SR);
          R.Bottom := SR.Bottom;
        end;
      end;
    until (LeftForm = nil) and (RightForm = nil) and (TopForm = nil) and (BottomForm = nil);
  end;

  // Checking right sizing
  procedure CheckSizingRight;
  var
    I: Integer;
    R, SR: TRect;
    LeftForm, TopForm, RightForm, BottomForm: TfrmWork;
  begin
    FSizingDockRight := True;

    GetWindowRect(Handle, R);
    repeat
      LeftForm   := nil;
      RightForm  := nil;
      TopForm    := nil;
      BottomForm := nil;

      // Checking right sizing adjoin top and bottom
      RightForm := GetSizingRightFormRange(R.Right, R.Top, R.Bottom);
      if (RightForm <> nil) then
      begin
        GetWindowRect(RightForm.Handle, SR);
        if SR.Top < R.Top then
        begin
          R.Top := SR.Top;

          // Checking sizingdockleft adjoin top and bottom
          LeftForm := GetSizingLeftFormRange(R.Right, SR.Top, SR.Bottom);
          if LeftForm <> nil then
          begin
            GetWindowRect(LeftForm.Handle, SR);
            if SR.Top < R.Top then
            begin
              R.Top := SR.Top;

              // Checking left sizing adjoin top
              TopForm := GetSizingTopFormRight(R.Right, SR.Top);
            end;
          end;
        end
        else if SR.Bottom > R.Bottom then
        begin
          R.Bottom := SR.Bottom;

          // Checking sizingdockleft adjoin top and bottom
          LeftForm := GetSizingLeftFormRange(R.Right, SR.Top, SR.Bottom);
          if LeftForm <> nil then
          begin
            GetWindowRect(LeftForm.Handle, SR);
            if SR.Bottom > R.Bottom then
            begin
              R.Bottom := SR.Bottom;

              // Checking left sizing adjoin bottom
              BottomForm := GetSizingBottomFormRight(R.Right, SR.Bottom);
            end;
          end;
        end;
      end
      else
      begin
        // Checking right sizing adjoin top
        TopForm := GetSizingTopFormRight(R.Top, R.Right);
        if (TopForm <> nil) then
        begin
          GetWindowRect(TopForm.Handle, SR);
          R.Top := SR.Top;
        end;

        // Checking right sizing adjoin bottom
        BottomForm := GetSizingBottomFormRight(R.Bottom, R.Right);
        if (BottomForm <> nil) then
        begin
          GetWindowRect(BottomForm.Handle, SR);
          R.Bottom := SR.Bottom;
        end;
      end;
    until (LeftForm = nil) and (RightForm = nil) and (TopForm = nil) and (BottomForm = nil);
  end;

  // Checking top sizing
  procedure CheckSizingTop;
  var
    I: Integer;
    R, SR: TRect;
    LeftForm, TopForm, RightForm, BottomForm: TfrmWork;
  begin
    FSizingDockTop := True;

    GetWindowRect(Handle, R);
    repeat
      LeftForm   := nil;
      RightForm  := nil;
      TopForm    := nil;
      BottomForm := nil;

      // Checking top sizing adjoin left and right
      TopForm := GetSizingTopFormRange(R.Top, R.Left, R.Right);
      if (TopForm <> nil) then
      begin
        GetWindowRect(TopForm.Handle, SR);
        if SR.Left < R.Left then
        begin
          R.Left := SR.Left;

          // Checking sizingdockbottom adjoin left and right
          BottomForm := GetSizingBottomFormRange(R.Top, SR.Left, SR.Right);
          if BottomForm <> nil then
          begin
            GetWindowRect(BottomForm.Handle, SR);
            if SR.Left < R.Left then
            begin
              R.Left := SR.Left;

              // Checking top sizing adjoin left
              LeftForm := GetSizingLeftFormTop(R.Top, SR.Left);
            end;
          end;
        end
        else if SR.Right > R.Right then
        begin
          R.Right := SR.Right;

          // Checking sizingdockbottom adjoin left and right
          BottomForm := GetSizingBottomFormRange(R.Top, SR.Left, SR.Right);
          if BottomForm <> nil then
          begin
            GetWindowRect(BottomForm.Handle, SR);
            if SR.Right > R.Right then
            begin
              R.Right := SR.Right;

              // Checking top sizing adjoin right
              RightForm := GetSizingRightFormTop(R.Top, SR.Right);
            end;
          end;
        end;
      end
      else
      begin
        // Checking top sizing adjoin left
        LeftForm := GetSizingLeftFormTop(R.Left, R.Top);
        if (LeftForm <> nil) then
        begin
          GetWindowRect(LeftForm.Handle, SR);
          R.Left := SR.Left;
        end;

        // Checking top sizing adjoin right
        RightForm := GetSizingRightFormTop(R.Right, R.Top);
        if (RightForm <> nil) then
        begin
          GetWindowRect(RightForm.Handle, SR);
          R.Right := SR.Right;
        end;
      end;
    until (LeftForm = nil) and (RightForm = nil) and (TopForm = nil) and (BottomForm = nil);
  end;

  // Checking bottom sizing
  procedure CheckSizingBottom;
  var
    I: Integer;
    R, SR: TRect;
    LeftForm, TopForm, RightForm, BottomForm: TfrmWork;
  begin
    FSizingDockBottom := True;

    GetWindowRect(Handle, R);
    repeat
      LeftForm   := nil;
      RightForm  := nil;
      TopForm    := nil;
      BottomForm := nil;

      // Checking bottom sizing adjoin left and right
      BottomForm := GetSizingBottomFormRange(R.Bottom, R.Left, R.Right);
      if (BottomForm <> nil) then
      begin
        GetWindowRect(BottomForm.Handle, SR);
        if SR.Left < R.Left then
        begin
          R.Left := SR.Left;

          // Checking sizingdocktop adjoin left and right
          TopForm := GetSizingTopFormRange(R.Bottom, SR.Left, SR.Right);
          if TopForm <> nil then
          begin
            GetWindowRect(TopForm.Handle, SR);
            if SR.Left < R.Left then
            begin
              R.Left := SR.Left;

              // Checking bottom sizing adjoin left
              LeftForm := GetSizingLeftFormBottom(R.Bottom, SR.Left);
            end;
          end;
        end
        else if SR.Right > R.Right then
        begin
          R.Right := SR.Right;

          // Checking sizingdocktop adjoin left and right
          TopForm := GetSizingTopFormRange(R.Bottom, SR.Left, SR.Right);
          if TopForm <> nil then
          begin
            GetWindowRect(TopForm.Handle, SR);
            if SR.Right > R.Right then
            begin
              R.Right := SR.Right;

              // Checking top sizing adjoin right
              RightForm := GetSizingRightFormBottom(R.Bottom, SR.Right);
            end;
          end;
        end;
      end
      else
      begin
        // Checking bottom sizing adjoin left
        LeftForm := GetSizingLeftFormBottom(R.Left, R.Bottom);
        if (LeftForm <> nil) then
        begin
          GetWindowRect(LeftForm.Handle, SR);
          R.Left := SR.Left;
        end;

        // Checking top sizing adjoin right
        RightForm := GetSizingRightFormBottom(R.Right, R.Bottom);
        if (RightForm <> nil) then
        begin
          GetWindowRect(RightForm.Handle, SR);
          R.Right := SR.Right;
        end;
      end;
    until (LeftForm = nil) and (RightForm = nil) and (TopForm = nil) and (BottomForm = nil);
  end;

begin
  if Msg.HitTest in [//htBorder,
                     htTopLeft, htTopRight, htBottomLeft, htBottomRight,
                     htLeft, htRight, htTop, htBottom] then
  begin
    FSizing      := True;
{    FSizingRules := [];
    if (Msg.HitTest in [htTopLeft, htBottomLeft, htLeft]) then FSizingRules := FSizingRules + [srLeft];
    if (Msg.HitTest in [htTopRight, htBottomRight, htRight]) then FSizingRules := FSizingRules + [srRight];
    if (Msg.HitTest in [htTopLeft, htTopRight, htTop]) then FSizingRules := FSizingRules + [srTop];
    if (Msg.HitTest in [htBottomLeft, htBottomRight, htBottom]) then FSizingRules := FSizingRules + [srBottom];

    CheckDock(Left, Top, Width, Height); }
    
    
    FSizingLeftMin := 0;

    FSizingLeft   := (Msg.HitTest in [htTopLeft, htBottomLeft, htLeft]);
    FSizingRight  := (Msg.HitTest in [htTopRight, htBottomRight, htRight]);
    FSizingTop    := (Msg.HitTest in [htTopLeft, htTopRight, htTop]);
    FSizingBottom := (Msg.HitTest in [htBottomLeft, htBottomRight, htBottom]);

    if FSizingLeft then CheckSizingLeft;
    if FSizingRight then CheckSizingRight;
    if FSizingTop then CheckSizingTop;
    if FSizingBottom then CheckSizingBottom;

{      for I := 0 to Screen.FormCount - 1 do
      begin
        if Screen.Forms[I] = Self then Continue;
        if not (Screen.Forms[I] is TfrmWork) then Continue;

        if TfrmWork(Screen.Forms[I]).FSizingDockLeft then
          ShowMessage(Format('FSizingDockLeft : %s', [Screen.Forms[I].Name]));
        if TfrmWork(Screen.Forms[I]).FSizingDockRight then
          ShowMessage(Format('FSizingDockRight : %s', [Screen.Forms[I].Name]));
        if TfrmWork(Screen.Forms[I]).FSizingDockTop then
          ShowMessage(Format('FSizingDockTop : %s', [Screen.Forms[I].Name]));
        if TfrmWork(Screen.Forms[I]).FSizingDockBottom then
          ShowMessage(Format('FSizingDockBottom : %s', [Screen.Forms[I].Name]));
      end; }

    
{    for I := 0 to Screen.FormCount - 1 do
    begin
      if Screen.Forms[I] = Self then Continue;
      if not (Screen.Forms[I] is TfrmWork) then Continue;
//      ShowMessage((Screen.Forms[I].Name));

      if FSizingLeft then
      begin
        CheckDockLeftTop(Left, Top);
        CheckDockLeftBottom(Left, Top + Height);
      end;
      
{      FormHandle := Screen.Forms[I].Handle;
      if (FormHandle <> 0) and (IsWindowVisible(FormHandle)) and
         (GetWindowRect(FormHandle, SR)) then
      begin
        TfrmWork(Screen.Forms[I]).FSizingDockLeft   := ((FSizingLeft) and 
                                                        (Left = SR.Left) and
                                                        (Left = SR.Left) and
                                                        ) or
                                                       (FSizingRight and ((Left + Width) = SR.Left));
        TfrmWork(Screen.Forms[I]).FSizingDockRight  := (FSizingLeft and (Left = SR.Right))or
                                                       (FSizingRight and ((Left + Width) = SR.Right));                                                      
        TfrmWork(Screen.Forms[I]).FSizingDockTop    := (FSizingTop and (Top = SR.Top)) or
                                                       (FSizingBottom and ((Top + Height) = SR.Top));
        TfrmWork(Screen.Forms[I]).FSizingDockBottom := (FSizingTop and (Top = SR.Bottom)) or
                                                       (FSizingBottom and ((Top + Height) = SR.Bottom));

//        if FSizingDockLeft or
//           FSizingDockRight or
//           FSizingDockTop or
//           FSizingDockBottom then break; 
      end; 
    end; }
  end
  else
  begin
    FSizing := False;
    for I := 0 to Screen.FormCount - 1 do
    begin
      TfrmWork(Screen.Forms[I]).FSizingDockLeft := False;
      TfrmWork(Screen.Forms[I]).FSizingDockRight := False;
      TfrmWork(Screen.Forms[I]).FSizingDockTop := False;
      TfrmWork(Screen.Forms[I]).FSizingDockBottom := False;
    end;
  end;

  inherited;
end;

procedure TfrmWork.WMEnterSizeMove(var Msg: TMessage);
begin
//  ShowMessage(IntToStr(Width));
  inherited;
//  ShowMessage(IntToStr(Width));
end;

procedure TfrmWork.WMExitSizeMove(var Msg: TMessage);
var
  I: Integer;
begin
  FSizing := False;
  for I := 0 to Screen.FormCount - 1 do
  begin
    TfrmWork(Screen.Forms[I]).FSizingDockLeft := False;
    TfrmWork(Screen.Forms[I]).FSizingDockRight := False;
    TfrmWork(Screen.Forms[I]).FSizingDockTop := False;
    TfrmWork(Screen.Forms[I]).FSizingDockBottom := False;
  end;

  inherited;
end;

procedure TfrmWork.WMTitleBarSystemMenuClick(Sender: TObject;
  DoubleClicked: Boolean);
begin
  if DoubleClicked then Close;
end;

procedure TfrmWork.WMMoving(var Msg: TWMMoving);
var
  I: Integer;
begin
  FSizing := False;
  FMoving := True;
  for I := 0 to Screen.FormCount - 1 do
  begin
    TfrmWork(Screen.Forms[I]).FSizingDockLeft := False;
    TfrmWork(Screen.Forms[I]).FSizingDockRight := False;
    TfrmWork(Screen.Forms[I]).FSizingDockTop := False;
    TfrmWork(Screen.Forms[I]).FSizingDockBottom := False;
  end;

  inherited;
end;

procedure TfrmWork.WMMove(var Msg: TWMMove);
begin
  FMoving := False;

  inherited;
end;

procedure TfrmWork.WMSizing(var Msg: TMessage);
var
  SizeFlag: DWORD;
  SizeRect: PRect;
begin
  SizeFlag := Msg.WParam;
  SizeRect := PRect(Msg.LParam);
  if SizeRect^.Right - SizeRect^.Left < GetSystemMetrics(SM_CXMIN) then
  begin
    if SizeFlag in [WMSZ_LEFT, WMSZ_TOPLEFT, WMSZ_BOTTOMLEFT] then
      SizeRect^.Left := SizeRect^.Right - GetSystemMetrics(SM_CXMIN)
    else
      SizeRect^.Right := SizeRect^.Left + GetSystemMetrics(SM_CXMIN);
    Msg.Result := 1;
  end;

  if SizeRect^.Bottom - SizeRect^.Top < GetSystemMetrics(SM_CYMIN) then
  begin
    if SizeFlag in [WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT] then
      SizeRect^.Top:= SizeRect^.Bottom - GetSystemMetrics(SM_CYMIN)
    else
      SizeRect^.Bottom := SizeRect^.Top + GetSystemMetrics(SM_CYMIN);
    Msg.Result := 1;
  end;

  inherited;
end;

procedure TfrmWork.WMSize(var Msg: TWMSize);
begin
  inherited;
end;

procedure TfrmWork.WMWindowPosChanging(var Msg: TWMWindowPosChanging);
var
  I: Integer;
  FormHandle: HWND;
  SR: TRect;
  SnapWidth, SnapHeight: Integer;
  MinWidth, MinHeight: Integer;
  MaxWidth, MaxHeight: Integer;
begin
  inherited;

  if (Msg.WindowPos^.flags and SWP_NOOWNERZORDER) = 0 then exit;

  if (FSizing) then
  begin
    FSnapped := False;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if Screen.Forms[I] = Self then Continue;
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      FormHandle := Screen.Forms[I].Handle;
      // test window
      if (FormHandle <> 0) and (IsWindowVisible(FormHandle)) then
      begin
        if (GetWindowRect(FormHandle, SR)) then
        begin
          if (TfrmWork(Screen.Forms[I]).FSizingDockLeft) or (TfrmWork(Screen.Forms[I]).FSizingDockRight) or
             (TfrmWork(Screen.Forms[I]).FSizingDockTop) or (TfrmWork(Screen.Forms[I]).FSizingDockBottom) then
          begin
            SnapWidth  := SR.Right - SR.Left;
            SnapHeight := SR.Bottom - SR.Top;

            MinWidth := GetSystemMetrics(SM_CXMIN);
            MinWidth := Max(MinWidth, Screen.Forms[I].Constraints.MinWidth);

            MaxWidth := 0;
            MaxWidth := Max(MaxWidth, Screen.Forms[I].Constraints.MaxWidth);

            MinHeight := GetSystemMetrics(SM_CYMIN);
            MinHeight := Max(MinHeight, Screen.Forms[I].Constraints.MinHeight);

            MaxHeight := 0;
            MaxHeight := Max(MaxHeight, Screen.Forms[I].Constraints.MaxHeight);

            // Minimum and maximum size checking
            if (FSizingLeft) then
            begin
              if (TfrmWork(Screen.Forms[I]).FSizingDockLeft) then
              begin
                SnapWidth := SnapWidth - (Msg.WindowPos.x - SR.Left);
                if (SnapWidth <= MinWidth) then
                begin
                  SR.Left := Msg.WindowPos.x;
                  Msg.WindowPos.x := Msg.WindowPos.x - (MinWidth - SnapWidth);
                  Msg.WindowPos.cx := Msg.WindowPos.cx - (Msg.WindowPos.x - SR.Left);
                end
                else if (MaxWidth > 0) and (SnapWidth >= MaxWidth) then
                begin
                  SR.Left := Msg.WindowPos.x;
                  Msg.WindowPos.x := Msg.WindowPos.x - (MaxWidth - SnapWidth);
                  Msg.WindowPos.cx := Msg.WindowPos.cx - (Msg.WindowPos.x - SR.Left);
                end;
              end
              else if (TfrmWork(Screen.Forms[I]).FSizingDockRight) then
              begin
                SnapWidth  := Msg.WindowPos.x - SR.Left;
                if (SnapWidth <= MinWidth) then
                begin
                  SnapWidth := MinWidth;
                  Msg.WindowPos.cx := Msg.WindowPos.cx - ((SR.Left + SnapWidth) - Msg.WindowPos.x);
                  Msg.WindowPos.x  := SR.Left + SnapWidth;
                end
                else if (MaxWidth > 0) and (SnapWidth >= MaxWidth) then
                begin
                  SnapWidth := MaxWidth;
                  Msg.WindowPos.cx := Msg.WindowPos.cx - ((SR.Left + SnapWidth) - Msg.WindowPos.x);
                  Msg.WindowPos.x  := SR.Left + SnapWidth;
                end;
              end;
            end
            else if (FSizingRight) then
            begin
              if (TfrmWork(Screen.Forms[I]).FSizingDockLeft) then
              begin
                SnapWidth  := SnapWidth - (Msg.WindowPos.x + Msg.WindowPos.cx - SR.Left);
                if (SnapWidth <= MinWidth) then
                begin
                  SR.Left  := SR.Right - MinWidth;
                  SnapWidth := MinWidth;
                  Msg.WindowPos.cx := SR.Left - Msg.WindowPos.x;
                end
                else if (MaxWidth > 0) and (SnapWidth >= MaxWidth) then
                begin
                  SR.Left  := SR.Right - MaxWidth;
                  SnapWidth := MaxWidth;
                  Msg.WindowPos.cx := SR.Left - Msg.WindowPos.x;
                end;
              end
              else if (TfrmWork(Screen.Forms[I]).FSizingDockRight) then
              begin
                SnapWidth := SnapWidth + (Msg.WindowPos.x + Msg.WindowPos.cx - SR.Right);
                if (SnapWidth <= MinWidth) then
                begin
                  SnapWidth := MinWidth;
                  Msg.WindowPos.cx := (SR.Left + SnapWidth) - Msg.WindowPos.x;
                end
                else if (MaxWidth > 0) and (SnapWidth >= MaxWidth) then
                begin
                  SnapWidth := MaxWidth;
                  Msg.WindowPos.cx := (SR.Left + SnapWidth) - Msg.WindowPos.x;
                end;
              end;
            end;

            if (FSizingTop) then
            begin
              if (TfrmWork(Screen.Forms[I]).FSizingDockTop) then
              begin
                SnapHeight := SnapHeight - (Msg.WindowPos.y - SR.Top);
                if (SnapHeight <= MinHeight) then
                begin
                  SR.Top := Msg.WindowPos.y;
                  Msg.WindowPos.y := Msg.WindowPos.y - (MinHeight - SnapHeight);
                  Msg.WindowPos.cy := Msg.WindowPos.cy - (Msg.WindowPos.y - SR.Top);
                end
                else if (MaxHeight > 0) and (SnapHeight >= MaxHeight) then
                begin
                  SR.Top := Msg.WindowPos.y;
                  Msg.WindowPos.y := Msg.WindowPos.y - (MaxHeight - SnapHeight);
                  Msg.WindowPos.cy := Msg.WindowPos.cy - (Msg.WindowPos.y - SR.Top);
                end;
              end
              else if (TfrmWork(Screen.Forms[I]).FSizingDockBottom) then
              begin
                SnapHeight := Msg.WindowPos.y - SR.Top;
                if (SnapHeight <= MinHeight) then
                begin
                  SnapHeight := MinHeight;
                  Msg.WindowPos.cy := Msg.WindowPos.cy - ((SR.Top + SnapHeight) - Msg.WindowPos.y);
                  Msg.WindowPos.y  := SR.Top + SnapHeight;
                end
                else if (MaxHeight > 0) and (SnapHeight >= MaxHeight) then
                begin
                  SnapHeight := MaxHeight;
                  Msg.WindowPos.cy := Msg.WindowPos.cy - ((SR.Top + SnapHeight) - Msg.WindowPos.y);
                  Msg.WindowPos.y  := SR.Top + SnapHeight;
                end;
              end
            end
            else if (FSizingBottom) then
            begin
              if (TfrmWork(Screen.Forms[I]).FSizingDockTop) then
              begin
                SnapHeight  := SnapHeight - (Msg.WindowPos.y + Msg.WindowPos.cy - SR.Top);
                if (SnapHeight <= MinHeight) then
                begin
                  SR.Top  := SR.Bottom - MinHeight;
                  SnapHeight := MinHeight;
                  Msg.WindowPos.cy := SR.Top - Msg.WindowPos.y;
                end
                else if (MaxHeight > 0) and (SnapHeight >= MaxHeight) then
                begin
                  SR.Top  := SR.Bottom - MaxHeight;
                  SnapHeight := MinHeight;
                  Msg.WindowPos.cy := SR.Top - Msg.WindowPos.y;
                end;
              end
              else if (TfrmWork(Screen.Forms[I]).FSizingDockBottom) then
              begin
                SnapHeight := SnapHeight + (Msg.WindowPos.y + Msg.WindowPos.cy - SR.Bottom);
                if (Msg.WindowPos.cy > 0) and
                   (SnapHeight <= MinHeight) then
                begin
                  SnapHeight := MinHeight;
                  Msg.WindowPos.cy := (SR.Top + SnapHeight) - Msg.WindowPos.y;
                end
                else if (MaxHeight > 0) and (SnapHeight >= MaxHeight) then
                begin
                  SnapHeight := MaxHeight;
                  Msg.WindowPos.cy := (SR.Top + SnapHeight) - Msg.WindowPos.y;
                end;
              end;
            end;

            continue;
          end;

        //  WMTitleBar1.Caption := Format('%d, %d : %d, %d', [Msg.WindowPos.X, Msg.WindowPos.Y, SR.Right, SR.Top]);
          // Right-> Left, SR.Right
          if ((Msg.WindowPos.X <= (SR.Right + FThresh)) and
              (Msg.WindowPos.X >= (SR.Right - FThresh))) then
          begin
            if (InRange(SR.Top, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(SR.Bottom, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(Msg.WindowPos.y, SR.Top - FThresh, SR.Bottom + FThresh)) or
               (InRange(Msg.WindowPos.y + Msg.WindowPos.cy, SR.Top - FThresh, SR.Bottom + FThresh)) then
            begin
              FSnapped := True;
              if (Msg.WindowPos.cx <> Width) then
              begin
                SnapWidth := Msg.WindowPos.x + Msg.WindowPos.cx - SR.Right;// - Msg.WindowPos.x;

                if (Msg.WindowPos.cx > 0) and
                   (SnapWidth <= Constraints.MinWidth) then
                begin
                  SnapWidth := Constraints.MinWidth;
                  Msg.WindowPos.cx := SnapWidth;
                end
                else if (Constraints.MaxWidth > 0) and (SnapWidth >= Constraints.MaxWidth) then
                begin
                  SnapWidth := Constraints.MaxWidth;
                  Msg.WindowPos.cx := SnapWidth;
                end
                else
                begin
                  Msg.WindowPos.cx := SnapWidth;
                  Msg.WindowPos.x := SR.Right;
                end;
              end;
            end;
          end
          // Right-> Left, SR.Left
          else if ((Msg.WindowPos.X <= (SR.Left + FThresh)) and
                   (Msg.WindowPos.X >= (SR.Left - FThresh))) then
          begin
            if (InRange(SR.Top, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(SR.Bottom, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(Msg.WindowPos.y, SR.Top - FThresh, SR.Bottom + FThresh)) or
               (InRange(Msg.WindowPos.y + Msg.WindowPos.cy, SR.Top - FThresh, SR.Bottom + FThresh)) then
            begin
              FSnapped := True;
              if (Msg.WindowPos.cx <> Width) then
              begin
                SnapWidth := Msg.WindowPos.x + Msg.WindowPos.cx - SR.Left;// - Msg.WindowPos.x;

                if (Msg.WindowPos.cx > 0) and
                   (SnapWidth <= Constraints.MinWidth) then
                begin
                  SnapWidth := Constraints.MinWidth;
                  Msg.WindowPos.cx := SnapWidth;
                end
                else if (Constraints.MaxWidth > 0) and (SnapWidth >= Constraints.MaxWidth) then
                begin
                  SnapWidth := Constraints.MaxWidth;
                  Msg.WindowPos.cx := SnapWidth;
                end
                else
                begin
                  Msg.WindowPos.cx := SnapWidth;
                  Msg.WindowPos.x := SR.Left;
                end;
              end;
            end;
          end;

          // Left -> Right, SR.Left
          if ((Msg.WindowPos.x + Msg.WindowPos.cx) >= (SR.Left - FThresh)) and
             ((Msg.WindowPos.x + msg.WindowPos.cx) <= (SR.Left + FThresh)) then
          begin
            if (InRange(SR.Top, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(SR.Bottom, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(Msg.WindowPos.y, SR.Top - FThresh, SR.Bottom + FThresh)) or
               (InRange(Msg.WindowPos.y + Msg.WindowPos.cy, SR.Top - FThresh, SR.Bottom + FThresh)) then
            begin
              FSnapped := True;
              if (Msg.WindowPos.cx <> Width) then
              begin
                SnapWidth := SR.Left - Msg.WindowPos.x;

                if (Msg.WindowPos.cx > 0) and
                   (SnapWidth <= Constraints.MinWidth) then
                begin
                  SnapWidth := Constraints.MinWidth;
                  Msg.WindowPos.cx := SnapWidth;
                end
                else if (Constraints.MaxWidth > 0) and (SnapWidth >= Constraints.MaxWidth) then
                begin
                  SnapWidth := Constraints.MaxWidth;
                  Msg.WindowPos.cx := SnapWidth;
                end
                else
                begin
                  Msg.WindowPos.cx := SnapWidth;
                end;
              end;
            end;
          end
          // Left -> Right, SR.Right
          else if ((Msg.WindowPos.x + Msg.WindowPos.cx) >= (SR.Right - FThresh)) and
                  ((Msg.WindowPos.x + msg.WindowPos.cx) <= (SR.Right + FThresh)) then
          begin
            if (InRange(SR.Top, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(SR.Bottom, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(Msg.WindowPos.y, SR.Top - FThresh, SR.Bottom + FThresh)) or
               (InRange(Msg.WindowPos.y + Msg.WindowPos.cy, SR.Top - FThresh, SR.Bottom + FThresh)) then
            begin
              FSnapped := True;
              if (Msg.WindowPos.cx <> Width) then
              begin
                SnapWidth := SR.Right - Msg.WindowPos.x;

                if (Msg.WindowPos.cx > 0) and
                   (SnapWidth <= Constraints.MinWidth) then
                begin
                  SnapWidth := Constraints.MinWidth;
                  Msg.WindowPos.cx := SnapWidth;
                end
                else if (Constraints.MaxWidth > 0) and (SnapWidth >= Constraints.MaxWidth) then
                begin
                  SnapWidth := Constraints.MaxWidth;
                  Msg.WindowPos.cx := SnapWidth;
                end
                else
                begin
                  Msg.WindowPos.cx := SnapWidth;
                end;
              end;
            end;
          end;

          // Top -> Bottom, SR.Bottom
          if ((Msg.WindowPos.y <= (SR.Bottom + FThresh)) and
              (Msg.WindowPos.y >= (SR.Bottom - FThresh))) then
          begin
            if (InRange(SR.Left, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(SR.Right, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(Msg.WindowPos.x, SR.Left - FThresh, SR.Right + FThresh)) or
               (InRange(Msg.WindowPos.x + Msg.WindowPos.cx, SR.Left - FThresh, SR.Right + FThresh)) then
            begin
              FSnapped := True;
              if (Msg.WindowPos.cy <> Height) then
              begin
                SnapHeight := Msg.WindowPos.y + Msg.WindowPos.cy - SR.Bottom;

                if (Msg.WindowPos.cy > 0) and
                   (SnapHeight <= Constraints.MinHeight) then
                begin
                  SnapHeight := Constraints.MinHeight;
                  Msg.WindowPos.cy := SnapHeight;
                end
                else if (Constraints.MaxHeight > 0) and (SnapHeight >= Constraints.MaxHeight) then
                begin
                  SnapHeight := Constraints.MaxHeight;
                  Msg.WindowPos.cy := SnapHeight;
                end
                else
                begin
                  Msg.WindowPos.cy := SnapHeight;
                  Msg.WindowPos.y := SR.Bottom;
                end;
              end;
            end;
          end
          // Top -> Bottom, SR.Top
          else if ((Msg.WindowPos.y <= (SR.Top + FThresh)) and
                   (Msg.WindowPos.y >= (SR.Top - FThresh))) then
          begin
            if (InRange(SR.Left, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(SR.Right, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(Msg.WindowPos.x, SR.Left - FThresh, SR.Right + FThresh)) or
               (InRange(Msg.WindowPos.x + Msg.WindowPos.cx, SR.Left - FThresh, SR.Right + FThresh)) then
            begin
              FSnapped := True;
              if (Msg.WindowPos.cy <> Height) then
              begin
                SnapHeight := Msg.WindowPos.y + Msg.WindowPos.cy - SR.Top;

                if (Msg.WindowPos.cy > 0) and
                   (SnapHeight <= Constraints.MinHeight) then
                begin
                  SnapHeight := Constraints.MinHeight;
                  Msg.WindowPos.cy := SnapHeight;
                end
                else if (Constraints.MaxHeight > 0) and (SnapHeight >= Constraints.MaxHeight) then
                begin
                  SnapHeight := Constraints.MaxHeight;
                  Msg.WindowPos.cy := SnapHeight;
                end
                else
                begin
                  Msg.WindowPos.cy := SnapHeight;
                  Msg.WindowPos.y := SR.Top;
                end;
              end;
            end;
          end;

          // Bottom -> Top, SR.Bottom
          if ((Msg.WindowPos.y + Msg.WindowPos.cy) <= (SR.Bottom + FThresh)) and
             ((Msg.WindowPos.y + msg.WindowPos.cy) >= (SR.Bottom - FThresh)) then
          begin
            if (InRange(SR.Left, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(SR.Right, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(Msg.WindowPos.x, SR.Left - FThresh, SR.Right + FThresh)) or
               (InRange(Msg.WindowPos.x + Msg.WindowPos.cx, SR.Left - FThresh, SR.Right + FThresh)) then
            begin
              FSnapped := True;
              if (Msg.WindowPos.cy <> Height) then
              begin
                SnapHeight := SR.Bottom - Msg.WindowPos.y;

                if (Msg.WindowPos.cy > 0) and
                   (SnapHeight <= Constraints.MinHeight) then
                begin
                  SnapHeight := Constraints.MinHeight;
                  Msg.WindowPos.cy := SnapHeight;
                end
                else if (Constraints.MaxHeight > 0) and (SnapHeight >= Constraints.MaxHeight) then
                begin
                  SnapHeight := Constraints.MaxHeight;
                  Msg.WindowPos.cy := SnapHeight;
                end
                else Msg.WindowPos.cy := SnapHeight;
              end;
            end;
          end
          // Bottom -> Top, SR.Top
          else if ((Msg.WindowPos.y + Msg.WindowPos.cy) <= (SR.Top + FThresh)) and
                  ((Msg.WindowPos.y + msg.WindowPos.cy) >= (SR.Top - FThresh)) then
          begin
            if (InRange(SR.Left, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(SR.Right, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(Msg.WindowPos.x, SR.Left - FThresh, SR.Right + FThresh)) or
               (InRange(Msg.WindowPos.x + Msg.WindowPos.cx, SR.Left - FThresh, SR.Right + FThresh)) then
            begin
              FSnapped := True;
              if (Msg.WindowPos.cy <> Height) then
              begin
                SnapHeight := SR.Top - Msg.WindowPos.y;

                if (Msg.WindowPos.cy > 0) and
                   (SnapHeight <= Constraints.MinHeight) then
                begin
                  SnapHeight := Constraints.MinHeight;
                  Msg.WindowPos.cy := SnapHeight;
                end
                else if (Constraints.MaxHeight > 0) and (SnapHeight >= Constraints.MaxHeight) then
                begin
                  SnapHeight := Constraints.MaxHeight;
                  Msg.WindowPos.cy := SnapHeight;
                end
                else Msg.WindowPos.cy := SnapHeight;
              end;
            end;
          end;
        end;
      end;
    end;

    // Full screen
    for I := 0 to Screen.MonitorCount - 1 do
    begin
      SR := GetMonitorWorkArea(Screen.Monitors[I].Handle);
      if (Msg.WindowPos.x <> Left) and
         ((Msg.WindowPos.x) >= (SR.Left - FThresh)) and
         ((Msg.WindowPos.x) <= (SR.Left + FThresh)) then
      begin
        FSnapped := True;
        Msg.WindowPos.cx := Msg.WindowPos.x + Msg.WindowPos.cx - SR.Left;
        Msg.WindowPos.x := SR.Left;
      end;

      if ((Msg.WindowPos.x + Msg.WindowPos.cx) <> (Left + Width)) and
         ((Msg.WindowPos.x + Msg.WindowPos.cx) >= (SR.Right - FThresh)) and
         ((Msg.WindowPos.x + Msg.WindowPos.cx) <= (SR.Right + FThresh)) then
      begin
        FSnapped := True;
        Msg.WindowPos.cx := SR.Right - Msg.WindowPos.x;
      end;

      if (Msg.WindowPos.y <> Top) and
         ((Msg.WindowPos.y) >= (SR.Top - FThresh)) and
         ((Msg.WindowPos.y) <= (SR.Top + FThresh)) then
      begin
        FSnapped := True;
        Msg.WindowPos.cy := Msg.WindowPos.y + Msg.WindowPos.cy - SR.Top;
        Msg.WindowPos.y := SR.Top;
      end;

      if ((Msg.WindowPos.y + Msg.WindowPos.cy) <> (Top + Height)) and
         ((Msg.WindowPos.y + Msg.WindowPos.cy) >= (SR.Bottom - FThresh)) and
         ((Msg.WindowPos.y + Msg.WindowPos.cy) <= (SR.Bottom + FThresh)) then
      begin
        FSnapped := True;
        Msg.WindowPos.cy := SR.Bottom - Msg.WindowPos.y;
      end;
    end;
  end
  else if (FMoving) then
  begin
    FSnapped := False;
    for I := 0 to Screen.FormCount - 1 do
    begin
      if Screen.Forms[I] = Self then Continue;
      if Screen.Forms[I].Active then Continue;
      if Screen.Forms[I].Tag <> 0 then Continue;
      if not (Screen.Forms[I] is TfrmWork) then Continue;

      FormHandle := Screen.Forms[I].Handle;
      // test window
      if (FormHandle <> 0) and (IsWindowVisible(FormHandle)) then
      begin
        if (GetWindowRect(FormHandle, SR)) then
        begin
          // Right-> Left, SR.Right
          if ((Msg.WindowPos.X <= (SR.Right + FThresh)) and
              (Msg.WindowPos.X >= (SR.Right - FThresh))) then
          begin
            if (InRange(SR.Top, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(SR.Bottom, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(Msg.WindowPos.y, SR.Top - FThresh, SR.Bottom + FThresh)) or
               (InRange(Msg.WindowPos.y + Msg.WindowPos.cy, SR.Top - FThresh, SR.Bottom + FThresh)) then
            begin
              FSnapped := True;
              Msg.WindowPos.x := SR.Right;
            end;
          end
          // Right-> Left, SR.Left
          else if ((Msg.WindowPos.X <= (SR.Left + FThresh)) and
                   (Msg.WindowPos.X >= (SR.Left - FThresh))) then
          begin
            if (InRange(SR.Top, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(SR.Bottom, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(Msg.WindowPos.y, SR.Top - FThresh, SR.Bottom + FThresh)) or
               (InRange(Msg.WindowPos.y + Msg.WindowPos.cy, SR.Top - FThresh, SR.Bottom + FThresh)) then
            begin
              FSnapped := True;
              Msg.WindowPos.x := SR.Left;
            end;
          end;

          // Left -> Right, SR.Left
          if ((Msg.WindowPos.x + Msg.WindowPos.cx) >= (SR.Left - FThresh)) and
                  ((Msg.WindowPos.x + msg.WindowPos.cx) <= (SR.Left + FThresh)) then
          begin
            if (InRange(SR.Top, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(SR.Bottom, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(Msg.WindowPos.y, SR.Top - FThresh, SR.Bottom + FThresh)) or
               (InRange(Msg.WindowPos.y + Msg.WindowPos.cy, SR.Top - FThresh, SR.Bottom + FThresh)) then
            begin
              FSnapped := True;
              Msg.WindowPos.x := SR.Left - Msg.WindowPos.cx;
            end;
          end
          // Left -> Right, SR.Right
          else if ((Msg.WindowPos.x + Msg.WindowPos.cx) >= (SR.Right - FThresh)) and
                  ((Msg.WindowPos.x + msg.WindowPos.cx) <= (SR.Right + FThresh)) then
          begin
            if (InRange(SR.Top, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(SR.Bottom, Msg.WindowPos.y, Msg.WindowPos.y + Msg.WindowPos.cy)) or
               (InRange(Msg.WindowPos.y, SR.Top - FThresh, SR.Bottom + FThresh)) or
               (InRange(Msg.WindowPos.y + Msg.WindowPos.cy, SR.Top - FThresh, SR.Bottom + FThresh)) then
            begin
              FSnapped := True;
              Msg.WindowPos.x := SR.Right - Msg.WindowPos.cx;
            end;
          end;

          // Top -> Bottom, SR.Bottom
          if ((Msg.WindowPos.y <= (SR.Bottom + FThresh)) and
              (Msg.WindowPos.y >= (SR.Bottom - FThresh))) then
          begin
            if (InRange(SR.Left, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(SR.Right, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(Msg.WindowPos.x, SR.Left - FThresh, SR.Right + FThresh)) or
               (InRange(Msg.WindowPos.x + Msg.WindowPos.cx, SR.Left - FThresh, SR.Right + FThresh)) then
            begin
              FSnapped := True;
              Msg.WindowPos.y := SR.Bottom;
            end;
          end
          // Top -> Bottom, SR.Top
          else if ((Msg.WindowPos.y <= (SR.Top + FThresh)) and
                   (Msg.WindowPos.y >= (SR.Top - FThresh))) then
          begin
            if (InRange(SR.Left, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(SR.Right, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(Msg.WindowPos.x, SR.Left - FThresh, SR.Right + FThresh)) or
               (InRange(Msg.WindowPos.x + Msg.WindowPos.cx, SR.Left - FThresh, SR.Right + FThresh)) then
            begin
              FSnapped := True;
              Msg.WindowPos.y := SR.Top;
            end;
          end;

          // Bottom -> Top, SR.Bottom
          if ((Msg.WindowPos.y + Msg.WindowPos.cy) <= (SR.Bottom + FThresh)) and
                  ((Msg.WindowPos.y + msg.WindowPos.cy) >= (SR.Bottom - FThresh)) then
          begin
            if (InRange(SR.Left, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(SR.Right, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(Msg.WindowPos.x, SR.Left - FThresh, SR.Right + FThresh)) or
               (InRange(Msg.WindowPos.x + Msg.WindowPos.cx, SR.Left - FThresh, SR.Right + FThresh)) then
            begin
              FSnapped := True;
              Msg.WindowPos.y := SR.Bottom - Msg.WindowPos.cy;
            end;
          end
          // Bottom -> Top, SR.Top
          else if ((Msg.WindowPos.y + Msg.WindowPos.cy) <= (SR.Top + FThresh)) and
                  ((Msg.WindowPos.y + msg.WindowPos.cy) >= (SR.Top - FThresh)) then
          begin
            if (InRange(SR.Left, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(SR.Right, Msg.WindowPos.x, Msg.WindowPos.x + Msg.WindowPos.cx)) or
               (InRange(Msg.WindowPos.x, SR.Left - FThresh, SR.Right + FThresh)) or
               (InRange(Msg.WindowPos.x + Msg.WindowPos.cx, SR.Left - FThresh, SR.Right + FThresh)) then
            begin
              FSnapped := True;
              Msg.WindowPos.y := SR.Top - Msg.WindowPos.cy;
            end;
          end;
        end;
      end;
    end;

    // Full screen
    for I := 0 to Screen.MonitorCount - 1 do
    begin
      SR := GetMonitorWorkArea(Screen.Monitors[I].Handle);
      if ((Msg.WindowPos.x) >= (SR.Left - FThresh)) and
         ((Msg.WindowPos.x) <= (SR.Left + FThresh)) then
      begin
        FSnapped := True;
        Msg.WindowPos.x := SR.Left;
      end;

      if ((Msg.WindowPos.x + Msg.WindowPos.cx) >= (SR.Right - FThresh)) and
              ((Msg.WindowPos.x + Msg.WindowPos.cx) <= (SR.Right + FThresh)) then
      begin
        FSnapped := True;
        Msg.WindowPos.x := SR.Right - Msg.WindowPos.cx;
      end;

      if ((Msg.WindowPos.y) >= (SR.Top - FThresh)) and
         ((Msg.WindowPos.y) <= (SR.Top + FThresh)) then
      begin
        FSnapped := True;
        Msg.WindowPos.y := SR.Top;
      end;

      if ((Msg.WindowPos.y + Msg.WindowPos.cy) >= (SR.Bottom - FThresh)) and
              ((Msg.WindowPos.y + Msg.WindowPos.cy) <= (SR.Bottom + FThresh)) then
      begin
        FSnapped := True;
        Msg.WindowPos.y := SR.Bottom - Msg.WindowPos.cy;
      end;
    end;
  end;
end;

procedure TfrmWork.WMWindowPosChanged(var Msg: TWMWindowPosChanged);
var
  I: Integer;
  FormHandle: HWND;
  SR: TRect;
  SnapWidth, SnapHeight: Integer;
begin
  inherited;

  if (FScreenWorkAreaChanged) then
  begin
    SetWindowPos(Handle, 0, FChangeLeft, FChangeTop, FChangeWidth, FChangeHeight, SWP_NOACTIVATE);
//    Application.BringToFront;
//    SetBounds(FChangeLeft, FChangeTop, FChangeWidth, FChangeHeight);
    FScreenWorkAreaChanged := False;
    exit;
  end;

  if (Msg.WindowPos^.flags and SWP_NOOWNERZORDER) = 0 then exit;

  if not FSizing then exit;
  for I := 0 to Screen.FormCount - 1 do
  begin
    if Screen.Forms[I] = Self then Continue;
    if not (Screen.Forms[I] is TfrmWork) then Continue;

    FormHandle := Screen.Forms[I].Handle;
    // test window
    if (FormHandle <> 0) and (IsWindowVisible(FormHandle)){ and
       ((Msg.WindowPos.X <> 0) or (Msg.WindowPos.Y <> 0))} then
    begin
      if (GetWindowRect(FormHandle, SR)) then
      begin
        if (TfrmWork(Screen.Forms[I]).FSizingDockLeft) or (TfrmWork(Screen.Forms[I]).FSizingDockRight) or
           (TfrmWork(Screen.Forms[I]).FSizingDockTop) or (TfrmWork(Screen.Forms[I]).FSizingDockBottom) then
        begin
          SnapWidth  := SR.Right - SR.Left;
          SnapHeight := SR.Bottom - SR.Top;

          if (FSizingLeft) then
          begin
            if (TfrmWork(Screen.Forms[I]).FSizingDockLeft) then
            begin
              SnapWidth := SnapWidth - (Msg.WindowPos.x - SR.Left);
              SR.Left  := Msg.WindowPos.x;
            end
            else if (TfrmWork(Screen.Forms[I]).FSizingDockRight) then
            begin
              SnapWidth := Msg.WindowPos.x - SR.Left;
            end;
          end
          else if (FSizingRight) then
          begin
            if (TfrmWork(Screen.Forms[I]).FSizingDockRight) then
            begin
              SnapWidth := SnapWidth + (Msg.WindowPos.x + Msg.WindowPos.cx - SR.Right);
            end
            else if (TfrmWork(Screen.Forms[I]).FSizingDockLeft) then
            begin
              SnapWidth := SnapWidth - (Msg.WindowPos.x + Msg.WindowPos.cx - SR.Left);
              SR.Left  := Msg.WindowPos.x + Msg.WindowPos.cx;
            end;
          end;

          if (FSizingTop) then
          begin
            if (TfrmWork(Screen.Forms[I]).FSizingDockTop) then
            begin
              SnapHeight := SnapHeight - (Msg.WindowPos.y - SR.Top);
              SR.Top    := Msg.WindowPos.y;
            end
            else if (TfrmWork(Screen.Forms[I]).FSizingDockBottom) then
            begin
              SnapHeight := Msg.WindowPos.y - SR.Top;
            end;
          end
          else if (FSizingBottom) then
          begin
            if (TfrmWork(Screen.Forms[I]).FSizingDockBottom) then
            begin
              SnapHeight := SnapHeight + (Msg.WindowPos.y + Msg.WindowPos.cy - SR.Bottom);
            end
            else if (TfrmWork(Screen.Forms[I]).FSizingDockTop) then
            begin
              SnapHeight  := SnapHeight - (Msg.WindowPos.y + Msg.WindowPos.cy - SR.Top);
              SR.Top  := Msg.WindowPos.y + Msg.WindowPos.cy;
            end
          end;

{    TfrmWork(Screen.Forms[I]).FSizingDockLeft := False;
    TfrmWork(Screen.Forms[I]).FSizingDockRight := False;
    TfrmWork(Screen.Forms[I]).FSizingDockTop := False;
    TfrmWork(Screen.Forms[I]).FSizingDockBottom := False;}
//          Screen.Forms[I].Tag := 1;
          SetWindowPos(FormHandle, 0, SR.Left, SR.Top, SnapWidth, SnapHeight, SWP_NOACTIVATE);
//          Screen.Forms[I].SetBounds(SR.Left, SR.Top, SnapWidth, SnapHeight);
//          Screen.Forms[I].Tag := 0;
//          break;
        end;
      end;
    end;
  end;
end;

procedure TfrmWork.WMDisplayChange(var Message: TWMDisplayChange);
var
  WidthFactor, HeightFactor: Double;
  L, T, W, H: Integer;
begin
  WidthFactor := Screen.WorkAreaWidth / FScreenWorkAreaWidth;
  HeightFactor := Screen.WorkAreaHeight / FScreenWorkAreaHeight;

  L := Trunc(Left * WidthFactor);
  T := Trunc(Top * HeightFactor);
  W := Trunc(Width * WidthFactor);
  H := Trunc(Height * HeightFactor);

  if (W < Constraints.MinWidth) then W := Constraints.MinWidth;
  if (H < Constraints.MinHeight) then H := Constraints.MinHeight;

  if (L < 0) then L := 0;
  if (T < 0) then T := 0;
  if ((L + W) > Screen.WorkAreaRect.Right) then
    L := Screen.WorkAreaRect.Right - W;
  if ((T + H) > Screen.WorkAreaRect.Bottom) then
    T := Screen.WorkAreaRect.Bottom - H;

  FChangeLeft   := L;
  FChangeTop    := T;
  FChangeWidth  := W;
  FChangeHeight := H;

  FScreenWorkAreaChanged := True;
//  Width := W;
//  Height := H;
//  SetBounds(L, T, W, H);
  inherited;

  FScreenWorkAreaWidth := Screen.WorkAreaWidth;
  FScreenWorkAreaHeight := Screen.WorkAreaHeight;

  Message.Result := 0;
end;

procedure TfrmWork.FormCanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
begin
  Resize := True;//(NewWidth >= GetSystemMetrics(SM_CXMIN)) and (NewHeight >= GetSystemMetrics(SM_CYMIN));
end;

procedure TfrmWork.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FClosing := True;
  Action := caFree;
end;

procedure TfrmWork.FormCreate(Sender: TObject);
var
  style: string;
begin
  FClosing := False;
  FSnapped := False;
  FThresh  := 30; //get threshold
  FSizing  := False;

//  WMPanel.DoubleBuffered := True;
  FPanelInstance := MakeObjectInstance(PanelWndProc);

  FPanelHandle := WMPanel.Handle;
  FDefPanelProc := Pointer(GetWindowLong(FPanelHandle, GWL_WNDPROC));
  SetWindowLong(FPanelHandle, GWL_WNDPROC, LongInt(FPanelInstance));
//       TStyleManager.LoadFromFile('OnyxBlue.vsf');
//  for style in TStyleManager.StyleNames do
//    ShowMessage(Style);

//  TStyleManager.SetStyle('Onyx Blue');
//  WMTitleBar.DoubleBuffered := True;

  FScreenWorkAreaWidth := Screen.WorkAreaWidth;
  FScreenWorkAreaHeight := Screen.WorkAreaHeight;

  FChangeWidth  := Width;
  FChangeHeight := Height;
  FScreenWorkAreaChanged := False;

  FSaveWidth := pnlDesc.ClientWidth;
  FSaveHeight := pnlDesc.ClientHeight;
end;

procedure TfrmWork.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
//
end;

procedure TfrmWork.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
//
end;

procedure TfrmWork.FormActivate(Sender: TObject);
begin
//
end;

procedure TfrmWork.FormDeactivate(Sender: TObject);
begin
//
end;

procedure TfrmWork.FormDestroy(Sender: TObject);
begin
  if FPanelHandle <> 0 then
    SetWindowLong(FPanelHandle, GWL_WNDPROC, LongInt(FDefPanelProc));
  FreeObjectInstance(FPanelInstance);
end;

procedure TfrmWork.FormResize(Sender: TObject);
var
  L, T, R, B, W, H, WW, HH, I: Integer;
begin
  if not Combine then
  begin
  FCombineLeft := Left;
  FCombineWidth := Width;
  FCombineHeight := Height + WMTitleBar.Height;
  end;

{  WW := pnlDesc.ClientWidth;
  HH := pnlDesc.ClientHeight;
  for I := 0 to (pnlDesc.ControlCount - 1) do
  begin
    ChangeDims(pnlDesc.Controls[i], WW, FSaveWidth, HH, FSaveHeight);
  end;
  FSaveWidth  := WW;
  FSaveHeight := HH; }
end;

procedure TfrmWork.ScaleCtrls(Control: TWinControl; WW, PW, HH, PH: Integer);
var
  I: Integer;
begin
  for I := 0 to Control.ControlCount - 1 do
    ChangeDims(Control.Controls[i], WW, PW, HH, PH);
end;

procedure TfrmWork.ChangeDims(Control: TControl; WW, PW, HH, PH: Integer);
var
  L, T, W, H: integer;
begin
  if Control is TWinControl then
    (Control as TWinControl).DisableAlign;
  try
    if Control is TWinControl then
      ScaleCtrls(Control as TWinControl, WW, PW, HH, PH);
    L := MulDiv(Control.Left, WW, PW);
    T := MulDiv(Control.Top, HH, PH);
    W := MulDiv(Control.Left + Control.Width, WW, PW) - L;
    H := MulDiv(Control.Top + Control.Height, HH, PH) - T;
    { THackControl(Control).Font.Size := (PW div WW) * THackControl(Control).Font.Size;}
    Control.SetBounds(L, T, W, H);
  finally
    if Control is TWinControl then
      (Control as TWinControl).DisableAlign;
  end;
end;

end.
