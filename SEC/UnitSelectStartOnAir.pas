unit UnitSelectStartOnAir;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitSingleForm, Vcl.Imaging.pngimage,
  WMTools, WMControls, Vcl.ExtCtrls, AdvOfficeButtons, Vcl.StdCtrls,
  System.Actions, Vcl.ActnList,
  UnitCommons, UnitConsts, UnitDCSDLL;

type
  TSelectOnAirStart = (
    SO_CURRENT,
    SO_FIRST
    );

  TfrmSelectStartOnAir = class(TfrmSingle)
    Label5: TLabel;
    Label2: TLabel;
    Label1: TLabel;
    Label3: TLabel;
    lblCurrentEvent1: TLabel;
    lblCurrentEvent2: TLabel;
    lblNextEvent1: TLabel;
    lblNextEvent2: TLabel;
    aorCurrentStart: TAdvOfficeRadioButton;
    aorNextStart: TAdvOfficeRadioButton;
    wmibStart: TWMImageButton;
    wmibCancel: TWMImageButton;
    actSelectStartOnAir: TActionList;
    actStart: TAction;
    actCancel: TAction;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure actStartExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
  private
    { Private declarations }
    FCurrItem: PCueSheetItem;
    FNextItem: PCueSheetItem;

    FSelectOnAirStart: TSelectOnAirStart;

    procedure Initialize;
    procedure Finalize;

    procedure GetPropertyValues;
    procedure SetPropertyValues;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent; ACurrItem, ANextItem: PCueSheetItem);

    property SelectOnAirStart: TSelectOnAirStart read FSelectOnAirStart;
  end;

var
  frmSelectStartOnAir: TfrmSelectStartOnAir;

implementation

uses UnitSEC;

{$R *.dfm}

procedure TfrmSelectStartOnAir.actCancelExecute(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TfrmSelectStartOnAir.actStartExecute(Sender: TObject);
begin
  inherited;
  SetPropertyValues;
end;

constructor TfrmSelectStartOnAir.Create(AOwner: TComponent; ACurrItem, ANextItem: PCueSheetItem);
begin
  inherited Create(AOwner);

  FCurrItem := ACurrItem;
  FNextItem := ANextItem;
end;

procedure TfrmSelectStartOnAir.Initialize;
begin
  lblCurrentEvent1.Caption := '';
  lblCurrentEvent2.Caption := '';

  lblNextEvent1.Caption := '';
  lblNextEvent2.Caption := '';

  FSelectOnAirStart := SO_CURRENT;
  aorCurrentStart.Checked := True;
end;

procedure TfrmSelectStartOnAir.Finalize;
begin

end;

procedure TfrmSelectStartOnAir.FormCreate(Sender: TObject);
begin
  inherited;
  Initialize;
  GetPropertyValues;
end;

procedure TfrmSelectStartOnAir.GetPropertyValues;
var
  CurrStartTime, NextStartTime: TEventTime;
  CurrDurationTC, NextDurationTC: TTimecode;

  procedure GetEventInfo(AItem: PCueSheetItem; var AStartTime: TEventTime; var ADurationTC: TTimecode);
  var
    Source: PSource;
    SourceHandle: PSourceHandle;
    R: Integer;
    I: Integer;
    StartTime: TEventTime;
    DurationTC: TTimecode;
  begin
    if (AItem = nil) then exit;

    AStartTime  := AItem^.StartTime;
    ADurationTC := AItem^.DurationTC;

    Source := GetSourceByName(String(AItem^.Source));
    if (Source <> nil) and (Source^.Handles <> nil) then
    begin
      for I := 0 to Source^.Handles.Count - 1 do
      begin
        SourceHandle := Source^.Handles[I];
        if (SourceHandle^.DCS <> nil) and (SourceHandle^.DCS^.Alive) and
           (SourceHandle^.Handle > INVALID_DEVICE_HANDLE) then
        begin
          R := frmSEC.DCSEventThread.GetEventInfo(SourceHandle, AItem^.EventID, StartTime, DurationTC);
//          R := DCSGetEventInfo(SourceHandle^.DCS^.ID, SourceHandle^.Handle, AItem^.EventID, StartTime, DurationTC);
          if (R = D_OK) then
          begin
            AStartTime  := StartTime;
            ADurationTC := DurationTC;
            break;
          end;
        end;
      end;
    end;
  end;

begin
  if (FCurrItem <> nil) then
  begin
    GetEventInfo(FCurrItem, CurrStartTime, CurrDurationTC);

    with FCurrItem^ do
    begin
      lblCurrentEvent1.Caption := Format('%s-%s', [String(Title), String(SubTitle)]);
      lblCurrentEvent2.Caption := Format('Source = %s, Start time = %s, Duration = %s', [String(Source), TimecodeToString(CurrStartTime.T), TimecodeToString(CurrDurationTC)]);
    end;
  end;

  if (FNextItem <> nil) then
  begin
    GetEventInfo(FNextItem, NextStartTime, NextDurationTC);

    with FNextItem^ do
    begin
      lblNextEvent1.Caption := Format('%s-%s', [String(Title), String(SubTitle)]);
      lblNextEvent2.Caption := Format('Source = %s, Start time = %s, Duration = %s', [String(Source), TimecodeToString(NextStartTime.T), TimecodeToString(NextDurationTC)]);
    end;
  end;
end;

procedure TfrmSelectStartOnAir.SetPropertyValues;
begin
  if (aorCurrentStart.Checked) then
    FSelectOnAirStart := SO_CURRENT
  else if (aorNextStart.Checked) then
    FSelectOnAirStart := SO_FIRST;
end;

end.
