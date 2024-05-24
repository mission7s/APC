unit UnitEditEvent;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitSingleForm, Vcl.Imaging.pngimage,
  WMTools, WMControls, Vcl.ExtCtrls, System.ImageList, Vcl.ImgList,
  System.Actions, Vcl.ActnList, Vcl.Mask, Vcl.ComCtrls, Vcl.StdCtrls,
  UnitCommons, UnitConsts;

type
  TfrmEditEvent = class(TfrmSingle)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    cbStartMode: TComboBox;
    dpStartDate: TDateTimePicker;
    edTitle: TEdit;
    cbSource: TComboBox;
    cbTransitionType: TComboBoxEx;
    edNotes: TEdit;
    edMediaId: TEdit;
    cbTransitionRate: TComboBox;
    Label12: TLabel;
    cbEventMode: TComboBox;
    Label13: TLabel;
    cbInput: TComboBoxEx;
    Label14: TLabel;
    cbOutput: TComboBoxEx;
    Label15: TLabel;
    edSubTitle: TEdit;
    wmibSearchMedia: TWMImageSpeedButton;
    meDurationTC: TMaskEdit;
    meInTC: TMaskEdit;
    meOutTC: TMaskEdit;
    actEditEvent: TActionList;
    actOK: TAction;
    actCancel: TAction;
    actSearchMedia: TAction;
    ilsTransitionType: TImageList;
    wmibCancel: TWMImageButton;
    wmibOK: TWMImageButton;
    Label16: TLabel;
    cbProgramType: TComboBoxEx;
    meStartTime: TMaskEdit;
    cbSourceLayer: TComboBox;
    lblSourceLayer: TLabel;
    Label17: TLabel;
    cbFinishAction: TComboBox;
    procedure actCancelExecute(Sender: TObject);
    procedure actOKExecute(Sender: TObject);
    procedure actSearchMediaExecute(Sender: TObject);
    procedure cbStartModeExit(Sender: TObject);
    procedure cbStartModeSelect(Sender: TObject);
    procedure cbTransitionTypeExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cbInputSelect(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cbSourceSelect(Sender: TObject);
    procedure meDurationTCExit(Sender: TObject);
    procedure meInTCExit(Sender: TObject);
    procedure meOutTCExit(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    FModified: Boolean;

    FSelectIndex: Integer;
    FEditMode: TEditMode;

    FEventMode: TEventMode;
    FInput: TInputType;

    FChannelCueSheet: PChannelCueSheet;

    FProgItem: PCueSheetItem;
    FCurrItem: PCueSheetItem;
    FAddItem: PCueSheetItem;

    FMediaDurTC: TTimecode;

    FSaveSourceIndex: Integer;

    function CheckPropertyValues: Boolean;

    procedure Initialize;
    procedure Finalize;

    procedure GetPropertyValues;
    function SetPropertyValues: Boolean;

    procedure SetEditMode(Value: TEditMode);

    procedure ClearTransitionType;

    procedure PopulateEventMode;
    procedure PopulateStartMode(AEventMode: TEventMode);
    procedure PopulateInput;
    procedure PopulateOutput(AInput: TInputType);
    procedure PopulateSource;
    procedure PopulateSourceLayer(ASource: PSource);
    procedure PopulateTransitionType;
    procedure PopulateTransitionRate;
    procedure PopulateFinishAction;
    procedure PopulateProgramType;

    procedure CalcDurTC(ADurTC: TTimecode);
    procedure CalcInTC(AInTC: TTimecode);
    procedure CalcOutTC(AOutTC: TTimecode);
  public
    { Public declarations }
    constructor Create(AOwner: TComponent; AIndex: Integer; AEditMode: TEditMode; AChannelCueSheet: PChannelCueSheet; ACurrItem: PCueSheetItem = nil; AProgItem: PCueSheetItem = nil); overload;
    constructor Create(AOwner: TComponent; AIndex: Integer; AEditMode: TEditMode; AChannelCueSheet: PChannelCueSheet; AEventMode: TEventMode = EM_MAIN; ACurrItem: PCueSheetItem = nil; AProgItem: PCueSheetItem = nil); overload;

//    property EditMode: TEditMode read FEditMode write SetEditMode;
    property AddItem: PCueSheetItem read FAddItem;
  end;

var
  frmEditEvent: TfrmEditEvent;

implementation

uses UnitSEC, UnitChannel;

{$R *.dfm}

constructor TfrmEditEvent.Create(AOwner: TComponent; AIndex: Integer; AEditMode: TEditMode; AChannelCueSheet: PChannelCueSheet; ACurrItem: PCueSheetItem = nil; AProgItem: PCueSheetItem = nil);
begin
  inherited Create(AOwner);

  FSelectIndex := AIndex;
  FEditMode := AEditMode;

  FChannelCueSheet := AChannelCueSheet;
  FProgItem := AProgItem;
  FCurrItem := ACurrItem;
  if (FCurrItem <> nil) then
    FEventMode := FCurrItem^.EventMode;

  if (FCurrItem <> nil) then
    FInput := FCurrItem^.Input
  else
    FInput := GV_SettingEventColumnDefault.Input;

  FSaveSourceIndex := -1;
end;

constructor TfrmEditEvent.Create(AOwner: TComponent; AIndex: Integer; AEditMode: TEditMode; AChannelCueSheet: PChannelCueSheet; AEventMode: TEventMode; ACurrItem: PCueSheetItem; AProgItem: PCueSheetItem);
begin
  Create(AOwner, AIndex, AEditMode, AChannelCueSheet, ACurrItem, AProgItem);
  FEventMode := AEventMode;
  FInput := GV_SettingEventColumnDefault.Input;
end;

procedure TfrmEditEvent.Initialize;
begin
  PopulateEventMode;
  PopulateStartMode(FEventMode);
  PopulateInput;
  PopulateSource;
  PopulateSourceLayer(nil);
  PopulateTransitionType;
  PopulateTransitionRate;
  PopulateFinishAction;
  PopulateProgramType;

  WMTitleBar.Caption := Format('%s event - %s', [EditModeNames[FEditMode], EventModeNames[FEventMode]]);

  cbEventMode.ItemIndex := -1;
  cbStartMode.ItemIndex := -1;
  dpStartDate.Date      := Date;
  meStartTime.Text      := INIT_TIMECODE;

  cbInput.ItemIndex  := -1;
  cbOutput.ItemIndex := -1;

  edTitle.Clear;
  edTitle.MaxLength := TITLE_LEN;

  edSubTitle.Clear;
  edSubTitle.MaxLength := SUBTITLE_LEN;

  cbSource.ItemIndex := -1;
  cbSourceLayer.ItemIndex := -1;

  edMediaId.Clear;
  edMediaId.MaxLength := MEDIAID_LEN;

  meDurationTC.Text := INIT_TIMECODE;
  meInTC.Text       := INIT_TIMECODE;
  meOutTC.Text      := INIT_TIMECODE;

  cbTransitionType.ItemIndex := -1;
  cbTransitionRate.ItemIndex := -1;

  cbFinishAction.ItemIndex := -1;

  edNotes.Clear;
  edNotes.MaxLength := NOTES_LEN;
end;

procedure TfrmEditEvent.meDurationTCExit(Sender: TObject);
begin
  inherited;
  if ((Sender as TMaskEdit).Modified) then
    CalcDurTC(StringToTimecode(meDurationTC.Text));
end;

procedure TfrmEditEvent.meInTCExit(Sender: TObject);
begin
  inherited;
  if ((Sender as TMaskEdit).Modified) then
    CalcInTC(StringToTimecode(meInTC.Text));
end;

procedure TfrmEditEvent.meOutTCExit(Sender: TObject);
begin
  inherited;
  if ((Sender as TMaskEdit).Modified) then
    CalcOutTC(StringToTimecode(meOutTC.Text));
end;

procedure TfrmEditEvent.Finalize;
begin
  ClearTransitionType;
end;

function TfrmEditEvent.CheckPropertyValues: Boolean;
var
  ErrorString: String;
  ErrorControl: TWinControl;
  Source: PSource;
begin
  Result := False;

  ErrorControl := nil;
  ErrorString := '';
  try
    if (cbEventMode.ItemIndex < 0) then
    begin
      ErrorString := SEnterEventMode;
      ErrorControl := cbEventMode;
    end
    else
    begin
      if (FEventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
      begin
        if (cbStartMode.ItemIndex < 0) then
        begin
          ErrorString := SEnterStartMode;
          ErrorControl := cbStartMode;
        end
        else if (cbInput.ItemIndex < 0) then
        begin
          ErrorString := SEnterInput;
          ErrorControl := cbInput;
        end
        else if (cbOutput.ItemIndex < 0) then
        begin
          ErrorString := SEnterOutput;
          ErrorControl := cbOutput;
        end
        else if (Trim(edTitle.Text) = '') then
        begin
          ErrorString := SEnterTitle;
          ErrorControl := edTitle;
        end
        else if (cbSource.ItemIndex < 0) then
        begin
          ErrorString := SEnterSource;
          ErrorControl := cbSource;
        end
        else if (cbSourceLayer.Visible) and (cbSourceLayer.ItemIndex < 0) then
        begin
          ErrorString := SEnterSourceLayer;
          ErrorControl := cbSourceLayer;
        end
        else if (cbTransitionType.ItemIndex < 0) then
        begin
          ErrorString := SEnterTransitionType;
          ErrorControl := cbTransitionType;
        end
        else if (cbTransitionRate.ItemIndex < 0) then
        begin
          ErrorString := SEnterTransitionRate;
          ErrorControl := cbTransitionRate;
        end
        else if (cbFinishAction.ItemIndex < 0) then
        begin
          ErrorString := SEnterFinishAction;
          ErrorControl := cbFinishAction;
        end
        else if (cbProgramType.ItemIndex < 0) then
        begin
          ErrorString := SEnterProgramType;
          ErrorControl := cbProgramType;
        end;
      end
      else if (FEventMode in [EM_PROGRAM, EM_COMMENT]) then
      begin
        if (Trim(edTitle.Text) = '') then
        begin
          ErrorString := SEnterTitle;
          ErrorControl := edTitle;
        end
      end;
    end;
  finally
    if ErrorString <> '' then
    begin
      MessageBeep(MB_ICONWARNING);
      MessageBox(Handle, PChar(ErrorString), PChar(Application.Title), MB_OK or MB_ICONWARNING or MB_TOPMOST);
      if (ErrorControl.CanFocus) then ErrorControl.SetFocus;
    end
    else
      Result := True;
  end;
end;

procedure TfrmEditEvent.GetPropertyValues;
var
  PItem: PCueSheetItem;
  StartTime: TEventTime;
begin
  case FEditMode of
    EM_INSERT:
    begin
      with GV_SettingEventColumnDefault do
      begin
//        FEditItem.EventID.ChannelID

        cbEventMode.ItemIndex := cbEventMode.Items.IndexOfObject(TObject(FEventMode));
        if (FEventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
        begin
          if (FEventMode = EM_MAIN) then
            cbStartMode.ItemIndex := cbStartMode.Items.IndexOfObject(TObject(StartMode))
          else
            cbStartMode.ItemIndex := cbStartMode.Items.IndexOfObject(TObject(SM_SUBBEGIN));

          if (FCurrItem <> nil) then
          begin
            if (FEventMode = EM_MAIN) then
            begin
              if (FSelectIndex >= 0) then
                // 현재 선택된 이벤트가 있는 경우 선택한 메인 이벤트의 시각시각으로 설정
                StartTime := FCurrItem^.StartTime
              else
                // 현재 선택된 이벤트가 없는 경우 마지막 이벤트의 End Time으로 설정
                StartTime := GetEventEndTime(FCurrItem^.StartTime, FCurrItem^.DurationTC);
            end
            else if (FEventMode in [EM_JOIN, EM_SUB]) then
            begin
              // 현재 선택된 이벤트에서 Join, Sub이벤트를 삽입 할 경우 Start Date는 현재 메인 이벤트의 Start Date 으로 설정
              // Start Time은 0으로 설정
              StartTime.D := FCurrItem^.StartTime.D;
              StartTime.T := 0;
            end;
            dpStartDate.Date := StartTime.D;
            meStartTime.Text := TimecodeToString(StartTime.T);
          end
          else
          begin
{            // 선택된 이벤트가 없을 시 큐시트의 일자와 00시로 설정
            with TfrmChannel(Owner).ChannelCueSheet^ do
            begin

              dpStartDate.Date := OnAirDateToDate(OnairDate);
              dpStartTime.Time := 0;
            end; }

            // 선택된 이벤트가 없을 시 현재일자와 00시로 설정
            dpStartDate.Date := Date;
            meStartTime.Text := INIT_TIMECODE;
          end;

          cbInput.ItemIndex := cbInput.Items.IndexOfObject(TObject(FInput));
          PopulateOutput(Input);

          if (FInput in [IT_MAIN, IT_BACKUP]) then
            cbOutput.ItemIndex := cbOutput.Items.IndexOfObject(TObject(OutputBkgnd))
          else
            cbOutput.ItemIndex := cbOutput.Items.IndexOfObject(TObject(OutputKeyer));


//          edTitle.Clear;
//          edSubTitle.Clear;
          cbSource.ItemIndex := cbSource.Items.IndexOfObject(TObject(Source));
          cbSourceLayer.ItemIndex := -1;
          edMediaId.Clear;

          if (FEventMode in [EM_JOIN, EM_SUB]) and (FCurrItem <> nil) then
          begin
            meDurationTC.Text := TimecodeToString(FCurrItem^.DurationTC);
            meInTC.Text       := TimecodeToString(FCurrItem^.InTC);
            meOutTC.Text      := TimecodeToString(FCurrItem^.OutTC);

            FMediaDurTC := FCurrItem^.DurationTC;
          end
          else
          begin
            meDurationTC.Text := TimecodeToString(DurationTC);
            meInTC.Text       := TimecodeToString(InTC);
            meOutTC.Text      := TimecodeToString(OutTC);

            FMediaDurTC := DurationTC;
          end;

  //        {$IFNDEF DEBUG}
          cbTransitionType.ItemIndex := Integer(TransitionType);// cbTransitionType.Items.IndexOfObject(TObject(TransitionType));
  //        {$ENDIF}
          cbTransitionRate.ItemIndex := cbTransitionRate.Items.IndexOfObject(TObject(TransitionRate));

          cbFinishAction.ItemIndex := cbFinishAction.Items.IndexOfObject(TObject(FinishAction));

          cbProgramType.ItemIndex := cbProgramType.Items.IndexOfObject(TObject(ProgramType));
          edNotes.Clear;
        end;

        edTitle.Clear;
        edSubTitle.Clear;
      end;
    end;
    EM_UPDATE:
    begin
      if (FCurrItem = nil) then exit;

      cbEventMode.ItemIndex := cbEventMode.Items.IndexOfObject(TObject(FCurrItem^.EventMode));

      PopulateStartMode(FCurrItem^.EventMode);
      if (FEventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
      begin
        cbStartMode.ItemIndex := cbStartMode.Items.IndexOfObject(TObject(FCurrItem^.StartMode));
        dpStartDate.Date      := FCurrItem^.StartTime.D;
        meStartTime.Text      := TimecodeToString(FCurrItem^.StartTime.T);

        cbInput.ItemIndex     := cbInput.Items.IndexOfObject(TObject(FInput));
        PopulateOutput(FCurrItem^.Input);

        cbOutput.ItemIndex    := cbOutput.Items.IndexOfObject(TObject(FCurrItem^.Output));

        cbSource.ItemIndex    := cbSource.Items.IndexOfObject(TObject(GetSourceByName(String(FCurrItem^.Source))));

        PopulateSourceLayer(GetSourceByName(String(FCurrItem^.Source)));
        cbSourceLayer.ItemIndex := cbSourceLayer.Items.IndexOfObject(TObject(FCurrItem^.SourceLayer));

        edMediaId.Text        := String(FCurrItem^.MediaId);
        meDurationTC.Text     := TimecodeToString(FCurrItem^.DurationTC);
        meInTC.Text           := TimecodeToString(FCurrItem^.InTC);
        meOutTC.Text          := TimecodeToString(FCurrItem^.OutTC);

  //        {$IFNDEF DEBUG}
        cbTransitionType.ItemIndex := Integer(FCurrItem^.TransitionType);// cbTransitionType.Items.IndexOfObject(TObject(TransitionType));
  //        {$ENDIF}
        cbTransitionRate.ItemIndex := cbTransitionRate.Items.IndexOfObject(TObject(FCurrItem^.TransitionRate));

        cbFinishAction.ItemIndex := cbFinishAction.Items.IndexOfObject(TObject(FCurrItem^.FinishAction));

        cbProgramType.ItemIndex    := cbProgramType.Items.IndexOfObject(TObject(FCurrItem^.ProgramType));
        edNotes.Text               := String(FCurrItem^.Notes);
      end;
      edTitle.Text    := String(FCurrItem^.Title);
      edSubTitle.Text := String(FCurrItem^.SubTitle);
    end;
    EM_DELETE: ;
  end;

  cbEventMode.Enabled := False;
  cbStartMode.Enabled := (FEventMode in [EM_MAIN, EM_SUB]);
  dpStartDate.Enabled := (FEventMode in [EM_MAIN]) and (TStartMode(cbStartMode.Items.Objects[cbStartMode.ItemIndex]) <> SM_AUTOFOLLOW);
  meStartTime.Enabled := (FEventMode in [EM_MAIN, EM_SUB]) and (TStartMode(cbStartMode.Items.Objects[cbStartMode.ItemIndex]) <> SM_AUTOFOLLOW);

//  cbStartMode.Visible := (FEventMode <> EM_COMMENT);
//  dpStartDate.Visible := cbStartMode.Visible;
//  meStartTime.Visible := cbStartMode.Visible;

  cbInput.Enabled  := (FEventMode in [EM_MAIN, EM_JOIN, EM_SUB]);
  cbOutput.Enabled := cbInput.Enabled;

  edSubTitle.Enabled      := (FEventMode in [EM_PROGRAM, EM_MAIN, EM_JOIN, EM_SUB]);
  cbSource.Enabled        := cbInput.Enabled;
  edMediaId.Enabled       := cbInput.Enabled;
  actSearchMedia.Enabled  := cbInput.Enabled;

  meDurationTC.Enabled  := cbStartMode.Enabled and (FInput <> IT_NONE);
  meInTC.Enabled        := meDurationTC.Enabled;
  meOutTC.Enabled       := meDurationTC.Enabled;

  cbTransitionType.Enabled := cbInput.Enabled;
  cbTransitionRate.Enabled := cbInput.Enabled;

  cbFinishAction.Enabled := cbInput.Enabled;

  cbProgramType.Enabled := cbInput.Enabled;

  edNotes.Enabled := cbInput.Enabled;
end;

function TfrmEditEvent.SetPropertyValues: Boolean;
var
  ErrorControl: TWinControl;
  Source: PSource;
  SaveItem: PCueSheetItem;
begin
  Result := True;

  case FEditMode of
    EM_INSERT:
    begin
      FAddItem := New(PCueSheetItem);
      FillChar(FAddItem^, SizeOf(TCueSheetItem), #0);

      with TfrmChannel(Owner) do
      begin
        FAddItem^.EventMode := FEventMode;

        if (FEventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
        begin
          FAddItem^.StartMode := TStartMode(cbStartMode.Items.Objects[cbStartMode.ItemIndex]);
          FAddItem^.StartTime.D := dpStartDate.Date;
          FAddItem^.StartTime.T := StringToTimecode(meStartTime.Text);
        end
        else
        begin
          FAddItem^.StartTime.D := 0;
          FAddItem^.StartTime.T := 0;
        end;
//          EventStatus;
//          DeviceStatus;

        StrPCopy(FAddItem^.Title, edTitle.Text);

        if (FEventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
        begin
          StrPCopy(FAddItem^.SubTitle, edSubTitle.Text);

          FAddItem^.Input  := TInputType(cbInput.Items.Objects[cbInput.ItemIndex]);
          FAddItem^.Output := Byte(cbOutput.Items.Objects[cbOutput.ItemIndex]);

          StrPCopy(FAddItem^.Source, cbSource.Text);

          if (cbSourceLayer.ItemIndex < 0) then
            FAddItem^.SourceLayer := 0
          else
            FAddItem^.SourceLayer := Integer(cbSourceLayer.Items.Objects[cbSourceLayer.ItemIndex]);

          StrPCopy(FAddItem^.MediaId, edMediaId.Text);

          FAddItem^.DurationTC := StringToTimecode(meDurationTC.Text);
          FAddItem^.InTC       := StringToTimecode(meInTC.Text);
          FAddItem^.OutTC      := StringToTimecode(meOutTC.Text);

          FAddItem^.TransitionType := TTRType(cbTransitionType.ItemsEx[cbTransitionType.ItemIndex].Data^);
          FAddItem^.TransitionRate := TTRRate(cbTransitionRate.Items.Objects[cbTransitionRate.ItemIndex]);

          FAddItem^.FinishAction := TFinishAction(cbFinishAction.Items.Objects[cbFinishAction.ItemIndex]);

          FAddItem^.ProgramType := Byte(cbProgramType.Items.Objects[cbProgramType.ItemIndex]);
          StrPCopy(FAddItem^.Notes, edNotes.Text);
        end;

        // Validation checking
        if (FEventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
        begin
          ErrorControl := nil;
          if (not IsValidStartDate(FAddItem, FAddItem^.StartTime.D, False)) then
          begin
            Result := False;
            ErrorControl := dpStartDate;
          end
          else if (not IsValidStartTime(FAddItem, FAddItem^.StartTime.T, False)) then
          begin
            Result := False;
            ErrorControl := meStartTime;
          end
          else if (not IsValidDuration(FAddItem, FAddItem^.DurationTC, False)) then
          begin
            Result := False;
            ErrorControl := meDurationTC;
          end
          else if (not IsValidInTC(FAddItem, FAddItem^.InTC, False)) then
          begin
            Result := False;
            ErrorControl := meInTC;
          end
          else if (not IsValidOutTC(FAddItem, FAddItem^.OutTC, False)) then
          begin
            Result := False;
            ErrorControl := meOutTC;
          end
          else
          begin
            Source := GetSourceByName(String(FAddItem^.Source));
            if (Source <> nil) and
               (Source^.SourceType in [ST_VSDEC, ST_CG]) then
            begin
              if (not IsValidMediaId(edMediaId.Text, False)) then
              begin
                Result := False;
                ErrorControl := edMediaId;
              end;
            end;
          end;

          if (not CheckCueSheetEditPossibleLockTimeByIndex(FSelectIndex)) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SENotPossibleEditLockTime), PChar(Application.Title), MB_OK or MB_ICONWARNING or MB_TOPMOST);
            Result := False;
            exit;
          end
          else if (not CheckCueSheetEditPossibleLocationByIndex(FSelectIndex)) then
          begin
            MessageBeep(MB_ICONWARNING);
            MessageBox(Handle, PChar(SENotPossibleEditBeforeLocationIncorrect), PChar(Application.Title), MB_OK or MB_ICONWARNING or MB_TOPMOST);
            Result := False;
            exit;
          end;
        end;

        if (not Result) then
        begin
          SetFocus;
          if (ErrorControl.CanFocus) then ErrorControl.SetFocus;
          Dispose(FAddItem);
          exit;
        end;

        FChannelCueSheet^.LastSerialNo := FChannelCueSheet^.LastSerialNo + 1;

        with FAddItem^.EventID do
        begin
          ChannelID := FChannelCueSheet^.ChannelID;
          StrCopy(OnAirDate, FChannelCueSheet^.OnairDate);
          OnAirFlag := FChannelCueSheet^.OnairFlag;
          OnAirNo   := FChannelCueSheet^.OnairNo;
          SerialNo  := FChannelCueSheet^.LastSerialNo;
        end;

        if (FEventMode = EM_PROGRAM) then
        begin
          FChannelCueSheet^.LastProgramNo := FChannelCueSheet^.LastProgramNo + 1;
          FAddItem^.ProgramNo := FChannelCueSheet^.LastProgramNo;

          FChannelCueSheet^.LastGroupNo := FChannelCueSheet^.LastGroupNo + 1;
          FAddItem^.GroupNo := FChannelCueSheet^.LastGroupNo;

          if (FCurrItem <> nil) then
          begin
            if (FSelectIndex >= 0) then
            begin
              FAddItem^.DisplayNo := FCurrItem^.DisplayNo;
              LastDisplayNo := FAddItem^.DisplayNo;
            end
            else
            begin
              LastDisplayNo := LastDisplayNo + 1;
              FAddItem^.DisplayNo := LastDisplayNo;
            end;
          end
          else
          begin
            LastDisplayNo := LastDisplayNo + 1;
            FAddItem^.DisplayNo := LastDisplayNo;
          end;
        end
        else if (FEventMode = EM_MAIN) then
        begin
          if (FProgItem <> nil) then
          begin
            FAddItem^.ProgramNo := FProgItem^.ProgramNo;
          end
          else
          begin
            FChannelCueSheet^.LastProgramNo := FChannelCueSheet^.LastProgramNo + 1;
            FAddItem^.ProgramNo := FChannelCueSheet^.LastProgramNo;
          end;

          FChannelCueSheet^.LastGroupNo := FChannelCueSheet^.LastGroupNo + 1;
          FAddItem^.GroupNo := FChannelCueSheet^.LastGroupNo;

          if (FCurrItem <> nil) then
          begin
            if (FSelectIndex >= 0) then
            begin
              FAddItem^.DisplayNo := FCurrItem^.DisplayNo;
              LastDisplayNo := FAddItem^.DisplayNo;
            end
            else
            begin
              LastDisplayNo := LastDisplayNo + 1;
              FAddItem^.DisplayNo := LastDisplayNo;
            end;
          end
          else
          begin
            LastDisplayNo := 0;
            FAddItem^.DisplayNo := LastDisplayNo;
          end;
        end
        else if (FEventMode in [EM_JOIN, EM_SUB]) then
        begin
          if (FCurrItem = nil) then
          begin
            FChannelCueSheet^.LastProgramNo := FChannelCueSheet^.LastProgramNo + 1;
            FAddItem^.ProgramNo := FChannelCueSheet^.LastProgramNo;

            FChannelCueSheet^.LastGroupNo := FChannelCueSheet^.LastGroupNo + 1;
            FAddItem^.GroupNo := FChannelCueSheet^.LastGroupNo;

            LastDisplayNo := LastDisplayNo + 1;
            FAddItem^.DisplayNo := LastDisplayNo;
          end
          else
          begin
            FAddItem^.ProgramNo := FCurrItem^.ProgramNo;
            FAddItem^.GroupNo   := FCurrItem^.GroupNo;
            FAddItem^.DisplayNo := FCurrItem^.DisplayNo;
          end;
        end
        else if (FEventMode = EM_COMMENT) then
        begin
          if (FCurrItem = nil) then
          begin
            FChannelCueSheet^.LastProgramNo := FChannelCueSheet^.LastProgramNo + 1;
            FAddItem^.ProgramNo := FChannelCueSheet^.LastProgramNo;

            FChannelCueSheet^.LastGroupNo := FChannelCueSheet^.LastGroupNo + 1;
            FAddItem^.GroupNo := FChannelCueSheet^.LastGroupNo;

            LastDisplayNo := LastDisplayNo + 1;
            FAddItem^.DisplayNo := LastDisplayNo;
          end
          else
          begin
{            if (FCurrItem^.EventMode = EM_PROGRAM) or (FProgItem = nil) then
            begin
              FChannelCueSheet^.LastProgramNo := FChannelCueSheet^.LastProgramNo + 1;
              FAddItem^.ProgramNo := FChannelCueSheet^.LastProgramNo;

              FChannelCueSheet^.LastGroupNo := FChannelCueSheet^.LastGroupNo + 1;
              FAddItem^.GroupNo := FChannelCueSheet^.LastGroupNo;

              LastDisplayNo := LastDisplayNo + 1;
              FAddItem^.DisplayNo := LastDisplayNo;
            end
            else }
            begin
              FAddItem^.ProgramNo := FCurrItem^.ProgramNo;
              FAddItem^.GroupNo   := FCurrItem^.GroupNo;
              FAddItem^.DisplayNo := FCurrItem^.DisplayNo;
            end;
          end;
        end;
      end;

    //        GroupNo
    //        LastCount := LastCount + 1;

//      with TfrmChannel(Owner) do
//        LastSerialNo := LastSerialNo + 1;
    end;
    EM_UPDATE:
    begin
      SaveItem := New(PCueSheetItem);
      try
        Move(FCurrItem^, SaveItem^, SizeOf(TCueSheetItem));

        FAddItem := FCurrItem;

        if (FAddItem^.EventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
        begin
          FAddItem^.StartMode := TStartMode(cbStartMode.Items.Objects[cbStartMode.ItemIndex]);
          FAddItem^.StartTime.D := dpStartDate.Date;
          FAddItem^.StartTime.T := StringToTimecode(meStartTime.Text);
        end
        else
        begin
          FAddItem^.StartTime.D := 0;
          FAddItem^.StartTime.T := 0;
        end;

        StrPCopy(FAddItem^.Title, edTitle.Text);

        if (FAddItem^.EventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
        begin
          StrPCopy(FAddItem^.SubTitle, edSubTitle.Text);

          FAddItem^.Input  := TInputType(cbInput.Items.Objects[cbInput.ItemIndex]);
          FAddItem^.Output := Byte(cbOutput.Items.Objects[cbOutput.ItemIndex]);

          StrPCopy(FAddItem^.Source, cbSource.Text);

          if (cbSourceLayer.ItemIndex < 0) then
            FAddItem^.SourceLayer := 0
          else
            FAddItem^.SourceLayer := Integer(cbSourceLayer.Items.Objects[cbSourceLayer.ItemIndex]);

          StrPCopy(FAddItem^.MediaId, edMediaId.Text);

          FAddItem^.DurationTC := StringToTimecode(meDurationTC.Text);
          FAddItem^.InTC       := StringToTimecode(meInTC.Text);
          FAddItem^.OutTC      := StringToTimecode(meOutTC.Text);

          FAddItem^.TransitionType := TTRType(cbTransitionType.ItemsEx[cbTransitionType.ItemIndex].Data^);
          FAddItem^.TransitionRate := TTRRate(cbTransitionRate.Items.Objects[cbTransitionRate.ItemIndex]);

          FAddItem^.FinishAction := TFinishAction(cbFinishAction.Items.Objects[cbFinishAction.ItemIndex]);

          FAddItem^.ProgramType := Byte(cbProgramType.Items.Objects[cbProgramType.ItemIndex]);
          StrPCopy(FAddItem^.Notes, edNotes.Text);
        end;

        // Validation checking
        if (FAddItem^.EventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
        begin
          with TfrmChannel(Owner) do
          begin
            ErrorControl := nil;
            if (not IsValidStartDate(FAddItem, FAddItem^.StartTime.D, False)) then
            begin
              Result := False;
              ErrorControl := dpStartDate;
            end
            else if (not IsValidStartTime(FAddItem, FAddItem^.StartTime.T, False)) then
            begin
              Result := False;
              ErrorControl := meStartTime;
            end
            else if (not IsValidDuration(FAddItem, FAddItem^.DurationTC, False)) then
            begin
              Result := False;
              ErrorControl := meDurationTC;
            end
            else if (not IsValidInTC(FAddItem, FAddItem^.InTC, False)) then
            begin
              Result := False;
              ErrorControl := meInTC;
            end
            else if (not IsValidOutTC(FAddItem, FAddItem^.OutTC, False)) then
            begin
              Result := False;
              ErrorControl := meOutTC;
            end
            else
            begin
              Source := GetSourceByName(String(FAddItem^.Source));
              if (Source <> nil) and
                 (Source^.SourceType in [ST_VSDEC, ST_CG]) then
              begin
                if (not IsValidMediaId(edMediaId.Text, False)) then
                begin
                  Result := False;
                  ErrorControl := edMediaId;
                end;
              end;

              if (not CheckCueSheetEditPossibleLockTimeByIndex(FSelectIndex)) then
              begin
                MessageBeep(MB_ICONWARNING);
                MessageBox(Handle, PChar(SENotPossibleEditLockTime), PChar(Application.Title), MB_OK or MB_ICONWARNING or MB_TOPMOST);
                Result := False;
                exit;
              end
              else if (not CheckCueSheetEditPossibleLocationByIndex(FSelectIndex)) then
              begin
                MessageBeep(MB_ICONWARNING);
                MessageBox(Handle, PChar(SENotPossibleEditBeforeLocationIncorrect), PChar(Application.Title), MB_OK or MB_ICONWARNING or MB_TOPMOST);
                Result := False;
                exit;
              end;
            end;
          end;

          if (not Result) then
          begin
            SetFocus;
            if (ErrorControl.CanFocus) then ErrorControl.SetFocus;
            Move(SaveItem^, FAddItem^, SizeOf(TCueSheetItem));
            exit;
          end;
        end;
      finally
        Dispose(SaveItem);
      end;
    end;
  end;
end;

procedure TfrmEditEvent.SetEditMode(Value: TEditMode);
begin
  FEditMode := Value;
  GetPropertyValues;
end;

procedure TfrmEditEvent.ClearTransitionType;
var
  I: Integer;
begin
  inherited;

  for I := cbTransitionType.ItemsEx.Count - 1 downto 0 do
  begin
    if (cbTransitionType.ItemsEx[I].Data <> nil) then
      FreeMemory(cbTransitionType.ItemsEx[I].Data);
  end;


  cbTransitionType.ItemsEx.Clear;
end;

procedure TfrmEditEvent.PopulateEventMode;
var
  I: TEventMode;
begin
  cbEventMode.Items.Clear;
  for I := EM_PROGRAM to EM_COMMENT do
  begin
    cbEventMode.Items.AddObject(EventModeNames[I], TObject(I));
  end;
end;

procedure TfrmEditEvent.PopulateStartMode(AEventMode: TEventMode);
var
  I: TStartMode;
begin
  cbStartMode.Items.Clear;
  case AEventMode of
    EM_MAIN:
    begin
      for I := SM_ABSOLUTE to SM_LOOP do
      begin
        cbStartMode.Items.AddObject(StartModeNames[I], TObject(I));
      end;
    end;
    EM_JOIN:
    begin
      for I := SM_SUBBEGIN to SM_SUBBEGIN do
      begin
        cbStartMode.Items.AddObject(StartModeNames[I], TObject(I));
      end;
    end;
    EM_SUB:
    begin
      for I := SM_SUBBEGIN to SM_SUBEND do
      begin
        cbStartMode.Items.AddObject(StartModeNames[I], TObject(I));
      end;
    end;
  end;
end;

procedure TfrmEditEvent.PopulateInput;
var
  I: TInputType;
begin
  cbInput.Items.Clear;
  for I:= IT_NONE to IT_AMIXER2 do
  begin
    cbInput.Items.AddObject(InputTypeNames[I], TObject(I));
  end;
end;

procedure TfrmEditEvent.PopulateOutput(AInput: TInputType);
var
  I: TOutputBkgndType;
  J: TOutputKeyerType;
begin
  cbOutput.Items.Clear;
  if (AInput in [IT_MAIN, IT_BACKUP]) then
  begin
    for I := OB_NONE to OB_BOTH do
    begin
      cbOutput.Items.AddObject(OutputBkgndTypeNames[I], TObject(I));
    end;
  end
  else
  begin
    for J := OK_NONE to OK_AD do
    begin
      cbOutput.Items.AddObject(OutputKeyerTypeNames[J], TObject(J));
    end;
  end;
end;

procedure TfrmEditEvent.PopulateSource;
var
  I: Integer;
begin
  cbSource.Items.Clear;
  for I := 0 to GV_SourceList.Count - 1 do
  begin
    if (not (GV_SourceList[I]^.SourceType in [ST_ROUTER, ST_MCS])) then
      cbSource.Items.AddObject(GV_SourceList[I]^.Name, TObject(GV_SourceList[I]));
  end;
end;

procedure TfrmEditEvent.PopulateSourceLayer(ASource: PSource);
var
  I: Integer;
begin
  cbSourceLayer.Items.Clear;

  cbSourceLayer.Items.AddObject(Format('%d', [0]), TObject(0));
  cbSourceLayer.ItemIndex := 0;

  cbSourceLayer.Visible := (ASource <> nil) and (ASource^.SourceType = ST_CG) and (ASource^.NumLayer > 0);
  lblSourceLayer.Visible := cbSourceLayer.Visible;

  if (ASource = nil) then
  begin
    exit;
  end;

  if (ASource^.SourceType = ST_CG) then
  begin
    for I := 1 to ASource^.NumLayer - 1 do
    begin
      cbSourceLayer.Items.AddObject(Format('%d', [I]), TObject(I));
    end;
  end;
end;

procedure TfrmEditEvent.PopulateTransitionType;
var
  I: TTRType;
  Data: Pointer;
//  ExItem: TComboExItem;
begin
  inherited;

  ClearTransitionType;
  for I := TT_CUT to TT_MIX do
  begin
    Data := GetMemory(SizeOf(Integer));
    PInteger(Data)^ := Integer(I);

    cbTransitionType.ItemsEx.AddItem(TRTypeNames[I], Integer(I), Integer(I), Integer(I), 0, Data);
//    ExItem := cbTransitionType.ItemsEx.Add;
//    ExItem.Caption := TransitionTypeNames[I];
//    ExItem.ImageIndex := Integer(I);

//    cbTransitionType.Items.AddObject(TransitionTypeNames[I], TObject(I));
  end;
end;

procedure TfrmEditEvent.PopulateTransitionRate;
var
  I: TTRRate;
begin
  inherited;

  cbTransitionRate.Items.Clear;
  for I := TR_CUT to TR_SLOW do
  begin
    cbTransitionRate.Items.AddObject(TRRateNames[I], TObject(I));
  end;
end;

procedure TfrmEditEvent.PopulateFinishAction;
var
  I: TFinishAction;
begin
  inherited;

  cbFinishAction.Items.Clear;
  for I := FA_NONE to FA_EJECT do
  begin
    cbFinishAction.Items.AddObject(FinishActionNames[I], TObject(I));
  end;
end;

procedure TfrmEditEvent.PopulateProgramType;
var
  I: Integer;
begin
  cbProgramType.Items.Clear;

  for I := 0 to GV_ProgramTypeList.Count - 1 do
  begin
    cbProgramType.Items.AddObject(GV_ProgramTypeList[I]^.Name, TObject(GV_ProgramTypeList[I]^.Code));
  end;
end;

procedure TfrmEditEvent.CalcDurTC(ADurTC: TTimecode);
begin
  meDurationTC.Text := TimecodeToString(ADurTC);
  meOutTC.Text := TimecodeToString(GetPlusTimecode(StringToTimecode(meInTC.Text), FrameToTimeCode(TimecodeToFrame(ADurTC) - 1)));
end;

procedure TfrmEditEvent.CalcInTC(AInTC: TTimecode);
begin
  meInTC.Text := TimecodeToString(AInTC);
  meOutTC.Text := TimecodeToString(GetPlusTimecode(AInTC, FrameToTimeCode(TimecodeToFrame(StringToTimecode(meDurationTC.Text) - 1))));
end;

procedure TfrmEditEvent.CalcOutTC(AOutTC: TTimecode);
begin
  meOutTC.Text := TimecodeToString(AOutTC);
  meDurationTC.Text := TimecodeToString(GetDurTimecode(StringToTimecode(meInTC.Text), FrameToTimeCode(TimecodeToFrame(AOutTC) + 1)));
end;

procedure TfrmEditEvent.actCancelExecute(Sender: TObject);
begin
  inherited;
  Close;
end;

procedure TfrmEditEvent.actOKExecute(Sender: TObject);
begin
  inherited;
  ModalResult := mrNone;

  if (not CheckPropertyValues) then exit;
  if (not SetPropertyValues) then exit;

  ModalResult := mrOK;
end;

procedure TfrmEditEvent.actSearchMediaExecute(Sender: TObject);
var
  Source: PSource;
  FileName: String;
begin
  inherited;

  with TfrmChannel(Owner) do
  begin
    Source := GetSourceByName(cbSource.Text);
    if (Source = nil) then exit;

    if (Source^.SourceType in [ST_VSDEC, ST_CG]) and (Source^.UsePCS) then
    begin
      FileName := edMediaId.Text;
      if (OpenDialogMediaFile(NAM_COL_CUESHEET_MEDIA_ID, FileName)) then
      begin
        edMediaId.Text := FileName;
        FMediaDurTC := GetMediaDuration(FileName);
        CalcDurTC(FMediaDurTC);
      end;
    end;
  end;
end;

procedure TfrmEditEvent.cbInputSelect(Sender: TObject);
var
  Input: TInputType;
begin
  inherited;
  Input := TInputType(cbInput.Items.Objects[cbInput.ItemIndex]);
  if ((FInput in [IT_MAIN, IT_BACKUP]) and (not (Input in [IT_MAIN, IT_BACKUP]))) or
     (not (FInput in [IT_MAIN, IT_BACKUP]) and (Input in [IT_MAIN, IT_BACKUP])) then
  begin
    PopulateOutput(Input);

    if (Input in [IT_MAIN, IT_BACKUP]) then
      cbOutput.ItemIndex := cbOutput.Items.IndexOfObject(TObject(GV_SettingEventColumnDefault.OutputBkgnd))
    else
      cbOutput.ItemIndex := cbOutput.Items.IndexOfObject(TObject(GV_SettingEventColumnDefault.OutputKeyer));
  end;

  FInput := Input;
  if (FInput = Input) then
  begin
    meDurationTC.Text := INIT_TIMECODE;
    meInTC.Text       := INIT_TIMECODE;
    meOutTC.Text      := INIT_TIMECODE;
  end;

  meDurationTC.Enabled := cbStartMode.Enabled and (FInput <> IT_NONE);
  meInTC.Enabled       := meDurationTC.Enabled;
  meOutTC.Enabled      := meDurationTC.Enabled;
end;

procedure TfrmEditEvent.cbSourceSelect(Sender: TObject);
var
  Source: PSource;
begin
  inherited;

  if (FSaveSourceIndex <> cbSource.ItemIndex) then
  begin
    Source := GetSourceByName(cbSource.Text);

    PopulateSourceLayer(Source);
  end;

  FSaveSourceIndex := cbSource.ItemIndex;
end;

procedure TfrmEditEvent.cbStartModeExit(Sender: TObject);
begin
  inherited;
  with (Sender as TComboBox) do
  begin
    if (Items.IndexOf(Trim(Text)) < 0) then
      Text := Items[ItemIndex];
  end;
end;

procedure TfrmEditEvent.cbStartModeSelect(Sender: TObject);
var
  StartMode: TStartMode;
  StartTime: TEventTime;
begin
  inherited;
  StartMode := TStartMode(cbStartMode.Items.Objects[cbStartMode.ItemIndex]);

  dpStartDate.Enabled := (FEventMode in [EM_MAIN]) and (StartMode <> SM_AUTOFOLLOW);
  meStartTime.Enabled := (FEventMode in [EM_MAIN, EM_SUB]) and (StartMode <> SM_AUTOFOLLOW);

  if (FCurrItem <> nil) then
  begin
    if (FEditMode = EM_INSERT) then
      // 현재 선택된 이벤트에서 삽입 할 경우 이전 메인 이벤트의 End Time으로 설정
      StartTime := GetEventEndTime(FCurrItem^.StartTime, FCurrItem^.DurationTC)
    else
      // 수정인 경우 현재 이벤트의 시작시각으로 설정
      StartTime := FCurrItem^.StartTime;
    
    dpStartDate.Date := StartTime.D;
    meStartTime.Text := TimecodeToString(StartTime.T);
  end
  else
  begin
    // 선택된 이벤트가 없을 시 시각 변경하지 않음

{    // 선택된 이벤트가 없을 시 현재 시각으로 설정
    dpStartDate.Date := Date;
    meStartTime.Text := INIT_TIMECODE; }
  end;
end;

procedure TfrmEditEvent.cbTransitionTypeExit(Sender: TObject);
begin
  inherited;
  with (Sender as TComboBoxEx) do
  begin
    if (Items.IndexOf(Trim(Text)) < 0) then
      Text := Items[ItemIndex];
  end;
end;

procedure TfrmEditEvent.FormCreate(Sender: TObject);
begin
  inherited;

  Initialize;
  GetPropertyValues;

//  DateTimePicker1.Color := clLime;
//  DateTimePicker1.CalColors.BackColor := clRed;

//  CMEdit1.StyleElements := [seFont, seClient, seBorder];
end;

procedure TfrmEditEvent.FormDestroy(Sender: TObject);
begin
  inherited;

  ClearTransitionType;
end;

procedure TfrmEditEvent.FormShow(Sender: TObject);
begin
  inherited;

  case FEventMode of
    EM_PROGRAM:
      edTitle.SetFocus;
    EM_MAIN:
      cbStartMode.SetFocus;
    EM_JOIN:
      cbInput.SetFocus;
    EM_SUB:
      cbStartMode.SetFocus;
    EM_COMMENT:
      edTitle.SetFocus;
  end;
end;

end.
