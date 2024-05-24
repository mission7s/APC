unit UnitCISApcXmlConverter;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Generics.Collections, System.Generics.Defaults, System.Types, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, WMControls, Xml.VerySimple,
  UnitCommons, UnitCISXmlConsts, UnitConsts, Vcl.StdCtrls, AdvOfficePager,
  AdvOfficePagerStylers, Vcl.Menus;

type
  TProcessFileInfo = record
    FileName: String;
    FileSize: Int64;
    FileTime: int64;
  end;
  PProcessFileInfo = ^TProcessFileInfo;
  TProcessFileList = TList<PProcessFileInfo>;

  TCueSheetFileInfo = record
    ChannelId: Word;
    OnairDate: array[0..DATE_LEN] of Char;
    OnairNo: Integer;
  end;
  PCueSheetFileInfo = ^TCueSheetFileInfo;
  TCueSheetFileList = TList<PCueSheetFileInfo>;

  TProcessThread = class;

  TfrmCISApcXmlConverter = class(TForm)
    wtiConverter: TWMTrayIcon;
    Label1: TLabel;
    mmXMLProceededLog: TMemo;
    Label2: TLabel;
    mmXMLProceededList: TMemo;
    pmTray: TPopupMenu;
    Show1: TMenuItem;
    N1: TMenuItem;
    Close1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Show1Click(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
    FProcessThread: TProcessThread;
    FIsTerminate: Boolean;
  protected
    procedure WndProc(var Message: TMessage); override;
  public
    { Public declarations }
    procedure Initialize;
    procedure Finalize;
  end;

  TProcessThread = class(TThread)
  private
    FForm: TfrmCISApcXmlConverter;

    FProcessDate: TDate;

    FCueSheetFileList: TCueSheetFileList;

    FSourceFileName: String;
//    FChannelID: Word;
    FChannelCueSheet: PChannelCueSheet;

    FLastDisplayNo: Integer;
    FLastSerialNo: Integer;
    FLastProgramNo: Integer;
    FLastGroupNo: Integer;
    FLastCount: Integer;

    function GetTargetXMLFileName(AFileName: String): String;

    function AddCueSheetFile(AChannel: Word; AOnAirDate: String): PCueSheetFileInfo;
    function GetCueSheetFileByOnAirDate(AChannel: Word; AOnAirDate: String): PCueSheetFileInfo;

    procedure DeleteCueSheetFileByOnAirDate(AOnAirDate: String);
    procedure ClearCueSheetFileList;

    procedure DoXMLConvert;
    procedure DoProcess;
  protected
    procedure Execute; override;
  public
    constructor Create(AForm: TfrmCISApcXmlConverter);
    destructor Destroy; override;

    procedure ClearCueSheetList;

    procedure WriteProceededLog(ALogStr: String);
    procedure WriteProceededList(AFileName: String);

    function LoadXMLPlayList(AFileName: String): Boolean;
    function SaveXMLPlayList(AFileName: String): Boolean;
  end;

var
  frmCISApcXmlConverter: TfrmCISApcXmlConverter;

function CompareFileTime(A, B: Pointer): Integer;

implementation

uses System.DateUtils;

{$R *.dfm}

procedure TfrmCISApcXmlConverter.WndProc(var Message: TMessage);
var
  LogStr: String;
begin
  inherited;

  case Message.Msg of
    WM_ADD_PROCEEDED_LOG:
    begin
      LogStr := String(Message.wParam);
      mmXMLProceededLog.Lines.Add(LogStr);
    end;
    WM_ADD_PROCEEDED_LIST:
    begin
      LogStr := String(Message.wParam);
      mmXMLProceededList.Lines.Add(LogStr);
    end;
  end;
end;

procedure TfrmCISApcXmlConverter.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  CanClose := FIsTerminate;
  if not FIsTerminate then
  begin
    Hide;
    ShowWindow(Application.Handle, SW_HIDE);
  end;
end;

procedure TfrmCISApcXmlConverter.FormCreate(Sender: TObject);
begin
  Initialize;
end;

procedure TfrmCISApcXmlConverter.FormDestroy(Sender: TObject);
begin
  Finalize;
end;

procedure TfrmCISApcXmlConverter.Initialize;
begin
  FIsTerminate := False;

  GV_SendPathList := TSendPathList.Create;
  GV_ConvertChannelList := TConvertChannelList.Create;
  GV_DeviceMapList := TDeviceMapList.Create;

  LoadConfig;

  SetEventDeadlineHour(30);

  FProcessThread := TProcessThread.Create(Self);
  FProcessThread.Start;
end;

procedure TfrmCISApcXmlConverter.Show1Click(Sender: TObject);
begin
  if IsIconic(Application.Handle) then ShowWindow(Application.Handle, SW_RESTORE);
  SetForeGroundWindow(Application.Handle);
  Show;
end;

procedure TfrmCISApcXmlConverter.Close1Click(Sender: TObject);
begin
  FIsTerminate := True;
  Close;
end;

procedure TfrmCISApcXmlConverter.Finalize;
begin
  FProcessThread.Terminate;
  FProcessThread.WaitFor;
  FreeAndNil(FProcessThread);

  ClearDeviceMapList;
  ClearConvertChannelList;
  ClearSendPathList;

  FreeAndNil(GV_DeviceMapList);
  FreeAndNil(GV_ConvertChannelList);
  FreeAndNil(GV_SendPathList);
end;

{ TProcessThread }

constructor TProcessThread.Create(AForm: TfrmCISApcXmlConverter);
begin
  FForm := AForm;

  FProcessDate := Trunc(IncHour(Now, HoursPerDay - EventDeadlineHour));

  FCueSheetFileList := TCueSheetFileList.Create;

  FChannelCueSheet := New(PChannelCueSheet);
  FillChar(FChannelCueSheet^, SizeOf(TChannelCueSheet), #0);
  FChannelCueSheet^.ChannelId := 0;
  FChannelCueSheet^.CueSheetList := TCueSheetList.Create;

  FLastDisplayNo := 0;
  FLastSerialNo  := 0;
  FLastProgramNo := 0;
  FLastGroupNo   := 0;

  FLastCount  := 0;

  FreeOnTerminate := False;
  inherited Create(True);
end;

destructor TProcessThread.Destroy;
begin
  ClearCueSheetFileList;
  FreeAndNil(FCueSheetFileList);

  ClearCueSheetList;
  Dispose(FChannelCueSheet);

  inherited Destroy;
end;

procedure TProcessThread.ClearCueSheetList;
var
  I: Integer;
begin
  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  for I := FChannelCueSheet^.CueSheetList.Count - 1 downto 0 do
    Dispose(FChannelCueSheet^.CueSheetList[I]);

  FChannelCueSheet^.CueSheetList.Clear;
end;

procedure TProcessThread.WriteProceededLog(ALogStr: String);
begin
  ALogStr := Format('[%s] %s', [FormatDateTime('ddddd t', Now), ALogStr]);
  SendMessage(frmCISApcXmlConverter.Handle, WM_ADD_PROCEEDED_LOG, NativeInt(PChar(ALogStr)), 0);
end;

procedure TProcessThread.WriteProceededList(AFileName: String);
begin
  SendMessage(frmCISApcXmlConverter.Handle, WM_ADD_PROCEEDED_LIST, NativeInt(PChar(AFileName)), 0);
end;

function TProcessThread.LoadXMLPlayList(AFileName: String): Boolean;
var
  P: PCueSheetFileInfo;

  Xml: TXmlVerySimple;
  DataNode: TXmlNode;
  EventListNode, EventInfoNode: TXmlNode;
  EventType, EventModeName: String;
  I, J: Integer;
  Item: PCueSheetItem;

  LogStr: String;
begin
  Result := False;

  if (FChannelCueSheet = nil) then exit;
  if (FChannelCueSheet^.CueSheetList = nil) then exit;

  ClearCueSheetList;

  EventType := '';

  Xml := TXmlVerySimple.Create;
  try
    Xml.LoadFromFile(AFileName);

    LogStr := Format('Load start xml path = %s, file = %s', [ExtractFilePath(AFileName), ExtractFileName(AFileName)]);
    WriteProceededLog(LogStr);

//    DataNode := Xml.DocumentElement.Find('data');
//    if (DataNode <> nil) then
    begin
      DataNode := Xml.DocumentElement.Find(CIS_XML_SOURCE_ID);
      if (DataNode <> nil) then
        FChannelCueSheet^.ChannelId := StrToIntDef(DataNode.NodeValue, 0)
      else
      begin
        WriteProceededLog('Not exist ''channel'' node.');
        exit;
      end;

      if (GetConvertChannelByChannel(FChannelCueSheet^.ChannelId) = nil) then
      begin
        WriteProceededLog(Format('Not allowed convert channel. Channel Id = %d', [FChannelCueSheet^.ChannelId]));
        exit;
      end;

      DataNode := Xml.DocumentElement.Find(CIS_XML_SCHEDULE_DATE);
      if (DataNode <> nil) then
        StrPCopy(FChannelCueSheet^.OnairDate, DataNode.NodeValue)
      else
      begin
        WriteProceededLog('Not exist ''onairdate'' node.');
        exit;
      end;

      // Compare Onair Date
      if (OnAirDateToDate(FChannelCueSheet^.OnairDate) < FProcessDate) then
      begin
        WriteProceededLog(Format('Not allowed convert onair date. onairdate = %s, processdate = %s', [FChannelCueSheet^.OnairDate, FormatDateTime('ddddd t', FProcessDate)]));
        exit;
      end;

//      if (OnAirDateToDate(FChannelCueSheet^.OnairDate) <> Date) and
//         (OnAirDateToDate(FChannelCueSheet^.OnairDate) <> IncDay(Date)) then Continue;

      FChannelCueSheet^.OnairFlag := FT_REGULAR;

      P := GetCueSheetFileByOnAirDate(FChannelCueSheet^.ChannelId, String(FChannelCueSheet^.OnairDate));
      if (P <> nil) then
      begin
        P^.OnairNo := P^.OnairNo + 1;
      end
      else
        P := AddCueSheetFile(FChannelCueSheet^.ChannelId, String(FChannelCueSheet^.OnairDate));

      FChannelCueSheet^.OnairNo := P^.OnairNo;

      EventListNode := Xml.DocumentElement.Find(CIS_XML_EVENT_LIST);
      if (EventListNode <> nil) then
      begin

        FChannelCueSheet^.EventCount := EventListNode.ChildNodes.Count;

        LogStr := Format('Channel ID = %d, onairdate = %s, eventcount = %d', [FChannelCueSheet^.ChannelId, FChannelCueSheet^.OnairDate, FChannelCueSheet^.EventCount]);
        WriteProceededLog(LogStr);

        for I := 0 to EventListNode.ChildNodes.Count - 1 do
        begin
          EventInfoNode := EventListNode.ChildNodes[I];

          Item := New(PCueSheetItem);
          FillChar(Item^, SizeOf(TCueSheetItem), #0);
          with Item^ do
          begin
            EventID.ChannelID := FChannelCueSheet^.ChannelId;//FChannelID;
            StrCopy(EventID.OnAirDate, FChannelCueSheet^.OnairDate);
            EventID.OnAirFlag := FChannelCueSheet^.OnairFlag;
            EventID.OnAirNo   := FChannelCueSheet^.OnairNo;

            // Serial No
            DataNode := EventInfoNode.Find(CIS_XML_EVENT_NUMBER);
            // Play list를 처음 open 했을 경우에만 큐시트의 SerialNo로 저장
            if (DataNode <> nil) and (FLastCount <= 0) then
              FLastSerialNo := StrToIntDef(DataNode.NodeValue, 0)
            else
              FLastSerialNo := FLastSerialNo + 1;

            EventID.SerialNo := FLastSerialNo;

            // Event Type
            DataNode := EventInfoNode.Find(CIS_XML_EVENT_TYPE);
            EventMode := EM_COMMENT;
            if (DataNode <> nil) then
            begin
              if (UpperCase(DataNode.NodeValue) = 'T') or (UpperCase(DataNode.NodeValue) = 'U') then
                EventMode := EM_MAIN
              else if (UpperCase(DataNode.NodeValue) = 'S') then
                EventMode := EM_SUB;
            end;

            if (EventType <> '') and (EventMode = EM_MAIN) then
            begin
              Inc(FLastProgramNo);
              Inc(FLastGroupNo);
            end;

            if (DataNode <> nil) then
              EventType := DataNode.NodeValue;

            // Program No
            // Play list를 처음 open 했을 경우에만 큐시트의 ProgramNo로 저장
            if (FLastCount <= 0) then
            begin
              ProgramNo := FLastProgramNo;
            end
            else
            begin
              ProgramNo := ProgramNo + FLastProgramNo;
            end;

            // Group No
            // Play list를 처음 open 했을 경우에만 큐시트의 GroupNo로 저장
            if (FLastCount <= 0) then
            begin
              GroupNo := FLastGroupNo;
            end
            else
            begin
              GroupNo := GroupNo + FLastGroupNo;
            end;

            // Display No
            if (EventMode = EM_MAIN) then
              Inc(FLastDisplayNo);

            DisplayNo := FLastDisplayNo;

            // Start Mode
            DataNode := EventInfoNode.Find(CIS_XML_TIME_TYPE);
            StartMode := SM_AUTOFOLLOW;
            if (DataNode <> nil) then
            begin
              if (DataNode.NodeValue = '') then
              begin
                // 추후 적용 예정
{                if (UpperCase(EventType) = 'U') then
                  StartMode := SM_MANUAL
                else if (UpperCase(EventType) = 'T') then
                  StartMode := SM_ABSOLUTE
                else }
                  StartMode := SM_ABSOLUTE;
              end
//              else if (UpperCase(DataNode.NodeValue) = 'U') then
//                StartMode := SM_AUTOFOLLOW // 임시  SM_MANUAL
              else if (UpperCase(DataNode.NodeValue) = 'BEGIN') then
                StartMode := SM_SUBBEGIN
              else if (UpperCase(DataNode.NodeValue) = 'END') then
                StartMode := SM_SUBEND
              else
                StartMode := SM_AUTOFOLLOW;
            end;

            // Start Time
            DataNode := EventInfoNode.Find(CIS_XML_START_TIME);
            if (DataNode <> nil) then
              StartTime := DateTimecodeStrToEventTime(String(EventID.OnAirDate) + ' ' + DataNode.NodeValue);

            DataNode := EventInfoNode.Find(CIS_XML_EVENT_MODE);
            EventModeName := '';
            Input := IT_MAIN;
            Output := Integer(OB_BOTH);
            if (DataNode <> nil) then
            begin
              if (UpperCase(DataNode.NodeValue) = 'VA') then
              begin
                Input := IT_MAIN;
                Output := Integer(OB_BOTH);
              end
              else if (UpperCase(DataNode.NodeValue) = 'BAK') then
              begin
                EventMode := EM_JOIN;
                Input := IT_BACKUP;
                Output := Integer(OB_BOTH);
              end
              else if (UpperCase(DataNode.NodeValue) = 'KEY') then
              begin
                Input := IT_KEYER1;
                Output := Integer(OK_ON);
              end;

              EventModeName := DataNode.NodeValue;
            end;

            // Title
            DataNode := EventInfoNode.Find(CIS_XML_TITLE);
            if (DataNode <> nil) then
              StrPCopy(Title, StringReplace(DataNode.NodeValue, #10, #32, []));

            // Sub Title
            DataNode := EventInfoNode.Find(CIS_XML_MEMO);
            if (DataNode <> nil) then
              StrPCopy(SubTitle, StringReplace(DataNode.NodeValue, #10, #32, []));

            // Source
            DataNode := EventInfoNode.Find(CIS_XML_EVENT_SOURCE);
            if (DataNode <> nil) then
              StrPCopy(Source, GetDeviceNameByChannelEvent(FChannelCueSheet^.ChannelId, EventModeName, StrToIntDef(DataNode.NodeValue, 0)));

            if (Input in [IT_MAIN, IT_BACKUP]) then
            begin
              // Media Id
              DataNode := EventInfoNode.Find(CIS_XML_CLIP_ID);
              if (DataNode <> nil) then
              begin
                StrPCopy(MediaId, DataNode.NodeValue);

                if (Input = IT_BACKUP) then
                begin
                  if (UpperCase(Copy(String(MediaId), 1, 3)) = '3ST') or
                     (UpperCase(Copy(String(MediaId), 1, 3)) = '4ST') then
                  begin
                    Dispose(Item);
                    continue;
                  end;

{                  if (UpperCase(Copy(String(MediaId), 1, 3)) = '4ST') then
                  begin
                    case FChannelCueSheet^.ChannelId of
                      1: StrCopy(Source, 'EBS1_FS1');
                      15: StrCopy(Source, 'EBS2_FS1');
                    end;
                  end
                  else if (UpperCase(Copy(String(MediaId), 1, 3)) = '3ST') then
                  begin
                    case FChannelCueSheet^.ChannelId of
                      1: StrCopy(Source, 'EBS1_FS3');
                      15: StrCopy(Source, 'EBS2_FS3');
                    end;
                  end; }
                end;
              end;
            end
            else if (Input in [IT_KEYER1..IT_KEYER4]) then
            begin
              // Media Id
              DataNode := EventInfoNode.Find(CIS_XML_LOGO_ID);
              if (DataNode <> nil) then
                StrPCopy(MediaId, StringReplace(DataNode.NodeValue, 'PLS2', 'PLUS2', [rfReplaceAll]) + '.tcf');
            end;

            // Duration
            DataNode := EventInfoNode.Find(CIS_XML_DURATION);
            if (DataNode <> nil) then
              DurationTC := StringToTimecode(DataNode.NodeValue);

            // In TC
            DataNode := EventInfoNode.Find(CIS_XML_SOM);
            if (DataNode <> nil) then
              InTC := StringToTimecode(DataNode.NodeValue);

            // Out TC
            OutTC := GetMinusTimecode(GetPlusTimecode(InTC, DurationTC), StringToTimecode('00:00:00:01'));

            // Video Type
            VideoType := VT_HD;

            // Audio Type
            AudioType := AT_STEREO;

            // Closed Caption
            DataNode := EventInfoNode.Find(CIS_XML_CAPTION);
            ClosedCaption := CC_NONE;
            if (DataNode <> nil) then
            begin
              if (UpperCase(DataNode.NodeValue) = 'Y') then
                ClosedCaption := CC_EXIST;
            end;

            // Voice Add
            DataNode := EventInfoNode.Find(CIS_XML_VOICE_ADD);
            VoiceAdd := VA_KOREA_SCREEN;
            if (DataNode <> nil) then
            begin
              if (UpperCase(DataNode.NodeValue) = 'N') then
                VoiceAdd := VA_NONE;
            end;

            // Transition Type
            // Ttansition Rate
            if (Input in [IT_MAIN, IT_BACKUP]) then
            begin
              TransitionType := TT_CUT;
              TransitionRate := TR_FAST;
            end
            else
            begin
              TransitionType := TT_CUT;
              TransitionRate := TR_SLOW;
            end;

            // Finish Action
            FinishAction := FA_NONE;
            DataNode := EventInfoNode.Find(CIS_XML_CONTROL);
            if (DataNode <> nil) and (Input in [IT_KEYER1..IT_KEYER4]) then
            begin
              if (UpperCase(DataNode.NodeValue) = 'STP') then
                FinishAction := FA_STOP
              else if (UpperCase(DataNode.NodeValue) = 'EJT') then
                FinishAction := FA_EJECT;
            end;

            // Program Type
            // 추후 event class로 분류
            ProgramType := 0;

            // Notes
            StrPCopy(Notes, '');
          end;

          FChannelCueSheet^.CueSheetList.Add(Item);
        end;
      end;
    end;

    Result := True;

    LogStr := Format('Load succeeded xml path = %s, file = %s', [ExtractFilePath(AFileName), ExtractFileName(AFileName)]);
    WriteProceededLog(LogStr);
  finally
    FreeAndNil(Xml);
  end;

  FLastCount     := FChannelCueSheet^.CueSheetList.Count;

{  if (FLastCount > 0) then
  begin
    FLastProgramNo := FChannelCueSheet^.CueSheetList[FLastCount - 1]^.ProgramNo;
    FLastGroupNo   := FChannelCueSheet^.CueSheetList[FLastCount - 1]^.GroupNo;
  end
  else  }
  begin
    FLastProgramNo := 0;
    FLastGroupNo   := 0;
  end;
end;

function TProcessThread.SaveXMLPlayList(AFileName: String): Boolean;
var
  Xml: TXmlVerySimple;
  DocType: TXmlNode;
  DataNode: TXmlNode;
  EventsNode, EventNode: TXmlNode;
  I: Integer;
  Item: PCueSheetItem;

  LogStr: String;
begin
  Result := False;

  LogStr := Format('Save start xml file = %s', [ExtractFileName(AFileName)]);
  WriteProceededLog(LogStr);

  // Create a XML document first, and save it
  Xml := TXmlVerySimple.Create;
  try
    DocType := Xml.AddChild('DocType', ntDocType);
    DocType.Text := 'data [' + #13#10 +
                    '<!ELEMENT data (channelid, onairdate, onairflag, onairno, events)>' + #13#10 +
                    '<!ELEMENT channelid (#PCDATA)>' + #13#10 +
                    '<!ELEMENT onairdate (#PCDATA)>' + #13#10 +
                    '<!ELEMENT onairflag (#PCDATA)>' + #13#10 +
                    '<!ELEMENT onairno (#PCDATA)>' + #13#10 +
                    '<!ELEMENT eventcount (#PCDATA)>' + #13#10 +
                    '<!ELEMENT events (event)>' + #13#10 +
                    '<!ELEMENT event (#PCDATA)>' + #13#10 +
                    ']';

    DataNode := Xml.AddChild('data');
    DataNode.Addchild(XML_CHANNEL_ID).Text := Format('%d', [FChannelCueSheet^.ChannelId]);
    DataNode.Addchild(XML_ONAIR_DATE).Text := String(FChannelCueSheet^.OnairDate);
    DataNode.Addchild(XML_ONAIR_FLAG).Text := Char(FChannelCueSheet^.OnairFlag);
    DataNode.Addchild(XML_ONAIR_NO).Text := Format('%d', [FChannelCueSheet^.OnairNo]);
    DataNode.Addchild(XML_EVENT_COUNT).Text := Format('%d', [FChannelCueSheet^.CueSheetList.Count]);

    EventsNode := DataNode.Addchild(XML_EVENTS);
    for I := 0 to FChannelCueSheet^.CueSheetList.Count - 1 do
    begin
      Item := FChannelCueSheet^.CueSheetList[I];
      EventNode := EventsNode.AddChild(XML_EVENT);
      EventNode.SetAttribute(XML_ATTR_SERIAL_NO, Format('%d', [I]));
      EventNode.SetAttribute(XML_ATTR_PROGRAM_NO, Format('%d', [Item^.ProgramNo]));
      EventNode.SetAttribute(XML_ATTR_GROUP_NO, Format('%d', [Item^.GroupNo]));
      EventNode.SetAttribute(XML_ATTR_EVENT_MODE, Format('%d', [Ord(Item^.EventMode)]));

      if (Item^.EventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
      begin
        EventNode.SetAttribute(XML_ATTR_START_MODE, Format('%d', [Ord(Item^.StartMode)]));
        EventNode.SetAttribute(XML_ATTR_START_TIME, EventTimeToString(Item^.StartTime));
        EventNode.SetAttribute(XML_ATTR_INPUT, Format('%d', [Ord(Item^.Input)]));
        EventNode.SetAttribute(XML_ATTR_OUTPUT, Format('%d', [Ord(Item^.Output)]));
      end;

      EventNode.SetAttribute(XML_ATTR_TITLE, String(Item^.Title));

      if (Item^.EventMode <> EM_COMMENT) then
      begin
        EventNode.SetAttribute(XML_ATTR_SUB_TITLE, String(Item^.SubTitle));
      end;

      if (Item^.EventMode in [EM_MAIN, EM_JOIN, EM_SUB]) then
      begin
        EventNode.SetAttribute(XML_ATTR_SOURCE, String(Item^.Source));
        EventNode.SetAttribute(XML_ATTR_MEDIA_ID, String(Item^.MediaId));
        EventNode.SetAttribute(XML_ATTR_DURATION_TC, TimecodeToString(Item^.DurationTC));
        EventNode.SetAttribute(XML_ATTR_IN_TC, TimecodeToString(Item^.InTC));
        EventNode.SetAttribute(XML_ATTR_OUT_TC, TimecodeToString(Item^.OutTC));
        EventNode.SetAttribute(XML_ATTR_VIDEO_TYPE, Format('%d', [Ord(Item^.VideoType)]));
        EventNode.SetAttribute(XML_ATTR_AUDIO_TYPE, Format('%d', [Ord(Item^.AudioType)]));
        EventNode.SetAttribute(XML_ATTR_CLOSED_CAPTION, Format('%d', [Ord(Item^.ClosedCaption)]));
        EventNode.SetAttribute(XML_ATTR_VOICE_ADD, Format('%d', [Ord(Item^.VoiceAdd)]));
        EventNode.SetAttribute(XML_ATTR_TRANSITION_TYPE, Format('%d', [Ord(Item^.TransitionType)]));
        EventNode.SetAttribute(XML_ATTR_TRANSITION_RATE, Format('%d', [Ord(Item^.TransitionRate)]));
        EventNode.SetAttribute(XML_ATTR_FINISH_ACTION, Format('%d', [Ord(Item^.FinishAction)]));
        EventNode.SetAttribute(XML_ATTR_PROGRAM_TYPE, Format('%d', [Ord(Item^.ProgramType)]));
        EventNode.SetAttribute(XML_ATTR_NOTES, String(Item^.Notes));
      end;
    end;

    Xml.SaveToFile(AFileName);
    Result := True;

    LogStr := Format('Save succeed xml file = %s', [ExtractFileName(AFileName)]);
    WriteProceededLog(LogStr);

    WriteProceededList(ExtractFileName(AFileName));
  finally
    FreeAndNil(Xml);
  end;
end;

function TProcessThread.GetTargetXMLFileName(AFileName: String): String;
var
  I: Integer;
  FilePrefix: String;
begin
  Result := AFileName;

  I := 0;
  FilePrefix := StringReplace(Result, '.xml', '', [rfReplaceAll]);
  while FileExists(Result) do
  begin
    Result := Format('%s (%d)%s', [FilePrefix, I + 1, '.xml']);
    Inc(I);
  end;
end;

function TProcessThread.AddCueSheetFile(AChannel: Word; AOnAirDate: String): PCueSheetFileInfo;
begin
  Result := New(PCueSheetFileInfo);
  FillChar(Result^, SizeOf(TCueSheetFileInfo), #0);
  with Result^ do
  begin
    ChannelId := AChannel;
    StrPCopy(OnairDate, AOnAirDate);
    OnairNo   := 0;
  end;

  FCueSheetFileList.Add(Result);
end;

function TProcessThread.GetCueSheetFileByOnAirDate(AChannel: Word; AOnAirDate: String): PCueSheetFileInfo;
var
  I: Integer;
  P: PCueSheetFileInfo;
begin
  Result := nil;

  for I := 0 to FCueSheetFileList.Count - 1 do
  begin
    P := FCueSheetFileList[I];
    if (P^.ChannelId = AChannel) and (String(P^.OnairDate) = AOnAirDate) then
    begin
      Result := P;
      break;
    end;
  end;
end;

procedure TProcessThread.DeleteCueSheetFileByOnAirDate(AOnAirDate: String);
var
  I: Integer;
  P: PCueSheetFileInfo;
begin
  for I := FCueSheetFileList.Count - 1 downto 0 do
  begin
    P := FCueSheetFileList[I];
    if (String(P^.OnairDate) < AOnAirDate) then
    begin
      Dispose(P);
      FCueSheetFileList.Remove(P);
    end;
  end;
end;

procedure TProcessThread.ClearCueSheetFileList;
var
  I: Integer;
begin
  for I := FCueSheetFileList.Count - 1 downto 0 do
  begin
    Dispose(FCueSheetFileList[I]);
  end;
  FCueSheetFileList.Clear;
end;

procedure TProcessThread.DoXMLConvert;
var
  SourceStream: TFileStream;

  I: Integer;
  SendPath: PSendPath;
  TargetFileName: String;
begin
  try
    SourceStream := TFileStream.Create(FSourceFileName, fmOpenWrite or fmShareExclusive);
    try
    finally
      SourceStream.Free;
    end;

    if (LoadXMLPlayList(FSourceFileName)) then
    begin
      for I := 0 to GV_SendPathList.Count - 1 do
      begin
        SendPath := GV_SendPathList[I];
        if (SendPath <> nil) then
        begin
          with SendPath^ do
            if (not MakeNetConnection(FilePath, UserId, Password)) then continue;

          TargetFileName := String(SendPath^.FilePath) + ExtractFileName(FSourceFileName);
          TargetFileName := GetTargetXMLFileName(TargetFileName);

//          ShowMessage(Format('TargetFileName, %d, %s', [I, TargetFileName]));

          if (not SaveXMLPlayList(TargetFileName)) then
          begin
            with GV_FailPath do
              if (not MakeNetConnection(FilePath, UserId, Password)) then continue;

            if (FileExists(String(GV_FailPath.FilePath) + ExtractFileName(FSourceFileName))) then
              DeleteFile(String(GV_FailPath.FilePath) + ExtractFileName(FSourceFileName));

            MoveFile(PChar(FSourceFileName), PChar(String(GV_FailPath.FilePath) + ExtractFileName(FSourceFileName)));
            exit;
          end;
        end;
      end;
    end
    else
    begin
      with GV_FailPath do
        if (not MakeNetConnection(FilePath, UserId, Password)) then exit;

      if (FileExists(String(GV_FailPath.FilePath) + ExtractFileName(FSourceFileName))) then
        DeleteFile(String(GV_FailPath.FilePath) + ExtractFileName(FSourceFileName));

      MoveFile(PChar(FSourceFileName), PChar(String(GV_FailPath.FilePath) + ExtractFileName(FSourceFileName)));
      exit;
    end;
  except
    with GV_FailPath do
      if (not MakeNetConnection(FilePath, UserId, Password)) then exit;

    MoveFile(PChar(FSourceFileName), PChar(String(GV_FailPath.FilePath) + ExtractFileName(FSourceFileName)));
    exit;
  end;

  with GV_SuccessPath do
    if (not MakeNetConnection(FilePath, UserId, Password)) then exit;

  if (FileExists(String(GV_SuccessPath.FilePath) + ExtractFileName(FSourceFileName))) then
    DeleteFile(String(GV_SuccessPath.FilePath) + ExtractFileName(FSourceFileName));

  MoveFile(PChar(FSourceFileName), PChar(String(GV_SuccessPath.FilePath) + ExtractFileName(FSourceFileName)));
end;

function CompareFileTime(A, B: Pointer): Integer;
begin
  Result := PProcessFileInfo(A)^.FileTime - PProcessFileInfo(B)^.FileTime;
end;

procedure TProcessThread.DoProcess;
var
  CheckDate: TDate;

  P: PProcessFileInfo;
  ProcessFileList: TProcessFileList;
  Comparison: TComparison<PProcessFileInfo>;

  SR: TSearchRec;
  FileAttrs: Integer;

  I: Integer;
  SrcFileName, DstFileName: String;
  TrgMXFFileName: String;
  CopyData: TCopyDataStruct;
begin
  CheckDate := Trunc(IncHour(Now, HoursPerDay - EventDeadlineHour));
  if (FProcessDate <> CheckDate) then
  begin
    DeleteCueSheetFileByOnAirDate(FormatDateTime('YYYYMMDD', FProcessDate));
    FProcessDate := CheckDate;
  end;

  with GV_SourcePath do
    if (not MakeNetConnection(FilePath, UserId, Password)) then exit;

  // XML Process
  FileAttrs := faAnyFile;
  if (FindFirst(String(GV_SourcePath.FilePath) + '*.xml', FileAttrs, SR) = S_OK) then
  begin
    ProcessFileList := TProcessFileList.Create;
    try
      repeat
        P := New(PProcessFileInfo);

        P^.FileName := SR.Name;
        P^.FileSize := SR.Size;
        P^.FileTime := SR.Time;

        ProcessFileList.Add(P);

{  //      if (SR.Attr and FileAttrs) = SR.Attr then
        begin
          FSourceFileName := String(GV_SourcePath.FilePath) + SR.Name;

  //        ShowMessage(Format('FSourceFileName, %s', [FSourceFileName]));
  //        FTargetFileName := GV_TargetPath + SR.Name;
  //        FTargetFileName := GetTargetXMLFileName(FTargetFileName);

          DoXmlConvert;

        end; }
      until (FindNext(SR) <> 0) or (Terminated);
      FindClose(SR);

      ProcessFileList.Sort(
        TComparer<PProcessFileInfo>.Construct(
          function(const A, B: PProcessFileInfo): Integer
          begin
            Result := A^.FileTime - B^.FileTime;
          end
        )
      );

      for I := 0 to ProcessFileList.Count - 1 do
      begin
        P := ProcessFileList[I];

        FSourceFileName := String(GV_SourcePath.FilePath) + P^.FileName;
        DoXmlConvert;
      end;
    finally
      for I := ProcessFileList.Count - 1 downto 0 do
      begin
        P := ProcessFileList[I];
        Dispose(P);
      end;
      ProcessFileList.Clear;
      FreeAndNil(ProcessFileList);
    end;
  end;
end;

procedure TProcessThread.Execute;
begin
  while not Terminated do
  begin
    DoProcess;

    Sleep(GV_ProcessInterval * 1000);
  end;
end;

end.
