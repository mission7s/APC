unit UnitAlarmThread;

interface

uses System.Types, System.Classes, System.SysUtils, Winapi.Windows, Vcl.Forms,
     WMTools, UnitCommons;

type
  TRiffChunk = packed record
    Signature: array[0..3] of AnsiChar;
    ChunkSize: DWORD;  // filesize - 8
    Format:    array[0..3] of AnsiChar;
  end;

  TFmtSubChunk = packed record
    ID:            array[0..3] of AnsiChar;
    Size:          DWORD;
    AudioFmt:      Word;
    Channels:      Word;
    SampleRate:    DWORD;
    ByteRate:      DWORD;
    BlockAlign:    Word;
    BitsperSample: Word;
  end;

  TFmtExtSubChunk = packed record
    Size:          Word;
    BitsperSample: Word;
    ChannelMask: DWORD;    // which channels are present in stream
    SubFormat: TGUID;
  end;

  TDataSubChunk = packed record
    ID:   array[0..3] of AnsiChar;
    Size: DWORD;
  end;

  TAlarmThread = class(TThread)
  private
    FAudioOut: TWMAudioOut;

    FFileName: String;
    FDuration: Cardinal;
    FInterval: Cardinal;

    FReadBuffer: PByte;
    FReadSize: Integer;

    FWaveFileName: String;
    FWaveFileStream: TFileStream;

    FWaveHeaderSize: Integer;
    FWavePosition: Int64;
    FWaveHeaderData: PByte;
    FWaveDataSize: DWORD;

    FFormatTag: Word;       { format type }
    FChannels: Word;        { number of channels (i.e. mono, stereo, etc.) }
    FSamplesPerSec: DWORD;  { sample rate }
    FAvgBytesPerSec: DWORD; { for buffer estimation }
    FBlockAlign: Word;      { block size of data }
    FBitsPerSample: Word;   { number of bits per sample of mono data }

    FLoadFmtSubChunkSize: Word;
    FValidBitsPerSample: Word;
    FChannelMask: DWORD;    // which channels are present in stream
    FSubFormat: TGUID;

    FEventFinished: THandle;

    function AudioOutFillBuffer(Buffer: PAnsiChar; var Size: Integer): Boolean;

    procedure LoadFromFile(AFileName: String);
  protected
    procedure DoAlarm;

    procedure Execute; override;
  public
    constructor Create(AFileName: String; ADuration, AInterval: TTimecode);
    destructor Destroy; override;
  end;

implementation

{ TAlarmThread }

constructor TAlarmThread.Create(AFileName: String; ADuration, AInterval: TTimecode);
begin
  FFileName := AFileName;
  FDuration := TimecodeToMilliSec(ADuration);
  FInterval := TimecodeToMilliSec(AInterval);

  FEventFinished := CreateEvent(nil, True, False, nil);

  LoadFromFile(FFileName);

  FAudioOut := TWMAudioOut.Create(Application.MainForm);
  FAudioOut.CallbackType := acCallback;
  FAudioOut.OnFillBuffer := AudioOutFillBuffer;
  FAudioOut.BufferSize := 16380;
  FAudioOut.BitsPerSample := FBitsPerSample;
  FAudioOut.Channels := FChannels;
  FAudioOut.SamplesPerSec := FSamplesPerSec;
  FAudioOut.Active := True;

  FReadSize := FAudioOut.BufferSize;
  GetMem(FReadBuffer, FReadSize);

  FreeOnTerminate := True;
  inherited Create(True);
end;

destructor TAlarmThread.Destroy;
begin
  FAudioOut.StopAtOnce;
  FAudioOut.Active := False;
  FreeAndNil(FAudioOut);

  if (FWaveFileStream <> nil) then
    FreeAndNil(FWaveFileStream);

  if FReadBuffer <> nil then
    FreeMem(FReadBuffer);

  CloseHandle(FEventFinished);

  inherited Destroy;
end;

procedure TAlarmThread.LoadFromFile(AFileName: String);
type
  TChunk = record
    ID: array[0..3] of AnsiChar;
    Size: DWORD;
  end;
var
  ReadChunk: TChunk;

  RiffChunk: TRiffChunk;
  FmtChunk: TFmtSubChunk;
  FmtExtChunk: TFmtExtSubChunk;

  DataChunk: TDataSubChunk;
  ReadSize: Longint;
begin
  if (FWaveFileStream <> nil) then FreeAndNil(FWaveFileStream);

  FWaveFileStream := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyNone);
  try

  //  loaded:=True;

    // riff chunk
    ReadSize := FWaveFileStream.Read(ReadChunk, SizeOf(ReadChunk));
    if (ReadChunk.ID <> 'RIFF') then exit;

    RiffChunk.ChunkSize := ReadChunk.Size;

    ReadSize := FWaveFileStream.Read(RiffChunk.Format, SizeOf(RiffChunk.Format));
    if (RiffChunk.Format <> 'WAVE') then exit;

    // fmt chunk
    while (FWaveFileStream.Position < FWaveFileStream.Size) do
    begin
      ReadSize := FWaveFileStream.Read(ReadChunk, SizeOf(ReadChunk));
      if (ReadChunk.ID = 'fmt ') then break;

      FWaveFileStream.Seek(ReadChunk.Size, soCurrent);
    end;

    if (ReadChunk.ID <> 'fmt ') then exit;
    FmtChunk.Size := ReadChunk.Size;

    ReadSize := FWaveFileStream.Read(FmtChunk.AudioFmt, SizeOf(FmtChunk.AudioFmt));
    ReadSize := FWaveFileStream.Read(FmtChunk.Channels, SizeOf(FmtChunk.Channels));
    ReadSize := FWaveFileStream.Read(FmtChunk.SampleRate, SizeOf(FmtChunk.SampleRate));
    ReadSize := FWaveFileStream.Read(FmtChunk.ByteRate, SizeOf(FmtChunk.ByteRate));
    ReadSize := FWaveFileStream.Read(FmtChunk.BlockAlign, SizeOf(FmtChunk.BlockAlign));
    ReadSize := FWaveFileStream.Read(FmtChunk.BitsperSample, SizeOf(FmtChunk.BitsperSample));

    FFormatTag           := FmtChunk.AudioFmt;
    FChannels            := FmtChunk.Channels;
    FSamplesPerSec       := FmtChunk.SampleRate;
    FAvgBytesPerSec      := FmtChunk.ByteRate;
    FBlockAlign          := FmtChunk.BlockAlign;
    FBitsPerSample       := FmtChunk.BitsperSample;
    FLoadFmtSubChunkSize := FmtChunk.Size;

    // extra fmt
    if (FmtChunk.Size >= 18) then
    begin
      ReadSize := FWaveFileStream.Read(FmtExtChunk, SizeOf(FmtExtChunk));
      FValidBitsPerSample := FmtExtChunk.BitsperSample;
      FChannelMask := FmtExtChunk.ChannelMask;
      FSubFormat := FmtExtChunk.SubFormat;
    end
    else
    begin
      FValidBitsPerSample := FmtChunk.BitsperSample;
      FChannelMask := 0;
      FSubFormat := GUID_NULL;
    end;

    // data chunk
    while (FWaveFileStream.Position < FWaveFileStream.Size) do
    begin
      ReadSize := FWaveFileStream.Read(ReadChunk, SizeOf(ReadChunk));
      if (ReadChunk.ID = 'data') then break;

      FWaveFileStream.Seek(ReadChunk.Size, soCurrent);
    end;

    if (ReadChunk.ID <> 'data') then exit;

    FWaveDataSize := ReadChunk.Size;

//    FLoadHeaderSize := FLoadFileStream.Position;
    FWaveHeaderSize := FWaveFileStream.Size - ReadChunk.Size;

    GetMem(FWaveHeaderData, FWaveHeaderSize);
    try
      FWaveFileStream.Position := 0;
      ReadSize := FWaveFileStream.Read(FWaveHeaderData^, FWaveHeaderSize);
    finally
      FreeMem(FWaveHeaderData);
    end;



 {   ReadSize := FLoadFileStream.Read(RiffChunk, SizeOf(RiffChunk));
    if (RiffChunk.Signature <> 'RIFF') then exit;
    if (RiffChunk.Format <> 'WAVE') then exit;

    ReadSize := FLoadFileStream.Read(FmtChunk, SizeOf(FmtChunk));
    if (FmtChunk.ID <> 'fmt ') then exit;

    FFormatTag           := FmtChunk.AudioFmt;
    FChannels            := FmtChunk.Channels;
    FSamplesPerSec       := FmtChunk.SampleRate;
    FAvgBytesPerSec      := FmtChunk.ByteRate;
    FBlockAlign          := FmtChunk.BlockAlign;
    FBitsPerSample       := FmtChunk.BitsperSample;
    FLoadFmtSubChunkSize := FmtChunk.Size;

    FLoadHeaderSize := SizeOf(TRiffChunk) +
                       SizeOf(TFmtSubChunk) +
                       SizeOf(TDataSubChunk);

    if (FmtChunk.Size = 40) then
    begin
      ReadSize := FLoadFileStream.Read(FmtExtChunk, SizeOf(FmtExtChunk));
      FValidBitsPerSample := FmtExtChunk.BitsperSample;
      FChannelMask := FmtExtChunk.ChannelMask;
      FSubFormat := FmtExtChunk.SubFormat;

      //ShowMessage(IntToStr(SizeOf(FmtExtChunk)));

      FLoadHeaderSize := FLoadHeaderSize + SizeOf(TFmtExtSubChunk);
    end;

    ReadSize := FLoadFileStream.Read(DataChunk, SizeOf(DataChunk));

    FLoadDataSize := DataChunk.Size; }
  except
    //on Ex: EFOpenError do
  end;
end;

function TAlarmThread.AudioOutFillBuffer(Buffer: PAnsiChar; var Size: Integer): Boolean;
var
  Channels: Integer;
  Samples: Integer;
	UnsignedRange, MaxAvailableSignedValue: Int64;
  BytesPerSample: Integer;
  ReadBuf: PByte;

  S, C, B: Integer;
  X: Int64;
  Pos: Int64;
begin
  inherited;

  if (Terminated) then
  begin
      SetEvent(FEventFinished);
    Result := False;
    exit;
  end;

  { This will happen on an error }
//  if (Size <= 0) or (not FAudioOut.Active) then
  if (Size <= 0) then
  begin
    Result := False;
    exit;
  end;

  { Read in a buffer }
  FillMemory(Buffer, Size, 0);

  try

    if (FWaveFileStream = nil) then
    begin
      Inc(FWavePosition, Size);
      Result := True;
      exit;
    end;

    FillMemory(FReadBuffer, FReadSize, 0);

    Pos := FWaveHeaderSize + FWavePosition;

    FWaveFileStream.Position := Pos;

    Size := FWaveFileStream.Read(FReadBuffer^, FReadSize);

    if (Size = 0) then
    begin
      SetEvent(FEventFinished);
      Result := False;
      exit;
    end;

    Channels := FAudioOut.Channels;
    Samples := Size Div (FAudioOut.WaveFmtEx.Format.nBlockAlign);
    UnsignedRange := Int64(1) shl FAudioOut.ValidBitsperSample;
    MaxAvailableSignedValue := UnsignedRange div 2 - 1;

    BytesPerSample := FAudioOut.ValidBitsperSample div 8;

    ReadBuf := FReadBuffer;
    for S := 0 to Samples - 1 do
    begin
      for C := 0 to Channels - 1 do
      begin
        X := 0;
        if (FAudioOut.ValidBitsperSample = 8) and (FAudioOut.BitsperSample = 8) then
        begin
          X := PByte(ReadBuf)^;
          PByte(Buffer)^ := ReadBuf^;
          Inc(Buffer);
          Inc(ReadBuf);
        end
        else if (FAudioOut.ValidBitsperSample <= 64) and (FAudioOut.BitsperSample <= 64) then
        begin
          for B := 0 to BytesPerSample - 1 do
          begin
            X := X + PByte(ReadBuf)^ shl (B * 8);
            Buffer^ := PAnsiChar(ReadBuf)^;//X shr (B * 8);
            Inc(Buffer);
            Inc(ReadBuf);
          end;
          if (FAudioOut.BitsperSample > FAudioOut.ValidBitsperSample) then
          begin
            // shift padded bits
            X := X shr (FAudioOut.BitsperSample - FAudioOut.ValidBitsperSample);
          end;
        end;

        // restore twos-complement arithmetics
        if (X > MaxAvailableSignedValue) then
          X := UnsignedRange - X;
        X := Abs(X);
      end;
    end;

    Inc(FWavePosition, FReadSize);

    Result := True;
  finally
  end;
end;

procedure TAlarmThread.DoAlarm;
begin
  if (not FAudioOut.Active) then
  begin
    System.SysUtils.Beep;
    exit;
  end;

  ResetEvent(FEventFinished);
  try
    FWavePosition := 0;
    FAudioOut.Start;
  finally
    WaitForSingleObject(FEventFinished, INFINITE);
  end;

  FAudioOut.StopAtOnce;
end;

procedure TAlarmThread.Execute;
var
  WaitInterval, LastProcTime: Int64;
  Frequency, StartCount, StopCount: Int64; // minimal stop watch

  IntervalStart, IntervalStop: Int64;

  SystemTime: TSystemTime;
begin
  QueryPerformanceFrequency(Frequency); // this will never return 0 on Windows XP or later
  if (Frequency = 0) then Frequency := 1;

  LastProcTime := 0;

  SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_TIME_CRITICAL);

  QueryPerformanceCounter(StartCount);
  while not Terminated do
  begin
    DoAlarm;
    QueryPerformanceCounter(StopCount);

    LastProcTime := 1000 * (StopCount - StartCount) div Frequency; // ElapsedMilliSeconds

    if (not Terminated) and (LastProcTime < FDuration) then
    begin
      QueryPerformanceCounter(IntervalStart);
      repeat
        QueryPerformanceCounter(IntervalStop);
  //      Sleep(0);
      until (Terminated or ((1000 * (IntervalStop - IntervalStart) div Frequency) >= FInterval));
    end
    else
      break;

    QueryPerformanceCounter(StopCount);
    LastProcTime := 1000 * (StopCount - StartCount) div Frequency; // ElapsedMilliSeconds

    if (Terminated) or (LastProcTime >= 10000) then break;
  end;

  SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_NORMAL);
end;

end.
