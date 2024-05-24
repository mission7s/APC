unit UnitTypeConvert;

interface

uses Winapi.Windows, System.SysUtils;

type
  TTypeConvWord = packed record
     case vType: Word of
        0:(vtWord: Word);
        1:(c2, c1: AnsiChar);
  end;

  TTypeConvInt64 = record
     case vType: Int64 of
        0:(c8, c7, c6, c5, c4, c3, c2, c1: AnsiChar);
        1:(vtInt64: Int64);
  end;

  TTypeConvDouble = record
     case vType: Int64 of
        0:(c8, c7, c6, c5, c4, c3, c2, c1: AnsiChar);
        1:(vtDouble: Double);
  end;

  TTypeConvInt = record
     case vType: Integer of
        0:(vtSingle: single);
        1:(c4, c3, c2, c1: AnsiChar);
        2:(vtInteger: Integer);
        3:(vtDWord: DWord);
  end;

  TTypeConvSmallInt = record
    case vType: SmallInt of
      0:(vtSmallInt: SmallInt);
      1:(c2, c1: AnsiChar);
      2:(vtWord: Word);
  end;

  TTypeConvSingle = record
  	case vType: Integer of
	    0:(c4, c3, c2, c1: AnsiChar);
     	1:(vtSingle: Single);
  end;

  TTypeConvBoolean = record
  	case vType: Integer of
	    0:(c1: AnsiChar);
     	1:(vtBoolean: Boolean);
  end;

  function PAnsiCharToWord(AValue: PAnsiChar): Word;
  function PAnsiCharToWordR(AValue: PAnsiChar): Word;
  function PAnsiCharToDWord(AValue: PAnsiChar): DWord;
  function PAnsiCharToDWordR(AValue: PAnsiChar): DWord;
  function PAnsiCharToSmallInt(AValue: PAnsiChar): SmallInt;
  function PAnsiCharToSmallIntR(AValue: PAnsiChar): SmallInt;
  function PAnsiCharToInt(AValue: PAnsiChar): Integer;
  function PAnsiCharToInt2(AValue: PAnsiChar): Integer;
  function PAnsiCharToInt3(AValue: PAnsiChar): Integer;
  function PAnsiCharToInt6(AValue: PAnsiChar): Integer;
  function PAnsiCharToIntR(AValue: PAnsiChar): Integer;
  function PAnsiCharToIntR2(AValue: PAnsiChar): Integer;
  function PAnsiCharToIntR3(AValue: PAnsiChar): Integer;
  function PAnsiCharToIntR6(AValue: PAnsiChar): Integer;
  function PAnsiCharToInt64(AValue: PAnsiChar): Int64;
  function PAnsiCharToSingle(AValue: PAnsiChar): Single;
  function PAnsiCharToDouble(AValue: PAnsiChar): double;
  function PAnsiCharToBool(AValue: PAnsiChar): Boolean;
  function PAnsiCharToHexCode(AValue: PAnsiChar; ADataSize: Integer): AnsiString;
  function PAnsiCharToString(AValue: PAnsiChar): AnsiString;

  function WordToAnsiString(AValue: Word): AnsiString;
  function WordToAnsiStringR(AValue: Word): AnsiString;
  function WordToAnsiStringWithCRC(AValue: SmallInt; var ACrc: Integer): AnsiString;
  function DWordToAnsiString(AValue: DWord): AnsiString;
  function DWordToAnsiStringR(AValue: DWord): AnsiString;
  function DWordToAnsiStringWithCrc(AValue: DWord; var ACrc: Integer): AnsiString;
  function SmallIntToAnsiString(AValue: SmallInt): AnsiString;
  function SmallIntToAnsiStringR(AValue: SmallInt): AnsiString;
  function IntToAnsiString(AValue: Integer): AnsiString;
  function IntToAnsiString2(AValue: Integer): AnsiString;
  function IntToAnsiString3(AValue: Integer): AnsiString;
  function IntToAnsiString6(AValue: Integer): AnsiString;
  function IntToAnsiStringR(AValue: Integer): AnsiString;
  function IntToAnsiStringR2(AValue: Integer): AnsiString;
  function IntToAnsiStringR3(AValue: Integer): AnsiString;
  function IntToAnsiStringR6(AValue: Integer): AnsiString;
  function Int64ToAnsiString(AValue: Int64): AnsiString;
  function SingleToAnsiString(AValue: Single): AnsiString;
  function DoubleToAnsiString(AValue: Double): AnsiString;
  function BoolToAnsiString(AValue: Boolean): AnsiString;

  function IntToBCD(AValue: Integer): Integer;
  function BCDToInt(AValue: Integer): Integer;

implementation

function PAnsiCharToWord(AValue: PAnsiChar): Word;
var
  vtType: TTypeConvWord;
begin
  vtType.C2 := AValue[0];
  vtType.C1 := AValue[1];
  Result := vtType.vtWord;
end;

function PAnsiCharToWordR(AValue: PAnsiChar): Word;
var
  vtType : TTypeConvWord;
begin
  vtType.C1 := AValue[0];
  vtType.C2 := AValue[1];
  Result := vtType.vtWord;
end;

function PAnsiCharToDWord(AValue: PAnsiChar): DWord;
var
  vtType: TTypeConvInt;
begin
  vtType.C4 := AValue[0];
  vtType.C3 := AValue[1];
  vtType.C2 := AValue[2];
  vtType.C1 := AValue[3];
  Result := vtType.vtDWord;
end;

function PAnsiCharToDWordR(AValue: PAnsiChar): DWord;
var
  vtType : TTypeConvInt;
begin
  vtType.C1 := AValue[0];
  vtType.C2 := AValue[1];
  vtType.C3 := AValue[2];
  vtType.C4 := AValue[3];
  Result := vtType.vtDWord;
end;

function PAnsiCharToSmallInt(AValue: PAnsiChar): SmallInt;
var
  vtType: TTypeConvSmallInt;
begin
  vtType.C2 := AValue[0];
  vtType.C1 := AValue[1];
  Result := vtType.vtSmallInt;
end;

function PAnsiCharToSmallIntR(AValue: PAnsiChar): SmallInt;
var
  vtType: TTypeConvSmallInt;
begin
  vtType.C1 := AValue[0];
  vtType.C2 := AValue[1];
  Result := vtType.vtSmallInt;
end;

function PAnsiCharToInt(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C4 := AValue[0];
  vtType.C3 := AValue[1];
  vtType.C2 := AValue[2];
  vtType.C1 := AValue[3];
  Result := vtType.vtInteger;
end;

function PAnsiCharToInt2(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C4 := AValue[0];
  vtType.C3 := AValue[1];
  Result := vtType.vtInteger;
end;

function PAnsiCharToInt3(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C4 := AValue[0];
  vtType.C3 := AValue[1];
  vtType.C2 := AValue[2];
  Result := vtType.vtInteger;
end;

function PAnsiCharToInt6(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt64;
begin
  vtType.C8 := AValue[0];
  vtType.C7 := AValue[1];
  vtType.C6 := AValue[2];
  vtType.C5 := AValue[3];
  vtType.C6 := AValue[4];
  vtType.C3 := AValue[5];
  Result := vtType.vtInt64;
end;

function PAnsiCharToIntR(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C1 := AValue[0];
  vtType.C2 := AValue[1];
  vtType.C3 := AValue[2];
  vtType.C4 := AValue[3];
  Result := vtType.vtInteger;
end;

function PAnsiCharToIntR2(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C3 := AValue[0];
  vtType.C4 := AValue[1];
  Result := vtType.vtInteger;
end;

function PAnsiCharToIntR3(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt;
begin
  vtType.C2 := AValue[0];
  vtType.C3 := AValue[1];
  vtType.C4 := AValue[2];
  Result := vtType.vtInteger;
end;

function PAnsiCharToIntR6(AValue: PAnsiChar): Integer;
var
  vtType: TTypeConvInt64;
begin
  vtType.C3 := AValue[0];
  vtType.C4 := AValue[1];
  vtType.C5 := AValue[2];
  vtType.C6 := AValue[3];
  vtType.C7 := AValue[4];
  vtType.C8 := AValue[5];
  Result := vtType.vtInt64;
end;

function PAnsiCharToInt64(AValue: PAnsiChar): Int64;
var
  vtType: TTypeConvInt64;
begin
  vtType.C8 := AValue[0];
  vtType.C7 := AValue[1];
  vtType.C6 := AValue[2];
  vtType.C5 := AValue[3];
  vtType.C4 := AValue[4];
  vtType.C3 := AValue[5];
  vtType.C2 := AValue[6];
  vtType.C1 := AValue[7];
  Result := vtType.vtInt64;
end;

function PAnsiCharToSingle(AValue: PAnsiChar): Single;
var
	vtType: TTypeConvSingle;
begin
	vtType.c4 := AValue[0];
	vtType.c3 := AValue[1];
  vtType.c2 := AValue[2];
  vtType.c1 := AValue[3];

  Result := vtType.vtSingle;
end;

function PAnsiCharToDouble(AValue: PAnsiChar): Double;
var
  vtType: TTypeConvDouble;
begin
  vtType.C8 := AValue[0];
  vtType.C7 := AValue[1];
  vtType.C6 := AValue[2];
  vtType.C5 := AValue[3];
  vtType.C4 := AValue[4];
  vtType.C3 := AValue[5];
  vtType.C2 := AValue[6];
  vtType.C1 := AValue[7];
  Result := vtType.vtDouble;
end;

function PAnsiCharToBool(AValue: PAnsiChar): Boolean;
var
  vtType: TTypeConvBoolean;
begin
  vtType.C1 := AValue[0];
  Result := vtType.vtBoolean;
end;

function PAnsiCharToHexCode(AValue: PAnsiChar; ADataSize: Integer): AnsiString;
var
  I: Integer;
  V: array[0..4] of AnsiChar;
  A: PAnsiChar;
begin
//  V := AnsiStrAlloc(ADataSize);
//  StrCopy(V, AValue);
//  Move(AValue^, V^, ADataSize);
//  for I := 0 to ADataSize - 1 do
//  begin
//    Result := Result + ansi FormatFloat('00', Ord(AValue[I]));
//  end;
////  Result := '12345645y398573295732050935402hgfkjdsfjldfjlkasfnkladsf';
//  StrDispose(V);
//  exit;

  for I := 0 to ADataSize - 1 do
  begin
//    A := System.AnsiStrings.StrFmt(V, '0X%0.2x ', [Ord(AValue[I])]);
    Result := Result + {AnsiString(A);//System.AnsiStrings.}Format('0X%0.2x ', [Ord(AValue[I])]);
  end;
end;

function PAnsiCharToString(AValue: PAnsiChar): AnsiString;
begin
  SetLength(Result, StrLen(AValue) + 1);
  CopyMemory(PAnsiChar(Result), AValue, StrLen(AValue) + 1);
end;

function WordToAnsiString(AValue: Word): AnsiString;
var
  vtType: TTypeConvWord;
begin
  vtType.vtWord := AValue;
  Result := vtType.c2 + vtType.c1;
end;

function WordToAnsiStringR(AValue: Word): AnsiString;
var
  vtType: TTypeConvWord;
begin
  vtType.vtWord := AValue;
  Result := vtType.c1 + vtType.c2;
end;

function WordToAnsiStringWithCRC(AValue: SmallInt; var ACrc: Integer): AnsiString;
var
  vtType: TTypeConvWord;
begin
  vtType.vtWord := AValue;
  ACrc := ACrc + Ord(vtType.C2) + Ord(VtType.C1);
  Result := vtType.c2 + vtType.c1;
end;

function DWordToAnsiString(AValue: DWord): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtDWord := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function DWordToAnsiStringR(AValue: DWord): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtDWord := AValue;
  Result := vtType.c1 + vtType.c2 + vtType.c3 + vtType.c4;
end;

function DWordToAnsiStringWithCrc(AValue: DWord; var ACrc: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtDWord := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
  ACrc := Ord(vtType.c4) + Ord(vtType.c3) + Ord(vtType.c2) + Ord(vtType.c1);
end;

function SmallIntToAnsiString(AValue: SmallInt): AnsiString;
var
  vtType: TTypeConvSmallInt;
begin
  vtType.vtSmallInt := AValue;
  Result := vtType.c2 + vtType.c1;
end;

function SmallIntToAnsiStringR(AValue: SmallInt): AnsiString;
var
  vtType: TTypeConvSmallInt;
begin
  vtType.vtSmallInt := AValue;
  Result := vtType.c1 + vtType.c2;
end;

function IntToAnsiString(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function IntToAnsiString2(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c4 + vtType.c3;
end;

function IntToAnsiString3(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2;
end;

function IntToAnsiString6(AValue: Integer): AnsiString;
var
	vtType: TTypeConvInt64;
begin
	vtType.vtInt64 := AValue;
  Result := vtType.c8 + vtType.c7 + vtType.c6 + vtType.c5 +
   			    vtType.c4 + vtType.c3;
end;

function IntToAnsiStringR(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c1 + vtType.c2 + vtType.c3 + vtType.c4;
end;

function IntToAnsiStringR2(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c3 + vtType.c4;
end;

function IntToAnsiStringR3(AValue: Integer): AnsiString;
var
  vtType: TTypeConvInt;
begin
  vtType.vtInteger := AValue;
  Result := vtType.c2 + vtType.c3 + vtType.c4;
end;

function IntToAnsiStringR6(AValue: Integer): AnsiString;
var
	vtType: TTypeConvInt64;
begin
	vtType.vtInt64 := AValue;
  Result := vtType.c3 + vtType.c4 + vtType.c5 + vtType.c6 +
   			    vtType.c7 + vtType.c8;
end;

function Int64ToAnsiString(AValue: Int64): AnsiString;
var
	vtType: TTypeConvInt64;
begin
	vtType.vtInt64 := AValue;
  Result := vtType.c8 + vtType.c7 + vtType.c6 + vtType.c5 +
   			    vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function SingleToAnsiString(AValue: Single): AnsiString;
var
	vtType: TTypeConvSingle;
begin
	vtType.vtSingle := AValue;
  Result := vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function DoubleToAnsiString(AValue: Double): AnsiString;
var
	vtType: TTypeConvDouble;
begin
	vtType.vtDouble := AValue;
  Result := vtType.c8 + vtType.c7 + vtType.c6 + vtType.c5 +
   			    vtType.c4 + vtType.c3 + vtType.c2 + vtType.c1;
end;

function BoolToAnsiString(AValue: Boolean): AnsiString;
var
	vtType: TTypeConvBoolean;
begin
	vtType.vtBoolean := AValue;
  Result := vtType.c1;
end;

function IntToBCD(AValue: Integer): Integer;
begin
  Result := AValue + 6 * (AValue div 10);
end;

function BCDToInt(AValue: Integer): Integer;
begin
  Result := AValue - 6 * (AValue div 16);
end;

end.
