unit Unit9;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, UnitCommons;

type
  TForm9 = class(TForm)
    Edit1: TEdit;
    Edit2: TEdit;
    Button1: TButton;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form9: TForm9;

implementation

{$R *.dfm}

procedure TForm9.FormCreate(Sender: TObject);
var
  DurTC: TTimecode;
  PlsTC, MinTC: TTimecode;
begin
  DurTC := GetDurTimecode(StringToTimecode('04:45:57:00'), StringToTimecode('15:40:57:00'));
  ShowMessage(TimecodeToString(DurTC));

  PlsTC := GetPlusTimecode(StringToTimecode('04:45:57:00'), DurTC);
  ShowMessage(TimecodeToString(PlsTC));

  MinTC := GetMinusTimecode(StringToTimecode('00:14:09:00'), DurTC);
  ShowMessage(TimecodeToString(MinTC));
end;

end.
