unit UnitStart;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitSingleForm, WMTools, WMControls,
  Vcl.ExtCtrls, Vcl.Imaging.pngimage, Vcl.ComCtrls, Vcl.StdCtrls,
  System.Actions, Vcl.ActnList,
  UnitCommons, UnitConsts;

type
  TfrmStart = class(TfrmSingle)
    lblChecking: TLabel;
    lblStatus: TLabel;
    aLstStart: TActionList;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure DisplayStatus(AStatusStr: String);
  end;

var
  frmStart: TfrmStart;

implementation

{$R *.dfm}

procedure TfrmStart.DisplayStatus(AStatusStr: String);
begin
  lblStatus.Caption := AStatusStr;
  Application.ProcessMessages;
end;

procedure TfrmStart.FormCreate(Sender: TObject);
begin
  inherited;

  lblChecking.Caption := 'Schedule event controller starting...';
  lblStatus.Caption := '';

  WMIBClose.Visible := False;
end;

procedure TfrmStart.FormDestroy(Sender: TObject);
begin
  inherited;
  //
end;

procedure TfrmStart.FormShow(Sender: TObject);
begin
  inherited;
  //
end;

end.
