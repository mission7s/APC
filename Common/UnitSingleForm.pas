unit UnitSingleForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, VCl.Forms, Winapi.MultiMon,
  Vcl.Dialogs, WMTools, WMControls, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Imaging.pngimage;

type
  TfrmSingle = class(TForm)
    WMPanel: TWMPanel;
    WMTitleBar: TWMTitleBar;
    pnlDesc: TWMPanel;
    WMIBClose: TWMImageSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure WMIBCloseClick(Sender: TObject);
    procedure WMTitleBarSystemMenuClick(Sender: TObject;
      DoubleClicked: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSingle: TfrmSingle;

implementation

{$R *.dfm}

procedure TfrmSingle.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmSingle.FormShow(Sender: TObject);
begin
  WMPanel.SetFocus;
end;

procedure TfrmSingle.WMIBCloseClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmSingle.WMTitleBarSystemMenuClick(Sender: TObject;
  DoubleClicked: Boolean);
begin
  if DoubleClicked then Close;
end;

end.
