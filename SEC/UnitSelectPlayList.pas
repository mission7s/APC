unit UnitSelectPlayList;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UnitSingleForm, Vcl.ComCtrls,
  Vcl.Imaging.pngimage, WMTools, WMControls, Vcl.ExtCtrls, System.Actions,
  Vcl.ActnList, Vcl.StdCtrls, Vcl.Mask, AdvSpin;

type
  TfrmSelectPlaylist = class(TfrmSingle)
    dpOnairDate: TDateTimePicker;
    Label12: TLabel;
    cbChannel: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    cbOnairFlag: TComboBox;
    Label3: TLabel;
    aseOnairNo: TAdvSpinEdit;
    Label4: TLabel;
    edFileName: TEdit;
    wmibDecrease1Second: TWMImageSpeedButton;
    wmibOK: TWMImageButton;
    wmibCancel: TWMImageButton;
    actSelectPlaylist: TActionList;
    actOK: TAction;
    actCancel: TAction;
    actSelectFileName: TAction;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSelectPlaylist: TfrmSelectPlaylist;

implementation

{$R *.dfm}

end.
