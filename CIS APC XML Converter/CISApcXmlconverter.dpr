program CISApcXmlconverter;

uses
  Vcl.Forms,
  UnitCISApcXmlConverter in 'UnitCISApcXmlConverter.pas' {frmCISApcXmlConverter},
  UnitCommons in '..\Common\UnitCommons.pas',
  UnitBaseSerial in '..\Common\UnitBaseSerial.pas',
  UnitTypeConvert in '..\Common\UnitTypeConvert.pas',
  UnitConsts in 'UnitConsts.pas',
  UnitCISXmlConsts in 'UnitCISXmlConsts.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfrmCISApcXmlConverter, frmCISApcXmlConverter);
  Application.Run;
end.
