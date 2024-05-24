object frmCISApcXmlConverter: TfrmCISApcXmlConverter
  Left = 0
  Top = 0
  Caption = 'CIS Xml Converter'
  ClientHeight = 523
  ClientWidth = 1384
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Century Gothic'
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  DesignSize = (
    1384
    523)
  PixelsPerInch = 96
  TextHeight = 17
  object Label1: TLabel
    Left = 479
    Top = 11
    Width = 116
    Height = 17
    Caption = 'XML Proceeded log'
  end
  object Label2: TLabel
    Left = 8
    Top = 11
    Width = 113
    Height = 17
    Caption = 'XML Proceeded list'
  end
  object mmXMLProceededLog: TMemo
    Left = 479
    Top = 34
    Width = 897
    Height = 481
    Anchors = [akLeft, akTop, akRight, akBottom]
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 0
    WordWrap = False
  end
  object mmXMLProceededList: TMemo
    Left = 8
    Top = 34
    Width = 465
    Height = 481
    Anchors = [akLeft, akTop, akBottom]
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object wtiConverter: TWMTrayIcon
    Icon.Data = {
      0000010001002020200000000000A81000001600000028000000200000004000
      000001002000000000000010000000000000000000000000000000000000FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF002727E2132727E2632727E2A42727E2D22727E2F32727E2FF2727
      E2FF2727E2F32727E2D22727E2A42727E2632727E213FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF002727
      E22D2727E2A72727E2FA2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FA2727E2A72727E22EFFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF002727E20D2727E2962727
      E2FE2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FE2727
      E2972727E20DFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF002727E22B2727E2DB2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2DB2727E22BFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF002727E2382727E2F12727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2F12727E238FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF002727E22B2727E2F12727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2F12727E22BFFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF002727E20D2727E2DB2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2DB2727E20DFFFFFF00FFFFFF00FFFF
      FF00FFFFFF002727E2962727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E297FFFFFF00FFFFFF00FFFF
      FF002727E22E2727E2FE2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FE2727E22EFFFFFF00FFFF
      FF002727E2A72727E2FF2727E2FF2727E2FF6E8BEDFF6E8BEDFF5977EAFF2727
      E2FF2727E2FF2727E2FF2727E2FF6582ECFF6E8BEDFF6E8BEDFF718DEEFF6E8B
      EDFF6E8BEDFF4661E7FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF3F59
      E6FF6E8BEDFF6E8BEDFF6885ECFF272AE2FF2727E2FF2727E2A7FFFFFF002727
      E2132727E2FA2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFB5C8F7FF2727
      E2FF2727E2FF2727E2FF5471EAFFFFFFFFFFFFFFFFFFD5E1FAFF5F7CEAFFFFFF
      FFFFFFFFFFFFE3EBFCFF3041E4FF2727E2FF2727E2FF2727E2FF2B37E3FFD4E0
      FAFFFFFFFFFFFFFFFFFF6E8BEDFF2727E2FF2727E2FF2727E2FA2727E2132727
      E2632727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFB5C8F7FF2727
      E2FF2727E2FF282DE2FFCBD9FAFFFFFFFFFFFFFFFFFF5673EAFF2727E2FF98B0
      F4FFFFFFFFFFFFFFFFFFA6BCF5FF2727E2FF2727E2FF2727E2FF90AAF3FFFFFF
      FFFFFFFFFFFFB0C3F7FF2728E2FF2727E2FF2727E2FF2727E2FF2727E2632727
      E2A42727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFB5C8F7FF2727
      E2FF2727E2FF6684ECFFFFFFFFFFFFFFFFFFAEC2F7FF2727E2FF2727E2FF2C38
      E4FFD4E0FAFFFFFFFFFFFFFFFFFF6380EBFF2727E2FF5270E9FFFFFFFFFFFFFF
      FFFFE5EDFDFF3245E4FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2A42727
      E2D22727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFB5C8F7FF2727
      E2FF2B37E3FFDFE8FCFFFFFFFFFFF5F8FFFF3A52E6FF2727E2FF2727E2FF2727
      E2FF4763E7FFF8FBFFFFFFFFFFFFEBF1FDFF415CE6FFDCE7FCFFFFFFFFFFFFFF
      FFFF5875EAFF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2D22727
      E2F32727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFF2F6FFFFD6E3
      FCFFDFE8FCFFFFFFFFFFFFFFFFFFC9D8FAFF3B53E6FF2727E2FF2727E2FF2727
      E2FF2727E2FF7C97F0FFFFFFFFFFFFFFFFFFF4F7FFFFFFFFFFFFFFFFFFFF94AD
      F3FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2F32727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF7F9FFFF5A77EAFF2727E2FF2727
      E2FF2727E2FF282DE2FFBDCFF8FFFFFFFFFFFFFFFFFFFFFFFFFFD2DFFAFF2B37
      E3FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFDBE6FCFF88A3
      F2FF88A3F2FF8DA7F2FFC4D5FAFFFFFFFFFFFFFFFFFFE8EFFDFF2A35E3FF2727
      E2FF2727E2FF2727E2FF8DA7F2FFFFFFFFFFFFFFFFFFFFFFFFFFA3BAF5FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2F32727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFB5C8F7FF2727
      E2FF2727E2FF2727E2FF2728E2FFB4C7F7FFFFFFFFFFFFFFFFFF4E6BE8FF2727
      E2FF2727E2FF4E6BE8FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF5D7A
      EAFF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2F32727
      E2D22727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFB5C8F7FF2727
      E2FF2727E2FF2727E2FF2727E2FF8AA4F2FFFFFFFFFFFFFFFFFF5674EAFF2727
      E2FF2E3DE4FFDCE7FCFFFFFFFFFFFFFFFFFF7894EEFFF8FBFFFFFFFFFFFFE8EF
      FDFF3448E4FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2D22727
      E2A42727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFB5C8F7FF2727
      E2FF2727E2FF292FE2FF4D6AE8FFE2EAFCFFFFFFFFFFFFFFFFFF3C55E6FF2727
      E2FF9EB6F5FFFFFFFFFFFFFFFFFF93ACF3FF2727E2FF7F9AF0FFFFFFFFFFFFFF
      FFFFB4C7F7FF272AE2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2A42727
      E2632727E2FF2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFB5C8F7FF2727E2FF5E7C
      EAFFFFFFFFFFFFFFFFFFD9E4FCFF2C39E4FF2727E2FF2931E2FFCBD9FAFFFFFF
      FFFFFFFFFFFF718EEDFF2727E2FF2727E2FF2727E2FF2727E2FF2727E2632727
      E2132727E2FA2727E2FF2727E2FF2727E2FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFB7CAF7FF3143E4FF3549E4FFE8EF
      FDFFFFFFFFFFFFFFFFFF4F6DE8FF2727E2FF2727E2FF2727E2FF4661E7FFFAFC
      FFFFFFFFFFFFF4F7FFFF3F59E6FF2727E2FF2727E2FF2727E2FA2727E213FFFF
      FF002727E2A72727E2FF2727E2FF2727E2FF6E8BEDFF6E8BEDFF6E8BEDFF6E8B
      EDFF6E8BEDFF6E8BEDFF6684ECFF4560E7FF2727E2FF2727E2FF4B67E8FF6E8B
      EDFF6E8BEDFF5D7AEAFF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF5875
      EAFF6E8BEDFF6E8BEDFF526FE9FF2727E2FF2727E2FF2727E2A7FFFFFF00FFFF
      FF002727E22E2727E2FE2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FE2727E22EFFFFFF00FFFF
      FF00FFFFFF002727E2972727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E297FFFFFF00FFFFFF00FFFF
      FF00FFFFFF002727E20D2727E2DB2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2DB2727E20DFFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF002727E22B2727E2F12727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2F12727E22BFFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF002727E2382727E2F12727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2F12727E238FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF002727E22B2727E2DB2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2DB2727E22BFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF002727E20D2727E2972727
      E2FE2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FE2727
      E2972727E20DFFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF002727
      E22E2727E2A72727E2FA2727E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727
      E2FF2727E2FF2727E2FF2727E2FF2727E2FF2727E2FA2727E2A72727E22EFFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFF
      FF00FFFFFF002727E2132727E2632727E2A42727E2D22727E2F32727E2FF2727
      E2FF2727E2F32727E2D22727E2A42727E2632727E213FFFFFF00FFFFFF00FFFF
      FF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFFFFF00FFF0
      0FFFFF8001FFFE00007FFC00003FF800001FF000000FE0000007C0000003C000
      0003800000018000000180000001000000000000000000000000000000000000
      0000000000000000000000000000800000018000000180000001C0000003C000
      0003E0000007F000000FF800001FFC00003FFE00007FFF8001FFFFF00FFF}
    PopupMenu = pmTray
    Left = 896
    Top = 8
  end
  object pmTray: TPopupMenu
    Left = 968
    Top = 8
    object Show1: TMenuItem
      Caption = 'Show'
      OnClick = Show1Click
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object Close1: TMenuItem
      Caption = 'Exit'
      OnClick = Close1Click
    end
  end
end
