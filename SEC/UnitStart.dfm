inherited frmStart: TfrmStart
  Caption = 'frmStart'
  ClientHeight = 136
  ClientWidth = 372
  FormStyle = fsStayOnTop
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  ExplicitWidth = 372
  ExplicitHeight = 136
  PixelsPerInch = 96
  TextHeight = 17
  inherited WMPanel: TWMPanel
    Width = 372
    Height = 136
    ExplicitWidth = 372
    ExplicitHeight = 108
    inherited WMTitleBar: TWMTitleBar
      Width = 366
      Caption = 'Starting...'
      ExplicitWidth = 366
      inherited WMIBClose: TWMImageSpeedButton
        Left = 333
        ExplicitLeft = 333
      end
    end
    inherited pnlDesc: TWMPanel
      Width = 366
      Height = 72
      ExplicitTop = 64
      ExplicitWidth = 366
      ExplicitHeight = 72
      object lblChecking: TLabel
        Left = 73
        Top = 14
        Width = 280
        Height = 19
        Anchors = [akLeft, akTop, akRight]
        AutoSize = False
        Caption = 'Schedule event controller starting...'
        EllipsisPosition = epEndEllipsis
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clSilver
        Font.Height = -14
        Font.Name = 'Century Gothic'
        Font.Style = []
        ParentFont = False
      end
      object lblStatus: TLabel
        Left = 73
        Top = 36
        Width = 280
        Height = 19
        Anchors = [akLeft, akTop, akRight]
        AutoSize = False
        Caption = 'Status'
        EllipsisPosition = epEndEllipsis
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clSilver
        Font.Height = -14
        Font.Name = 'Century Gothic'
        Font.Style = []
        ParentFont = False
      end
    end
  end
  object aLstStart: TActionList
    Left = 323
    Top = 61
  end
end
