inherited frmChannel: TfrmChannel
  Caption = 'frmChannel'
  ClientHeight = 869
  ClientWidth = 1400
  Font.Name = 'Tahoma'
  ExplicitWidth = 1400
  ExplicitHeight = 869
  PixelsPerInch = 96
  TextHeight = 14
  inherited WMPanel: TWMPanel
    Width = 1400
    Height = 869
    ExplicitWidth = 1400
    ExplicitHeight = 869
    inherited WMTitleBar: TWMTitleBar
      Width = 1390
      Font.Name = 'Tahoma'
      ExplicitWidth = 1390
      inherited WMIBClose: TWMImageSpeedButton
        Left = 1359
        ExplicitLeft = 1359
      end
      inherited WMIBMinimize: TWMImageSpeedButton
        Left = 1323
        ExplicitLeft = 1323
      end
    end
    inherited pnlDesc: TWMPanel
      Width = 1390
      Height = 801
      ExplicitWidth = 1390
      ExplicitHeight = 801
      object AdvSplitter1: TAdvSplitter
        Left = 1
        Top = 546
        Width = 1388
        Height = 6
        Cursor = crVSplit
        Align = alBottom
        AutoSnap = False
        Color = 2824214
        ParentColor = False
        ResizeStyle = rsUpdate
        Appearance.BorderColor = clNone
        Appearance.BorderColorHot = clNone
        Appearance.Color = 2497307
        Appearance.ColorTo = 2497307
        Appearance.ColorHot = 2497307
        Appearance.ColorHotTo = 2497307
        DblClickAction = dbaOpenClose
        GripStyle = sgNone
        ExplicitTop = 539
        ExplicitWidth = 1077
      end
      object acgPlaylist: TAdvColumnGrid
        Left = 1
        Top = 97
        Width = 1388
        Height = 449
        Cursor = crDefault
        Align = alClient
        BevelWidth = 4
        BorderStyle = bsNone
        Color = 2497307
        ColCount = 12
        Ctl3D = False
        DefaultRowHeight = 26
        DoubleBuffered = True
        DrawingStyle = gdsClassic
        FixedColor = 5060404
        RowCount = 300
        Font.Charset = ANSI_CHARSET
        Font.Color = 12105912
        Font.Height = -14
        Font.Name = 'Century Gothic'
        Font.Style = []
        GradientEndColor = 788230
        GradientStartColor = 788230
        Options = [goFixedVertLine, goFixedHorzLine, goHorzLine, goColSizing, goColMoving, goRowSelect, goThumbTracking, goFixedHotTrack]
        ParentCtl3D = False
        ParentDoubleBuffered = False
        ParentFont = False
        ScrollBars = ssNone
        TabOrder = 0
        StyleElements = [seBorder]
        OnDrawCell = acgPlaylistDrawCell
        OnFixedCellClick = acgPlaylistFixedCellClick
        OnGetEditText = acgPlaylistGetEditText
        OnKeyPress = acgPlaylistKeyPress
        OnSelectCell = acgPlaylistSelectCell
        OnSetEditText = acgPlaylistSetEditText
        ActiveRowColor = 11304545
        GridLineColor = 788230
        GridFixedLineColor = 788230
        HoverRowCells = [hcNormal, hcSelected]
        OnGetDisplText = acgPlaylistGetDisplText
        OnCustomCellDraw = acgPlaylistCustomCellDraw
        OnGetCellColor = acgPlaylistGetCellColor
        OnGetCellBorder = acgPlaylistGetCellBorder
        OnGetCellBorderProp = acgPlaylistGetCellBorderProp
        OnRowUpdate = acgPlaylistRowUpdate
        OnExpandNode = acgPlaylistExpandNode
        OnBeforeExpandNode = acgPlaylistBeforeExpandNode
        OnBeforeContractNode = acgPlaylistBeforeContractNode
        OnClickCell = acgPlaylistClickCell
        OnDblClickCell = acgPlaylistDblClickCell
        OnCanEditCell = acgPlaylistCanEditCell
        OnEllipsClick = acgPlaylistEllipsClick
        OnCheckBoxClick = acgPlaylistCheckBoxClick
        OnComboChange = acgPlaylistComboChange
        OnComboDropDown = acgPlaylistComboDropDown
        OnEditingDone = acgPlaylistEditingDone
        OnEditCellDone = acgPlaylistEditCellDone
        OnEditChange = acgPlaylistEditChange
        OnSelectionChanged = acgPlaylistSelectionChanged
        HighlightColor = 8869383
        ActiveCellFont.Charset = DEFAULT_CHARSET
        ActiveCellFont.Color = clWindowText
        ActiveCellFont.Height = -11
        ActiveCellFont.Name = 'Century Gothic'
        ActiveCellFont.Style = [fsBold]
        CellNode.NodeType = cnGlyph
        CellNode.NodeColor = 16765315
        CellNode.NodeIndent = 20
        CellNode.ExpandGlyph.Data = {
          82020000424D820200000000000042000000280000000C0000000C0000000100
          20000300000040020000000000000000000000000000000000000000FF0000FF
          0000FF000000FFD1834AFFD183E7FFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183E6FFD18349FFD183E4FFD183FFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183E6FFD183FFFFD183FFFFD183FFFFD183FFFFD183FFCA9E5EFFCA9E
          5EFFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFD9AC68FF845B2EFF845B2EFFD9AD69FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFE6B971FF885F31FF835A2DFF835A
          2DFF885F31FFE6B972FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFF0C3
          79FF8F6535FF835A2DFFB3884EFFB3884EFF835A2DFF8F6535FFF0C379FFFFD1
          83FFFFD183FFFFD183FFF8CA7EFF986E3CFF835A2DFFA57A44FFFDCF81FFFDCF
          81FFA47A44FF835A2DFF986E3CFFF8CA7EFFFFD183FFFFD183FFA87D47FF835A
          2DFF996F3CFFF8CA7EFFFFD183FFFFD183FFF8CA7EFF986F3CFF835A2DFFA87D
          47FFFFD183FFFFD183FFE7BA72FF9F7540FFF0C379FFFFD183FFFFD183FFFFD1
          83FFFFD183FFF0C379FF9E7440FFE7BA72FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183E7FFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183E6FFD1834AFFD183E6FFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83E6FFD18349}
        CellNode.ContractGlyph.Data = {
          82020000424D820200000000000042000000280000000C0000000C0000000100
          20000300000040020000000000000000000000000000000000000000FF0000FF
          0000FF000000FFD1834AFFD183E7FFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183E4FFD1834AFFD183E6FFD183FFFFD1
          83FFE7BA72FFA87D47FFF8CA7EFFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183E7FFD183FFFFD183FFFFD183FF9F7540FF835A2DFF986E3CFFF0C3
          79FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFF0C379FF996F3CFF835A2DFF8F6535FFE6B971FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFF8CA7EFFA57A44FF835A
          2DFF885F31FFD9AC68FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183FFFDCF81FFB3884EFF835A2DFF845B2EFFCA9E5EFFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFDCF81FFB388
          4EFF835A2DFF845B2EFFCA9E5EFFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFF8CA7EFFA47A44FF835A2DFF885F31FFD9AD69FFFFD183FFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFF0C379FF986F3CFF835A2DFF8F65
          35FFE6B972FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83FF9E7440FF835A2DFF986E3CFFF0C379FFFFD183FFFFD183FFFFD183FFFFD1
          83FFFFD183FFFFD183E6FFD183FFFFD183FFE7BA72FFA87D47FFF8CA7EFFFFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183E6FFD18349FFD183E6FFD1
          83FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD183FFFFD1
          83E6FFD18349}
        CellNode.TreeColor = 8608301
        ColumnHeaders.Strings = (
          'NO'
          'Start type'
          'Start Time'
          'Duration'
          'Title'
          'Media'
          'Device'
          'Status'
          'Transition'
          'Finish action'
          'SOM'
          'Description')
        ControlLook.FixedGradientHoverFrom = 13619409
        ControlLook.FixedGradientHoverTo = 12502728
        ControlLook.FixedGradientHoverMirrorFrom = 12502728
        ControlLook.FixedGradientHoverMirrorTo = 11254975
        ControlLook.FixedGradientDownFrom = 8816520
        ControlLook.FixedGradientDownTo = 7568510
        ControlLook.FixedGradientDownMirrorFrom = 7568510
        ControlLook.FixedGradientDownMirrorTo = 6452086
        ControlLook.FixedGradientDownBorder = 14007466
        ControlLook.CheckAlwaysActive = True
        ControlLook.ControlStyle = csFlat
        ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
        ControlLook.DropDownHeader.Font.Color = clWindowText
        ControlLook.DropDownHeader.Font.Height = -11
        ControlLook.DropDownHeader.Font.Name = 'Tahoma'
        ControlLook.DropDownHeader.Font.Style = []
        ControlLook.DropDownHeader.Visible = True
        ControlLook.DropDownHeader.Buttons = <>
        ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
        ControlLook.DropDownFooter.Font.Color = clWindowText
        ControlLook.DropDownFooter.Font.Height = -11
        ControlLook.DropDownFooter.Font.Name = 'Tahoma'
        ControlLook.DropDownFooter.Font.Style = []
        ControlLook.DropDownFooter.Visible = True
        ControlLook.DropDownFooter.Buttons = <>
        DisabledFontColor = 9076349
        EditWithTags = True
        EnhTextSize = True
        Filter = <>
        FilterDropDown.Font.Charset = DEFAULT_CHARSET
        FilterDropDown.Font.Color = clWindowText
        FilterDropDown.Font.Height = -11
        FilterDropDown.Font.Name = 'Tahoma'
        FilterDropDown.Font.Style = []
        FilterDropDown.TextChecked = 'Checked'
        FilterDropDown.TextUnChecked = 'Unchecked'
        FilterDropDownClear = '(All)'
        FilterEdit.TypeNames.Strings = (
          'Starts with'
          'Ends with'
          'Contains'
          'Not contains'
          'Equal'
          'Not equal'
          'Clear')
        FixedColWidth = 36
        FixedRowHeight = 26
        FixedFont.Charset = DEFAULT_CHARSET
        FixedFont.Color = 16765315
        FixedFont.Height = -14
        FixedFont.Name = 'Century Gothic'
        FixedFont.Style = []
        Flat = True
        FloatFormat = '%.2f'
        HideFocusRect = True
        HoverButtons.Buttons = <>
        HoverButtons.Position = hbLeftFromColumnLeft
        HTMLSettings.ImageFolder = 'images'
        HTMLSettings.ImageBaseName = 'img'
        Look = glListView
        MouseActions.DisjunctRowSelect = True
        MouseActions.NodeAllExpandContract = True
        MouseActions.RowSelect = True
        Navigation.KeepHorizScroll = True
        PrintSettings.DateFormat = 'dd/mm/yyyy'
        PrintSettings.Font.Charset = DEFAULT_CHARSET
        PrintSettings.Font.Color = clWindowText
        PrintSettings.Font.Height = -11
        PrintSettings.Font.Name = 'Tahoma'
        PrintSettings.Font.Style = []
        PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
        PrintSettings.FixedFont.Color = clWindowText
        PrintSettings.FixedFont.Height = -11
        PrintSettings.FixedFont.Name = 'Tahoma'
        PrintSettings.FixedFont.Style = []
        PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
        PrintSettings.HeaderFont.Color = clWindowText
        PrintSettings.HeaderFont.Height = -11
        PrintSettings.HeaderFont.Name = 'Tahoma'
        PrintSettings.HeaderFont.Style = []
        PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
        PrintSettings.FooterFont.Color = clWindowText
        PrintSettings.FooterFont.Height = -11
        PrintSettings.FooterFont.Name = 'Tahoma'
        PrintSettings.FooterFont.Style = []
        PrintSettings.PageNumSep = '/'
        ScrollColor = 8608301
        ScrollType = ssMetro
        SearchFooter.Color = clBtnFace
        SearchFooter.FindNextCaption = 'Find &next'
        SearchFooter.FindPrevCaption = 'Find &previous'
        SearchFooter.Font.Charset = DEFAULT_CHARSET
        SearchFooter.Font.Color = clWindowText
        SearchFooter.Font.Height = -11
        SearchFooter.Font.Name = 'Tahoma'
        SearchFooter.Font.Style = []
        SearchFooter.HighLightCaption = 'Highlight'
        SearchFooter.HintClose = 'Close'
        SearchFooter.HintFindNext = 'Find next occurrence'
        SearchFooter.HintFindPrev = 'Find previous occurrence'
        SearchFooter.HintHighlight = 'Highlight occurrences'
        SearchFooter.MatchCaseCaption = 'Match case'
        SearchFooter.ResultFormat = '(%d of %d)'
        SelectionColor = 3812392
        SelectionTextColor = clHighlightText
        ShowSelection = False
        ShowDesignHelper = False
        SortSettings.DefaultFormat = ssAutomatic
        SortSettings.HeaderColorTo = 16579058
        SortSettings.HeaderMirrorColor = 16380385
        SortSettings.HeaderMirrorColorTo = 16182488
        VAlignment = vtaCenter
        Version = '3.1.6.1'
        WordWrap = False
        Columns = <
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taCenter
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'NO'
            HeaderAlignment = taCenter
            HeaderFont.Charset = DEFAULT_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = DEFAULT_CHARSET
            PrintFont.Color = clWindowText
            PrintFont.Height = -11
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 36
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taCenter
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Start type'
            HeaderAlignment = taCenter
            HeaderFont.Charset = DEFAULT_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = DEFAULT_CHARSET
            PrintFont.Color = clWindowText
            PrintFont.Height = -11
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 68
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taCenter
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Start Time'
            HeaderAlignment = taCenter
            HeaderFont.Charset = DEFAULT_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = DEFAULT_CHARSET
            PrintFont.Color = clWindowText
            PrintFont.Height = -11
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 140
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taCenter
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Duration'
            HeaderAlignment = taCenter
            HeaderFont.Charset = DEFAULT_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = DEFAULT_CHARSET
            PrintFont.Color = clWindowText
            PrintFont.Height = -11
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 90
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taLeftJustify
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = ANSI_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Title'
            HeaderAlignment = taCenter
            HeaderFont.Charset = ANSI_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = ANSI_CHARSET
            PrintFont.Color = clWhite
            PrintFont.Height = -12
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 140
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taLeftJustify
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = ANSI_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Media'
            HeaderAlignment = taCenter
            HeaderFont.Charset = ANSI_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = ANSI_CHARSET
            PrintFont.Color = clWhite
            PrintFont.Height = -12
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 120
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taCenter
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = ANSI_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Device'
            HeaderAlignment = taCenter
            HeaderFont.Charset = ANSI_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = ANSI_CHARSET
            PrintFont.Color = clWhite
            PrintFont.Height = -12
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 68
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taLeftJustify
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = ANSI_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Status'
            HeaderAlignment = taCenter
            HeaderFont.Charset = ANSI_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = ANSI_CHARSET
            PrintFont.Color = clWhite
            PrintFont.Height = -12
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 48
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taCenter
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = ANSI_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Transition'
            HeaderAlignment = taCenter
            HeaderFont.Charset = ANSI_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = ANSI_CHARSET
            PrintFont.Color = clWhite
            PrintFont.Height = -12
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 64
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taLeftJustify
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = ANSI_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Finish action'
            HeaderAlignment = taCenter
            HeaderFont.Charset = ANSI_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = ANSI_CHARSET
            PrintFont.Color = clWhite
            PrintFont.Height = -12
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 88
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taLeftJustify
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = ANSI_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'SOM'
            HeaderAlignment = taCenter
            HeaderFont.Charset = ANSI_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = ANSI_CHARSET
            PrintFont.Color = clWhite
            PrintFont.Height = -12
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 64
          end
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taLeftJustify
            Borders = []
            BorderPen.Color = clSilver
            ButtonHeight = 18
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 2497307
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = ANSI_CHARSET
            Font.Color = 12105912
            Font.Height = -14
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Description'
            HeaderAlignment = taCenter
            HeaderFont.Charset = ANSI_CHARSET
            HeaderFont.Color = 16765315
            HeaderFont.Height = -14
            HeaderFont.Name = 'Century Gothic'
            HeaderFont.Style = []
            MinSize = 0
            MaxSize = 0
            Password = False
            PrintBorders = [cbTop, cbLeft, cbRight, cbBottom]
            PrintColor = clWhite
            PrintFont.Charset = ANSI_CHARSET
            PrintFont.Color = clWhite
            PrintFont.Height = -12
            PrintFont.Name = 'Tahoma'
            PrintFont.Style = []
            ReadOnly = False
            ShowBands = False
            SortStyle = ssAutomatic
            SpinMax = 0
            SpinMin = 0
            SpinStep = 1
            Tag = 0
            Width = 140
          end>
        ColWidths = (
          36
          68
          140
          90
          140
          120
          68
          48
          64
          88
          64
          140)
      end
      object WMPanel2: TWMPanel
        Left = 1
        Top = 1
        Width = 1388
        Height = 96
        ColorHighLight = 2497307
        ColorShadow = 2497307
        Align = alTop
        ParentColor = True
        TabOrder = 1
        DesignSize = (
          1388
          96)
        object lblPlayListFileName: TLabel
          Left = 0
          Top = 71
          Width = 1265
          Height = 19
          AutoSize = False
          Caption = 'D:\CIAB\SEC\CueSheet\Work\P20210601.xml'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = 12105912
          Font.Height = -14
          Font.Name = 'Century Gothic'
          Font.Style = []
          ParentFont = False
        end
        object wmibFreezeOnAir: TWMImageSpeedButton
          Left = 953
          Top = 9
          Width = 60
          Height = 60
          Action = frmSEC.actControlFreezeOnAir
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000006BC4944415478DAED5B8B3B5469187FDD721B
            13A28B3B112995A2C88A4A29934B852EB249BBD5FE07FBEC1FB0CFFE07BBCF6E
            EDB2114229145DB445C925239592A418AA756D5C22C5BEEFA7A33317FB34D6A1
            9DE6F73CE7F19DF99C99F7F77DEFF77B7FDF99397A3E2B03A1B141EA01003FE1
            11868718B40B723CAEE0F1BD87A76FA39E898989A79E9EDE6D1B5B7B4B0BB135
            181818CC7680D38AF7EFDF439FBC1B3A3BDA7AC7C6C6028870AEED7C875D9656
            B6B31D9BA0E8E97E85A4DBCF10E1D76EEE2BC4DA36B3CAA0997EFAA4AE9F088F
            616ECF763C3302D42AD011D666E8086B3B7484B51DD342382828007C967B4345
            4535486BEB0409D4C7671904AD5F07F5F58FA0B4AC1CD031CD0E617FBFD5E0E7
            377E6D4B8B0C0A2F140B4278FBF62DE0E2ECC4DA35D2BB6C70679CB033061081
            811046464620EF5C217476760942D8CACA1276C6EC00636363765E5C7C159E36
            3F9B39C2F4C1FBF7C58189C97800E7F32F425B5BBB2064392C5AB800A2A32580
            1B1D181E1E868CCC5C78F3E6CDCC100E09F90ABC977AB27655550D54DF91AAFC
            8F9191119BF9A960B26B7D7D5742C03A3FD67EF4E8315CFBAB5478C273E78A61
            DFDE5836D23D3DBD703AFB2C8C8E8E4EF4EBEBEBC3D62D9BC0D5D5196EDEBA0D
            75750F340A6AF9B2A5101CBC1E9A9B9FC3A5CB252AEF1D1B1B03F3ACAD987065
            64E6C0EBD7726109077F1508CB519509178B2EC3B3672D93F63F7EFC04AE965C
            D788F0A68D1BC0D3D383B5EFDDAF873254653E9C9C1C4112B195B51F3C780837
            4A6F0947984638E9E07EB6866976B34E9F5128110E0E7610B9633B6B0F0E0E42
            EE997CE8EFEFD788B0B9B919ECDA19052291393BCF2FB80832D9477DA0CC8A8F
            DB09D638CBB49653FFCC605B3F41083B393A804412CEDAB7A9EE6289E0071247
            E936CF9A9D9FCD2B80972F5F694496C38205F39174246B77757543764E9EC2C0
            FAAE5A010101FEAC5D58580C2DAD32610807ACF347E158C1DAA492BDBDBD137D
            F6F6761015393EBB531114656C0C0D062FAF25AC7DEEFC05686F7F31D1676969
            893AB29BB5A5D23A1CFC2A610847454620B145580E862025355DA16F030ACD32
            141C02A57A77778FCAF514E8FAC0B5E0E868CFCE29556FDEAA5018380E94B27B
            E277B1F67D5CCBA54A6B39E96002989A9A60397C8165F1823084130FEC656BAB
            BDFD258E7AA1425F3C0647EA49AA792A235BE55A9148C4529EABDD1C86868621
            27370FFAFA54D77AC2FE38108BC52CADA91AF0111D25013BBB85A8110370322D
            5318C247BE4D627735D5A9EFD1238798A85139292ABEA2722D3F03943199DA6E
            DB1606AE2ECE4C947EFD2D45A16FF3A61058B2C45D6DDFB411FEEED861F65739
            C568106830FE2DF8D8DDD1606B6BA3F67D3B3A3A7196CFA9BC1EB22108BCBDBD
            589B48F1D5985FFE7EFEE584B084EB1F36C0F5EB6513AF93421F3B9ACCDA0D0D
            8D5072ED86CAB5DCFA5707754B8411E6393A6552A1286A4B3F889A6084BF39FC
            35B37D4D4DCDCC05F17128E9005B9F32F4D4F9E8AD9541B341B3A20E942D9435
            CA8844D57740F527CF9C927A4AA16F5B7818737323EFDEC1F1E3A9C210E68449
            9D8870C19107FE23255DC50CF02D271FEA2C24819649F2A144303434503B88A4
            E0A4E4648032B37285214C012F5EECCA823BF1FB497887A3CB816F062E5D2A81
            A6A7CD2AD753EA3B3B3B82A3C378596A95B5C1F3E7AD6A37F4F439F47984F2F2
            4AA8BD7B4F61F0929313C1C8D0506DB64D1B617E5A161416416B6BDB449FC8DC
            1C1212E259305D588373D01D29CF9A26A0C171717182D09060C8C66CEA1F1850
            E837369EC38C500FD6F04FDDA0684C98EF70D4B929BE90DCA9A985CACA3B5326
            CC81BCF5C0C0E0A4FD73E618C1DBB79FB60D9DD26E89CC838DCD3C2616696959
            681C8626FA4C4D4DD9DA220744A054A394D304B4D15FB366153460AD6F6C6C52
            E95FEBBF1AACD1AF979596ABCCBA2084A92E527D24A89B45DA314922C2596A6B
            B2BE38846FDD0C6E6E2E2097CB21FD94A263A301A5DD1AA1A2B21A6A6AEE6AF4
            DE53224C440E24EC61A9464A4CBE5979134E6BCFDDDD0D6A6BEF697C9F8B7357
            6435D3D2B314FA2C2C44ECB309D5D552A8AAAE119E30C1C36331846D0E65ED57
            7F77401E6E05FF8B407DF684093B705FEC88FB6382A6771EFE9784693DC5C7C5
            809999193B57AE955A4798406A1D132D617653862622BFA048BB0913E8768C97
            A707DCC7B426CBA9F584A71B3AC23AC23AC23AC23AC23AC23AC29F2761BAAB92
            98B857BB08D3F749B4DFE67EC0C207DDF6898E8A609696BE447BA1E197759F25
            6121A123ACEDD011D676E8086B3B38C25FDC431E5FDC633C5EE860CAE9412DB1
            D81AF4B56CA6477166E51F1FD40AD4FBF0281E7DCDFE231EF4F3588BD90E729A
            D187C7653C7E40AD6AF8077088416E55C792E10000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000004CF4944415478DAED9BCB52225718C70F2872
            11908BA2A8999A8D71F206C96E16E322DBBC41F202D9A6F200A9BC4CB2CDC259
            CC2E79844C59959AC5C84D41AE0A2A90F33BA4E3A13D4DD90C0D0CF6BFAA8B96
            43C3F7FFCE77F99FD376E0FB1F7E14A7A77F1C09217E95C71B7924C56AA1298F
            B7F2F8E9E4E4DBB3C0E1175F1E0783C1BF8E8E5EA5F2F97D110A85166DE04C71
            7777278AC582383BFBBB3E180CBE86F06FC7C75F7DF7E2C5CB45DBE6293E7CF8
            47927EFF3B841BAF5FBF49AEDACCDAC14CBF7BF7B60DE1A18CED45DB3317C85A
            257CC2AB0C9FF0AAC327BCEA9809E1546A4B6C6E6E8A66B3255AAD962786C6E3
            71B1B595149DCEB5A8D7EB8B239C4C26E59150E7DD6E575C5E563D219CCD6645
            341A51E738B5D168CE9F70241211DBDB59752E35AAB8B8B8546AC60BA0027776
            B685D4FCEAEF6AB5266E6E6EE647981FDEDBDBFDDF00C8F67A3D4FC85A088737
            A483B7452010500E2E972BA2DFEFCF87703A9D52790B9ACDA6CA5F3B306C381C
            4E45CEE9DA4422A1721990CF575757DE135E5F5F17BBBB396514215CA95C8C19
            C7FB994C5AE65C5416988668B7DBAE8C8AC73765214CA990ADD5AE1E7D772EB7
            A3429CF799E5FBFB7B6F09630C46816AB52A0DEB3A8E5F5F5F2BA3DD209D4ECB
            E889A9739C85D374E8B5A3DDEEB8AADAAE09E3E17C7E4FE52EB38B877584C361
            555C00F9C5ECBBCDB3B5B535358BBC02537D20C2986572B9582C3D39755C13D6
            BD4B6BB0F75DCB9091A117D2D05B57642D6C6C6C28D2C0E4D84482BEBCA5CE69
            85B4444F085330281CA05C2E4B631EF2479FDD690A8A1D7A68DB673914A28EEC
            AA73377DD935610841ACDF27948A63637AEE3223A69E4CC14399F11D0012E4A8
            A9F01029440C30E52AA945D8F31D38C413C2937EC40A678C2F95CA8FAEE53A3E
            63F56E0BE461A542B57D9CEBF47A9C640AEB07E7F7551E7B42F8E0605F152E53
            F5B5C6682728213BF408B0C3A9DA5A9292A2747E5E181BA3F5C56231E3D8CC08
            1F1E1E180D84288427194F11A21899707B7BAB2ABA1DBAC081945E8D75077EFC
            78EE2DE14EA7238B52DD38E6D47BAD1034C1290F75C2765293C66646787F3FAF
            72D014B6D69893F1968232818820329C9C64CAD36C36A3D4DC60301485824721
            6D1526531141D84722614731A04B4E1D2609697D1E27F2DAEDF664BFBD7CB22D
            33239CC96464A1882AE30A85E29891BAB09FB47CA30885C3A3B56DAFD77D244D
            1F3E1755B3081A8D86ECB70F9A5C778653919C09613D2CED0A87B6431B715A54
            4C034893ABA6A520E98383113F4F5DA0B8264C4F841430A9A9F165634B2D1D3F
            153872921E873869E4096190CBE5647B192DCFC855FDC746C23F275F47E2A256
            ABC9AAED6E6782853EE9C175547C3BD856425AA2D0E6B201C00C3293C0348B14
            2E0483DBFCB260555F9362C391F97C5E9D9B162F9E10860861CD6C3A2DC2294C
            D1684C19E4769FCB525748CD5269BC15E929354DCA4CBDC583A4A3C50054127D
            F7530BD4521306AC8B591F03B73B0F9F2561FBCE84BD57AE1C61A0EF17CF6A23
            7EA909035640EC4C10D6B3D8885F7AC2B3864FD827EC13F609FB847DC23EE1E5
            248C8CE566C088B0F9BEF4674718D5C63692D33FB0206591B4DCAA757BB36E29
            097B099FF0AAC327BCEAF009AF3A2CC2CFEE218F67F718CFAB6030F8E7337950
            EB9BC07F8FE21DCBB15FE471228FC4A28D9C3110DCA7F2F859D6AAF7FF02DD30
            46BE0C780E900000000049454E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD97200000B624944415478DAE55B099014E5157E3DB303
            2E820A2210EFE001411404411664572BA095A42A26A6C85546138CA14A2131E4
            28430CC69084252486787305E211444B725459B140230405C18808082CC88260
            40040CC7CA023B3B9DF77CFDE67FFFDF7FCF36E02E5BF1AFEAEA7FA6BBFFFF7D
            EF7EAF6782739E0E01C7457854E3310C8F53E0FF6BECC7E3793CEEC463638080
            7BE0E4153C4E3BD19435F3D88BC72002FC0C4E6E38D1D4B4D0984780F7414A35
            0E8E72F5F018A96AC67DEA087093F7BB04044D50E4AE9896A096D8A749C0B2A7
            6C1E38DF1727A1BDA1CC65F5A68849DAA7380FEC854370CE69F7290558131138
            1B07601365111081D7C4A4011CB86B97D82B541FE4731AD089806360156899CB
            7D1A6CE81090869016DDAB146097808C73F671BDA0362E780849E4BCDA2BE380
            CE04A5F7D27BEA736AC041020119D95C7D5724220208D1998828A404EDEE457B
            64A23994D82B74F670F7F26D57127011281E598798ACC783362A8ECBD1E81052
            4A9B68FD6CA000077E09CB5E9AA9F2396C62BF186057BAD9C010E1CE89C0EFF7
            0218713EC083EB01666F325C1702F28E144425652F005BB26581CD54BAF6355C
            7FEC25007FD90A30798D0DB031B4E7C284242927021602046436638891F3AD17
            038CBC909F5BBA0BC1BF6A13902FD89F0B1AADDA50EF25EBD37EC2D4DF5E0130
            B42BDFFE1832F5A11A9B9979015E301A25D24E0538E348948EB288805C74BEAA
            0BC0A4FEFCCC07798031CB006AF61B4EE72302F28AFB2ED745C23EB065EABBEE
            1D5083AEC47430C7F7DFF53AC08BEFF29A0D1148CD5C2D6997C71660AF74A3CD
            09702E3A9FDE16393D14E0D41C03F8D16B00AFEDE185F5E61AB4CFB65C5F218C
            2D73984CF75DD611E0DE01FCF94003C04D2F01BC7788F76A88F6CC6B2927D872
            0CB078E16C6080EA731B3CFFB837C0E7CE362AF678ADD940A449F7D6E51D098B
            5AABA1C18A84DB97F173450716D1F515B4E59117F173CFFD0760E26A80231150
            0D381F2A26A701ECAA168125A0A4CEE7B507787408DFB3ED0380D1CB58ADB4BA
            DE7519DBDC947500736AE39E5A7674C311012407F84374508B7702FC7295ED44
            898E3F0C04F8647B0672F3CB005BEA78FF0F8117E2A65412B07053DB504ED96E
            DB2C3AA64F015C7F2EDF4F042DDF6D673CA3D0917DFE1C238509ABFCC98806AC
            C30F31EB3367F1F5BF6F0398BAC1CEBAFA9F0E70771F73FDDEB500871B8D941B
            1C7312AD8B012E725A54594958A4DB0E6D765E15AA5C8EA5FBBDE5B664C9CE26
            5CCE0BEF390C703B4AFFDD7A4F26A46DD801DE19FDC3438300BA9CC4F7FC0C1D
            D4AAFFAA5C1B1F9E820EECBC93D164D096BFB4089D668392B2A3D6AEB7B600BB
            EA2C4E2A17A9F4E03300AAFB1BDB9DB7D5F6EEBF1BC0EA460BFE0043D49ABD0A
            2C948EC38162C025A701FC7E007FBFB98ED7D26AF905D4B09B2FE0F99DE83097
            EC62B00D05E3C492D43A06D80D416581017C5B4F761C34280C6DAF3744F44622
            EFE9CBF3E777A0AABD194F27DD922E56FF2AD3A24463D827F8F3F8D7997932CE
            6C07F0C0953C9FBB053562BD02EC09554D02D6762BDE99EC770A72BD6F2780FD
            A842DF7AD926F63B68BBD79D699841CEA4E080391B09BD05BDECE59DF8BB95A8
            AA33D046DF39187766DD5153EE8F40FD037DC1F40D36E366A3E33CB50DAEF13E
            C01DAFB21D0B50F79C08D88DBB650A745B3CE656B16D11B7EF5E694B8862E4B9
            6857DB0FB2ED368636D8AEE5A8F2983175C8D98CA2983A1609DE79C8B63562FE
            B40A806EE5CC3CCAE2F49880DAD4BB23C7E22F2F3A4EC039053CA76C78FE709E
            2FC22CE7FEF536014F56F1FDE4B5293EBA9E78540FA301EE206F4EDE58674741
            E4B1077666C209941EDFC568714D37BE366CBE51692BFB5289C831017EF13AA3
            6233379ACDE9DA9F2B793E7F3BC02335B6932169FD069DDD051DFC80371DC0D8
            FBEF38E0DB7B1A268D5868E2BD9890243F55CF3533E005E894A6D6D8809EAAE2
            F94294FE7DEBE2807F8EB1F3D28E7EC06422147EB447257AC6A014AF8D005FFF
            4F7BCD31C88CE1D1B54A07705E01D7797C6AC062C3CF7E1A6331A67D4BDE437B
            5C6BDBF0AC216C9FABD1118D5FA9241549EBB3677175E51BD3509D9F7D475554
            D173BFEA07D00799B4F708E7CE3AACFD14D5BD02C3643DDAEEB50B1070A31D8E
            527B69377796384C8009D4F9ED3936920A6AC0E3FB70E2519FE794EF50A36DC7
            C44CCAC189483DA8ACAC5E1D278C348AB486A2037962AA90042C0D4A4E28F978
            1B69F9FA6296A64F95538525EDA1755AF90BF48C955D79911B17B31AC9F82226
            033776E7F9A4359C0BEB944EF6188880FB456169050259BECB4855C7EC4A7448
            E32EE539F98B67DE36D788A6A7AF063829CB26F4931571B0F930DE10F0022EA6
            961903B84D64C394D88FEEC90FDEF306C01BEF1B223A614A38B582EFA7304245
            45BEE00793947888367C480B1E155873DFD10B9DD72B00BB0FABE77181F22C57
            4E5B30C57DA2366EBF1A70626AA95B3A5AA52597262F4B019FC60BE8B81E7042
            D36D3D8C2399B39925E36BEDD0676FC35D370AA333D5DE7B0EDB4C2BD6DD21D3
            B6EF88A9891B95FDFA6AE278B5E4A985B5A79E8E52BCF014B6D16F2FE1C44188
            A08E0475262803122D20C66867945429C99CFCC037304F7E01C3DBC29D710D20
            B3A17CFD3E64F68E83066843C16EF5E4952A7BABA5226030CECB974F536938B6
            17DF3F773317FF521CD0C27DA962EAC7CF501B66DC0ABB93E905AB344BFCC40E
            CCD3473AE92B31F2C94A63D77FDA6480BAB5B0CF7E6380DD7E5651CA51A822D0
            E51896E6547219479B8D5ACAC449AD4B1B0DC2ECE81A4CFC1F4582D6ED33DD8E
            A45E96742769BF89C8AC215D3865FCE64B36604A4F674526456BCFD8E8F7CE49
            A5A11770E000CE7A549B92014AFB6810A0B1CBA318182D9E77329E4675ADC854
            B09D64265ADF05AC1BF05D318FFFA3023C7DA35DF44BA2D110FA1D563260301C
            9776A91BA6A81018D0999FFBDB36AC5FD7C6FBD179A55ED25F1286EAA6BB7690
            D50A305564DAA911E0994D004E4A29BD80B55A6BBBCA06106B9F924ACFBE8ABD
            280DF2D84FD4DA126E0CE34D3CDD92D5EB4B463709010F8E008B0D0BE82ECD05
            D86D9D5A47C6CC2F466FFD30663C27A35D2FDBCD75B07E0BA0C18AF310C96ACD
            C92A73A956806F514E8B6821C0330627DBB0062C612915607128BA9B9875A42E
            1E96C208552D4F6126B4619F899122513D17C25D53D12DE089FDB995448029EC
            B9129E1E017E2C92B0EBB4F285A304ACA5ACBDB67E7398551A10F5D5EC5797A1
            DDA3D6DD0C377DD51D16024C3DB38A08F0AD4B6C9A08F034059824ECF6A58F19
            B0482393001E54C220711814E802C4BB9502B84C495603A67C9DBCB406AC9D96
            05B896DB431F096057CAD6DCC98C64682927BD994F6A039752691770B3A87412
            68701820F7E81CD995B656671ABE9A3B8DD312C0AE9796D6AC804CEA65A5062C
            AA5D6480BAA0B91FAA49A8BED309BFDB28D42D6189F3F446D2CDB4647F9D7850
            5A498D031DEB4BB56753017641F92A9C989435139C92D0ADB9B31ED59EAC00DF
            B4D8DEA75B3B53ADD1CB776A3369C9962AFC5303F601D7E0DDE15B4DE7D0EE4B
            76019D896C78F215FCEE79673D3719F4A05CFAF1A13C9FF516773A259DD46093
            5EA41D15E052E09300FA9E738B139D7949CA4A6F1E09F057FFE548B89C0B171A
            33DFE2EEA8FB2B032B77FFA8001FEBD0A560528A49EF94A83CA497702316D9CF
            9F81363CEF6A9ED39B888737C47FE660A9736B01AC13193783BB01EBED715889
            FD752B4AFB4DFB79BA9F1A0C94D28EC60A6DC51ED3460A43BB06765F939E10C0
            42B4FE3994786E4972C871F97EA6A487C47C1D6BDD9F39F8ECF78400D6D598EE
            7AB8357292930C154091AA2FE9695580937E53A9F3F5A624ECFE0ACF97D9B52A
            C0DE39D87DAE44C0AE54232EB8995DAB010C503A7B4B0B5803D5008FFAC7A52D
            0DDAFA4E7F2835128082F35DAB00ECC3E403DF045E330F93AFF900A7FE93474B
            003F9E914272751FCBBFF1D0EBB1A5F0F1F8A3564510FD158FFE9DF66B3C86E3
            D1E178566D85E3001E0BF0188747CDFF002521B7A55821FB630000000049454E
            44AE426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000006D74944415478DAED5B59505457103DC8E6B0
            89A85191615111545051501611051404595CC0058888899A54FEF2934AE53395
            CA4FAAF291AAA4124D30828080A280026E5110641110104582289B1A4516D911
            C9ED8B436623E5101E98714ED52BFA4DF366FADCDBF774DF37F3B41C57BAA1FD
            AB7C5B00DFB2C3971D26502F74B2E3123BBE30FDDABD564B9C3C6C87A1C19BA2
            4B3F98EA556543ABB773AA039C500C8B4C30E0E0875EDFCFDAA1ADEB4A845345
            39DFEFD02F3A35D5B1098A7EF748F47A7F7A9A0877CCF8CEDF44DD66561E34D3
            1D9F677511E16196DB531DCFA48069153484D5191AC2EA0E0D6175C78410F6F0
            7085A3C332141696A0ACBC4290401D1D97C3C37D1DAAABEF2137AF00C3C3C353
            43D8C579359C9D9DB8DDD0D084CCF3D98210DEBA7533ACAD2CB95D5A769B0FEE
            A413B6620104B04008838383483B9B89E7CF5B05213C73A629B6876E83BEBE3E
            3FCFCEBE8C07F50F278F307DF0BEBD61983E7D248073E917D0DCDC22085909E6
            CF9B8B909040686969A1BFBF1F0989A9E8EDED9D1CC25E5EEBB16CA91DB78B8B
            4B5172AB4CE17F747575F9CC8F07635DEBE4B412AEEB9CB97DEFDE7D5CFD2357
            78C233669860EF9E5D7CA4DBDADA712AF90C5EBF7E3DEA9F366D1AB66CF6868D
            8D156EE4DF4445C51D95827258BE149E9EEEA8AF7F849C8B5714DE7BD7AE50CC
            329BC9852B2131051D1D6FBFE9191761CFF56E7060AA4CB89075110F1F368CE9
            BF7FFF4F5CBE724D25C2DE9B36C0CECE96DB9555D5C863AA2C0D4B4B310203B6
            70FBCE9DBBB89E9B2F1C611AE1E8FDFBF81AA6D94D3A755AA6445858982368DB
            566EF7F4F420F5743ABABABA54226C6868801DDB83616464C8CFD3332EA0A9E9
            1F7DA0CC0A0FDB0E3336CBB4968FFF9E80A1A12161085B8A2D1018E8C7ED9B54
            775989900E248CD26D96193F3F939681274F9EAA445682B9733F60A483B8DDDA
            FA02C929693203EBB46A055C5D5DB89D99998D86C6266108BBAE7361C2B182DB
            A492EDEDEDA3BE050BCC111C3432BBE31114796CDAE8097BFB25DC3E7BEE3C5A
            5A1E8FFA4C4D4D998EECE4765959051BFC6261080707053062F35939E843ECF1
            7819DF062634CB99E01028D55FBC6853B89E0275775B0BB178013FA754BD915F
            2833701250CAEE0EDFC1ED2AB69673E5D672F4FE088844D359397CCCCAE27961
            084745EEE16BABA5E5091BF54C195F380B8ED49354F36442B2C2B54646463CE5
            25B55B82BEBE7EA4A4A6E1E54BC5B51EB12F0C2626263CADA91A4823243810E6
            E6F3984674E3445CA230840F7D1C0D6D6D6DA5EA7BF8D0012E6A544EB2B22F29
            5C2B9D01F2184B6DFDFD7D61636DC545E9E75F62657C3EDE5E58B264B152DF84
            11FEE4C841FE573EC568106830FE2DF85D3B433067CE6CA5EFFBECD97336CB67
            155EF7DAE08165CBECB94DA4A4D558BAFCFDF8D331610957DFADC1B56B79A3AF
            93421F391CC3ED9A9A5A5CB97A5DE15AC9FA5706654B841396EAE8E4496D64A2
            B6F48DA80946F8A3831FF2B6AFAEAE9E7741D238101DC9D76713EBA9D3596F2D
            0F9A0D9A1565A06CA1AC914710537D0BA6FED433C71E3F29E3F3F7F3E5DDDCE0
            AB57387AF4B8308425C2A44C4424C1510FFC5B6CBC423320DD724A43590B49A0
            651273200A3A3ADA4A0791149C949C1AA0C4A454610853C08B16D9F0E08EFD7A
            02AFD8E84A20DD0CE4E45C41DD837A85EB29F5ADACC4105B8C94A5C6A6663C7A
            D4A874434F9F439F4728282842F9ED4A99C18B898982AE8E8ED26C9B30C2D269
            99919985C6C6E6519F91A1212222C27930ADAC06A7B0EE487ED654010D8EB5B5
            25367A79229965535777B78C5F5F5F8F37426DAC86BFED064565C2D21D8EB26E
            4A5A486E9596A3A8E8D6B8094B40BD757777CF987E3D3D5D0C0CBCDD36745CBB
            256A1E66CF9EC5C5222E2E89350E7DA33E9148C4D7167540044A354A3955401B
            FD356B56A186D5FADADA3A05FF5A97D53063FD7A5E6E81C2AC0B4298EA22D547
            82B259A41D5360801F4F6D55D697047E5B7CB070A1353A3B3B117F52B663A301
            A5DD1AA1B0A804A5A5B7557AEF711126229111BB79AA911253DF2CBF09A7B5B7
            78F142949757AA7C9F4BD25D51AB19179F24E3333636E29F4D282929437149A9
            F08409B6B68BE0EBB391DB4FFF7A8634B615FC2F02F5CE13266C63FB6231DB1F
            1354BDF3F0BF244CEB293C2C14060606FC5CBE56AA1D6102A9756848206F379B
            5813919E91A5DE8409743BC6DECE16552CADA9E5547BC2130D0D610D610D610D
            610D610D610DE1779330DD55898ADAA35E84E9FB24DA6F4B7EC0220DBAED1312
            1CC05B5AFA12EDB18A5FD6BD9384858486B0BA434358DDA121ACEE90107EEF1E
            F2783F1EE3718B44AFCFC8633CF6181A2CE00F6A556641ABEFE554C736A1189E
            6E8C01477FC9835A6E5A6F1EC5A3AFD9BF6107FD3CD678AA839C60D00C5E64C7
            974CAB6AFE06B346A1147184D5860000000049454E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000006BC4944415478DAED5B8B3B5469187FDD721B
            13A28B3B112995A2C88A4A29934B852EB249BBD5FE07FBEC1FB0CFFE07BBCF6E
            EDB2114229145DB445C925239592A418AA756D5C22C5BEEFA7A33317FB34D6A1
            9DE6F73CE7F19DF99C99F7F77DEFF77B7FDF99397A3E2B03A1B141EA01003FE1
            11868718B40B723CAEE0F1BD87A76FA39E898989A79E9EDE6D1B5B7B4B0BB135
            181818CC7680D38AF7EFDF439FBC1B3A3BDA7AC7C6C6028870AEED7C875D9656
            B6B31D9BA0E8E97E85A4DBCF10E1D76EEE2BC4DA36B3CAA0997EFAA4AE9F088F
            616ECF763C3302D42AD011D666E8086B3B7484B51DD342382828007C967B4345
            4535486BEB0409D4C7671904AD5F07F5F58FA0B4AC1CD031CD0E617FBFD5E0E7
            377E6D4B8B0C0A2F140B4278FBF62DE0E2ECC4DA35D2BB6C70679CB033061081
            811046464620EF5C217476760942D8CACA1276C6EC00636363765E5C7C159E36
            3F9B39C2F4C1FBF7C58189C97800E7F32F425B5BBB2064392C5AB800A2A32580
            1B1D181E1E868CCC5C78F3E6CDCC100E09F90ABC977AB27655550D54DF91AAFC
            8F9191119BF9A960B26B7D7D5742C03A3FD67EF4E8315CFBAB5478C273E78A61
            DFDE5836D23D3DBD703AFB2C8C8E8E4EF4EBEBEBC3D62D9BC0D5D5196EDEBA0D
            75750F340A6AF9B2A5101CBC1E9A9B9FC3A5CB252AEF1D1B1B03F3ACAD987065
            64E6C0EBD7726109077F1508CB519509178B2EC3B3672D93F63F7EFC04AE965C
            D788F0A68D1BC0D3D383B5EFDDAF873254653E9C9C1C4112B195B51F3C780837
            4A6F0947984638E9E07EB6866976B34E9F5128110E0E7610B9633B6B0F0E0E42
            EE997CE8EFEFD788B0B9B919ECDA19052291393BCF2FB80832D9477DA0CC8A8F
            DB09D638CBB49653FFCC605B3F41083B393A804412CEDAB7A9EE6289E0071247
            E936CF9A9D9FCD2B80972F5F694496C38205F39174246B77757543764E9EC2C0
            FAAE5A010101FEAC5D58580C2DAD32610807ACF347E158C1DAA492BDBDBD137D
            F6F6761015393EBB531114656C0C0D062FAF25AC7DEEFC05686F7F31D1676969
            893AB29BB5A5D23A1CFC2A610847454620B145580E862025355DA16F030ACD32
            141C02A57A77778FCAF514E8FAC0B5E0E868CFCE29556FDEAA5018380E94B27B
            E277B1F67D5CCBA54A6B39E96002989A9A60397C8165F1823084130FEC656BAB
            BDFD258E7AA1425F3C0647EA49AA792A235BE55A9148C4529EABDD1C86868621
            27370FFAFA54D77AC2FE38108BC52CADA91AF0111D25013BBB85A8110370322D
            5318C247BE4D627735D5A9EFD1238798A85139292ABEA2722D3F03943199DA6E
            DB1606AE2ECE4C947EFD2D45A16FF3A61058B2C45D6DDFB411FEEED861F65739
            C568106830FE2DF8D8DDD1606B6BA3F67D3B3A3A7196CFA9BC1EB22108BCBDBD
            589B48F1D5985FFE7EFEE584B084EB1F36C0F5EB6513AF93421F3B9ACCDA0D0D
            8D5072ED86CAB5DCFA5707754B8411E6393A6552A1286A4B3F889A6084BF39FC
            35B37D4D4DCDCC05F17128E9005B9F32F4D4F9E8AD9541B341B3A20E942D9435
            CA8844D57740F527CF9C927A4AA16F5B7818737323EFDEC1F1E3A9C210E68449
            9D8870C19107FE23255DC50CF02D271FEA2C24819649F2A144303434503B88A4
            E0A4E4648032B37285214C012F5EECCA823BF1FB497887A3CB816F062E5D2A81
            A6A7CD2AD753EA3B3B3B82A3C378596A95B5C1F3E7AD6A37F4F439F47984F2F2
            4AA8BD7B4F61F0929313C1C8D0506DB64D1B617E5A161416416B6BDB449FC8DC
            1C1212E259305D588373D01D29CF9A26A0C171717182D09060C8C66CEA1F1850
            E837369EC38C500FD6F04FDDA0684C98EF70D4B929BE90DCA9A985CACA3B5326
            CC81BCF5C0C0E0A4FD73E618C1DBB79FB60D9DD26E89CC838DCD3C2616696959
            681C8626FA4C4D4DD9DA220744A054A394D304B4D15FB366153460AD6F6C6C52
            E95FEBBF1AACD1AF979596ABCCBA2084A92E527D24A89B45DA314922C2596A6B
            B2BE38846FDD0C6E6E2E2097CB21FD94A263A301A5DD1AA1A2B21A6A6AEE6AF4
            DE53224C440E24EC61A9464A4CBE5979134E6BCFDDDD0D6A6BEF697C9F8B7357
            6435D3D2B314FA2C2C44ECB309D5D552A8AAAE119E30C1C36331846D0E65ED57
            7F77401E6E05FF8B407DF684093B705FEC88FB6382A6771EFE9784693DC5C7C5
            809999193B57AE955A4798406A1D132D617653862622BFA048BB0913E8768C97
            A707DCC7B426CBA9F584A71B3AC23AC23AC23AC23AC23AC23AC29F2761BAAB92
            98B857BB08D3F749B4DFE67EC0C207DDF6898E8A609696BE447BA1E197759F25
            6121A123ACEDD011D676E8086B3B38C25FDC431E5FDC633C5EE860CAE9412DB1
            D81AF4B56CA6477166E51F1FD40AD4FBF0281E7DCDFE231EF4F3588BD90E729A
            D187C7653C7E40AD6AF8077088416E55C792E10000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object wmibAssignNext: TWMImageSpeedButton
          Left = 1016
          Top = 9
          Width = 60
          Height = 60
          Action = frmSEC.actControlAssignNextEvent
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000005AB4944415478DAED9BFB53135714C70FCF04
            6CD2F00850E9B4BC0308424029880A42145AA76AB1204EB133F597F64FE8F40F
            E8F43F687FA8BFB4B60A22386D4728C607280F7934401112823C9CC101021192
            CCF086DE7335DB2D25E4B59B84C0776687BBC926B91FCEB9E79E7BEEAE4F7A46
            1E6835AA4400F88E1C0A7288C1BB642087921C5F27CAE45A1FA15028F3F1F1E9
            0897464B44E250F0F3F373770739D5FAFA3A180D7A98D54DCE6F6E6EE622F06D
            69C4BB659210A9BBFBC6AB5EE9A709F4CB3A045E884B382CF636CB6E155A7A74
            A4DF84C09BC4B7DDDD1F9788C42AD807F666ED037BBBF681BD5DBC03930C0EA4
            D270888E3E08919152785B2C0691E82D080808A0EFAFAEAE82D16882058301A6
            A7753039F91274BA592019D1EE02160802212D2D159265892016DB979A1B08BC
            5AA385818141585E5EF16C605F5F5FC8C848872CF961080C0CFCCF7BAB6B6BA0
            9FD383C16884E5A5657A6D406000FD8784864818AB9BB5B2B202AADE7EE8EB1B
            A05992C70187858582A2B81042434398D716171741A31981E7A3633BBA2AC2A3
            EBC7C5C64052523C04070733EFCDCF2F8052F91074B3739E031C1F1F0B45A70A
            C0DFDF8F01EDEE56C1E09006363636ECFA2E844F494E82A347B3212848485F43
            0B3F6A7E02C3C323EE074E3B9402274E1C63CEB1534F5ADB9D1E7F18078EE7E7
            118B2730AF3D7EDC0603CF86DC078C963D73BA88396F69698567836AAB9FF3F7
            F7A77FD7C8B8B6A6D4D464283899CF9C2BEF3F02ADF6B9EB8171CC967D728E71
            E3A6A60774AC5AD38103C170A9A28C4E59D535F5603299AC7E06C7764949316D
            A37BD7D5FF0EB30E8C698781719C957F7A810950B65A161513F31E7C587A9AB6
            1BFF54C2D8D8844D9F635B5AAF7F05B76AEFD81D1F1C0696CB3320F78323AFBF
            84B817BA99AD8A8D7D1F4A4B147603A38A8B0A9831DDF1B41B54AA3EFE810502
            01547D5641E7D9C5C525B871B39604A86597006320BB5C594EA337CED3D77FA9
            B62B383A047C245B4EA68C2CDA76246A3A038C3A445CFBE41BD7EEECEA819E9E
            5EFE8031D07C7EA592260668DD9FAFDFB43B0B721618E307F621282888A6A1BF
            DEA8B539F7B61BF8E03B5170FEFC59DAEEEBFB1BDADA3B69BBB4540141422134
            342A616969895760545E6E0E6466A6D3F6EDBADF606646C70F704E4E36646765
            D276FD9D3F606A6A1AE2E3C85C7CE6F55C8C6319FFE33B4173011C191941A6C4
            8F69BBA3A38BE6DCBC00E33444AEA76D84C56901CFD1C5CC655E6BD05C00A35B
            5FFDA28A2E38C6C627A0B151C90FB025492412A828BF60133417C0A8B2B27310
            1921A57372754D1D3FC068E14CB2FCC32081418B2DCCA0D879AF2568AE80158A
            42484C88A745841FAFFDC40FF0575F5EA591DA5661FA87199125E07612F47A49
            F07344987561F685FAFE876B9E0B8CEE7FB9F222737EFF41B343CB3E970073E1
            D228BC0ED34467A05DE2D296644FD0E20ADA6541CBD969890B68F6B4343EF102
            1A1AEEF103CC45E2C1057404B1EC456261943DAB26A7524B95AA9FFC58176D9F
            FDA804044201DCBDDB6417AC25E87BCA873032326AF17A97A5965C2C1E6C819E
            23E3B2C6C2B84477BE527589F681F7C503CAD9E5E14EC21A59963C837EE7D090
            66DB6B5CBA3C44FDBF00708BF31D02CBBFED8602008A5DE2C1008381C615724B
            8907B5B588D7DCD20A833616F11C554A8A0C0A0B8ED3B6CB8B78287699168306
            2E04C6C75FF002CBAE74BAA54C6B16BB108FD02D2488716D69B665516E2BC49B
            B575AB45AD1E86B6F6A79C6CB5E41FCB05992C9179CDED5B2D666DB799D6D9F5
            1785777C332D8B16EA501EB5996696A5ED52B55A0BA363E356B74BC3C3C3689A
            BA2BB64BD91DDF69437C8E74DA683279C786385B7BE69687AD62DFD412151501
            629168DB9B5AF03688A9A999DD7B538BA76A1FD8DBB50FECEDDA07F6769981F7
            DC431E7BEE319E649211B5E3835A627128F87A99A53788650DFF3EA895E7F3E6
            513C1979EF5B72605941E4EE4E722C2339706BE21B12AB34FF00CAD5296E0A99
            49140000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD9720000043F4944415478DAED9BCB521A4114861BB928
            C270535134B1B231983748762E74916DDE2079816C537980545E26D966A10B77
            C923C4B22AE5C27801B90C37956BCEDFA471302233D37321037F551750304C7F
            7D4E9F73BA1B7C6FDFBD67FBFBDFB618639FA9ED528B316FA942ED80DA87BDBD
            D7C7BE274F9F67E7E6E67E6C6D6D27329975160C06DDEEA0A56AB55AECFCFC8C
            1D1FFF2C77BBDD9700FE92CDBE78B3B9F9CCEDBED9AA93935F047DF415C0EACE
            CE6ECC6B96BD2F58FAF0F0A006E01EF9B6DBFD714414ABD80CD8CB9A017B5D33
            60AFCB11E05028C4E6E7E7E931C8028100F3FBFD8C2A3BFE1E553EACD3E9B076
            BBC39ACD26BBBDBDE58FFF1D3080A2D1085B5C5CE49046D46EB759A3D160B55A
            9D0FC84403FB7C3E028D3245890EAC28D4EDF678B5038B02843ECA3F8F0141FB
            F7F35D56ADD608BCC67ABDDEE401A3344DA592438B0FC0C15AD7D7378FBA2AC0
            715D381C26AF0873B7176AB5DAAC542AD2F5ADC901462793C924EFB800AD56AB
            AC5E6F18B60EBE035321165306E0F88E52A9CC07CF7560CCD5442231788D4E95
            CBAAF4FC838B2712710E2F542E97F9DC760D18964DA55283D7B042BD3EBE43C2
            13F4583F128990F7DC0D68B158326D692960CCB9747A65D0F942A14873F57AEC
            7570D3D5D5347F7E7999E3EE3F4E98DB4B4BA9C120E572791E001D0306246045
            80D26BD97EE717A8F34B860609D25A1AB080361A1F4C032B8AC2E2F1FED617DC
            0B6EA6575A6B1901869005C49C56D50A0F8CB6032398ACADADF247B823DCD248
            809201C63D311D302D70CF8B8B4B43F736058C74118BF5AD6B266ACA00435AD7
            AE542AD4F45BD9147026B3C647B8D3C1085F189E47B2C0881FF030F4016528AC
            6C1B3016012B2BCBFC39CA3E5555F9730421B85BA15018EB62B2C0503C1EE7E5
            2B84E0A577C1611818AE0C9786F2F93CAD6E9A43007AE69515C05881214B4046
            82976160A42151E46354E1CE780D3717F9781CB415C0B817EE897BE37A7C8F2D
            C0A3072240239ED6056D0530040BC3D2C8C9C814B600C3C258FEF5A1862B2404
            116DDD3B0ADA2A609193F1FD6767E7F6006F6CAC0FACA84758D2E572C3A3AF05
            46D043F03323A426A428E8F4F4F7E40263B18FB422647631E008B0152E0DE173
            70491968475C7AF440E80F5A56413B16B464D39215D0C369E986173CB6006B0B
            0F51E1182D3CAC8076ACF0182E2DABFC66D0F272BFB4BCBA1A5F5AEA832E12F4
            E874E5586909C92E1EF4403F362F1D5D3C40B2CBC3C7A1C3940514BE7B326A07
            C5F1E5A1EC06808C5CD9008064B67864E4CA160F24B3896756AE6EE241DA6D5A
            DC18911539D10E69773A5DD9A615D26EC4A3233871B0DAD213B3112F74FFA805
            E749580579F2A845E8A1C334A40B58C273876942D61C972E70D8893F2ED5767C
            6A0EC4B59A9A9F3C3CA4BB1FB58408DEEFCD1FB54CAA66C05ED70CD8EB9A017B
            5D0278EAFEE431757FE3D9A6CAE7FB94FC51EB95EFEF5FF1B2F4DE276A7BD414
            B73B69B1B0F1B54FED23C5AAA33F22FD22BE7D7E8D2E0000000049454E44AE42
            6082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000009C04944415478DAE59B09AC56D51180E7FEFF
            7B0F442A568B0DB8E1C222B2750BA2426205D4186368A22658B744A3D245DBB4
            51AC0A2DD5D8C6D85A5B13B5A6DA6ADCD2E22E0A56416569526B4191A53C2885
            078F5696870ABC7FB99D7973E7BF73E69EFBEF0F1238C979F7BEBB9C7BBE3373
            66E62C7F70FC7321601A8AF91ECC93311F010757EAC2BC00F3AD98D706083C1C
            4F96623EF240D7AC97D34ECC6710F09FF1E45BBDF185A0CEF7C2DE83FE0B01EF
            8226A9B1060C8294EB65E09CF3D07FBDC1F4290137549E0F3290EB417C3F308D
            E04087315418FD09156833E1EB06B6A0163230D73341F23D0D505490A1064E81
            AF17BC66E072A099E89A3EDA06F0250B588CAEE9239806B10DD62BC025F52C03
            4A9294FB19732C97A81222E592B4433EAF045E0B74D5C0095805A341E53C1324
            D53A08FC2A1D7AD4B96861A373DD18F5405705EC83D560FA3C6BC075FF95321C
            60A39E16B410BAD7ECFD5AA12B029783CD062E68E96824EE740105AD2B5A5261
            235101D6E08506A0CB02A7C16695245B1468D6033F6200C0578F061889C763FB
            011CD317A05F0B97FB791E60DB5E80CD9F03ACC468E0FD4F0056ED4A4216C232
            D76A844E05AE044B402D99E8682007B4024C3B11E0BCC10C594BEA40F8791D00
            73FF0DB02BE702E6A3F37CB17EE8B2C09560B34AC2F47F1FCC970D01987E32C0
            E12D6E797B0A00EB3F4589EE01D89DE37748D2830E0338F17080C3CCF39FA1F4
            9F5A0FF0CC0680BD8508AEC8D0852AA06B0216E95ACBDB62605B95848762707A
            FB188093FAC7E5ECE806988FD27A6F1BC09AAEB8F56D1FA63286E1FB671D0370
            EE2080A3FAC4656CFC0CE0E7CB59D545C2B9A20B9D0F93163D4DCAA9C0DAE064
            55D6B0728D2A79EB6896B0803EB18E553357F437A624FB712A97BAC295A7E0F0
            AD8DAF5119BFFC10E0F58E58A21ABAA0241D1A695704B6AAECC046C0AD4A9DA7
            9D0070D3C81864C11680075761949EF7F7271D71F92A25E5F447159F310260CA
            E0F8DEAF57E2706763A4D64895336AAEA1D3543B1558BB9E6C46C14647FAFF5C
            ACCC1D63F81D2AE53708FACAA6A4BBB1E5F78D34616F31F94C608E171E0770F3
            697123917ABFD1C19024DD5C316E006DCDD3A4EC0596BE2B52CC2AC8D648C2D4
            671F3C83D5984AA08ABCB3CDE35BC1EDBB5FC2FEF9E8997C7EDD6280AD7BA06C
            A2E7267D1960D6183EEF4692194B015677B9C0A2DA79ADDAB500678C1A536E8B
            A0C9A23E3201604864A01EF818E0D5CD716020A0454F273A1B0DD39CAFF0F99D
            1F602375A654CC74AD8B50D23F18C9F736A0B5BF76315BFE7CA4DA7923E93429
            3BC0569DB581924CD0DF46B773ED507EE7ADAD00F77EE48F772D486080EFF807
            C0A24EF79D523D14B004393347C57DFA9135007F6A67098B01CB192917C36A81
            3D46AA2D3A1E8D2AF9E444F6B3BBBA59BD76E793E19F0DF4453D272AE0DB1178
            61A76B686C5FD651DC1118D03C7E365B6FF2D3972D648FE093723E74352E15D8
            51E7C8488974FB6401AE3915E0AA53F8F987B1955FDBCCE7BE10B0E0B1D216F8
            EDAD491594C611E9EA80E7E213D888517A742DC063EBDCBE6C7DB4347C5960ED
            7345BA94FBA0549F9DC452EEC268E9FA255C7068809DB85741F400A301FAD9B8
            24B085B6C0D2BDDAB0D19F9EC8C109C5E0972F4243E6B1D8B914B54E00EBD051
            AC720F307EE8EB3808F8D537F8D997FE83AADDCE2FFD088D497F54B75F6070B0
            639F1BF76A232646EBA7E3E23EFC76673262D2A914AB67628F71C3700E6129DD
            805D6AC58EA45AEB80241558C6B6BA70516582A60F4D3F29B2B058D97FED0618
            3F10038F48C5A85F7D7719C0CEEEF88316F84C049E3D362A03ADF4C2AD6E7C5C
            50FD5EA4DCA2E200AAD3E823D1338CE7321EC26EF584C778E553FA7112D8BAA3
            8081C9DF0EC721DEC0BEDCAFD7EEE6C619D0C6526FCDC4D0DF5BC6C6A460D44A
            80EF8C80677DC056DAFA51795EABB51EACB4627D5E3C87071FEFA2EF9FF97EAC
            D225E030D9E85501B746A32082A608A96F96CFA531E8F9C138E2993DCE85D692
            D68663C2C01878F63F59C23D0302635DB58BF2A935053DA70DE011D855EFBA90
            F6E805D62EA9C5E392089406F3970C61C952ACACA76FC855108C2482FE0EF6AF
            EDDDAE5A59E0773A5D75162955EAC714D27E73104F229CBF20295D1D6B176A05
            16955E30953F566DA2D6BFF96FAE5A69600A1E9EDE105BF49C898701C01D8B2B
            E01FA2A1BCE8787E66D23C7F6CDD30F09B53DD09B94AA91D816F5A167F90208E
            C3C1FEEF27F0B7E8CBF77CC88301B1D21658249C0DDCD9955E05AE47A5E9FE8D
            E8A73FD9970C0026A32ADE322A86BE7B05C0FC2DAE91D1FED8CE8A3645A59B6D
            B47A60B57B023700A19878A6819ED791B4EA522F19BD09BC182D1A485C598FD1
            4A03AEC72D517CFDBF7D7130607DA1CCA44C4109DD369ACFE9FE5D2B3854B573
            30767C4EF57A21724B8BD12DDDE2714BFA581E181A0B3CC832931AD3B8B5BBE8
            EF937AFA682A4AFA27A3F9DBF4CC9CE53C95537A5624AC8EA763E0F1DBF1B1E1
            FB633B0377D71A78D4125A3EBF112DEC7A7E89FA234DC9DCB59CDD107D705F31
            FEB0F4C952C3AABE4F479AC3A20940812677F5D72D4A9DCDF3D70F03B87408FF
            DF5068D98CC1838478E5242C1032CEA5F3F38F654344F7D6A1E65CFD9E59B588
            B482EAF3643478A039ECE98D0E1E1A1D1EF6A896024F8B8F7D0B6F9351D2579C
            0C3017B5E7D54DAE546552805CD1F79B393CAC7502E0C6A5EC8234705E03EB01
            8419EB6AFFAA1BDBAABC489ABEF987B39A3C0150EB14CF9BD8D7EE5B194FF1E4
            C3A45BB013013E096B5089E6B24A1BE8F8E351ECC3C558353CC5E355EB2A26F1
            EEFF187DE8E6E43A90633852246C838A8CB9A7FBB89EC4BB6E09071D0D4DE239
            6E20A87E9A963E30072DEBE2FFFA673EECF48D4067C01F5458588AE0668D8DA7
            69C9F5AD6AE6346DA589F86C74D413F1F4119AAE7D7993BB705DF0F4612D65FB
            2DBB984E92EDF589F846965A5E47D5FEDD6A5E2114EB6DA76113DA04C655014F
            19CD189E5C6A99BB51C12A0355F7528B95723D8B69DB31D27A1CDDC56BB49856
            28BF8469830AB21317A054AF40D7F745B39826926DEA629A54422AD0C8722981
            5325176DE3E5CEB4BD06541E2D97D292CA9432CBA5026507060D2F97FA54BB19
            0BE2ED184175EE8D16C4819F19D48F17C4FB1DC805712DE57AB63CD00A01F56D
            3238F56C79A0E88DFA6AD7FEDCF2500D74DAA696925FC56747E0C8E66B47F108
            6770CAA6968E3D001FE100E0EFDB51757722042877665C5CAF6D6AA9049DB66D
            491AC4EEC8D365D92415B03BF00468BF6D5BAA06BA9A8D6976079EDDE32149AF
            3A1EB08D69E5A06D7878D06C3D4C858E2E1E949B4B2D7425F08362FB70B5E056
            85CB19AE84C1F23440A89EDBEF1BC42B81CB350D5FBA5641C2259014C8464035
            F021F7238F43EE673C23F064091C1A3FD49A10443FC5A35FA7DD8D790AE62F1C
            E89A3539E19005E663BE0DF3EAFF03527F7FB435DB02230000000049454E44AE
            426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000005C34944415478DAED9B5B5053571486FF0021
            093631A280D5692B220414CA4D2DE0050B51689DAAC58A3AC5CED497763A7DEB
            4BA7D3C74EA72F9DE94367DA87FAD2DAAAA8E8B41D418C17502E7229A2C84550
            D0193B284225C94C2011D3BDB6269E5203B99C93C4907FE60CFB9CDCF6C75A7B
            EDB5D73E479691998F875F362603F8861D7A7668105A32B2C3C08ECFB55F15F4
            CB5E396AD761CAD6AC327CAF8DEE3A0D99C518E80E8A2ABB4A036B7A092CFA4F
            1F22529E47C0C755B5DF95295A2A03DD3749355950014BD12755043C3EFFDB52
            4DA85976BAC8D2E39FD59809D8CE7C3BD0FDF18B58AC4218389415060E758581
            435D9203CB6432C4C52DC2D2A54B90901087F91A0DD4EA972097CBF9EB369B0D
            269319E34623EEDD1BC1DDBB7F6364E401EC76FB8B05AC5044233D7D255275C9
            D0683C4BCD8D0CBEB7AF1F5D5DDD989CB40637704444043233339093FD3AA2A3
            A3FFF39AEDD1238C8D8EC1683261726292BF571E2DE7FF90D8055AA7D51DB25A
            ADE8B872159D9D5D989A9A0A3EE0850B63A12FDE84D8D805CE6B168B057D7D03
            B8796B704657257872FDE589CB90929284989898679D7C380E83E13C461E8C06
            0F705252228ADE2C445454A413B4ADAD03DD3D7D78FCF8B147DF45F069A92958
            B326172A95925F230B5FA8BB841B3706020F9CBE2A0D1B363CFB3C75EA524393
            CFE38FE2C0FA75F9CCE22B9CD72E5E6C44D7F59EC0019365B76C2E729ED7D737
            E07A77EFAC9F8B8A8AE27F1FB1713D9B56AE4C45E1C675CE73C3D90BE8EFBFE9
            7F601AB365EF6E73BA716DED393E5667D3BC7931D85D5EC6A7AC23952760369B
            67FD0C8DED929262DE26F7AE3AF1071E7831A6BD06A671B6EBBD1DCE00E5AE65
            49CB96BD8AB74A37F376CD690306076FBBF539A1A5C7C6FEC1D163273D8E0F5E
            0367676722EF8DD5BC4DEE456EE6AE12135F436989DE6360527151A1734C375F
            6E434747A7F4C00A850215EF97F379D66299C0A1C3C758809AF40B3005B2BD7B
            76F1E84DF3F4C15F8F78141CBD025E9D9BCDA68C1CDEF6266AFA024C5AC55C7B
            E353D76E696D477BFB15E98029D07CB06F0F4F0CC8BABF1C3CEC7116E42B30C5
            0FEA834AA5E269E86F878EB99D7B7B0CBCE4E5C5D8BE7D2B6F77765E4363530B
            6F9796EAA1522A515D63C0C4C484A4C0A4FCBCB5C8CACAE0EDE355BFE3FEFD11
            6980D7AECD456E4E166F9F38F9278687EF2169399B8BB73C998B692CD37F7C26
            6831801312E2D994F80E6F3737B7F29C5B12609A8694CA27E91EC1D2B440E7E4
            629191916E418B014C6EBDFFC30ABEE0181CBA8D9A1A8334C0AEA4D56A51BE6B
            875BD0620093CACAB621213E8ECFC9472AABA401260B67B1E51F05090A5A4251
            0625CC7B5D418B05ACD76F42F28A245E44F8E9C0CFD2007FFCD17E1EA9DD15A5
            7F9411B9026E6241EF0A0B7EDE88B22ECABE483FFC78207881C9FDF7EED9E93C
            3F7BAECEAB659F5F80C5706912BD8FD2445FA0FDE2D2AEE449D0120BDA6F41CB
            D769490C68E1B43474FB0EAAABCF48032C46E22106743CB3EC4E66619227AB26
            9F52CB8E8EABECC75A797BEBDB255028153875AAD6235857D0670CE7313070CB
            E5FBFD965A8AB17870077A948DCB4A17E392DC795FC56EDE07C9170F245F9787
            33896A6439D999FC3B7B7AFA9EFB1EBF2E0F49FF2F001C157D87C0F56F07A000
            4012967828C050A0F1870252E2214D2FE2D5D537A0DBCD229EB74A4BD36153E1
            7ADEF67B118F242CD352D0A085C0D0D01D49608595CE8094691D1216E209BA9E
            0531B12D2DB42C2960857887A66FB5F4F6DE4063D36551B65AD615E441A74B76
            5E0BF8568B43CFDB4C6B69FD8BC37BBF9996C30B75A4A0DA4C73C8D576696F6F
            3F6E0D0ECDBA5DBA68D1429EA6BE10DBA5C28ECFB4213ECA3A6D329B4363435C
            A83973CBC374096F6A59BC381E1AB5FAB937B5D06D10C3C3F75FDC9B5A825561
            E050571838D415060E753980E7DC431E73E3319EFC0A588A9F3CC6938A295B13
            7F50EB5A0D6413A640F74D54D9956A58334A1D0F6AE5CB9E3E8AA763AF7DCD0E
            2A2BA803DD49914516A4AD892F58ACEAFB170DA2891447928CEE000000004945
            4E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000005AB4944415478DAED9BFB53135714C70FCF04
            6CD2F00850E9B4BC0308424029880A42145AA76AB1204EB133F597F64FE8F40F
            E8F43F687FA8BFB4B60A22386D4728C607280F7934401112823C9CC101021192
            CCF086DE7335DB2D25E4B59B84C0776687BBC926B91FCEB9E79E7BEEAE4F7A46
            1E6835AA4400F88E1C0A7288C1BB642087921C5F27CAE45A1FA15028F3F1F1E9
            0897464B44E250F0F3F373770739D5FAFA3A180D7A98D54DCE6F6E6EE622F06D
            69C4BB659210A9BBFBC6AB5EE9A709F4CB3A045E884B382CF636CB6E155A7A74
            A4DF84C09BC4B7DDDD1F9788C42AD807F666ED037BBBF681BD5DBC03930C0EA4
            D270888E3E08919152785B2C0691E82D080808A0EFAFAEAE82D16882058301A6
            A7753039F91274BA592019D1EE02160802212D2D159265892016DB979A1B08BC
            5AA385818141585E5EF16C605F5F5FC8C848872CF961080C0CFCCF7BAB6B6BA0
            9FD383C16884E5A5657A6D406000FD8784864818AB9BB5B2B202AADE7EE8EB1B
            A05992C70187858582A2B81042434398D716171741A31981E7A3633BBA2AC2A3
            EBC7C5C64052523C04070733EFCDCF2F8052F91074B3739E031C1F1F0B45A70A
            C0DFDF8F01EDEE56C1E09006363636ECFA2E844F494E82A347B3212848485F43
            0B3F6A7E02C3C323EE074E3B9402274E1C63CEB1534F5ADB9D1E7F18078EE7E7
            118B2730AF3D7EDC0603CF86DC078C963D73BA88396F69698567836AAB9FF3F7
            F7A77FD7C8B8B6A6D4D464283899CF9C2BEF3F02ADF6B9EB8171CC967D728E71
            E3A6A60774AC5AD38103C170A9A28C4E59D535F5603299AC7E06C7764949316D
            A37BD7D5FF0EB30E8C698781719C957F7A810950B65A161513F31E7C587A9AB6
            1BFF54C2D8D8844D9F635B5AAF7F05B76AEFD81D1F1C0696CB3320F78323AFBF
            84B817BA99AD8A8D7D1F4A4B147603A38A8B0A9831DDF1B41B54AA3EFE810502
            01547D5641E7D9C5C525B871B39604A86597006320BB5C594EA337CED3D77FA9
            B62B383A047C245B4EA68C2CDA76246A3A038C3A445CFBE41BD7EEECEA819E9E
            5EFE8031D07C7EA592260668DD9FAFDFB43B0B721618E307F621282888A6A1BF
            DEA8B539F7B61BF8E03B5170FEFC59DAEEEBFB1BDADA3B69BBB4540141422134
            342A616969895760545E6E0E6466A6D3F6EDBADF606646C70F704E4E36646765
            D276FD9D3F606A6A1AE2E3C85C7CE6F55C8C6319FFE33B4173011C191941A6C4
            8F69BBA3A38BE6DCBC00E33444AEA76D84C56901CFD1C5CC655E6BD05C00A35B
            5FFDA28A2E38C6C627A0B151C90FB025492412A828BF60133417C0A8B2B27310
            1921A57372754D1D3FC068E14CB2FCC32081418B2DCCA0D879AF2568AE80158A
            42484C88A745841FAFFDC40FF0575F5EA591DA5661FA87199125E07612F47A49
            F07344987561F685FAFE876B9E0B8CEE7FB9F222737EFF41B343CB3E970073E1
            D228BC0ED34467A05DE2D296644FD0E20ADA6541CBD969890B68F6B4343EF102
            1A1AEEF103CC45E2C1057404B1EC456261943DAB26A7524B95AA9FFC58176D9F
            FDA804044201DCBDDB6417AC25E87BCA873032326AF17A97A5965C2C1E6C819E
            23E3B2C6C2B84477BE527589F681F7C503CAD9E5E14EC21A59963C837EE7D090
            66DB6B5CBA3C44FDBF00708BF31D02CBBFED8602008A5DE2C1008381C615724B
            8907B5B588D7DCD20A833616F11C554A8A0C0A0B8ED3B6CB8B78287699168306
            2E04C6C75FF002CBAE74BAA54C6B16BB108FD02D2488716D69B665516E2BC49B
            B575AB45AD1E86B6F6A79C6CB5E41FCB05992C9179CDED5B2D666DB799D6D9F5
            1785777C332D8B16EA501EB5996696A5ED52B55A0BA363E356B74BC3C3C3689A
            BA2BB64BD91DDF69437C8E74DA683279C786385B7BE69687AD62DFD412151501
            629168DB9B5AF03688A9A999DD7B538BA76A1FD8DBB50FECEDDA07F6769981F7
            DC431E7BEE319E649211B5E3835A627128F87A99A53788650DFF3EA895E7F3E6
            513C1979EF5B72605941E4EE4E722C2339706BE21B12AB34FF00CAD5296E0A99
            49140000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object wmibTakeNext: TWMImageSpeedButton
          Left = 1079
          Top = 9
          Width = 60
          Height = 60
          Action = frmSEC.actControlStartNextEventImmediately
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000006524944415478DAED9BF953535714C7BF81B0
            A4400854AC1016655594D58E6CA254512B56ACB4EECB4C6B7F68FF834EFF804E
            FF8C765A69D5E288082D8A4051D6AA0852D945544008420801952DBDE7B62F06
            348124CF2444BECE1BEEDBE6DD4FCEB9E79E7B9E4F12979086CEF6C628003FB0
            2D9B6D723897346C2B63DBB75131499D124F4FCF18894452B72A40A9F091FBC3
            D5D5D5DE1D1455B3B3B318D78C6058D5A7D6E974A9045C10B03A384FE11760EF
            BEBD558D8E0C32E8FE8B043C161E192F7736CB2E1459FA4157B3968075CCB7ED
            DD1F9B88C52AAC003BB356809D5D2BC0CEAEB702CCB23604AEF9006161A158BD
            3A007E7EBE90C964FCDCDCDC1CB45A2D9E8D8CA2AF6F000F1FF6627C5CBB3C81
            5D5C5C10BB210609099B20972F3D1DEFED7D8C3B8D4D78FA7470F900AF6116CD
            DABE95595331EFF8F4F43454AA6778FEFC39666666E02A7585C2D717FEFE7EFC
            0732544747176E56D7E2E5CB29C7064E4A4A406ACA87AF20195827EB7C6B5B07
            86879F71375E28A9548AB0D010C4C6C6203858A93F4EEE5D5A5A0615BBCF2181
            333252111FB751BF4F56AAA96DE0165DAAC83BB66DCBC0FBCCEAFC07635E71A5
            B8F4ADB8B855C0C9C90948D9F29F656766665176BD023D3DBD6FBC9682589032
            102D2DADCC655FBE769E162E9999E9D8B03E9AEF4F4D4DA1E06211D46AB56300
            0705AEC18103FBF4B0978B4A30383864F4FA2FBF38090F0F0F3435B7A0A6A6DE
            E875E9695B58D08BE3EDD151352EFC7E89AF72EC0A4CD6387A244F1F89FF6463
            CE9865057DF3F519FEB7BDBD13E5155546AFA3296DF7AE1D080F5FCBF76FDFB9
            8B8686DBF6058E8FDF888CF454DEA6317BBDFCAF45EF592A30C9C3C31DC78E7E
            CEE76EB2EECFBF9C332B26880A4C163879E208BCBDBD7834CECF3F8FC9C9C53B
            630E30693D1BCB1F6565F2B69856361B58A90C42EEFEBDBC7DBFB51DD5D57588
            8C0CC7A3478F4D829B0B4C73F4E95347B99535E3E3387BF6BC7D80838202A164
            1BBFB9B39B47DEED6C4AD16834282C2C817662421460527A7A0A12E237F1F6B9
            F31731C2D2519B032F9497D77BF894456B0A60A6A02D010E0D09C6BE7D7B78BB
            82DDD3C6EEB53B30C9DBCB8B4D513926A12D01A61FF3F4A963BCDDD4748F2734
            3607FE2C2F97AF804C49AD1EC36FE70AA0D3E95E03EEECEA465959E5923B28DC
            47696A65E50DDB031F3F7608BEBEA657425AED04CEB2E86D98437F75E634DCDC
            DC789BA6319ACE96A2A8A808B8B200A6D18CA37FE0A9ED81297ABAB949E71D93
            79CAD858DBCD5D7A62621285978B3136A679ADE3D93BB3F4FBE6408B29ABC730
            BB17077273F872CF18ACA0E8E848ECDCB1DDAED0660353E454F8F9F2766B6B07
            A299E568A5B318AC31E8AAAA6AFC73BFCD71811359629FC6127CD2D56BE578F2
            A48FAD7062F0A0E7211F674BD142E87266E9761B59DA6C60B9DC07278E1FE6ED
            277DFD282AFAC3A2071B42D32AABB8E42A72F6EE62B9F31CEA1B6E995C79D914
            98F4094B0642986B932E5DBA82010B17EA1111EB782675F7EE3DE8D8BF8FF764
            EBCFD1345457F7375EBC78617F60C3B530551F0B0A0AAD5EB3AE5B17360F9844
            8582FAFA5B3C67379CD36D0E4CCACECE425464046F8B9114180253894798B349
            2AD5306EDCAC15C5CD2D06A6EAC5E14307F9329144A95F2D73414B2D61085C5C
            5CCAB3392A2119BEB316C3CDAD9A87030256F13958B006553D2A2AAB2C2AB31A
            02175E2E417FFF000F9054245C1B16AABFCE5A37B73AF1A0F1BC974557777777
            BE4F95893AD6214A28DE549E3507585068680832B7A6CE2BEE5BEAE6A2AC9614
            0A05EBECCE79457802EFEEEEE1AF538658E768DF546033054C22D74E4C8C4372
            5222A452CBDD5CB4370FD4A1CD9B13796262ECFF8A982AF62D062CC8DBDB1B19
            E929FA221F89B2BBFC5F2FD81658109564E2E262F96241EEE3233A30C58BCDC9
            89DCDA545F2391F7FCF853BE7D800D450B0A8AB6FECCD529AA53CE3C34A4B218
            981215AA96526140100D991B376B780DDBEEC0E6C81430C588CCCC34042B83F4
            C7262727515D538FAEAE07663DC7A18105F7A5D7AFC29B469A8A9A9B5B70EB76
            23A6A6A6CD7E8EC302CB649E56BBEFB201A642A0E19C6BA9FB2E1B6041D6BAEF
            B20216C37D1D1A98128ABC83FB79BBA6561CF77568605B6905D8D9B502ECEC5A
            01767609C0EFDC471EEFDC673CEB2512492D7DA82597FBC3C5C92C3DC72CAB79
            F5A1569AE4FF4FF162D8B9EFD9B68B6D3E563EC3D1446FF9AEB1ED3B16ABDAFF
            056363F55F83FF8DDA0000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000004C24944415478DAED9B496E1B4714868B1267
            8AE2A0918A626423CBBE41B2F3C25A649B1B2417C8D6F001825C26D966212FBC
            4B8E104340E055448A2229CE1A38A4BEB24B6A3064BB8712BB49F1070A2235A1
            BE7A43BDF71A8CFCF8D3CFE2F4F48F2321C4AF72BD966B532C975A72BD93EBCD
            C9C9F76791C3AF9F1FAFADADFD7574F4225F2A1D88582C16F4068DEAEEEE4E9C
            9FFF2BCECEFEBE1A8D46DF02FCDBF1F1CB1F9E3DFB26E8BD3DAA3E7EFC47427F
            F81DE0E6AB57AF3797CDB293C2D2EFDFBFEB003C96BE1DF47EE62299ABC40A78
            99B5025E76AD80975D8F069C482444329910F1785C44A351B1BEBEAEBE3F1E8F
            C570389477E240DCDCDC88EBEB6B31180C1613381289884C262D36363614A453
            01DD6AB5C5EDEDEDE20027127191CFE7FF5787CBDA555538C3E148BE1B732CEA
            3062B1A83A20AB7ABD9EB8BA6AAABF093570369B15B9DC438385DBB2F96EB7A7
            60793F296093C9A4F208BE6AE1DEF57A5D5AFB2E9CC0F97C4EB9B016A0CD66F3
            B3459D69D23BB070AD5693316EDEC57D015B2D8B15EBF586E8F7FB33A148649D
            4E77AACB62710E2F93C9DC435F5C548D2734CFC06C7E6767FB1EB65ABDB44D3A
            070725217B6E09DC51713A4BB95C4E1EE4278F211C809E16127305C61A7B7BBB
            F799B856ABCFB4ACD6E1E157EA2B2E8F27D8696BAB2852A9947A4DF66EB55AC1
            0213B3B89F5300B7C0780207CADD8D75CBE5B2AB9C601CB854DA579B198DC6A2
            52A9A842C22430227B170A05F5DAA4955D035B63B7DBEDAA784CA753B278B8B1
            05770B4CD8ECEFEFA983257195CB95E080597AF3BC2E14F26A5324AE59D06E81
            91F5CAAB542E54129B3BF0A4B000162781D9417B01A620D9DEDE52AF1B8D862A
            640207760AED0598FF4BBE40ED764715347307DEDDDD511D909DA6C5DC03705F
            958E4EA5FF8E7CD1685CCD1F9844F2A54E08EB026C2D1874E181B0309676A274
            3A2DE8310683A16A27E70E4CF69CEC720021D638086071E9C992908D178B85FB
            F76EA04DCA770C034BFC52F8CF820D13B46B6032A776E95EAF2B4BC0B4BA96BE
            043B0B9AB8243E430B4C614F818F483E141C5445FDBEF3514D9096760D8C7549
            5C88248255BDC80A4D97757959934D0377EE585E3FAD471BF7788A6112949E52
            54AB55CF8D3A256926B3A15A464497A4859B036E7ADCE309D85A4F9BEA596907
            ADC00858A04DC6B8E72C5D2C1695859089A2C00A0CA8BEB311EE4D9362C2CD3D
            035B7B56E4B7F4B30213CFF1784C8D90AC77BE0937F7750F536212CFDA1A4C3D
            B0B4970D598149842444122437422AF530D5F4EBE6BE0B0FE2998D6A68EE639A
            756A6637713D0D588B0449AB682D69BDBAB9916E89A13A316D1DC203CEDDCCC6
            D91496B13B003B60846B5303F87573634F1E3E6D28AB3635596B6B316BE610BC
            006B9133B0B61EF221371311E30FD3D810B36532F864576537DD740ACC616E6E
            66D524441F2CDE747E5E0E06D82A5C9C6C1B8DC6548CE37EB362CE0930BF8375
            F5CD80F83D62D9E9F82734CF87ED80C911B95C5E3D7ED5C2AA5C832447370A35
            F034F725F1F1B8A6DD6E7BBAFE420B4C08F875DF850126EB5A139E57F75D1860
            2DBFEEBB50C026DC37D4C0C42A236064CA7D430D3C2FAD80975D2BE065D70A78
            D9A5819FDC873C9EDCC7785EC8EEE4CF27F241ADEF229F3F8A772C7FF68B5C27
            726583DEA461B5E53A95EBADCC551FFE033BFAF4AF8928907C0000000049454E
            44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD97200000AB74944415478DAE59B7B8C9C5515C0CFCCEC
            6E1FDB52DB525AC076DB52680B7DD0DA2654B0ADA5101A135F053501C4441345
            89897FC8A33C540A285A4C7DC446A411108D514B408444DB62112B220A2D5DFA
            5A0AD2C72EA5A5BBDB279DDD99CF73E67C67EFB9E7BB333B333BDB1ABDC9CDF7
            CD37DFDCB9BF7BCE3DAF6F2635F6B711603B1FFB77B02FC67E06FC6FB5C3D8D7
            61BF157B4B0A8127E3C9DFB1BFEF74CFAC9F5B07F64B08780D9E7CB256A3A62A
            BC3F3AB5D08F137027F4418D35602A15BE2EAF2D9C7E1D45E1EB356E4709B8E2
            F1439029B99E72EFA7CC3D1A2CD26091FF5ADF53EB5611B005B59029733D6D24
            DE0314434525CEFB0BBC6C602D2D0B9A56C774600174F3241CC3E5CDB13FC1CB
            024EC01AD08CBA9E36EF156B1A321FD3E402D7F402D402BA57E0106C3A061289
            D2B54C00D6D30691161869460E301FA96EEEA9157449600D2BAA2AB002E89DAB
            85E80D58AB734E81E614AC5CAF257451600B9B52EA9B31A019B30069F319DBB4
            840B6A1CC3096077E45F1789D7023A085C4C8D35585D3A3E2AE042C7EB338763
            48330A60EA3080B18D00C31B783C0279E704C0BF8F02BC720860E30180D6E30E
            B270CC3BF85C6434C040D714B8182C01D5A5146CFCBA018F1F1D0B707513C039
            83CBFB72FAE21711FAB137019ADB7DD88294F36E21724AEAA221D54839012CD2
            2DEC43B52F05B23ECDE07572C43E1D25FAF58B009A86F8831FEF0678E30806B1
            5D00D91C7F961683EEAB33BABEB61560E5368CF4BB1CB886D6128F8CB4FB0CAC
            DD4E08B63EED54F9BAF300BE30C9F9DBF710ECCF6F33C0AE233C411B69356400
            3E3012E0AA7301668F70EFBD8DEA7ED726806D9D3E7477DEC113A89674A5D045
            81B53516B525A9D6C7477AFDD50B01968E739F7D164157B7007466C3E1A78EA8
            A44DC128FEA6A900138638ADB8F95F009BDB1DA8061615CF5729E520B0A87346
            19A58254D37C24E8EB49B2E7F367B2F8CDDF6DE6FD186A5331F19C8EFD997DAC
            B2D66AD0B85F9ECC12A77614A16FC484F5CD232CCD2E05DEA3EA91F3E155035B
            7516A324526D888167A13AAE9CC3F712EC3294C8CE234EAADA7550FBD57C80A1
            F5004FEC06787067788274FFE72F00B8A6895FBF750C5F6F4489E758B25D46D2
            B92AA51C04EE3154B1EAD6A75D6FAC0378F8526789EF7D15E01F075DA2601301
            6A4F5DCEC7F56D002B5E4B4A582FD29D33013E349AAFFD62172E500B8366F34E
            DAB98011EB1BB0A8B14838962EF5CF4C60F5A346C669E5563F2BA2A6A3230DBC
            8E809B93FB582FD4E07841470C60C8AB37001C782FBC97BBAB50EB04B00E177B
            0C550C3B0827F36B54CFB306B235FEE20BBC276DE8A813036A7F50C0DF6B4E4E
            50CEE5FE25B8976F9EC6E78FA29457EDF0F7B27655F99A02AB7D4BFD923301BE
            3F97EFFD23BA9D9FBF8EEA7716C0CB18357564FDA0A027E3C1F6740C4CAE8A8C
            9BC4CA21B5A6460BFD9B052C658AC4966EF0D53964B1FB0C2C06AB41F5B9083C
            6724BFFFFC7E0E386E9CCCFEF31B9B58F5243ED66AFDCC62077C7F33BFAF016D
            1584E671D314804F8DE7EBD73E8F46F1703220F1F6715F803301090FC8387009
            4046A204EE99053066900B1A08BA12E09091A346B1F88A397CBE1C0DE393BB7D
            97A42D75AE5AE0B40938EAB48411788002164D3813A1EFBE1860740C7DC72B0E
            5A06B62A9D8364A4A4F73E351AF7F78BF8FC976FE056DA1A03E7FDD8BA66C0DA
            2589741F9C8791D1B0D203EEC33DF79517796256C21BD0AADFB7C54D5043EB84
            5FDADF96F0F1893DB8A89B7D35D6C055AB742960922E0510E7F692091D3CC901
            8318146A6B16B2BBA16FBA1FFDF0BAD640012F4A5639AE3C87E7B107D3C997DE
            4D4AB5DF801B62E841D8CFA8F7557A18BEBE7D06EFE3771176D9CB007B8FBB89
            505B3486DD0CF9F8027433BB2800DF7DE9C8494750366DB455929A0087F6B035
            5A142E92D11AD7E8C34A562320640C2F3F1BE0160D8D925EDFEAA0ADFF166B1F
            45BDE7C7958497092B9D29E287E978D968CE6A90BD20A18508F1A50B18F67681
            35EAA6177331DE7FAB82FED17634687B15B4516BAFEC137049A15451C62A1BD8
            BA251D565E3B91FD2E358A9836B5F33EDBF80E40DBF1A441D19390F0F30ABCFF
            B6692E65A4D87A7D9B9B90AD5A860A0162A9BB23A8B8DE5556E05128E1A058C7
            37B2E1A2B6F910FB5DBD0FED3ED37B4BB60A815EA9D49B927DB2C077CEE08CE8
            91D7F95ADE2C9E40770554BB12E88A6269EA0FCCE5688B1A25EADB3B9D2AE6F2
            0658ED636B1F3E8C86EC134D008FBFC5D7964D77935C1B87AD14AE5AD8508A58
            89A44B664B16988EB311F607713C4DD5C7AFBDC4598DAE2B5BF701E00A0A9954
            B2DA7929C6E3B74DF7274645804711FAE97D7E7654482002EADDF39D06BA3C60
            70AE2913572435F85D33D9D550A32482524480E27E521BAE4CCAAFA4D0B81AF8
            44376765D2A82EB60A8DDB6B9D0E5457407425C4735545A45CB2E2A1A52CF52C
            82A62C66F50701460DE4CFFD0ED5F2A19DC9A70692AF466AEC8C81A6712F53C0
            DFC6486CD250808F8FE3EF920953B042F5B243591FB85861A09894CBAA6949C5
            5202119AE445C338551C1C4BE3AF68A91F40CB7DB83B29E1488D6DEBDBD4A9C2
            21C094286CEDE09CFB86F3780B69357F18D5FCA9BD5CF615696745CDF3A6DE55
            2E70A26AA926A84BB5B346606C3C9BCB3ED40EA13F5E8D13FA532B4F48BB27BD
            985AC2348E06BEE755B6D252549889DF71FD448EE4A4B560AAF8C36D005B3A18
            B64B75BDAF2B020EA9608F6F560B4081C8DD186D3535BA31DAB39C28FCF320EF
            BDF6933C316DBCEA944ACF37C0DB3BFD67CB743F554148CD1B949A531594AA21
            074EC4528E7C2987EA5D251FB5A4A188754DFB466720F6CFA2FA7D7A829B906D
            B7A00BFBCB7E7F3C91B005DED1E93F88932D42CFA8A8F03F6F941B77CF31806B
            9E7392EEADDE55D6B325BDF742D0E2BB47E08496367108699F2F11F073FBFDB1
            445BE68F717E58249C0E00D3E290943F36D6456EA44D4BD625F77145C02168FB
            5CD8AABA4C40E2E48968692F44C3367E0827196B76E3DEEC4806360DF11EBE63
            067F7EB9002B95A6495379E9BA895C6591B609A3BD15E8125B3AFBA8D2020C25
            A0ADAADBE7C13A1990660DA14878C118074C61E6B64EB780A4293720E8B4E16E
            1C4A567EBA9323B29379178C64B5D1820A818B41F7849F907C089E0A8C1199B1
            B4D597271A0B46733043ED5B9B591306D4F173AB8FBC9FEFA746304FEEC1787B
            17BABF6CC02DA938BB6CB7D42BB4485BED6DBD10FA332265FB34D20B68E2B85A
            802921A1A2C2E726F9EABB0533B39FECE0C84B071E5D2AFAB2BF1E282BF02817
            1A941AEBA3F701A136D15B4F681987AA0B51C2DFBC986F6F431773B6F2B9E4DB
            1F6AE1A71C3AB4CC99632889083D34AFF8775AA5C053817BE57EEBE2F4134902
            5E3ECBFF8CA8EF63A8BEC7BA95448BE4C655250F9580877E72E8BD1F58206BE1
            45A517A11BBB570153AEFDE3ED9C8D49552317F9C9824854FF20461BCAB2D2C3
            6AC02D7CF09E126E4D7E06F1B3797CFF2AB4BECFB6F993CF19A92632A380812A
            56E3AA1AB8187CE23D63EC32A9B02F2769EB705266651311FB0B804A7FC75513
            E07216C4FE4846FBF2BA94BFCFB5844205BD52522D057B4A818BFE74D1BC96E6
            3D89046775134F2DCA043DE5C0A060AD4B13E96BFF2D27BA305F4C7DCB853D65
            C0025DCE6FAC358095A42C004065523D6DC000BDBB330D2C2791BAD66F3F2EED
            4F680D1E824D40D7005403F7E94F1E7D052FB7D5482A476BFE379EFFF256F81B
            CF143C7901FE3FFEA8352F15FF158F1E91DD87FD0AEC434FF7CC6ADC30A184B5
            D89761DFF11F96C6EAA5E4CDA60C0000000049454E44AE426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000006704944415478DAED9BF953535714C7BF3161
            4981106CB5B2AB6C8A06103B02411405F7AA95D61D9D69ED0F6DA7BFF5974EA7
            3F763AFDA533FDA17F403BADB46A7144848A225094B0D40DA5229B880A28B884
            10088425BDE7DA97069490CD2444BECE1BEEDB7CF79373EEB9E79E374FA4484C
            83FA6B550C80EFD896CD36193C4B1AB695B2ED4BF937CA1651F809431CC6466A
            A4A53FCABD1B4A20D2695CDD4187CA209541BF7C1374D99FAB21F64A25E07CE9
            B91F727CEA8EBBBA6FAF54C3CA5CE8D67F769280FB02BFDF2CF334CB4E1659BA
            EF8BB35A023630DF76757F9C2216AB300BECC99A05F674CD027BBA5E09B04824
            42F082B711191981F9F3E72128281052A9949F1B1F1F8756ABC593A7CFD0D9D9
            8DBB773BD0DFAF9D99C073E6CC41FCD23824262E874C66793ADED1711F57AFD5
            E3E1C34733077801B368E6DAD5CC9AF209C7474646D0DBFB043A9D0EA3A3A310
            4BC490070662EEDC20FE0399AAB9B91597AAAA313CAC776FE0152B12919AF2CE
            FF900CAC8575BEF176331E3F7EC2DD78B22412092223C2111F1F87B0B050E371
            72EF929252F4B2FBDC12383D3D15098A65C67DB292AABA8E5BD4529177AC5993
            8E3799D5F90FC6BCE24C51C92B7171BB8093931391B2EAB9654747C7507AA11C
            EDED1D2FBD9682584868301A1A1A99CB0EBF705E2C1623234389A54B62F9BE5E
            AF47FEC942A8D56AF7000E095E809D3BB719614F1716E3D1A39E29AFFFE8C35C
            F8F8F8A0FE460354AADA29AF53A6AD62414FC1DBCF9EA971E28F53181B1B732D
            305963DFDE1C63243ECBC6DC549615F4E92747F8DFA6A6169495574E791D4D69
            1B37ACC7E2C50BF9FE95ABD7515777C5B5C00909CB90AE4CE56D1AB317CAFE9A
            F61E4B81493E3EDED8BFEF033E7793757FF9F5985531C1A1C06481DC837BE1EF
            EFC7A3715EDE710C0E4EDF196B80494BD8585E9799C1DB8EB4B2D5C0A1A121D8
            B17D0B6FDF6A6C4255550DA2A317E3DEBDFB66C1AD05A639FAF0A17DDCCA9AFE
            7E1C3DEA98F293D5C02121C108651BA9A5A58D47DEB56C4AD168342828288676
            60C021C024A532058909CB79FBD8F19378CAD251A7034F969FDF1B788F456B0A
            60E6A06D018E080FC3B66D9B78BB9CDD739BDDEB726092BF9F1F9BA2B69A85B6
            05987ECCC387F6F3767DFD4D9ED0381DF8FD9C1D7C0564F63F55F7E1F763F930
            180C2F00B7B4B6A1B4B4C2E20E0AF7519A5A5171D1F9C007F6EF4660A0F99590
            563B80A32C7A9BE6D01F1F390C2F2F2FDEA6698CA6334B14131305310B601A4D
            3FBABA1F3A1F98A2A7979764C231A9AF948DB58DDCA50706065170BA087D7D13
            6BDCD4F1ECAC4CE3BE35D08E94DD63D8D7D7173B776CE5CBBDA96005C5C64623
            6BFD5A97425B0D4C91531E14C8DB8D8DCD886596A395CE74B05341575656E19F
            5BB7DD17388925F6692CC1279D3B5F86070F3AD90A270E77DAEFF27166892643
            97314B3739C9D25603CB64013878600F6F3FE8EC4261E19F363DD8149A565945
            C5E7B075CB06963B8FA3B6EEB2D99597538149EFB264209CB936E9D4A933E8B6
            71A11E15B5886752D7AFDF8481FDDBBC29DB788EA6A19A9ABF313434E47A60D3
            B530551FF3F30BEC5EB32E5A14390198448582DADACB3C67379DD39D0E4CCACE
            CE444C74146F3B22293005A6128F3067937A7B1FE3E2A56A87B8B9CDC054BDD8
            B37B175F269228F5AB662E68AB254C818B8A4A78364725242A360872849BDB35
            0FCF9BF7169F83056B50D5A3BCA2D2A632AB2970C1E962747575F3004945C285
            9111C6EBEC7573BB130F1ACF5B5874F5F6F6E6FB5499A8611DA284E265E5596B
            8005454484236375EA84E2BEAD6EEE90D5925C2E679DCD9A508427F0B6B676FE
            3AA587758EF6CD053673C02472EDA424059257244122B1DDCD1DF6E6813AB472
            65124F4C4CC79DA9CC15FBA60316E4EFEF8F74658AB1C847A2EC2EEFB713CE05
            1644251985229E2F166401010E07A678B13239895B9BEA6B24F29E9F7ECE730D
            B0A9684141D1762E73758AEA9433F7F4F4DA0C4C890A554BA930208886CCC54B
            2A5EC37639B03532074C312223230D61A121C663838383A852D5A2B5F58E55CF
            716B60C17DE9F5ABF0A691A6A21B371A70F9CA35E8F523563FC76D81A5525FBB
            DD77C6005321D074CEB5D57D670CB0207BDD7746013BC27DDD1A98128A9C5DDB
            795B55ED18F7756B60676916D8D3350BECE99A05F67409C0AFDD471EAFC7673C
            69B9D0653DFF8C6709C646AAF9875A37CF423464D93BA29922836F00F48ACDC2
            875A69A2FF3EC58B63E7BE65DB06B605D8F90C771359F03CDBBE62B1AAE95FA6
            2155147CE125940000000049454E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000006524944415478DAED9BF953535714C7BF81B0
            A4400854AC1016655594D58E6CA254512B56ACB4EECB4C6B7F68FF834EFF804E
            FF8C765A69D5E288082D8A4051D6AA0852D945544008420801952DBDE7B62F06
            348124CF2444BECE1BEEDBE6DD4FCEB9E79E7B9E4F12979086CEF6C628003FB0
            2D9B6D723897346C2B63DBB75131499D124F4FCF18894452B72A40A9F091FBC3
            D5D5D5DE1D1455B3B3B318D78C6058D5A7D6E974A9045C10B03A384FE11760EF
            BEBD558D8E0C32E8FE8B043C161E192F7736CB2E1459FA4157B3968075CCB7ED
            DD1F9B88C52AAC003BB356809D5D2BC0CEAEB702CCB23604AEF9006161A158BD
            3A007E7EBE90C964FCDCDCDC1CB45A2D9E8D8CA2AF6F000F1FF6627C5CBB3C81
            5D5C5C10BB210609099B20972F3D1DEFED7D8C3B8D4D78FA7470F900AF6116CD
            DABE95595331EFF8F4F43454AA6778FEFC39666666E02A7585C2D717FEFE7EFC
            0732544747176E56D7E2E5CB29C7064E4A4A406ACA87AF20195827EB7C6B5B07
            86879F71375E28A9548AB0D010C4C6C6203858A93F4EEE5D5A5A0615BBCF2181
            333252111FB751BF4F56AAA96DE0165DAAC83BB66DCBC0FBCCEAFC07635E71A5
            B8F4ADB8B855C0C9C90948D9F29F656766665176BD023D3DBD6FBC9682589032
            102D2DADCC655FBE769E162E9999E9D8B03E9AEF4F4D4DA1E06211D46AB56300
            0705AEC18103FBF4B0978B4A30383864F4FA2FBF38090F0F0F3435B7A0A6A6DE
            E875E9695B58D08BE3EDD151352EFC7E89AF72EC0A4CD6387A244F1F89FF6463
            CE9865057DF3F519FEB7BDBD13E5155546AFA3296DF7AE1D080F5FCBF76FDFB9
            8B8686DBF6058E8FDF888CF454DEA6317BBDFCAF45EF592A30C9C3C31DC78E7E
            CEE76EB2EECFBF9C332B26880A4C163879E208BCBDBD7834CECF3F8FC9C9C53B
            630E30693D1BCB1F6565F2B69856361B58A90C42EEFEBDBC7DBFB51DD5D57588
            8C0CC7A3478F4D829B0B4C73F4E95347B99535E3E3387BF6BC7D80838202A164
            1BBFB9B39B47DEED6C4AD16834282C2C817662421460527A7A0A12E237F1F6B9
            F31731C2D2519B032F9497D77BF894456B0A60A6A02D010E0D09C6BE7D7B78BB
            82DDD3C6EEB53B30C9DBCB8B4D513926A12D01A61FF3F4A963BCDDD4748F2734
            3607FE2C2F97AF804C49AD1EC36FE70AA0D3E95E03EEECEA465959E5923B28DC
            47696A65E50DDB031F3F7608BEBEA657425AED04CEB2E86D98437F75E634DCDC
            DC789BA6319ACE96A2A8A808B8B200A6D18CA37FE0A9ED81297ABAB949E71D93
            79CAD858DBCD5D7A62621285978B3136A679ADE3D93BB3F4FBE6408B29ABC730
            BB17077273F872CF18ACA0E8E848ECDCB1DDAED0660353E454F8F9F2766B6B07
            A299E568A5B318AC31E8AAAA6AFC73BFCD71811359629FC6127CD2D56BE578F2
            A48FAD7062F0A0E7211F674BD142E87266E9761B59DA6C60B9DC07278E1FE6ED
            277DFD282AFAC3A2071B42D32AABB8E42A72F6EE62B9F31CEA1B6E995C79D914
            98F4094B0642986B932E5DBA82010B17EA1111EB782675F7EE3DE8D8BF8FF764
            EBCFD1345457F7375EBC78617F60C3B530551F0B0A0AAD5EB3AE5B17360F9844
            8582FAFA5B3C67379CD36D0E4CCACECE425464046F8B9114180253894798B349
            2AD5306EDCAC15C5CD2D06A6EAC5E14307F9329144A95F2D73414B2D61085C5C
            5CCAB3392A2119BEB316C3CDAD9A87030256F13958B006553D2A2AAB2C2AB31A
            02175E2E417FFF000F9054245C1B16AABFCE5A37B73AF1A0F1BC974557777777
            BE4F95893AD6214A28DE549E3507585068680832B7A6CE2BEE5BEAE6A2AC9614
            0A05EBECCE79457802EFEEEEE1AF538658E768DF546033054C22D74E4C8C4372
            5222A452CBDD5CB4370FD4A1CD9B13796262ECFF8A982AF62D062CC8DBDB1B19
            E929FA221F89B2BBFC5F2FD81658109564E2E262F96241EEE3233A30C58BCDC9
            89DCDA545F2391F7FCF853BE7D800D450B0A8AB6FECCD529AA53CE3C34A4B218
            981215AA96526140100D991B376B780DDBEEC0E6C81430C588CCCC34042B83F4
            C7262727515D538FAEAE07663DC7A18105F7A5D7AFC29B469A8A9A9B5B70EB76
            23A6A6A6CD7E8EC302CB649E56BBEFB201A642A0E19C6BA9FB2E1B6041D6BAEF
            B20216C37D1D1A98128ABC83FB79BBA6561CF77568605B6905D8D9B502ECEC5A
            01767609C0EFDC471EEFDC673CEB2512492D7DA82597FBC3C5C92C3DC72CAB79
            F5A1569AE4FF4FF162D8B9EFD9B68B6D3E563EC3D1446FF9AEB1ED3B16ABDAFF
            056363F55F83FF8DDA0000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object wmibIncrease1Second: TWMImageSpeedButton
          Left = 1142
          Top = 9
          Width = 60
          Height = 60
          Action = frmSEC.actControlIncrease1Second
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000006824944415478DAED9B0953535714C7FF4181
            0092514B704308DB2022844D0415DC58ACB64A85B6A84551A79DF61B74FA013A
            FD066D6DD5568BB8401410A880226A5D0694D6714331A21016410D01D914E83D
            575F4A6C300B090991FFCC1D6E5EEECB7BBF77CF3DF79C7B79A2707902EED7D7
            0503F88195645624702C6958A964E5DBE090A8FB22B1581C221289AE784917CC
            F494CCC6B469D36C7D8316D5D0D010BA35CFD0D9A1528F8C8CC4137081D4DB67
            CBCC59525BDF9B55F5FC593B836E51107057405084C4D17AF66D514F2B1B6EF4
            10F008B36D5BDFCF8488F92A4C013BB2A6801D5D53C08E2EAB003B3939C1C767
            017C7D7D30C75B0A894402B1D8957F4773614FCF0B3C7DFA0C2A550B940F1BD1
            DBDB3739815D5C9C111E1E86F025617073131B7D9E52D9886BD7FF4667E7D3C9
            031CE02F43626202DCDDDD758E0F0F0FA3BFBF9F9501B03816CEECA1B889C570
            7676FEDF6FDCBA750797AFD4E0E5CB97F60B4CE61B1FBF14F288253AC79B9B55
            A8A9AD435B5BBBDEF35C5D5D10152947686888D6DC49CF9FAB515A560E8DA6DB
            FE800976EDDA240407056A8FA954ADF8F3740506078DEF251A06CB13E2F8EF91
            FAFAFA70B2B0146AB5DABE8057AC884704BB594155E72EE0EEDD7B7ADBAE5E9D
            086FA9177B18957A7BCFCDCD0D9F66A6C3C3E3F59020C7965F50C8E1ED0298C6
            6C5ADA3AEDE7E25365CC8C5BF4B695483CB17DDB67BCFEE8511337597DA21EDE
            9A95C1BD3A8986C5A992D37CECDB1498C6DFD6AC4CDE2BA4B367AB517FAF61CC
            F6F3E6CE417AFA47BCDED2D286C2A29231DB528ABA2B67BBD6A919FAED09015E
            1A1B8DD8D8D7ED9BD95C5A5C5CF6CEF6A60093E6B2F69FBC694FA69D7BF818F7
            F6360126B3CBD9B98DF5B22B37B5FD070E197450A602930898C0491595556868
            50DA06D8DFDF0FEBD39279BDA9A9998F3143320778D6AC99C8FA3C83D71F3F6E
            4649A9E1EB580538297139C2C242795D71A218EDED4FAC024CCAD9B99D476CAF
            5E0D61DFFE8316316B938133333643CAA617BAF8CF7B0F18758EB9C01B3E4C85
            9FDF425ECFCF3F890E0B849E2603EFD99DCD62661716F0F7E2F7837956051EED
            1C2D358E4D06FEE6EB3DFC2F8580478E16980CDCD5A5C1E1BCE3469D171C1C88
            E475AB79FDFCF9BF70EBF6DDC9014C1E7DF7AE2FB49F8D757641410148495EC3
            EB172E5EC6CD9BB7271EF8AB2F737870F0E2452F0E1E32CEA44991F27024B078
            D914E8A8C8089E9890CEB000E49E050210938169AAA0298312F9BDBFFC66D2C5
            4C854E4D598BC0407F5E2F2C2C414B6BDBC403D398A2B145CACBCB87BAABCB6A
            D03BB2B76A93099A964CC9C02C06BC2824186BD624F17AC303252A2AAA4CBEE8
            DBD0B7EFD4A3BAFAA24E9B191E1EC8CECEE2F5F6271D50288AC60D6B1630250E
            3B776CE3E398CCFAD77DE60504313191885B1AC3EB9AEE6EE4E61ED3F97EFEFC
            79888808434747275AD954660973360B98B46AD54A2C0E0DE1757224E450CC11
            F57408B398D3E567AD92EC5B0CD8D373064F0F85DD464B3914BB05264547CBB1
            2C2E96D729D6CD3B92CF52B91E5BF3580F98D2C48D1B52F9FAB300AD3851C4D7
            9B2D215FDF85701FB5D4FBB0F1310606066C074C2207B6E9E30DF0F2FA407BAC
            B6B60E35B5D7C77553DEDE52646CD9A473EC384B1E2CB16E3DEE655A825E9F96
            C2BCEA5CED31EA0985A2D8E01C4D60E4AC46CFAF6439043BFA21DA15B0709394
            D54447C9211289F83143D98D309FBF1D78D0F9B48341A2E5DF44967FDB1DB020
            0A39C991D1AA88A11B14523F1AF3C78E9F78E743B15B6041346D9199BECBC98C
            0798A2B08040195C5D5CD1DAD636E6F2F084018FA5D89828C8E5AFB764680E17
            E6F1C1C141FE57ADEE42615129F3F8AFC604A687999991AEB33DF3CF8D9BB874
            E9AAFD016FDEB451C7B9E9D31FB947D1DDDD33267014F313F1CB62D1CA029DA6
            6615968485A2F24C35DF7AB53B60CA7C7C16CCE78E49C6C6B9BFCC8FAF3BD7D4
            5CE3DFD36A48EBA8CD377DC0C22A485F5F3FAE5CADE18E51B008BB031E2D73C7
            30CD08B40A121020E3C7E9FC728AC58D4C536D062C644B044130C6020BA28501
            326DDA87A28C8A36DEEC1A989C0F2DDF281F34E281F2A1D1C0D3A74F475C5C0C
            EAEA6E80A67C4A55493FFEB4CFBE818D913EE055492BB078F1229E8BD3D8A505
            C249D1C3C64826F3D52E12945754F13094E6E0958909F05DE8C3A7355AE73E57
            7D813BBC490F6C0D4D013BBAA6801D5D53C08E2E01F8BD7BC9E3BD7B8D67114B
            D92ED38B5A12C96C3839584F0FB39ED5FCF7A25682E8CDAB78B477F23D2B29AC
            78DAFA262D2CFA5FC70A56BE63BEAAFE5FC25FD85F2D7BEBF60000000049454E
            44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000004F64944415478DAED9BCB52134114863B10C8
            8D5C08091040CB0DA26FA03B17B270EB1BE80BB8B57C00CB97D1AD0B5CB8D347
            D0A2CA722124DC929000090924F6D7D871C4814C2793CC10F2577539C00CF4DF
            E79CFF5CC6045EBC7C2536363EAE0A21DEC9F554AE84182D54E4FA24D7EBF5F5
            679B81953BF7D7262626BEAEAE3E48E5724B626A6ACAEB0DBA8A66B329F2F96D
            B1B9F9ADDC6AB51E41F8FDDADAC3E777EFDEF37A6F03C5CF9F3F24E9EF1F207C
            F8E4C9D3C4A859F632B0F4E7CF9F8E20DC96BEEDF57E8602A955624C78943126
            3CEA18131E750C8470201010A1504884C321313D3D2D82C1A090959CFA59BBDD
            16E7E7E72A279E9E9E8A5AADAEBEBE91842135331313B1584C4C4E4E3A7EAE56
            AB894AA5AA0EE1C6108E442222954AFE47148BCAFA552D80F539186D712B8E8E
            8EC5E1E1A17AC6B78421904C26A46567FEF97EBD5E17D56A55BA6DC3F63908C7
            E333CA1BACE4B1F2C141519C9D9DF98F3064676753221A8D76BE475CB2616D51
            27E0B038347E1F20A6F7F6F60742BA2FC2B8B0D5B2A552491C1F9FD8DECBC120
            6057598F50989FCF764202D2BBBB7BAE0B5ACF8489D9B9B974E76B2C8275ED80
            4A2F2E2EA86B5C7D7FFFC0F63E2CBCB030AFEEEF76EF500913736C4C5BA3582C
            899393932BEF0F85A645369B55D71C0A877315209DCB2D76E2BADBEF1E0AE144
            2221575C5DD7EBA7D20AFBD7DE6F42F8F2FDB874A1B0E39A721B13B65A804DE4
            F385AE02654A18703FCF8162B128AD5CF38670241296B13BA7AE9DC6582F8499
            BE1036267F67208453A994AAA6002ADA6834BA3ED30B6190CBE5A44E5C78D2F6
            76DE15B736264CEA20BDF0C7B7B6B61D3DD32BE14C664ED6E3E13F87BB2B0FB7
            FFD2D398F0D2524EC52F6242FC0E9230C2884002B7E2D898F0CACAB2FA971270
            6767D798304507AAEB04D16844A4D317B9BE542ACBA2E6F86610C623F00C0DA7
            2264255C2E9755733174C2CBCB4B2A3599B834A05148269346A4E3F1B8AAB181
            5B05883161520529C344B47A254DE94A090BAE2B5D074A389D9EED7447C4A269
            4763429A024797AFA425930ECC35C2B15854763EB3EA1AD5443D4D7199346284
            285901510803723D39DF0D18134680D80871DC4F41604D3976CACD4C8C02E762
            F6D570C59D7B220CE86D995400840441E905589AF0C04B9ACDC14C385C211C0C
            4E4AF15AE84C28DC1214DF1206D694814BE392C31CB70E9D30D6A56B62F6AC49
            232C6E8D5AA9A1691C34985F7BA2D2562060994C4636137F5FA4572A153563EE
            07342734295650D5B971987D8F69214D8180AA6A6009ACDD2D47438C7BAC96C3
            73A8BBAD87E82BC27A93C434AAAB85AC5B77A3F3B95DE1A1E759D4D2F4DFBE23
            AC41C9497EA51CECB6413D17BBAE09B11639BE24AC41DA6AB5DAD78A4C3F84A9
            C238543C817468921287FABA14827A708FEB6BF7D707433C93D375E566479899
            3582667D3D53AD1EA97752BE239CCD66FE11373B140A0549FCFC4AC23AFF6BCB
            52F151E939B5F25009E38A10C6B0E170444D4029564865000B5B5FBED911D643
            01FD1CC26852CB7BF63F007A8D61C2801655F7C9A66F1B3D247CD12D31896422
            E994B00684716D62DAA47DF48C304A9E4824D5DB7F9653C2FA7D34420574CFFC
            EBD796BF093B811D61DD9A12B72CD4FA4658D80910353D24D0718AF0F15E9AE6
            026BA3CE4C4B7C1FC35E614C78D431263CEA18131E7568C2B7EE431EB7EE633C
            0F6489F6E5967C50EB71E0CF47F1D6E4CFDECAB52E57DCEB4DBA0CE6C61B72BD
            915AF5FD370F78E3AF1AFAC6790000000049454E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD97200000B034944415478DAE59B7DB0555515C0D7B9F7
            BD972028E08B0253447D203C14B440C008128D29756AC81299E1510AA5025998
            4A68991FA3961F69FDE1584D53D98C8523334D6533E2AB88E8F1210F44A508F9
            4AE24304E1F1F8F0DD7BCF6D2DD6DEF7ACBDCE3EF7DD773FACA933B3DFF9B867
            EF737E6BADBDF6DA6B9F179CF55C1E706BC2F230962BB09C06FF5B5B079697B0
            2CC2B22540E0E178B00A4BBFFFF49BD5783B84653C013F8F07D3ABD56AD0C3FB
            F3EF2DF452023E0C1598B1040C023FB43DD770F63C9FF75FAFC1D649C03D6E5F
            4306F2BA390FC4EF0E4C5EC19A3F798F006A01DE23601FA8844C896B314108C0
            BC812A765C2BF09281A5C62468CA5C4B0960795DD69520123034E7A14F08E2ED
            AA015E12700C5668B4002A8E53416926ED8012B8390ECD3128015403BA5B600B
            485BCA5CB06069019A0EE2E0B2AE048EC1E23E9717B0023E5F65E8A2C0529BF6
            38250003036D61D3810B9A02B72FFBFA2B41E54CA16BD93082CD0973D7265E2E
            7422B034C9423F55DAAC4B997DE00AC06A18C43E8078FF0CF32E58560940FE16
            2AE8AA0217854D31545D2060531174031E7FE40C80B18D00234E0718D41B07F9
            7A6E2F836FFDD609806D470036BC03B0621FC0FE132E244193966B059D082C1D
            54DA034BFB7A7B8EFB3E0875EDD9009FC1D2AFA1B487D39357EE0778662BC03F
            3A5C583ACE8571ADE7053840CFA163C056BB29E590D242937542A3743CF90300
            5F190170C6FBDCC6E9050F77617893E5E3DE75ACED5E69058EE5376F023CB519
            E0683682CD0868A9F1B0022D7B81A5294BF3B51AA53D81D2FE169C7A5C3BC4D5
            1A99EBAFB6036CEE7085681FD41B81A7639D4F0C06E85B1FD5DDD909F08D7680
            5DC75CE86C189DE78486C332B4EC054E05C9B00D768F2F7DD785005307457537
            22E883AF021CCFBAC2B3C73AF0A0EDAA0F01DC703E0B90B677D0226E5D03B0A3
            3332F14CE89AB836EDB281A576531E33AE0F22CD2EB800E0B343228027FF06D0
            BA271E5D513BF3F0DEF3FA027C0785B1F7B88AA381CDFC8971008DA64B9023FB
            521BC0DBEFBAC04EFFCE8313A5950DACB52B4DB8CE68780A6AF5DED111ECE2F5
            AC5D39690023BC41BD009E9EC0E72F1F00B8FF157FAC4C427E7A22DF6FEFBDFD
            6586CD784C3B9B7783978A80F57043B00DE6983CF0339300FA1B4FFCC86B007F
            D81B8FAA2C7C737F80872FE16BAF1FE23EAACDD03A2012ECB31F63E746DB43D8
            F6EF7745D0A4D58CD2B28CC67A0C2CC7DA24EDDED804D0721EDFBFFE20C0A275
            FE10D2B637AA1F9AF287F9FC35025E67E264A161093D12EF7F7C6C64DA3357B0
            4F70B41CBAC03D31EB02B00C1DEB0231FC98BE4BE0BD50F24BA7B067A5874CFF
            13C0B1ACAB550D7C216AF8BB1618CDFECE76BF094A2D1170B34938DDBF1160D9
            EE48CB59A3651D9C94052CCDD93A27EB99E97CCA07F105C67045EA63A42D3971
            771A36E522047ECC68EC5504BE635DB209DA6B434E05F8D1443E5EF336D7E90A
            230726CDDB06256503DB60A2CE63CE0B9B013E7D1657A4A183FA642E213D632D
            664CFFC84409981C519290E84D6C84B76432C0E9E827DE45A8AB5BD992B2CAAC
            6D605231B004B5C704FC43F4B6C34E63B3FAD44BC253E6E3430DB5476D8DEECF
            438E0496A16192A01EB818605C235FA3216AD3217E5E97E8C359E1AD4BF5D431
            60DB67D3C6A4E5D8FBDBA91833633F3E88E3E375CBC5C3201EDB5AE111F093DD
            006B8D93E324C738CB38C7FB70287B71B7DB7F33D506D60E8BCA1FA771A57F1E
            456FBD52794A619612F8E20111F0BF8E713D2D206D1D047C39FA8BC517F1F9E3
            9B009EDFE99A7155806D64D553E05CC203A9BD01183DFDFAF2C88393B35BDC1E
            0796937B326902BEDB003F81C04B76C64DB9E626FDE295DC970FA049CF58EE07
            D67D91EA5F7F2EC04DC3A2875A0F2F81F50B5F3F14FBAEA94343D30BBBE28147
            55807D4ECB969F5D8643461F7EE855ADC56730A0DABB6E681C7A9180D62FFCED
            313CEDA46DDE2A80F6832ACAB25AAE06B0F6D056D3DF1A1DCD8EBEB8924D5BA7
            646410E084A9583E7F0EC0CDC3A387AF35D0D29CEDF61C0E4B8DA7707BD37044
            389289C6DDC25E3DB72C601D78482D5F8D53B93B4671C5E57BD9D4641AC68E87
            F6E1B24D1BB6CE18EA42FF0E4DF5D1D7DD9722501A87A93E8DF573DB2247157A
            B2222194091C28603D71A009C3D28F333C49F99AD648CA895A06377B42FBD938
            DCCC69E2DFF7E07471E69FDD6169347AF6CFE1D4F3EF870136A029AF3BE04FFB
            54042C4D504F0D65F0717B334FDA695BB687E7B85ACB726CD6E95A7B3E131DD9
            27CF04B86703673A9C97322E3D9F77937B391FB0087E4ED6E909B0ECC769A565
            0B7E666F809F7F94E1A9E26D6B79D694040CE087B6894109E8BC1844097A3951
            C829709DBBB65B12786202C0CE98D2DA81051C05DD684C9262DD169CC2ED3B11
            37EB500007AAFDB41242ECC582C8EB671360E56FA52EC29594E2A9574E8CF259
            34E5A3FCF349E81CC02DAB39D5EACB2C6A60DF2A45E1F9582E6D8C52BD642994
            BF3ED8155991765AB9BC277A4B80F602C7321F2A63497BCA4351C8D86496D2E9
            657EBA15E0C75BDC5C535E80D8B6E5F28C5E621D8EED7D7F9C6BE673FFCA1950
            B942E14BF7F816E08A02276A3921274D8980872EC129E080A87E47865F704767
            24612D48AA4B137C8AAD299B619763E8FA0F2EE5849FDC6E6A6360DF0A45C693
            082896AC4F04D6CE45A67DEA8C96E91E3AA62166D6B9D18BDFBD9EB314B23FCA
            F6C8CB2F1AC581C737DB8549E3FED434B743C9063B5E53A475B2BB806BD676F6
            241301DD2DC9145D6A91EB4A6965E285E55273CFD03E1CFBD28BCE5AC12F68DB
            D2AB183734F1584C6B4CF356C7D79FE9FE2B06F16A066D0BD600BCD111F906EB
            A832A19BFA71A01578B7C050043AAD60A597A5342B2D975028A817C96D7D4ABC
            B718608209C4736D7B04BCC0007F15EFD97A840106A0339B3890339B1494B4ED
            779302B175A8528063A6ADA1C5B9FCDD272C0B49C3D88C73F87A83F1F8F4E463
            39BE9FFAF35DD815BA727CFF54B494F90678E15A160E859CDF1BEB2ECFFC723B
            C0639BE24E2C29A359F682B8FC9E4343CB70D27681A7C6BBCE4D6FF4127356F2
            6A8306A6E0661B3AC1E967731684D2BD146E5E83BEE03E8CE7DBDE4A5883CAC7
            C3CE6E8165DFD20E4D8EAD85FB9587B7FD7E602F1E5FC9C94D4293BC6C20E79D
            7FB18DEBEE3ECE792B3B6C4993B6C09370BAB87024AF48FEE40D8056748C47B2
            6CD2D2ACAD53EB31B086D65A4FB402703F87B09EDD0A81CCFB0BA60FDFBC2ABE
            3C43F74D1B1C392D32E9ED9D5CF7368CE527BC9FAF53FD7B5E6187D615BA4149
            A65CE0246D4B8D16C005B07454B258A735173DFA167CD12FAF52C2331A9E8613
            8BAF8DE4EBE4B4B68909C644049E7D3E3B481A0D66FF250E9CB41453D1876931
            612498B4FEF065304E40E65FC073EAE5FBDC36ACA06826F5F566BE367F356BF3
            14F4CC2D38DE2FD9C110CF4EE6DFC7BFE0597B0AFD4353C59F1E6A01F8223569
            DE32DA92DF7381AA47C1C99D26D940664FC064E2749DB449F13B79EBCD38676E
            111A965AAEC8A44B15829EF7CA6C47CCCB7B7C81FD8DF2597386F1F57BB19FBE
            799487A55B47B0F32381D112ED031BB97F4BADEA69634D808B4117FB34517B7F
            0015DC281F6137AB395F9F2D2BF0A826746C481342B095929C9E6F54B0B3211D
            6A86BA4089A165B5A10BD73C1EDD3BBC2968D9A6DD24B0ECAB654D1EAA096DC1
            63C2F008C0DEEB8BE4749B72CE2B21ED9CB8E4E9612DC1B500027D5F82F90340
            6C9222B527930DDD7D815B73E052E1E5EF498E4ADE1B5B8C9339AD04D8F71CB8
            3B01F83E8A49949407D2278CD8332BFD278F5AC16B0114DB7C2A4BD06267D5FF
            8DA7D642F00297DEE4C97FE3C1A816DAE0FFE31FB52604E65FF1285DF620962B
            B1F4ADA4D5FFC20DA3705886653196CDFF06F96151A5B6261467000000004945
            4E44AE426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD9720000069A4944415478DAED9B6950535714C7FF1121
            0424D5A9E0C612400691256C22A0A02208D556A9D016171475DA693BFDD62F9D
            4E3F763AFDD2997EE84C5B5BB5D5222A180504655144AD4B41692DAEC588B28B
            D6B0482088F49E8BC9101B4C02090991FFCC1B2E8FF792FB7BF7DC73CF399727
            0896C640F1C5793F005FB323911D62D8963AD951CE8ECFA67F19FB8FC02377D0
            1F03FD1745E5DF4D77A82D8140D969E90E9A548322315441C950267EA2809D7D
            34011F16957EBB5EF8C7214BF7CDACEA8BDD0C65C2C73202EE78ED9B14B1AD8D
            EC8BA291EEF8F44437010F32DBB6747FC645CC576112D89635096CEB9A04B675
            990578CA942970779F074F4F77CC727385582C86A3A390FF6D606000DDDD4FF0
            E8D1BF686A6A86FC6E3D7A7A941313D8C1C11EC1C181080E0A8448E468F07D72
            793D2E5FF9130F1F3E9A38C03EDE12C4C5C5C0C9C949EBFCB367CFD0DBDBCB8E
            3E0C0E0EC29E3D1491A323ECEDEDFFF719D7AEDDC0858B55E8EFEFB75E6032DF
            E8E845908604699D6F6C6C4255750D5A5BDB74DE27143A202C548A80007F8DB9
            931E3F56A0F878293A3BBBAC0F98601312E2E137DF5773AEA9A905274ACAA052
            193E4A340D6263A2F8E791944A258EE61743A1505817F09225D108619D55ABE2
            F459DCBC795BE7B5CB97C7C1CD75267B18E53A474F2412E19DF454383B0F4D09
            726C7987F339BC5500D39C4D4E5EA9F9BDF0D87166C6CD3AAF158B5DB069E3BB
            BC7DEF5E0337595DA211DE9091C6BD3A89A6C5B1A2123EF72D0A4CF36F43463A
            1F15D2A95395B875BB6EC4EBE7CC9E85D4D43779BBB9B915F90545235E6B6767
            876D599B344E4DDF678F0BF0A2C870444686F176235B4B0B0B8FBFF47A638049
            B3D9F56F3FBF9E4C3B7BFF21EEED2D024C6697B575231B652137B5DD7BF6E975
            50C602930898C04965E515A8AB935B06D8DBDB0B29C989BCDDD0D0C8E7983E8D
            0678C68CE9C8782F8DB7EFDF6F4451B1FEEF310B707C5C2C020303785B76A410
            6D6D0FCC024CCADABA89476C4F9F0E60D7EEBD26316BA381D3D3D6C1952D2FF4
            E53FEEDC63D03DA3055EFDC62A787979F0765EDE51B49B20F4341A78C7F64C16
            333BB080BF07BFEECD312BF070E768AA796C34F0471FEEE03F29043C70F0B0D1
            C01D1D9DD89F936BD07D7E7EBE485CB99CB7CF9CF91DD7AEDF9C18C0E4D1B76F
            DBACF9DD5067377FBE0F921257F0F6D97317505B7B7DFC813F783F8B07074F9E
            F460EF3EC34C9A142A0D460C8B978D810E0B0DE18909E9240B406E9B2000311A
            98960A5A322891DFF9D32F467D99B1D0AB9212E0EBEBCDDBF9F945686E691D7F
            609A5334B748393979507474980D7A4BE6064D3241CB92311998C98017F8FB61
            C58A78DEAEBB2347595985D15FFA22F4F51BB75059794EEB9A69CECEC8CCCCE0
            EDB607ED90C90AC60C3B2A604A1CB66ED9C8E73199F5CFBB461710444484226A
            51046F777675213B5B7B236FEEDC39080909447BFB43B4B0A5CC14E63C2A60D2
            B2654BB130C09FB7C9919043198D68A4FD99C594949E324BB26F32601797693C
            3DA4512699CAA1582D30293C5C8AC55191BC4DB16ECE813C96CA755B9AC77CC0
            9426AE59BD8AD79FD5D0B22305BCDE6C0A797A7AC06958A9F76EFD7DF4F5F559
            0E98440E6CED5BAB3173E6EB9A73D5D535A8AABE32A64EB9B9B9226DFD5AAD73
            B92C793045DD7ACC655A824E494E625E75B6E61C8D844C56A8778D26307256C3
            D757B21C821DFE10AD0A58DD49CA6AC2C3A4100804FC9CBEEC46BD9EBF1878D0
            FDB48341A2F26F5C5CACF501AB4521273932AA8AE8EBA03AF5A3397F28F7C84B
            1F8AD502AB45CB1699E9CB9CCC5880290AF3F19540E820444B6BEB88E5E17103
            1E49911161904A87B664680D57AFE32A956AA8338A0EE41714338FFF7444607A
            98E969A95ADB337F5DADC5F9F397AC0F78DDDA355ACE4D977ECB3E88AEAEEE11
            81C3989F885E1C891616E834343621283000E5272BF9D6ABD50153E6E33E6F2E
            774C1236CFBD255EBCEE5C557599FF9DAA212DC336DF7401ABAB204A652F2E5E
            AAE28E516D1156073C5CA39DC3B4225015C4C747C2CFD3FDA5148B1B98A65A0C
            589D2D1104C1180AAC161506C8B4691F8A322ADA78B36A60723E54BE91DFA9C7
            1DF95D8381A74E9D8AA8A808D4D45C052DF994AA92BEFF619775031B225DC0CB
            E29760E1C2053C17A7B94B05C20931C2864822F1D414094ACB2A78184A6BF0D2
            B818787AB8F3658DEADCA72BCF728737E181CDA149605BD724B0AD6B12D8D6A5
            067EE55EF278355EE389D90CE5CAA1D77858ACD67F81BFA8F5F709087ACDF37F
            8E96D2A0A30B54C129EA17B56204CF5FC5A3BD93AFD891C40E174B77D2C4A211
            2C63C7E7CC57DDFA0F052C381416D5E8380000000049454E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000006824944415478DAED9B0953535714C7FF4181
            0092514B704308DB2022844D0415DC58ACB64A85B6A84551A79DF61B74FA013A
            FD066D6DD5568BB8401410A880226A5D0694D6714331A21016410D01D914E83D
            575F4A6C300B090991FFCC1D6E5EEECB7BBF77CF3DF79C7B79A2707902EED7D7
            0503F88195645624702C6958A964E5DBE090A8FB22B1581C221289AE784917CC
            F494CCC6B469D36C7D8316D5D0D010BA35CFD0D9A1528F8C8CC4137081D4DB67
            CBCC59525BDF9B55F5FC593B836E51107057405084C4D17AF66D514F2B1B6EF4
            10F008B36D5BDFCF8488F92A4C013BB2A6801D5D53C08E2EAB003B3939C1C767
            017C7D7D30C75B0A894402B1D8957F4773614FCF0B3C7DFA0C2A550B940F1BD1
            DBDB3739815D5C9C111E1E86F025617073131B7D9E52D9886BD7FF4667E7D3C9
            031CE02F43626202DCDDDD758E0F0F0FA3BFBF9F9501B03816CEECA1B889C570
            7676FEDF6FDCBA750797AFD4E0E5CB97F60B4CE61B1FBF14F288253AC79B9B55
            A8A9AD435B5BBBDEF35C5D5D10152947686888D6DC49CF9FAB515A560E8DA6DB
            FE800976EDDA240407056A8FA954ADF8F3740506078DEF251A06CB13E2F8EF91
            FAFAFA70B2B0146AB5DABE8057AC884704BB594155E72EE0EEDD7B7ADBAE5E9D
            086FA9177B18957A7BCFCDCD0D9F66A6C3C3E3F59020C7965F50C8E1ED0298C6
            6C5ADA3AEDE7E25365CC8C5BF4B695483CB17DDB67BCFEE8511337597DA21EDE
            9A95C1BD3A8986C5A992D37CECDB1498C6DFD6AC4CDE2BA4B367AB517FAF61CC
            F6F3E6CE417AFA47BCDED2D286C2A29231DB528ABA2B67BBD6A919FAED09015E
            1A1B8DD8D8D7ED9BD95C5A5C5CF6CEF6A60093E6B2F69FBC694FA69D7BF818F7
            F6360126B3CBD9B98DF5B22B37B5FD070E197450A602930898C0491595556868
            50DA06D8DFDF0FEBD39279BDA9A9998F3143320778D6AC99C8FA3C83D71F3F6E
            4649A9E1EB580538297139C2C242795D71A218EDED4FAC024CCAD9B99D476CAF
            5E0D61DFFE8316316B938133333643CAA617BAF8CF7B0F18758EB9C01B3E4C85
            9FDF425ECFCF3F890E0B849E2603EFD99DCD62661716F0F7E2F7837956051EED
            1C2D358E4D06FEE6EB3DFC2F8580478E16980CDCD5A5C1E1BCE3469D171C1C88
            E475AB79FDFCF9BF70EBF6DDC9014C1E7DF7AE2FB49F8D757641410148495EC3
            EB172E5EC6CD9BB7271EF8AB2F737870F0E2452F0E1E32CEA44991F27024B078
            D914E8A8C8089E9890CEB000E49E050210938169AAA0298312F9BDBFFC66D2C5
            4C854E4D598BC0407F5E2F2C2C414B6BDBC403D398A2B145CACBCB87BAABCB6A
            D03BB2B76A93099A964CC9C02C06BC2824186BD624F17AC303252A2AAA4CBEE8
            DBD0B7EFD4A3BAFAA24E9B191E1EC8CECEE2F5F6271D50288AC60D6B1630250E
            3B776CE3E398CCFAD77DE60504313191885B1AC3EB9AEE6EE4E61ED3F97EFEFC
            79888808434747275AD954660973360B98B46AD54A2C0E0DE1757224E450CC11
            F57408B398D3E567AD92EC5B0CD8D373064F0F85DD464B3914BB05264547CBB1
            2C2E96D729D6CD3B92CF52B91E5BF3580F98D2C48D1B52F9FAB300AD3851C4D7
            9B2D215FDF85701FB5D4FBB0F1310606066C074C2207B6E9E30DF0F2FA407BAC
            B6B60E35B5D7C77553DEDE52646CD9A473EC384B1E2CB16E3DEE655A825E9F96
            C2BCEA5CED31EA0985A2D8E01C4D60E4AC46CFAF6439043BFA21DA15B0709394
            D54447C9211289F83143D98D309FBF1D78D0F9B48341A2E5DF44967FDB1DB020
            0A39C991D1AA88A11B14523F1AF3C78E9F78E743B15B6041346D9199BECBC98C
            0798A2B08040195C5D5CD1DAD636E6F2F084018FA5D89828C8E5AFB764680E17
            E6F1C1C141FE57ADEE42615129F3F8AFC604A687999991AEB33DF3CF8D9BB874
            E9AAFD016FDEB451C7B9E9D31FB947D1DDDD33267014F313F1CB62D1CA029DA6
            6615968485A2F24C35DF7AB53B60CA7C7C16CCE78E49C6C6B9BFCC8FAF3BD7D4
            5CE3DFD36A48EBA8CD377DC0C22A485F5F3FAE5CADE18E51B008BB031E2D73C7
            30CD08B40A121020E3C7E9FC728AC58D4C536D062C644B044130C6020BA28501
            326DDA87A28C8A36DEEC1A989C0F2DDF281F34E281F2A1D1C0D3A74F475C5C0C
            EAEA6E80A67C4A55493FFEB4CFBE818D913EE055492BB078F1229E8BD3D8A505
            C249D1C3C64826F3D52E12945754F13094E6E0958909F05DE8C3A7355AE73E57
            7D813BBC490F6C0D4D013BBAA6801D5D53C08E2E01F8BD7BC9E3BD7B8D67114B
            D92ED38B5A12C96C3839584F0FB39ED5FCF7A25682E8CDAB78B477F23D2B29AC
            78DAFA262D2CFA5FC70A56BE63BEAAFE5FC25FD85F2D7BEBF60000000049454E
            44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object wmibDecrease1Second: TWMImageSpeedButton
          Left = 1205
          Top = 9
          Width = 60
          Height = 60
          Action = frmSEC.actControlDecrease1Second
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD9720000065F4944415478DAED9B7953535718C69FB025
            804445821622B24841565944A8064551142B3842C7B54A75DA69BF41A71FA0D3
            6FD0FE63DB69B55337141554401471C1024691458422D51050B6101050597ADE
            A34945939A1B121222CFCC9D9C9BDC4BCE2FE73DE75D0E57141D9B82E6266528
            801FD891CE0E291C4B5A7694B2E3DBD0B0B86691442209138944953E32FF795E
            526F383B3BDBBA8316D5D8D81806B4BDE8EE6AD74C4C4C2413F04999AF7CFBBC
            F9325BF7CDAAEAEB7DC2A0D5F904DC1FBC3446EA6823FBB668A45B5B6A070978
            82D9B6ADFB332D626B1566811D59B3C08EAE5960479755809D9C9C2097FB2320
            408E85BE3248A552482462FE19F9C2C1C167E8E9E9457BBB1AAD0FDB3034343C
            3381DDDC5C111D1D89E8A848B8BB4B4CBEAFB5B50D35B7EFA0BBBB67E6000707
            0542A148818787C7A4F7FBFBB5E8EDED63A33884B1F171B84B246CC4BDE0E3B3
            E09D98BDBEBE11372BABF0F2E54BFB0526F34D4E5E81D89828FD7B0477AFAE01
            CDCD7F636060D0E07D04BB2460312223C3B9F9EBD4D7A741D1F96268B503F607
            4CB0EBD6A5227469083F1F672358537307CA3BB57CAE9A2A3FBF8FB076CD6ACC
            9DFB2A2B1D1E1EC6E9822268341AFB025EB52A19316CCEEA3A5974BE044F9F76
            19BC363CFC632C58E08DEA6A259E3F7FFECEE7AEAEAE485BAB404848103FA785
            EDC4C902FE77ED0298E66C46C67A3DECA9D3E7F87C3524B1588C035FECE5ED4A
            3647C9020C892C86463A2C2C949FAB54ED385778112C87B52DB058EC865D3B73
            D94AECCECD38FFD4597475751BBDDECB6B0EF6EED9C1DB34C255D5B78D5E4BD0
            D9599958B468213F2F2B2B47D38316DB02AF488C4762E2ABEBFFAAAAE1F3F6FF
            24045877FDCE1DB9707171E6A67DE48F63FC87B509308D40DEFEDDDC4C9F3D1B
            E29D79DF022514989494948084F8E5BC5D527A192D2DADB6010E0A5A824D19E9
            BC5D79AB1A4AE5DDF7DE630E304D97FDFB76412412E1D123150A8B2EDA063855
            F109F39DCB78FBF7C34799C90DBEF71E7380495B323378783A3A3A86433FFF66
            11B3160C9C9B930D99CC8705065A66CEC74DBAC75CE0B8E5313CA8219D38711A
            5D16083D05031F3CF0398B99DDF0B0ED1F5CB8506A55E00016896DC9DCC8DB96
            9AC78281BFF9FA207F6D686C4279F935C1C08DF71FE0CA950A93EE93B1783B37
            771B6F5FBD7A1DF50DF76706B08B8B0BF2F2F6C095BD924C5DEC28C1F8EC3570
            C5B59BA863F1F9B4037FF5651E0FFC29A5BB587CC9E42FF2F7F743E6E68DDCB7
            9A0A2D97FB61EBA79B79FB120B401E582000110CBC73470EE6CF9FC7B39A3F8F
            9E14F46542A1A3A222A0589DC2DB0505855077744E3F70FAFAB5080D7D951DFD
            F2EB118C8C8C580D3A63E37A040707F236B9A5172FA69E270B060E67817D5A5A
            2A6F5754DC401D4BDA85EA6D6843F132554FF6EDDBCDE7FD139681E5E79F9932
            AC59C09438EC671DA1794C958C63C74F9995CDBC095D7BAF1ED7AF574EFA9CAA
            22818101BCDDDDD5631173360B98B486A57011CBC278BB9CB98B0633DD05ADC2
            FE2CF96FBCDF641173B51A30F9554A0F69942999A744DD5A2519BB0026C5C7C7
            626552226F53C9B5E04C21837F616B1EEB01539A48619FAE0047D0E70A2F58AC
            C64C61A5C71BA5DE876D8F0C9686A60D98440B58D6D64C3E1749C3C32328BB7C
            95A5738FA7D4295F5F1972B6674D7AEF384B1E2C51B79E729996A037656C809F
            DF22FD7BAA76354A4A2E0BF6D124B21C82D5FD887607ACEB24957CE2E36279C2
            3E950ED2FDE4834954FE55B0FCDBEE8075A290931632AA8A58A2836F06397609
            AC13B92DF2AB535D648C01CFF1F444704820C46E6274747642A552DB16D89812
            13E2101313A937FBB7A5D1F433F75684D1D151A3C0F463E6E66CD3EF4692EED6
            D6E1C68D5BF6079C9DB565D2E26648878F1CD5EF4719028E63EB44F2CA4474B0
            50F3B1AA1D5191CB507AA99C6FBDDA1DB0A7A707E42C863636C2B473D1D1F944
            7F6E08983235CAD8C80556DEAAE2651F9D45D81DB0501902268FB0213D4D9F36
            52C0535C5C064D7FBF6302EB449B6E64DAF4DF05B4CD43F1BC4302537D8C7625
            94CA5AD0CCA05495F4E34F871C13784DEA2A444484F3ED1D9ABBB4E5E330234C
            0580A41509BC5DCC4255DA20271FBC5A918280C5729E9EAAD59DB8525E6174AB
            7646015B43B3C08EAE596047D72CB0A34B07FCC13DE4F1C13DC613CE52B69BF4
            A09654EA0D27071BE97136B2DAFF1ED44A11BD7E148FF64EBE67C7067678D9BA
            9316166D8B94B0E33BB65635FD0B934CB75FA30BFE320000000049454E44AE42
            6082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD972000004D34944415478DAED9BCD525A4914C71B51E4
            43101014D4A4B271CCBCC1CC2E8BB898EDBCC1CC0BCC766A1E606A5E26D96661
            16D9651E615256A5B2C808A222208AA8C0F4EF986608B926345CB870E55F752B
            97C0D5FEF5F9E873BA25F0CBAFBFA9FDFD573B4AA9BFF4F55C5F09E52FD5F4F5
            5A5FBFEFEDFD7410D87EF4DDEEC2C2C2DF3B3B4F93F9FCA65A5A5AF27A80AEEA
            E6E646150A87EAE0E09F4ABBDDFE01E017BBBBDFFFFCF8F113AFC736567DF8F0
            5E43BF7B0970F5D9B3E709BF59B65F58FACD9BD775803BDAB7BD1ECF44A47395
            9A03FB597360BF6B0EEC778D05381008A8E5E565150E2FAB5028A416171795AE
            E4E4BD4EA7A35AAD96AC89CD6653351A57F27A2681815A5989A9582CA682C1E0
            C0CF351A0D55AB9DCB24CC0C70241251C9E4EA17A0B7B7B702D26AB5F5AB8E4C
            0A16A7AAC3137A55AF5FA86AB52A5E30B5C00C7A7535A12DBBD2FD3F5C94C137
            1A971AB875EF73E17058BC01D73762724E4FCB32515307CCA053A9A48A46A3F2
            1ACB9C9F9FEBAB6E6525E29D9F83E5CD841D1F9F8C057A24605CD85896416299
            EBEB6BC7CFC6625171636255B7688E93974EA72434CCCF2B958E5D4F68430333
            B0B5B57477705FB30871BBB999977B62140F7052BFC75C5D5DA9939353EF8101
            D8D858970485EB027B9F6511AE9ACB6DC83D16AED56AF77E16E84C26A3DD3C24
            AFCBE533757979E92D702291D057FC13404D20BE261BE0BBCF07F5846E083CDE
            532C1EB996B9AD8119443E9F132B0F3A185BE0FE492D97CBDACA0D6F802391B0
            8EDD35B9AF566B9295BFA5618009179E6182DD8C656BE0643229D5142A148A03
            65D161805126B3266B351E74785870C5ADAD81D7D7B3521F939171E741342C70
            3CBEA28B9A55B92F954A3A318E5E7A5A03B3BC10BFD4BFACBBE304C6BA5819B9
            15C7D6C0DBDB5BF2EFC5C5853A3BAB5803DB3C170A2D698F5A977B9EE1D99900
            BECBEC78C65DB33068B2A33263BD47954A45EAF389036F6D6D0A808D4B23EA65
            DCD374488340D354508420B70A106B60669C99A7AB393A2A59FD325B68560356
            054435C786C1C48129F04DADCB52E1D408B8054DAD6E9A89617E972BC0743DA9
            544AEE878DAB7E6827776525C8E57212F7D4E9744E6EC81A9881505A32D861DC
            DA09BA5EAFEBC9AB7EF63E999D650999FD2F4F80112D1C3B156894E5825C0038
            D675C35DC706DCDBCD3050DC6D5C5B3253018CE2F1B8EC65215C8E2C3A292B79
            028C75E99ACC061CD074346E6DC910BFC1E042F735FBD79E64E95E91C0280C28
            0111B0C434EDDC28A239A149E915C9D18D7DEB91B7698166BD24F9189151A9C2
            86B1089E93CD66BB933875C06690C434ED9C595B4719A03996894623DD4A6BAA
            808D5866D896A13A726380BD45CE54021BB16CB5DB9D9193CC7DC06CFF30A978
            02E16353944CF4B814EBF71EC9F48BB59CE5CD6CE53801538191D08CDB23F6B9
            D9EF9E3AE06C36F359727352B158EC9E4739019BF5DF58968A8F5A7C502B4F14
            185704B8EFD0B02B2CDC6CFEBFA1EF044C224BA7D3B204B255C4B68FCDE6DE54
            FF05801370FF1994ED69E3CC011B018C6B13D336EDE3CC019BF368732047AB8A
            3E7EFCD79FC0A635256EB9C8D6BEB130C73A9C312113A7243ECEA5692EB036D9
            99FADD17313C0ECD81FDAE39B0DF3507F6BB0CF083FB92C783FB1ACF535DA2BD
            7D205FD4FA31F0E9AB78BBFABD3FF5B5A7AFB8D78374591C4FEEEBEB0F9DABDE
            FD0702D6C0AF37F8478C0000000049454E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD97200000AC34944415478DAE59B6B9054D511807B6676
            79AC820F343C9508080663120554489468C058F155A5B12A95145049611E9A88
            0A31E66159125F153196FE3065D4442504DF564989464C254674510841124160
            7949581505655956D899B963F7F6EDB97DFA9E3B8F9D594899537576EEB93BF7
            DEF39DEED3A7BBCF9DD4318F1700CBF1586FC33A056B7FF8749536AC2F62BD0E
            EB8614028FC18365580F3FD83DEBE1F211D6D309F8493CB8B85E774D55F9FDC2
            81857E8A8077430D6AAC0153293FB4B42D9CB40B05FFF91E28ED045CF5FD2D64
            4A9F0FDBBE81D070050D5670DBFA3BF52E5501FB402DA43E9F4ED0EF420855B0
            C70700BC62E0A2140D685A9D4B876D3B000EAC860A8F03F36907430FD401018E
            C11A40FAD4E7F4A7AF6828820CC2B673AC064483D70A5D16584B49A42790697B
            6C6053663ECB83020322A0F982DB0E0A7E69D7025D12D8C26A29661464C603AF
            81E55EB6E31A8C6ABE608EC1A3E6354227026B35165881CA18D08C1980B499DB
            BAF8D45920730AD6C2FB8C5BDD802B816D48879F0A986A2F3C3FE128AE630F03
            18DC848B7C23DF2F8B74EFED03D8B40760D52E8097DE03D8B9DF05CC05513B6F
            06A01ED089C0DA400968574D336411366CF743A84B86A3CB762CFAA8BD2A7B38
            3D79E90E80873702AC6F8BC0721E70811783D65DD58E018B74C5F26A583AD798
            E6E306F9C43A7910C0ACCF010CE8EDDEBCB50360EB5E800F3BB1C3D8CB437150
            06F70518D98FEF5304C7FACC3680DFAD4357281B0166836800EC1CEFAE94BDC0
            5A95B5BA3684D2145852DFCB4F00B8747874FD2E54D145FF05F8FBBB00EFEF73
            A7877490AE3B6500C079C3004E3E32BA766B3BC0CF57026CEB88244CD0399172
            101933B1E6D54AD90B2CD275600D3075FA575F00F8DA60BE8E3AB57013C0135B
            F9D819402549FD307AF24947005C89DA31B489CF91365CF93AC0967686CE2A70
            316C79A3DADD06D6D2D5D657E629A9A1405327BF393CEAE48DAB00D6B5790207
            3C3115076504AAF19F7140F6E4E2DE539F0CC0356301CE18C8E776A066FCA099
            35242752D6922EB86B764DC05ABA0D469505F82C9CB373BF14A9F0EC1500EF74
            C4DD482A64A11F9DCCC77F68410DD8E2EF205D7A35424F1DC2ED153B01E62C07
            E82C4492CE856A9DD3460C22D5EE36B0B5C004D9187E1E86000BCE0438A21777
            6416AADF863DF10849CA2054D507BFCCC70B50C27FDA140F0745D2F4DC79E301
            4E0C5311B7FE1B60F1767E4E67600C5910F7C6AA06D6DE9218A64623DD99C703
            CC18C9DFFF630B4324F9CC2471B2CA0F7D25029EBF31F4933D857A3210BFFF00
            0E50EF34ABF6B7FF01D0912F2DE56AD4BA085C9CBF62A03C73B709E7DAD367F1
            9A4BF36BDA52EE844F5A5D0308EC78CC0F8149BAB4E66A89F89203DF1B85A023
            F878EE1B002FB4BA7339ABA06B022EA5CE04FD559CBB379DCC17DEBF018DD0E6
            D2E11BDD7308022F3883DBF311F8A196C85A2769064D179AF7A461AF7D00F0D3
            150C9B356A2D963B5F0BB05E73AD3ACF41A372D1B17C21A9DA3B1FBBA3AB0377
            F1D468B9597866088CD27D7063692323ABC42DA7009C8AEEE97E843AEFC5B85A
            E7947AD70CDCA8E7AF82BE7F12C0E8FEEC41913AC73C1F034CF720E04742E087
            43E0721138F5E35B9F05B86C34B767BE0AB066B7BB2E0BB056EB9A8035A89C5B
            3C05DDC3068057D0FFBD7E55DCF3F1757C5853B42CF9809396A88947B394A9DC
            88F3F82FAD0A32A823B0F6A6B43AD33C7EE95CBEE859741DEF58E31A0EE9BCC0
            880124E0C742E0E77089B9FDCDB8E3AFDBD20FD2A47B27F2F97978CD935BE3FE
            752EA81138638C5535C0D699978E1F82167DD1D9007D337CFEBE0DEC715960DD
            5FBA9680EF9BC4ED3BF1598F6D89CFDD1E93B0040A4BCEE1CF97318EBDE18D48
            A5C568687099C374CFD3D0F8DC86EAD95BA0D7A3E5DE1C97AE1E2C8AA7EF181F
            A9F473DBA380A2AEC0A58C16ADA7C30F05781B43BEEFBEA2E2D74065300AAE6A
            8AD59F60A07FBF9E1D1147D20AF8125C0D668DE5F68F9601ACDC19391B391548
            D40538C9CBBAE18B0053C2E8E8E2BF8571AEF178C4DDD3D344EE3B0EA17FA3A0
            EFF540CB75BFC6F57EF240BED7D425187464233847D2B5026BC743FC6701BFE0
            18809F7D9E2FBC7B2D7A5D6F7347B594F31E298BBB4A653C42CF1B1741DFB41A
            E0F9ED6EA72851401E1DCDFB373FE2652967241B28D5D6835C157039D7F2C85E
            DC116A53BC3AB3D99DC7BEFC930E37C5B392F949D08F6E01B86B6D24592AB476
            4BA8B80681FFB5CB9FF2C9D7020CE0A65D1B3CF399EAB528E1F387F1F77F8BD6
            73D1B678A6D1B736EB1C351D8FE9CFE074FDDE5CF49DAE7E40942111185F72CF
            09132BCC7E78C3431BF86B7F7AE8216CBCC85AD3BCA240BDF5E378E631C919B1
            BB164EEE5AFD4F4AE081D3AAADA75025DB32152500AC94A78FE430914ACB1E8E
            89B551D119091B1509984EF9EA73BEDD0A6BAC62F3D83CAB94B4CBA7788C9433
            613E8B02F5F10322E8392B3864B4EAED8BA02CB093FFC6CF899FE1882915769A
            D6FD5D9D71D5D6C1834DDD26417B81ADC54E1B29D327A56EEE3A953D222AB444
            DDBC9AFDEC72C0006EEAB7088EE7C71E0E70CF696EBAE8B257395FA6E7B0F5B8
            6449B41B7025817D52B6F1B1ACD374BE3F0612B78E7353ADCBD141F8E54ADE51
            F0A551F5FDE53E19D526FF79543FB79364272851AFB54796A96C101FE052C9FA
            4460AB7A99B43BAFB574683ED3BC96F948B1F25B6DF1076AE32480765F8A0691
            DA67A38373C509FCFD2B96853B1310418B842531E024EA9574637E7AB9AD16BB
            4328D05A0BA81C872EE70FC77056E43B08BC76B71F16C0AC04065806E29C219C
            C5A4F213348A2D6D718B2DB0D900BCBB134135C05A122903EDB3B05206F5E175
            B52DEB97AC5860AD3599943A17DEF7EB087C55087CD5EB6C18E97EE4FC4C42A3
            D6D4C01B72CDEF73463317F85789A012E0986A2B6881B56F03A414B8CFE1F8FE
            6894FC88F87B1FD2DCB6972D3D759C06828029D94FE59AE5BCE378140EE69D13
            388928E591CD1CAA5A239694D1AC6843DC426BC741DEE7B06A6B5F6EA1F4D0B8
            019058A813D35E06F8605F1C783601B7F3CEE434B415FF4177F39F681C2F408F
            6F2EAE0CCD3B389369F35C3698A908D876DEBEC86207455F58746448D5FB7252
            8EE667D779A335DB3B386F256A7DEE50DE91D4C0E45FD396CCEE4EDEC5F86B2B
            6FDD7406AE5AEB10B22A600BADA55E3C56C03EADB0EF84945A87B5512460315A
            A4D29BDBF97F044C73980AA9392522C8A0C9CE84DE75EC167092B48BE70CBC4F
            9DED1B3ED667B77BD1A405DF40E0D927F2FDC8689184A54C3A1A60C628DED5A0
            E56AC6D238B00D2C8A06B49617D3628311FEB192F7BDBF65256B1D1DDA3BBE36
            8CBD7FFC1A4BB30F5AE6E92338BF45BD5E3899FF7FFAE278925EAB74B7817DE0
            7600B42A3B5206B58C81BBD4495B4BFC4204BEEE24FE3EA57808988C180D0449
            737F9EADF53A9CF7D39584B5946B52E94A07C107EE8B82B41ADB77BC48C2E4C0
            50229ECE51128F962D5A96C8905152900666F5879C31A1F96D735D49DBA97503
            B692F7BD8BA9A5EBCC69A3FA7AF7326D07302C3631A021BBE578D40BDA91BE6A
            58752FE5D525BDE0E64B0E38152A742DEB0D6E07C0919867AE8B2797311AA08B
            06B62FB8551D3CF474F12E7106BC9C27A7635EDF4BA9158787071A3CC9BAEB35
            DB0E4E115819251980726FE01E74E0E2718965CD7E070C8C06D58361610F3AB0
            85F682FBBE54A48E6FD3D8C1883DAFD61F79F414B8862F5792DE0AF294F6BAFF
            8CA7A707C10B5CF92DBB7EC64399A366F8FFF8A1D6C454F8533CFA75DA2D58A7
            62ED57CB5DFF070B7AE1B004EB2FB0AEFB04D49E26A54853DFF7000000004945
            4E44AE426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD9720000067A4944415478DAED9B5950935714C7FF9125
            0404B48A5A88C8520AB2CA22021A14418358D1113A6E54A94E3B6DA76F7DE974
            FAD8E9F4A5337DE84CFB62DB69B5534141AC209B28A28802A2C8226A916A5814
            D00491B09ADE733519D1A426212121F29FF926E74BBE8FDC5FEEB9F79E732E9F
            202C221EF2AFAB03007CC78E1476B8C1B634C08E72767C39EF9B845B82A5B9AA
            404C8CD588CA7F9CE7D854028172C0D20D34A95422378C864AA14CF95C0E3B87
            38023E262AFD61BBF0728EA5DB66568D246441B9FEB33C0256B87F9FEA666B3D
            FBB2A8A7155F140F12B08AF9B6A5DB332D62731566816D59B3C0B6AE59605B97
            5980E7CC9903B1D80BDEDE622C5EE401373737383909F967131313181C7C82FE
            FE87E8ECEC42FB9D0E0C0D296726B0A3A303C2C24210161A0291C849EFFBDADB
            3B507FE52AFAFAFA670EB09FAF0F249278383B3B4F7A5FA118C0C3878F582F0E
            61E2E953889C9C588FBB62E1C205B0B3B39B746D73732B2ED6D4626C6CCC7A81
            C97DE3E25622223C54F31EC15D6F6AC1AD5BFFE0F1E341ADF711EC32EFA50809
            09E2EEAFD6A34772149D2AC5C0C063EB0326D8F5EB1311F08E3F3F7FCA7AB0BE
            FE2A1AAE36F2B1AAAF3C3DDFC6BAB56BE0EEFE2C2B552A95385E5004B95C6E5D
            C0AB57C7219C8D5975238B4E95E1C1835EADD70605BD8B050BDE425D5D034646
            465EF9DCC1C10149EB24F0F7F7E5E734B11D3D56C0FFAE5500D398954A9335B0
            F9C74FF2F1AA4D42A110FB3FCCE2760D1BA3E401DA441E433D1D1818C0CF65B2
            4E9C2C2C814AA5B22CB050E8885D3B33D94C2CE26E9C97FF377A7BFB745EEFEA
            3A17597B76709B7AB8B6EE8ACE6B097A6B7A1A962C59CCCF2B2A2AD176F3B665
            8157C64421262692DB976BEBF9B8FD3F1902ACBE7EE78E4CD8DBDB71D73EFC67
            0EFF612D024C3D90BD6F3777D3274F8678635E3741190A4C8A8D8D4674D40A6E
            97959FC1EDDBED9601F6F55D8654690AB76B2ED5A1A1E1DA6BEF31069886CBBE
            BDBB20100870F7AE0C85452596014E9424B0B57339B7FF387484B9DCE06BEF31
            0698B4394DCAC3D3F1F1091CFCE57793B8B5C1C099195BE1E1B190050603CC9D
            73F5BAC758E0C815E13CA8211D3D7A1CBD26083D0D063EB0FF0316333BE24EC7
            BF282E2E372BB0378BC436A76DE4B6A9C6B1C1C09F7E7280BFB6B4B6A1B2F2BC
            C1C0AD376EE2ECD92ABDEEF360F17666E6366E9F3B7701CD2D376606B0BDBD3D
            B2B3F7C081BD92F49DEC28C178FF3970D5F98B6862F1F9B4037FFC51360FFC29
            A52B293DADF717797979226DD346BEB6EA0B2D167B62CB7B9BB87D9A0520374D
            1080180CBC734706E6CF9FC7B39ABF8E1C33E8CB0C850E0D0D86644D3CB70B0A
            0AD1D5DD33FDC029C9EB1010F02C3BFAF5B7C3181E1E361BB4746332FCFC7CB8
            4DCBD2E8E8D4F36483818358609F9494C8EDAAAA6A34B1A4DD50BD0CAD2D5EA6
            EAC9DEBDBBF9B8BFCF32B0BCBC135386350A9812877DAC21348EA99291939B6F
            5436F32274E3F5665CB85033E973AA8AF8F87873BBAFB7DF24EE6C1430692D4B
            E182970772BB922D172D462E17340B7BB1E4BFF5469B49DCD56CC0B4AE527A48
            BD4CC93C25EAE62AC9580530292A2A02AB6263B84D25D78213850C7ED4D23CE6
            03A63491C23E75018EA04F16169BACC64C61A5F30BA5DE3B1D77B59686A60D98
            441358FA96343E16494AE5302ACE9C63E9DCBD29356AD1220F646C4F9FF45E2E
            4B1E4C51B79E729996A053A51BE0E9B944F39EACB30B6565670C5EA349E43904
            ABFE11AD0E58DD482AF9444546F0847D2A0DA4FB690D2651F9572249B03E60B5
            28E4A4898CAA22A668E08B418E5502AB45CB16ADAB539D647401CF7571819FBF
            0F848E4274F7F44026EBB22CB02EC54447223C3C44E3F6AF3446AE60CB5B11C6
            C7C77502D38F9999B14DB31B49BAD6D884EAEA4BD607BC357DF3A4C94D9B0E1D
            3EA2D98FD2061CC9E689B85531E866A1E63D5927424396A3FC7425DF7AB53A60
            17176788590CADAB8769E7A2BBE7BEE65C1B30656A94B1D1125873A996977DD4
            1E6175C0864A1B30AD081B5292346923053CA5A515902B14B609AC166DBA916B
            D37F17D0360FC5F336094CF531DA95686868048D0C4A55493FFD7CD03681D726
            AE46707010DFDEA1B14B5B3E36D3C35400885D19CDED5216AAD20639ADC16B24
            F1F05E2AE6E96957570FCE5656E9DCAA9D51C0E6D02CB0AD6B16D8D6350B6CEB
            5203BF710F79BC198FF1C4674199FCEC311E16AB8D5DE40F6A5D2F8660786614
            D5F595CAC915A361A9EA07B5E205CF1FC5A3BD936FD9B1811DAE966EA489453D
            58C68EAFD85CD5F61FD60A1714EE1AA1210000000049454E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000003C0000
            003C08060000003AFCD9720000065F4944415478DAED9B7953535718C69FB025
            804445821622B24841565944A8064551142B3842C7B54A75DA69BF41A71FA0D3
            6FD0FE63DB69B55337141554401471C1024691458422D51050B6101050597ADE
            A34945939A1B121222CFCC9D9C9BDC4BCE2FE73DE75D0E57141D9B82E6266528
            801FD891CE0E291C4B5A7694B2E3DBD0B0B86691442209138944953E32FF795E
            526F383B3BDBBA8316D5D8D81806B4BDE8EE6AD74C4C4C2413F04999AF7CFBBC
            F9325BF7CDAAEAEB7DC2A0D5F904DC1FBC3446EA6823FBB668A45B5B6A070978
            82D9B6ADFB332D626B1566811D59B3C08EAE5960479755809D9C9C2097FB2320
            408E85BE3248A552482462FE19F9C2C1C167E8E9E9457BBB1AAD0FDB3034343C
            3381DDDC5C111D1D89E8A848B8BB4B4CBEAFB5B50D35B7EFA0BBBB67E6000707
            0542A148818787C7A4F7FBFBB5E8EDED63A33884B1F171B84B246CC4BDE0E3B3
            E09D98BDBEBE11372BABF0F2E54BFB0526F34D4E5E81D89828FD7B0477AFAE01
            CDCD7F636060D0E07D04BB2460312223C3B9F9EBD4D7A741D1F96268B503F607
            4CB0EBD6A5227469083F1F672358537307CA3BB57CAE9A2A3FBF8FB076CD6ACC
            9DFB2A2B1D1E1EC6E9822268341AFB025EB52A19316CCEEA3A5974BE044F9F76
            19BC363CFC632C58E08DEA6A259E3F7FFECEE7AEAEAE485BAB404848103FA785
            EDC4C902FE77ED0298E66C46C67A3DECA9D3E7F87C3524B1588C035FECE5ED4A
            3647C9020C892C86463A2C2C949FAB54ED385778112C87B52DB058EC865D3B73
            D94AECCECD38FFD4597475751BBDDECB6B0EF6EED9C1DB34C255D5B78D5E4BD0
            D9599958B468213F2F2B2B47D38316DB02AF488C4762E2ABEBFFAAAAE1F3F6FF
            24045877FDCE1DB9707171E6A67DE48F63FC87B509308D40DEFEDDDC4C9F3D1B
            E29D79DF022514989494948084F8E5BC5D527A192D2DADB6010E0A5A824D19E9
            BC5D79AB1A4AE5DDF7DE630E304D97FDFB76412412E1D123150A8B2EDA063855
            F109F39DCB78FBF7C34799C90DBEF71E7380495B323378783A3A3A86433FFF66
            11B3160C9C9B930D99CC8705065A66CEC74DBAC75CE0B8E5313CA8219D38711A
            5D16083D05031F3CF0398B99DDF0B0ED1F5CB8506A55E00016896DC9DCC8DB96
            9AC78281BFF9FA207F6D686C4279F935C1C08DF71FE0CA950A93EE93B1783B37
            771B6F5FBD7A1DF50DF76706B08B8B0BF2F2F6C095BD924C5DEC28C1F8EC3570
            C5B59BA863F1F9B4037FF5651E0FFC29A5BB587CC9E42FF2F7F743E6E68DDCB7
            9A0A2D97FB61EBA79B79FB120B401E582000110CBC73470EE6CF9FC7B39A3F8F
            9E14F46542A1A3A222A0589DC2DB0505855077744E3F70FAFAB5080D7D951DFD
            F2EB118C8C8C580D3A63E37A040707F236B9A5172FA69E270B060E67817D5A5A
            2A6F5754DC401D4BDA85EA6D6843F132554FF6EDDBCDE7FD139681E5E79F9932
            AC59C09438EC671DA1794C958C63C74F9995CDBC095D7BAF1ED7AF574EFA9CAA
            22818101BCDDDDD5631173360B98B486A57011CBC278BB9CB98B0633DD05ADC2
            FE2CF96FBCDF641173B51A30F9554A0F69942999A744DD5A2519BB0026C5C7C7
            626552226F53C9B5E04C21837F616B1EEB01539A48619FAE0047D0E70A2F58AC
            C64C61A5C71BA5DE876D8F0C9686A60D98440B58D6D64C3E1749C3C32328BB7C
            95A5738FA7D4295F5F1972B6674D7AEF384B1E2C51B79E729996A037656C809F
            DF22FD7BAA76354A4A2E0BF6D124B21C82D5FD887607ACEB24957CE2E36279C2
            3E950ED2FDE4834954FE55B0FCDBEE8075A290931632AA8A58A2836F06397609
            AC13B92DF2AB535D648C01CFF1F444704820C46E6274747642A552DB16D89812
            13E2101313A937FBB7A5D1F433F75684D1D151A3C0F463E6E66CD3EF4692EED6
            D6E1C68D5BF6079C9DB565D2E26648878F1CD5EF4719028E63EB44F2CA4474B0
            50F3B1AA1D5191CB507AA99C6FBDDA1DB0A7A707E42C863636C2B473D1D1F944
            7F6E08983235CAD8C80556DEAAE2651F9D45D81DB0501902268FB0213D4D9F36
            52C0535C5C064D7FBF6302EB449B6E64DAF4DF05B4CD43F1BC4302537D8C7625
            94CA5AD0CCA05495F4E34F871C13784DEA2A444484F3ED1D9ABBB4E5E330234C
            0580A41509BC5DCC4255DA20271FBC5A918280C5729E9EAAD59DB8525E6174AB
            7646015B43B3C08EAE596047D72CB0A34B07FCC13DE4F1C13DC613CE52B69BF4
            A09654EA0D27071BE97136B2DAFF1ED44A11BD7E148FF64EBE67C7067678D9BA
            9316166D8B94B0E33BB65635FD0B934CB75FA30BFE320000000049454E44AE42
            6082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object lblEventCount: TLabel
          Left = 1276
          Top = 71
          Width = 111
          Height = 19
          Alignment = taRightJustify
          Anchors = [akTop, akRight]
          AutoSize = False
          Caption = '10 / 100'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = 12105912
          Font.Height = -14
          Font.Name = 'Century Gothic'
          Font.Style = []
          ParentFont = False
        end
        object pnlPlayedTime: TWMPanel
          Left = 0
          Top = 9
          Width = 180
          Height = 60
          ColorHighLight = clBlack
          ColorShadow = clBlack
          Color = clBlack
          ParentBackground = False
          TabOrder = 0
          object imgPlayedTime: TImage
            Left = 1
            Top = 1
            Width = 178
            Height = 58
            Align = alClient
            AutoSize = True
            Picture.Data = {
              0954506E67496D61676589504E470D0A1A0A0000000D49484452000000B40000
              003C0806000000CF2E9087000009084944415478DAED9DE9521B471485AFF67D
              4362B30DD8D8061C675F5CC9FFE405F234A93C402A2F93BC40F23FA9C4899338
              4E30C4D806DBEC08ED3B4A9F16334842A3198D469619EE57D5460846D332675A
              676EDFBEEDA016B745FB56B4CF458B12C35C1CB2A2FD28DA57A2AD3BC43FCBA2
              FD2C5A7CDC3D6398213816ED5308FA3BD1BE1C776F18C602BE87A033C43683B1
              077908BA39EE5E308C55B0A0195BC182666C050B9AB1152C681BE1090728309B
              246F2C4C4E8F8B9A274DAA17CA54DE4B5369FF58FCA5C7FFA776077CE49F8C93
              3BE42787D34927B53A5533052A1F1C53B37132F4EBB3A06D42787E9A8257539A
              3FAF658B947EF4F48DEDE349AD41D9272FA99ACE0D750E16B40D0809A184FA88
              19D48B153AFAF3BFB1F633F5F18AFCE4E847E6F116558EB2A6CFC182BEE0E0A3
              7BE2DD9B7D7FE7A45AA7A3871BE26BCDF06B7AA321F9D5E5F749113A5C2D2136
              1B0D399A36CA156967AAD982FC6A046F2C44B1950561351C9ABF03DB71F8605D
              5A1133B0A02F38D15B57A527D5026286D56894AB7D5FC71DF4937F2A4E81549C
              1C3AA368374D21F092F0C0E5BD63F149D05FDCDE7898E242D4A4AD692ABE3AA0
              FCF35D53FF1F2CE80B4EEAA365727ADD3D7F6644CC188543D7A6C83711B1A43F
              95A31C155EECF51DB571AED8D2BCA6A8D15F8CD26660415F70A63EBBDBF3793D
              313B5C4E0ACF4D536076A2E7CFEBF9125573456A08EFDD1056458940E03897D7
              43AEA08FBC9120B9C3819EC797B68F28BFB5AB19B9D013F5DE4F8F4CFD7FB0A0
              2F38539FDE3D270A3D31C35EC496E6C815F0763C0FF116778FA87290A1937AC3
              D0F99D6E17F952310A4E4F489177BC5EA94A99B52D4D1BD24FD42CE84B0A6E08
              611B149A4288E947CF3445040F1B5B9EEFB8316B546A947FB633547401F826A2
              14BE3E432E9FE7AC3F274DCA3CDEA4EA715EF398D8F25CC7736C392E31A1AB93
              149A9F928FF5C4EC4B0AF1DC9EEB18118BAF0EA9B0B5278437FCA406C0644968
              6E8A825792674F0A8565D6B7A872D8FB82F18B113E7AFB9AFA7D69FB9072E202
              33757E62415F68E06993EFDD92424AFFA32D664F344489B7AEAB6286B7CDAEBF
              A0CA9013195AF812112952F4AF754292FDAB650B3D7F5F15B518D10FFF58979F
              1AA6FE3F88057DE1819881D6280B0B006BE2709FC692C5487EFCEF73AA891BBF
              5182A9F8F89D858EF31EFDF54453AC9ED31BCC61FAC582B63B0E0725EEDE204F
              A4251678DA638C94B9E26B39BD2712A4B8F864503C7B2D576A4DC18F28AF8405
              6D7302331314B931AB7E9FFDEF259591A8F41AC1C40F268014724FB7A9B47334
              9273B1A06D0C426AC90F96C4477ECB92940F32D2378F03F863F864D0AC637A7B
              CD7068701058D03606D186D0B549F9588AE80F033912C2A2C073EB4D950F8AD3
              E3A6E4FBB7D58BABF0625F4657AC86056D5370A398FA6849BD212B6CEE51E1E5
              BEEE71D26F4783325A927BBAA3199530437788F1E0B735CBC285EAFB2616B42D
              F127A3145D6A4D58C80CB6DFF53FE211629BBC77A7E339D814240A19CDD4EB87
              B4401F2EA9A1BCECDA16950F879BCC39F71E88056D4B301BA8241C210B0EC9F3
              7A4070A94F56CE3D8FC848515884E2F6E1D0236AF4E65599D50790C88459442B
              6141DB11E183273F5E56ED06C27458E6A48796A015AC9822474E34C27800BE7E
              FFFEAAA5213C16B40DE948FA1762D9FF65D5D0C8DA2DE8DCC6368585E7552E0C
              05E46540D8F55265E0BEC1DB4FDE5B91171DC0448BD10502865E9F58D0B6A33D
              EE2B2732FEDE30745CB7A091F186E7902F8D787647569C504D71E7340F64C0C5
              AD89B717D5891EABE3E22C681BD21EAE33EA9F412F412B20E534727D863CC232
              B483D52AB9CD5DB9B2DC28ED3EDAEAF01D0BDA866066508EA8349860FA095A01
              197B91851972B6A588022C08C86DBCA29A01FBD07EC19576D3F238AB6041DB90
              F67586F0BA884E18C188A0017C30CA1184AEA4C441673E44C696EF3F16B6BDBF
              A482B34999370D6037603BAC82056D43DAA79947216810989E3877C388F0DEC1
              AFFA37A01D82B6783A9E056D43228B5784E012F2B1D59603D973B034EDAB6400
              427A39148A31101E64CBC10C447B85A2B2104CD6A060FA09DAE9F55078615A1D
              F915CC4CBA44C505E73FBDE08A2F0F28BF69AE64412F58D036A4236C374009B0
              5E82461E737036252F1075F5C92966A7C5957C11C0613B4617ACFC48BCB3281F
              6304DDFFE55F43B371DD82CEAC6EB616BDFA3B57870F95B88459CC7B77D484FF
              F4C30D4B57CEB0A0ED48B768C4088D915A0FBDA96F4431F2C28FC3F79A9DAEC6
              C88C115ABEDE00179BE1B74E2C685B82725BDE44583EC6EA10AC12D14353D042
              21A5BD341584D71D3629BF3D465E4DE7E978F5B9A5EF9B056D53DA7DB4D1DCE3
              5E82C6C88E8B41AF669D11BA73B447B11C8C056D53700387F2B58AED301A8F9E
              78EF16B9833E597D09C7940F3396F5293893A4F08D56FC59C6ACEFAF5A52E4BC
              E37D130BDAB6841766D4822FC8B93878B0A62B205C088831D7F3654B5793C8FA
              211F2CA9F5A151E026FFDC5C3199BEE72116B46D415552B948F6749436EAA547
              41BB77C6E82C17C956CDD580EE070BDAE6B4AFE30308C58DAA5A9216A8A2145B
              9957BF37BABED10C2C689B83D139F1CE4DE98B015689A41F6DC82D2A5E07386F
              E2EEA2BADA1BE74D3F7C2247E991BC5F6241DB1EEC3C954029B053EB61B4AAFF
              B06042063167A5203B449CC60A15132B5D8CC282BE24C8CAA34B67656B475DDF
              AEBBAE1D40ADE88AC5ABBCBB61415F2290F219593C2B0B86BF3C220D46D34B8D
              22D34317663A966C617D62697734E5BFDA61415F32BAEBCC01B9DAE4D9CED005
              1C656AE9F59973DB54BCCE7A7A2CE84B08B66C43119AEE3D03B19A1B791AD8FC
              526FD58982C3E1206F2222F3AFB13B403B72334D6133AA16565FD2ED0FB1A02F
              25A83587D830BC7537987C41A23E46EC86B88143F23EF62704D8AF10B5EF5CE2
              46132332EA6C74A795027865C4BCCDEE37681616F42507A36A786E4A7337AB41
              817D41469ED69E2AA38605CD48B0E13D6C832C1FE6700C76B0B02728EB25ED4A
              663C42566041331DC03EC063634F164C8A2086ED10F64489612396DC143602B1
              644C9220C91F1ED9EA2423D3FD27163463232068E40746877D21867903C843D0
              DF89F6E5B87BC23016F03D048D250A3F89161F776F1866083073F399723BBB2C
              DA37A27D215A64DC3D639801402EEC0FA27D2DDAE3FF01B4A15AD463B9DD6E00
              00000049454E44AE426082}
            Transparent = True
            ExplicitLeft = 40
            ExplicitTop = -24
            ExplicitWidth = 180
            ExplicitHeight = 60
          end
          object lblPlayedTime: TLabel
            Left = 7
            Top = 25
            Width = 157
            Height = 27
            AutoSize = False
            Caption = '00:00:00'
            Font.Charset = ANSI_CHARSET
            Font.Color = 9485642
            Font.Height = -27
            Font.Name = 'Crystal'
            Font.Style = []
            ParentFont = False
          end
          object Label2: TLabel
            Left = 7
            Top = 3
            Width = 70
            Height = 17
            Caption = 'Played time'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 9485642
            Font.Height = -12
            Font.Name = 'Century Gothic'
            Font.Style = []
            ParentFont = False
          end
        end
        object pnlRemainingTime: TWMPanel
          Left = 182
          Top = 9
          Width = 180
          Height = 60
          ColorHighLight = clBlack
          ColorShadow = clBlack
          Color = clBlack
          ParentBackground = False
          TabOrder = 1
          object imgRamainingTime: TImage
            Left = 1
            Top = 1
            Width = 178
            Height = 58
            Align = alClient
            AutoSize = True
            Picture.Data = {
              0954506E67496D61676589504E470D0A1A0A0000000D49484452000000B40000
              003C0806000000CF2E9087000008494944415478DAED9DF956535714C6776692
              100CB32232D4228AB5D50E56DB7FDB17F069BAFA005D7D197D81F65F97735BE7
              5A5C2A2220A0C8103293747F37377012C8CD4DEE8D9193FD5BEB282A490EE6CB
              BEDFD9679F7D3D54628AC7EF3C7EE2D143827078D8E4F1278F5F78CC7AF89769
              1E3779C4DB3D334170C03A8F4B10F4551E57DA3D1B4170816B10F40689CD10F4
              20014117DB3D0B41700B11B4A0152268412B44D0825688A035A23BE4A5919E20
              C5C33E0AF83C5428106D6777683991A795448E8A9FC03B1D097869B0DBCF73F5
              9197FF9CDD29D2467A8756B7F3B453703E4111B4268CF70669341EACF9EF9BA9
              1D7AF836F5C9CE31C7C29E5DCDD08754DED16B88A035A09E9841325BA0BF1792
              6D9DE7C5B1A871E5A8092BF1DF9534BD4F362F6A11F421271AF4D2F9E311CBEF
              C9E68BF4602949997CFDB71A8288F073C6BB7CFCBB8FC2018F21429FD7633C0F
              0497CE172ABEDFAE808EF073CE0C87C9EBADFD3D05B61DF7DE240D2BD20C22E8
              43CED460170DB127AD054408AB91CE152C9F071F8CA1EE80E16FADA2E8EB0F59
              9A5FCF1A5F43A06786BB08D6F71DFBF4B7ECD37125B0A2378CC784C96311A817
              36B2F46A2DDBD4FF8708FA90F3DD892805FD07ABC38E9821E431B62C7D91DA1F
              8A3258543EE6E7C3220E4C0DF0872956F9B80F6C179EBFCB584658BCD6E9A1AE
              9AA2C67C11A59B41047DC8F971B2FBC0BFAF2766588871F6DDC77A022515A8B0
              22121C693759B8497E7C962D065C06AC464E112AA2EDE9A1FD1662792B6788DA
              0A4B51F34B5C7F9568EAFF43047DC8F961A27B9F28EA891951799AAD4A3858A9
              C4148B7889C508FB90B39942F3F3076320EA373E1811F3F96049604DEA6125EA
              EB2F45D01DC9F9910845437BC2CC73047DC462DEAEE1650F8AAA995C915EAE65
              1C65170016927EF6DF781EBB39EF7E53D4EA55229DC3C270BBA93988A00F39A3
              478234DE574AD9D513733F4752446635222E6CE4389A66A85E40C6864890C5BA
              6EFAE746F0721487CA0A35548E85E8A9813D512FF19C5EAC651A78853D44D087
              1C78E1AF8F47089AB112333212678FEE6517901E7BB69AA1351B5119517D861F
              0B56B6F234FB2E6D7B7E31BE7A7C613EF6C9727A7741594D59D410231684997C
              C1F66BA888A035C088805412E94184FC5EBA3012269F998E43247FB29CA2AD8C
              3DD1C01220BA03F8F33BF3F6EDC0182F3C4FF4EE5D41EE2FA62AF2D82AB1908F
              C558A484CD791D84085A73F0069F3B16A618476880FA8EC72CE6CD06ACC31916
              749F296864396EBFB62F685895AFD8E7973DFB16BFEEC3A554CB442782D61C64
              1F3EEB0FEDFEF9BFD534AD261A5BFCA9826E66C136C88F3D85859FC9CBF7195A
              DCCCB5E4E715416B4C80ADC837A3915DAB012143D08DE254D0E0142F4607CD1D
              CD1D8EF277D927E75DA8AEAB4604AD31AA7F8588EE2D242B3646ECE286A0AB3F
              5C6FD6B334672357DD2822684DC13A11DBE27E5340736B597AB3D19C80DC1034
              387E2448134A8A118B4BB783B4085A53B07B376DFA5614CEDF9D6FFE127F7638
              4CF1486951E944D048315E3C11D9CDCA3C5D4EDB4A1B3682085A53D4545BA3B9
              E36ACE1D0D534FD8B9A0C1445F882375C0F8FAFE62D2518AEE2044D01A8237F5
              FBB1E8AE5F4585DC7AAAF11DBE32AAA051EFF197838302D8D8C176376AB3B732
              CDCFC9EA6717416B865AF48FDDE69B7309475E5515347622FF69F3C9172B44D0
              1A8242FDA9C152EE191B190F969C9D2514410B6D0505FB27CC33864EFD331041
              0B6DE5F381100DC74A0B2FBBB5C956A88246C664C3F4E3F0C33EABB354541217
              0AFED1A64065A42760D498CCD9A8F46B0411B4865C60FF5C2EB6B7737AA41E58
              60FA7D1E47CF716B6E7B376DA86EC763B1FAD8C5F60A22680DF976344AA14049
              806B1C199FAE38B31C75DB0FD800054DE55D4A35A588C305771DA401AB11416B
              08EAA3CBC7AB600F1E398C80A58A39D4527B2A36678A46D1BEF5638BFC4D2B89
              7CC52EA59B69C06A44D01AF2A5522EEAB660DC40FDC0B991855111416B88DA5E
              003513B71AA85FFE185C52367D56D8E3CF3AF4F82A22680DC1696A348001B005
              37E6129F44A34680328ECB137BAD179EAF666839E15E6DB4085A43904943142C
              1701E184482327545AC9D158804E0EEC1D38B8F1CAD92EE6BE9F9D44D05A32C3
              11BAD7EC86B4B499A317EFDDBBAC3B015BF251D33FB7C20E89A03545DDFE6E55
              ED71A3187663BC7BB75D811B39F26A44D09A525D7BDCCA737C76C1EEE064F97C
              23ABEEF6FC76532768AC10416BCC645F8846CCDA632C0AEFB4404076F19947B0
              CA1B348B1B39A35B93DB88A035069D8E700CAB7C896F67D3736C751B8D21A9D4
              4A0187049AED016D85085A73D02D296EEECA01B5BFF3C7424D23B67A0E2268CD
              C1A51EC5456A515C2BCEF2D5021FA619A5C139AE12387AD5AA05AA08BA03303A
              7C2A1192CCC6E5CD345E6C84B8D94FAF6C79E0E3514B9DCCB97B8E504504DD21
              A887530DF85DC796F38A8BBB742AC3DD01A32E5B6D938B5CF8528B332D22E80E
              02AD7407AAEEC7326FFA59B74400414D2BE5A165D044FD59135D9B9A797D1174
              0731C551732816A8F8BB446687A367D6F1296C44E5C9FEA0E1DB55DC38066617
              117407820D0E5890EAD353383DF2762B67DCF8C7EEA20DDAC5163BDA8E45AA6E
              71016521D7FC31377444D01D0A36384EF687F659038362E996C5B8AD321AC1E0
              606CF99E2BB8A70A44D315F0500F2FFAD048BD3A2203F491C6C2B3950BC08310
              417738E8CE3FD61B32EE13EE0668BA3EBF9E6BBA8F9E5344D08201F2C528ED3C
              3062D701E938E4B561579C7468720311B45001AC083C766FD84F21BFC7B018D5
              F7314414DE642B824D12DC330563A7DDA57C262268412B20E80D1E3DED9E8820
              B8400282BECAE34ABB6722082E700D823ECDE3068F78BB6723080E58E771B96C
              F7A779FCC6E3671EB176CF4C101A608BC71F3C7EE5F1EC7FB193CCC55476FD3A
              0000000049454E44AE426082}
            Transparent = True
            ExplicitWidth = 180
            ExplicitHeight = 60
          end
          object lblRemainingTime: TLabel
            Left = 4
            Top = 26
            Width = 157
            Height = 27
            AutoSize = False
            Caption = '00:00:00'
            Font.Charset = ANSI_CHARSET
            Font.Color = 2649285
            Font.Height = -27
            Font.Name = 'Crystal'
            Font.Style = []
            ParentFont = False
          end
          object Label3: TLabel
            Left = 7
            Top = 3
            Width = 92
            Height = 17
            Caption = 'Remaining time'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 2649285
            Font.Height = -12
            Font.Name = 'Century Gothic'
            Font.Style = []
            ParentFont = False
          end
        end
        object pnlNextStart: TWMPanel
          Left = 364
          Top = 9
          Width = 180
          Height = 60
          ColorHighLight = clBlack
          ColorShadow = clBlack
          Color = clBlack
          ParentBackground = False
          TabOrder = 2
          object imgNextStart: TImage
            Left = 1
            Top = 1
            Width = 178
            Height = 58
            Align = alClient
            AutoSize = True
            Picture.Data = {
              0954506E67496D61676589504E470D0A1A0A0000000D49484452000000B40000
              003C0806000000CF2E9087000007C14944415478DAED9DF952DB5614878F24F6
              C5987DDF219090AD69A693FEDFBE409EA6D307E8F465921748FF6F3B5D42129A
              B0070824615F8C592DB9E72759D8065B76AC6B1CECF3CDDCCE34047127FE7439
              3AF7DC238D1C4679FCCAE3071E0112849BC3018FDF78FCC4634EE3FF8CF1F883
              47B0D03313041FECF17802A19FF1785AE8D90882029E43E87D923043280E0E21
              74B4D0B310045588D0425121420B4585082D1415227411515BAB51679B468180
              4615651A59FCC9868FA2B4B9C563DBA2E857F049575769D4D2AC516D8D46BA4E
              747E1EA5FD10D136CFCFB4FC5F5F842E127ABB75EAEAD0ED0F34150787517A3B
              637EB5733C8B446971C9A2BD7D7F3A8AD0450044E96651BC383A8ED2EBB78515
              FAD103C3FECD910E8838B760D2CE5EEE4A8AD0371CFCEABE7BDB20CDE3EF9CF2
              AFF5FFA64D3A3BCB7C3D8D2F54C3D7ACAFE3B0A09AA88A4384F23222DD70BE6E
              F13D711E213A398952F89828C42BFF118735D98433817A8DC6470DD23D266B72
              9C343965722892DBBF87087DC3191ED0A9B539FDEA7C768E50C3A29353EF8FB9
              BA5AA3B6168D9A9B34CF5534E5CFE070617B274A1B1CAB1F1F7BFF9C608346B7
              86BDA5FEF8D9A295B5DC026A11FA86F3E83EFF1A2F4F6D473632D7B0C8BDDD1A
              353678872CD9B2BB6FD187B5A81DE2A4A331A8D1E8507AA98F79BEAFA6720B8F
              44E81BCE936FCB52FE7926990DDD89BDDBDBAE3EA4E13B0EC3513A443871C2D7
              3A8B9219F3CBE0D0A3A242A39A2AA23A0E4BEA6AB594DFBFBE01B1D3672EBCA4
              C6F7FFF94F24A77F0F11FA86F3DDA3B22B52649219ABF2E8906EA7D01239E2B8
              787D83C3875D8B2259FA54C6F75373236E0C489E7CBD63BEDEDCA29576B54E27
              B5085DC2DCE307423C18BA9C733CFB6E36BD44C100C7B023C91241FCE5558B76
              7D64170004EDEFD1A9AA327E71E4C267E74DDA3B487DED26483D9CFC508BF94C
              4AC8519A20AFDBD7EDC4BF99646E6C70E47165C6DFC203D8DA278B2C059B1A00
              9B25DD9DC9F966483DFFDEA49DDDD4F36AE107D1E1C1B8D49F385C59FE200F85
              250962E1FB13862D9297CC002933ACD000E93184037E3732D2816C06C21A2376
              F740EAE959D3DEE049852B35D27FAFA622749A458A3115227411A0C712149956
              D98E769D063824409A6D66DEA27038BF1F3DB6E2C747742A8FA501236694DEBC
              35D3CA8A074CCCC8CFBC44E812A3A282EC073E5521462620E99DB17898130A3B
              5BF0F9AA2B11A185BC83CD9AD141E3E2FF97383EFEBC919F3B4A8416AE85C41D
              4D841EC862649B1AFC124468E15A403DC883BB0695194EECB1FAC9A2D58FEA57
              69115AC80A1416E1E1D34F56E4728AF1E51B53792C2F420B3648FF7575EAB660
              C84D273EB421D587941F78BF62D1FA666E1662DB1CB5276E2A6F76317D6E3A57
              4468C1069B21BD5DCEEA8902A3D985B8D43DFCE73D9DCED770F2656129F76575
              B05FA7F616E75A3B7BCECF5189082DD874B4E934D01BAFB84B943A5168ACCE58
              A57305C7C3EEC4567B6CEEFCF552EDA103115AB041613F76F69A8257A5C6EAED
              0AED675B1A200E7FFC305E5085533447C7EA1414A1850B20F5AD613DA9361A52
              878F4899D06062DCA0FA5AC7E8F92593B6B64568214FA4921A2B284A4E810AA1
              8706746A8BE5A455A7EF44E82204D904C4C465A96BFFEDB8D82B5D06299A1AB5
              2BF5D24085D02A63F2547317A18B8CBE1E9DBADAD51CA9BA8C0AA1DD2229B0B9
              63D1C27B115AF0A02FD6FF221FA8167A8B859E17A1052F9049C059C18AF2345F
              D7E225A7E9C0CE6065C5D5904345CC2B218770AD0CF6E9D4DE1A371E6714DD93
              E52A841EEAE787C2D8E6CADA67E730AD2A44682189CB32EF87A2F661D78E5675
              5909D44707EA9C1B6461D9A2CD2D115AC803A9649E9937ED76079D6D6A84465A
              F0F1C3783DC7D4B469B74C5085082DD820044028E002C9DECD99763F8E44D1FD
              0A8D16631363F1ADEFBF27D59E5E11A1051B142675C71ED412650689C5F97E85
              46BD48476CB5476B83E939A9E510F2003661D053031B2EE82B672678D6D9A153
              7FAC8ED94FC927322BDFDC332E0ECDAADEF60622B49011C4BD58A11122A02963
              AE206C41F80270AD7F5F9B49378E92B992082D5C03D88EC7112CB7B3E9C7758B
              5656E508967043E9EF8D674AD0C071722A92730F682F446821EF340434BA3D1A
              6F63808D146CA8E40311BAC4C083D975359901A8D89B18D72F4E7BA31415B9E7
              7CCD41842E215A9B351A1A30ECD749A01558A6AEFE7E4117D2DB633A5596BBFD
              F488A6DE99F6CE63BE10A14B88C4668D6823309DC7FE7697FBDAD92F04CAC329
              EFCB88D02504DE0F3832108F65D111149906D56DB9B071829AECC41ED48B2B16
              6DE4D8FEE04B10A14B0C3BECE837ECDCB20B1A28A2DF9CDFD51AAB327602DDF3
              8200DBDA8BCB266D2ADE404987085D82A0D6796448BFF2B62B6C45A33E19DD91
              B2ADAFC08D815ED0D83471C31917B4ED9D5FB4E820747D8A89D0254A7939D140
              9F4ECDC1AB95FE68A608094387CE7B52F07EC3E4970639D98BFA3AE7E6703318
              89E03D2D58F5F3916BF642842E7190234679685D8DE6FF62CCE151D4CE33EF1F
              14462B115AB081D878F1666350F77C29662AF070B9BB67D92FDE2C94C82E22B4
              9004420A9C26A9E750026FD7AAAA44259E66377304C8254738363E39250AF36A
              1CE2D004EF4D515D64942B22B4505440E87D1E81424F4410147008A19FF1785A
              E8990882029E43E8711EBFF308167A3682E0833D1EDFBBCFB3633C7EE1F1238F
              FA42CF4C10BE80108F173C7EE631F33FC40380C5DF3DAB370000000049454E44
              AE426082}
            Transparent = True
            ExplicitWidth = 180
            ExplicitHeight = 60
          end
          object lblNextStart: TLabel
            Left = 7
            Top = 25
            Width = 127
            Height = 27
            AutoSize = False
            Caption = '00:00:00:00'
            Font.Charset = ANSI_CHARSET
            Font.Color = 830136
            Font.Height = -27
            Font.Name = 'Crystal'
            Font.Style = []
            ParentFont = False
          end
          object Label5: TLabel
            Left = 7
            Top = 3
            Width = 61
            Height = 17
            Caption = 'Next start '
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 830136
            Font.Height = -12
            Font.Name = 'Century Gothic'
            Font.Style = []
            ParentFont = False
          end
        end
        object pnlNextDuration: TWMPanel
          Left = 547
          Top = 9
          Width = 180
          Height = 60
          ColorHighLight = clBlack
          ColorShadow = clBlack
          Color = clBlack
          ParentBackground = False
          TabOrder = 3
          object imgNextDuration: TImage
            Left = 1
            Top = 1
            Width = 178
            Height = 58
            Align = alClient
            AutoSize = True
            Picture.Data = {
              0954506E67496D61676589504E470D0A1A0A0000000D49484452000000B40000
              003C0806000000CF2E9087000007114944415478DAED9DF952DB5614878F24DB
              EC603643CC1A0881846C9DA4D3F6FFF605F2349D3E40A72F93BC40FB7F339926
              19B2117008940001C2629B1D2CA9E72745D806EF96706D9F6FE66686605B37D1
              A7EB73CFD53D52C86682DB1FDC7EE6D64E82503DC4B9FDC5ED576E1185FF98E4
              F68C5BB0D23D1384328872FB11423FE1F6B8D2BD110417780AA163246186501B
              EC4368B3D2BD1004B710A1859A4284166A0A115AA82944E81A425389BABB55EA
              6823F2FB15320CA2834393B6B64D3A3AAEFC6956D8B65EEE5F6F8F422DCD0AA9
              FCF369C2A478DCA42F9B261D1C94DF4711BA460876283436AA52C0A75CFA1D4E
              F0DABA419F578D8AF6F1F64D8DDADB948CBF73AB8F22740DD0155468625C2325
              C76B3012BE9CD12BDACF7BB7356A6E5272BE66B54CA945E82AC7EF277A7047E3
              7023BB28069FE10F119DE27BC59FEAEBC31CC2F0E88F7006E87C4D9C25888E4F
              4C3A3C24EB3311D614422040343DA95143207B5FF1496F67F5823FF322227495
              333CA052B85FCDFA7BC83CBFA0533456FC696E6951E8EE9496F7752B5F0C5A59
              2B6C546D6C50E8F6248746FEEC527FDD366861A9B4515A84AE72EEF3E8DCD490
              590EC81CF9A4D36E34FB29C6E46C30AC50B05DA5CD2D8316979322F97CFCF9D3
              1AF97DB9C3048CD2EFE70B0F67F2497D7AC6E1D1EBD2C22311BACAF9E1A12F63
              EC9C4F668410433CBAF785D4F3F7E395CF5F26C84C790BA446DC8B8C09B2143E
              CDCEA0B434F3DFF3C5809F973E1B148BA71FA789DF737C6CA67D562AF9A47EF6
              2251D2FF87085DE564123A9FCC10F4E6B86A4995CAF6AEC1EF2B3F13327E5DA5
              DE2ED58AB3E7170C3A3CCADC8F6C52A3FFB8B04A4184AE7230214C1513271332
              EFEC663EAD48EF4D8C69E7933C801C35320B3B51775478783F19A6E886DD9F6C
              313C2EAE5B37D5B4B00613C237B31272D42523432A5D0BD976E2442E2CEAB4B5
              93F9947676D8E93D2721E2E47E57795267B898A2EEEB5569743819CA60C4FDB8
              98FD22BB28F5325F5CE8572988D0554E430013439F15DFE692194C4D683CF973
              464E93221C0E44E3DE9C7E1C6782C31A279D982F75E8488D0BEBF53BDD1AD94B
              4184AE01905EC389DCCFB374DCDFC723E7A06A2DB2CC458C9273BD05F78B278D
              5313C99137A17328F15EA793D3CCAF57BF8541E57C5B88D0750616371289F2A4
              2986D6164CFC92614E7CDFA4F773DEAD588AD082E7747771F8713DB940F3FC55
              C2B30B4A8416AE84C1B04AFD218536B74C5A5EF1EEEB4184166A0A115AA82944
              68C1020B2DE16B76DA0C3960B34AAD10A1058B019679286CE7CD766386B564ED
              95D49D41857ABAEC783AE6721E5C84162CFA432A8D0E25D7C3BD921AB9E6470F
              7C561A0F79E91733BAABC710A1050BAC34DE1853A93BE8ADD438CEF7DFF9CEF3
              D233EF7457F73B8AD0C239900D77E17576782BF59D5B1AB536DB46471675DADE
              11A1058FC824F5D68EBD83C42DA9C746540AF5D89F5FCC6E9782FA4F2274CDA1
              69764C8C9BF333013173ADD4418AAE4E859A1A93B77462C97A76DE9D78171B0B
              06BE6D1B5BDF34AC0D026E2142D720C3832A85FBD4F23FE802D8A2F5E9DFF2E5
              BBC6328F0CD8FD2B67FF602644E81A24DFC6D952D989DAF174B93877FD018433
              1F1745682107488D61AF60C09FE5F74AF256CD6CA0204C6AB98193537B1749A2
              B49D5169A4E6BC37B031D78551DF4184162E815A1CD875E210DB3369EEA3EEDA
              1D723DDD0ADD18B5EFBE43FC8C38DA2D4468210DAF6576408D3B64541043CBC2
              8AE0091765C676A90F1EC8EC2522B46081BC30F2C30ED8CE351BD1ADD25FD584
              082D58609286C91AA8569981082D5860116664D0BE7D146504BC901931331655
              1A1B700CD3AAACE4FA31488416AE8810C7E763C3F6B7805B559A2E22420B5702
              3610A0B0A453F66B6DC3F0646FA1082D5C09A9159E5074E6D59B049D9DB97F1C
              115AF09C8E76C5AADAE4AC3B7E5EB3CB8F7981082D780AEED89B9E52C9A7D93A
              1FF2441015FAA52E875075207372F756F21114A857F7F6834E4747DE2927420B
              9E81E5EDF1D16465D45C657EDD4284AE43900F46517214535CDF3069E3AB37DF
              FF8D1C6EDCE17043E370038FBAD8F4E83869FF3612A1EB8EB656C57A1A95C3DE
              BE494B2B862B0FBEBC08D2752A1FCA8B8C462644E83A04CF45B937AD5D7A1404
              6A456F6CDA35A30BBD030EA33D6A41877A15ABC6339E04B0E5E2A6D76211A1EB
              143CDF1077D775052FDFE98F62E8AB5FCC8C55F411463406EC8702B5B6DA32A7
              3E23112509509AA05288D0750E72C4D883D892E109AFFFCC24D276A860828789
              5E2ED6394E5E5AAEDCFDA622B460D1CE62F771D880F20518708F4E4CEBD110A9
              A1C7A307DA793E3915ACFC4563064F2EDD2FED552C22B490062671C87E1C1C99
              97EEB8C3C6DB70BF4209FEFB9313FB6955FB3CA144CCFD7FD90420420B350584
              8E716BAF744704C105F621F4136E8F2BDD13417081A7107A8ADBDFDC8295EE8D
              20944194DB4FCE947592DBEFDC7EE1D656E99E094211EC71FB93DB6FDCE6FE03
              8DCBE3B6C5902B8F0000000049454E44AE426082}
            Transparent = True
            ExplicitWidth = 180
            ExplicitHeight = 60
          end
          object lblNextDuration: TLabel
            Left = 7
            Top = 25
            Width = 127
            Height = 27
            AutoSize = False
            Caption = '00:00:00:00'
            Font.Charset = ANSI_CHARSET
            Font.Color = 830136
            Font.Height = -27
            Font.Name = 'Crystal'
            Font.Style = []
            ParentFont = False
          end
          object Label7: TLabel
            Left = 3
            Top = 3
            Width = 81
            Height = 17
            Caption = 'Next duration'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 830136
            Font.Height = -12
            Font.Name = 'Century Gothic'
            Font.Style = []
            ParentFont = False
          end
        end
        object pnlOnairStatus: TWMPanel
          Left = 1275
          Top = 9
          Width = 112
          Height = 60
          ColorHighLight = 2497307
          ColorShadow = 2497307
          Anchors = [akTop, akRight]
          ParentBackground = False
          ParentColor = True
          TabOrder = 4
          object imgOnair: TImage
            Left = 1
            Top = 1
            Width = 110
            Height = 58
            Align = alClient
            AutoSize = True
            Picture.Data = {
              0954506E67496D61676589504E470D0A1A0A0000000D494844520000006E0000
              003C0806000000871D8F07000006154944415478DAED9D594C15571CC63F88A0
              8981E00EFA80441A1057F4418A4A34D200AE49436A5C121FDC835A7711E31235
              6ADC635CA246630CD507691FDC80D046A3561B63C52D2AA58AB189C59D627D10
              2B74BEFC9D3298A19CBBC8BD279C2F99DC997BE7CC9C7B7EF3DFCE9D7B6F481D
              80C123267E613D6CB296746B89845130AADA5A7EB496DC9F7F3A561E923A6262
              42586DED2F39BF974665545620F27D4DA03B68E4A2EAB0701447C7614F7C72D5
              FBD0D01482FBFEDBDF7EFDFA9B3FCA02DD372305E5C726615F7CFF1F08EEAFC2
              0B0591C6D2F4102D2F2B2DFB6F82ABB37C66A0FB63E481AC9C04069C8632E034
              9501A7A90C384D65C0692A034E5319709ACA80D354069CA632E0349501A7A90C
              384D65C069AACF03AE674F60C20460F870A0776F202A4A9EAFAA02EEDC01CE9D
              038E1F07EEDD0BF4FBD756FE059798086CD9028C1EADB6FFA953C0D2A5C0FDFB
              811E07EDE43F70B367033B7600AD5B7BD6EEDD3B60DE3CE0C081408F8556F20F
              B8F5EB81152B7CEBC9BA75C0AA55811E0F6DE43B385ADADEBDFEE9CDCC992DC7
              F26EDE04FAF605B66F07162DF2B8B96FE018D3D881F070E0C30760F162EB8883
              81EC6CB5F60505C0E5CBC0D6AD4068A8B8CDFEFDBD8B797575EECF8784F830BA
              9F510105C7E4C24E44E6CE0576EF06E6CF9758A7A2050B809D3BA5EDAE5DF5C7
              1C3B56BD0F4F9F029D3B03C5C5406666C3D78A8A808C0CE0D933A04B173F8C76
              F0C87B704CF9EFDE95F5F3E725F54F49012E5C00C2C280CA4A203ADABDADFD5A
              4D0D30742870F5AA9408C386C9EB49496AA502A1514D41B1E106ABF57921EFC1
              AD5D0BAC5C29EB43860057AE00A5A562FECB96495970FB36D0AB57C376ACE3B8
              0FCB804D9BC4650C1800A4A602172FCA3E2A890ADB11BEAA25111E2F987EFD64
              3B27473C045DD5C387B2EED49C39C09E3D9E8D896DE14ED9AE90AEDCE915B66D
              03162EACBF98F87E28F6CF76A36E5EC46770972E493CE39BEED1435C26DD1C01
              120261307EB13CA8AD9536761CE3EB84CC7D070D02C68C014E9F062A2A80EEDD
              E5D8B4C4C6E41C74D5F860B7B181D8DBB76EC9203941D983EA093C3BC67E6AD5
              7CDE3E870A388AFB36E11DBC07F7FAB5CC881C3E0C4C9D0A1C3C084C9B064C9A
              041CB38E75FD3A909C0C6CDE2C1644D14269698446B0DC373F5FDACE98011C39
              024C9922332CEDDA357E6E5ED93CB6A7718B56C77373F06C70941B209E2326A6
              DE429B3A2ED5587FDCE2B01B38027383EF5770F6159697076CDC28B18D56121B
              0BB46A053C78206EB14F9F86EDE83E390D161727DBB432BAC8B43460F97260C3
              86A63BEF742B9EC8D9CE69716EC7F974601B938A75DAFBA880538CC3BE83633C
              A355D916D6A18340B9760D3879121837AE613B3E47D7488B7BFC1878F142DA0E
              1CD830236D2E708DB95B5570AAB1D69318A720DF5D253BC0FAADB0503A45788C
              7B2F5F02D5D5404282C0A13A7604CACA80C848A07D7B203E5EA0D12D656589E5
              E6E636AFAB6C0A5C53714E75C09DE70E2838BA37669367CF02A34649FC62A6B9
              7AB53CEEDB07CC9A059497CB1B6707395884C59916AE73DF356B2483641C3C73
              061839523D39F12479682C39F1159CEA451434166797036FDF4AA7BB7593BA8E
              56C61A8F8FDC6696E8D4A34752A7D1EA3843121121DB4F9E48BADEB66DF39703
              BE8053D9CF3E5750807316E0F6AC09E719A74F97349F71EFD02171894EBD7A25
              592863230B76B6E11CA573F684C75699F6522DC0DD02BFBFC0A9F44335AB6C16
              70943DE5F5FCB9580D6B34BA50D5A4813E9FD9649B36721174EAE4FD94971B80
              FF9BF2F22738CAAD8EB3CFC1F30715384E32DFB82145764989401C3F1E387A54
              ADFDE4C9C0891312DBD2D3053C3B5EE6C537633D9D64F63738CA5F33270AF2FD
              631D16CEFBF7CB3A6BA2AE5D257B54112D9571C7AEF55AD2C73A3ECA7C90AAA9
              FC77EB022D8FC985B975A15914F89B85962CF12EA6B57005EEF63C4E449BBBBB
              BC96B921565319709ACA80D354069CA632E0349501A7A90C384D65C0692A034E
              5319709ACA80D354069CA632E0349501A7A96C70E6C7B43592F3C7B4CDCFD76B
              A4EF6293B0F7E3CFD72786D5D65EE11F4664FE5981887F8CE505A3DEB40A4751
              CC7F7F18F165C8C7BF6849B01EF85599AFAC2522D09D3472D51B6B29B1963C2B
              2729FB1723CBC0798063F9D20000000049454E44AE426082}
            Transparent = True
            ExplicitHeight = 60
          end
          object imgOffair: TImage
            Left = 1
            Top = 1
            Width = 110
            Height = 58
            Align = alClient
            AutoSize = True
            Picture.Data = {
              0954506E67496D61676589504E470D0A1A0A0000000D494844520000006E0000
              003C0806000000871D8F07000006684944415478DAED9D6B53534718C7FFE122
              F7700D8A16E208C8551441412B1405ABA8BDE88B76A6EFDA2FD0B79D7E804EBF
              4A3B53ED458B17502B740C152D784304B5800212AE81008A90EEB374D36348E4
              901C93ECB8FF9933E49C3DCBD9B3BF3CB7CD99C4F4E5575FA3A7DF9E0FE07BB6
              35B0CD0CA57094836DCD6CFBA6C06AE935EDAFFFA2C06482CD929A98624E8845
              648429D40354F2A2A565171CCE05D82767A75C2E5413B89F32D3124FA526C585
              7A6C4A3A34E19863F09CA709DC745E768659599A1C22CBEB1B1C9B25702EE633
              433D1EA57588E52450E024940227A9143849A5C0492A054E522970924A819354
              0A9CA452E024950227A9143849A5C0492A054E52BD15706673227272B2B03133
              1DC9C989888E8EE6C7171717313D3D8BE7A3E318181886C3311BEAFB97568682
              33272560E7CE426CDE9CA9EBFCA1A15174753D8063C619EA79904E8681CBCBCD
              41797911222222D6D56F797919B76EDDC7A3C783A19E0BA96408B81D3BB6A3B8
              2837A081DCBBDF87BB777B433D1FD22860706469151525860CA6A3E36ED02DEF
              A31375888F5F79D6A6B3B31B3D0FFF79E371A3547FA80A1919697EFFEF80C051
              4C3B72E400778F2E970B9D2C5E6564A422FBBD4DBAFA0F3E1DC1F8D8248F8B26
              9389BBCD0B17DAFC8A799F7FD6E8F5F80F3F36F9EC53B07D2B76ED2A5A758EAF
              E3462AA4E06A0E54B813118A53BD7DFDD8CE6EBA9CDDB41EFDCD06FD900D3A3F
              DF8ADDE5C5FC18252CAD6D37758F4158467FFF33D8DA6FBFD6565D5506AB750B
              E6E6E6F1DBD9ABABFA52BBC592B6AACDD7F17092DFE028E56F3C5AC35F8F8E4E
              E0CAD576A4A7A7E0D0C16A6681262C2CBC406C6C8CD7BEA28D2CACE5B20D1313
              D338585785CCCC34DEDE74BE5557A940D0486B4DB080EB6941EF24B8D2D27C94
              14E7F1D734F9E3E353F8F0F0FB48494942D7ED1E3C78F018478FD4F03A4EABE9
              E9199C67EEB0A8701BCACA0A303535838B97FEE4D0EB0F55F373F4242AE46A08
              86DEC9257864792D97DBDDFB228609B5314BDFBDBBD8EBF167CC13AC2561E15A
              095748AE5CEB153CDD31DDCFCA5CB6BBDDA8372F1230389A648A67B3CE399C3B
              F7077799E43A096073CB75A4A69AD150BFCF1DFF48228E51FBE4A483B7133072
              8DE4224F1CAF4342421CC658DCA337832F6D61D73AC0AEB59EF820FA68211869
              7122C67A5A351D1F1B9B5805C21738129DBB567CF51BDCA9930D7C45E4C993A7
              F8EBC61DECA92CC5B66DD9B0D9BAD03F30C4AD8FE091E5910591C8420B99A511
              34B232AB75339BA49D78CC32C91B2CA3ACDA5B86AD5BB7F01596D3679A7D5EDB
              5F57465666B74FB827CF28706BB96C6F71D81B3802E60DBEA1E0C43BECF69D87
              E8EE7EC4639BC592CA071FC12CEBF8F10FDC6E51ABA32C0B4D4E4EC2D9732B37
              495666B74FE2F2151B8A582D58C66AC2B506AF752BEB91673F23C009006F72A7
              E21C3DE0F466B2018313F14C58D8999F9B99BB8B67FBFBF98DB4796488221325
              8B733AE771F2D306B7056A335259C0E98DB5EB89716F159C70953D3D4F78FD56
              5B5B89AC4D165CB84840E63890C5C557F8BDE91A5EBC78C9FBC4C46CC0B1C65A
              D62F8ABBC2A4A4780E7C78C48E6BD73A78B242498B4CAE52EF847B5E3B64E044
              72323CCC26BDB583C72FCA34291BA49856595182DCDC1CCCCC3AD1D73BC03213
              203FCF8AC4C478F4F50DE0E6AD7B2829617D4AFEEF535BC3E067597427277AB3
              3D5F7D8C00A7F7DCB0B138510EBC7AB5845F7E6D415C5C2CAFEBC8CAA80EA3BF
              B44F59A256E41EA99DACEE58630DA2A2A2F8FEFCFC023EF9B89EED4706A51C30
              0A9C9E1827DE3461014E5B808B551391595249D0CDE2DEDE3D3BB06143F46BFD
              5EBE5CE45928B9442A05686D92D628B5AB274DCCBDEA59F6D25B80FB0AFCE196
              5506051C49241A14C3C86A969696B90BA5225C8F2829A16C32323292BF092806
              FABBE4E5ADA60BE69297B73A4E581A5D3FACC0691799479E8FA1B5F52672B2B3
              50C56E5C8F6CED5D181C1CE1B16DE3C6745E9C53F93013A44566A397BC8C5A39
              D1A3803FD6C965AEB192B948122D5FC5C5C570CBD123B2545AB7A4BA8E148A8F
              756495FA20555219F6E802591E2DD0AA471782A3903F2C44C5BB3F31ED5D57E8
              1ECFEB1F524F770520F540ACA452E024950227A9143849A5C0492A054E522970
              924A8193540A9CA452E024950227A9143849A5C0492A054E520970EACBB42592
              F6CBB4D5D7D74B24EDD7D7179A4CB8AE7E3022BCE5F18311FB4CFFFD444B016B
              FB8E6D87D9A6EFC148A5606B866D97D8F62DCB497AFE05845EDDEEF3314D4100
              00000049454E44AE426082}
            Transparent = True
            ExplicitLeft = -4
            ExplicitHeight = 60
          end
        end
        object pnlRemainingTarget: TWMPanel
          Left = 730
          Top = 9
          Width = 180
          Height = 60
          ColorHighLight = clBlack
          ColorShadow = clBlack
          Color = clBlack
          ParentBackground = False
          TabOrder = 5
          object imgRemainingTarget: TImage
            Left = 1
            Top = 1
            Width = 178
            Height = 58
            Align = alClient
            AutoSize = True
            Picture.Data = {
              0954506E67496D61676589504E470D0A1A0A0000000D49484452000000B40000
              003C0806000000CF2E9087000009874944415478DAED9DD95623D71586B76681
              668124A01B9AA91B4FB193955C24F7C90BF869B2F200597919FB0592FBE4262B
              B6DB43CF4CDD4C9A0724044295F31F51D5A5A18A2AA9844CB1BF5EC716024A25
              F1D7D17FF6DE67CB433D9E8AF10F31FE2C469C18E6FE5013E35F62FC558CD71E
              F19F1D31FE234672D667C630135011E38F10F437627C3DEBB3611807F81682AE
              12DB0CC61D34206865D667C1304EC182665C050B9A71152C68C655B0A05D4420
              1CA1482A47A1F918797D0152942E5DB59BD4AA16A8592B8ABFF4ECFFD4FE6098
              E6E20B1408CD93C7EBA56EE78ADACD3AB5EA4552BADD898FCF827609B1C54714
              4D2F51EF4F3ACC65AB4EC5C397BFDA73EC5E5F51E5648FDAE7D5891E8305ED02
              7A425936FD998E98A9F3FB3FCFF43C735B5FC9770E63142A1FBDA58B4665ECC7
              6041DF73F0D6BDF8E453329A99C175E7928A072FE4FFAD1E33286C4B2034473E
              61117CFE0079FD41F9BDAE38C6B5B009D79717C2CEB4E852D805D81A2BE098E9
              474FC9E3F11AFE8CA25CD3D9EE8FD28A8C030BFA9E935C5A179E74D1F0FB1071
              49588DCE55DBF438102F8E33174FDF328B0E03BBD0AA95C42848919B118A2428
              F5685B08CFF8023C2F9F502DFF7EACD783057DCFC96D8AB771FF68015A113366
              E3E8E20A8523CE94F25C9C57A85138329DB5C3D1242557B60C45DDB9BAA0BC98
              A5C781057DCF597EF68791F7DF26668FD727BCF70A4592591AB62B0A5D5D34C5
              42B2419DCB969C55C3D194FCCE45A32C176EFEE01C05E7A21408CF8FFCFDF3CA
              19D585B095EEF5C8C73717B542C7AFFE3BD6EBC182BEE72C3DFBFD9028E03F8B
              872F0CC50C7B0131F903E1BEFB21DE66254FAD7A49D8888E767F3CBB2A849F93
              B7CF2BA7543B3BD4BEE7F5F9692E96A6F964468ABCEF7862A6AD88459E910D31
              16350BFAC18205612014D1BE86104BEF5F1A8A487A588848B730BB16C2AFE50F
              0DA30B6682D60381C633ABE40B84B4FB100B47E4C2281C87993FB5B249FA591E
              E773B6FB7CACD783057DCF415C37B6F858DEBE4DCC104F5288473F23620126AD
              81629CD4B02A68800B455A99D492769F22FE558EDE49BB320A2C44934B1BA48A
              FAB6C73083057DCF41B62DB3FE8510924788F995A1988373314AAF3ED3C48CF0
              585988CC4A2203B32E3290E0BC7C2A67F3DBE8BD13888BC7E3EB3D9EF8573A7C
              25133CA35045AD280AE5F77E14B3B4B510E3D0EB412CE87B0F440D8C52C7BE40
              505893CFC8EBF5CBAFBB5DCCE4AFC5C2EFDCD2F16125522BDBF276F9E88DE5C4
              0752F1E9C74FFB1EB7B0FFB3A158F1F3C0EA798D7C2D8805ED723CB4B0B643C1
              70547E056B81991C110C3B20BC07AC2651541009493F7EA679F6CB8B06150F90
              829F8EEC58D02E0761B978764DFBBA72B24B2D142ADD212846EA79E41EB5B303
              19D69B062C681783905A7643F8EB9BB77C54B4558E7767722EC9E50D9A8B2DC8
              DB8AB01E32BDAD0B0D3A050BDAC5C41656282A06E889E82799A69E0548A76737
              3ED72EAE46F188EA62380D0BDAA5C0B366B7BED41664F5C27B6A944E667A4E7D
              21465C606F7F300D178EF5BC8905ED4AC2B114A596B7E46D59C126C4D3354843
              DF155EAF4F5E646A28AF7CFC962EEAE5098FDA0F0BDAA5201BA8D65F346B05AA
              9EEC8D751C9F3F28C376282305281B45D8CE6A29EA2089A5759ABFA90E44A205
              5944276141BB120FE5B6BFD2EC06C274ED66CDDE1190F1CB3CA6483243A38B8F
              F254CFBFB76D1942F37119C603B01DA76FBE272725C8827621BDA2FFCFE46D08
              EEF4CD77B68487AC234487ECA219C8FAE162516CEC55C48592DBFEAD169746A2
              C56E6CDBF4F8C482761DFAB86F2F91F1C2D6EFC73133EB6A31AE3B6D610F7A29
              F27034216CC8C7E2A3718AF117D63ED1123D4EC7C559D02E04C541D1742F5C67
              D73FCBF0DAE66FB4191462AB9EEE69B33066EF446E5D5E340033FFD9BBE7B6C2
              817A1FDD281DC9E228A76041BB9044EE09CD2732F2366C01F6FDE9E988851D6A
              9E47A19FDDB1F0C3CE9141BB02B16736BE900B466036CBA2E868B0EE1A7B0B55
              3BD3ACE6C505B3EFD8736741BB90CCFAE743C5F683181519E97790233D8D34F5
              28904EEFED76C12C7B2C66D90F433FA32F6A32029B0AF27B3F39F6DC59D02EC4
              8AA02BC7EF46CED2B18C10744A15B4715DB2BE46BA511682CE0F0B1A3B5992CB
              9BA6E7C182666E45EF51D18F63308A809AE9F3328A8386FFF4B02AB02CF2774D
              36ABC272A856029601D661188F585C66E5962F3D88C2F86FAAF72689918F8205
              ED42F4B6C1AE47450F8EECE697A4ED1E1911C5E88F82286251F883ECD56115BD
              C737B22BE3C28276217D61BB315A80E905075070AFFA6DF862B5101F8CB3A85B
              58DDD116851CB6636E05825B5CFB54DE4684E2E4CDFF6C356A448B8345213AD5
              1618013B5310178B62A746C4E3A1A5EDDF7D4CAC1CFC32D10E95A1C3130BDA7D
              0C88062D0DECEE50412111BCB85A0F3208EA30E07DED163C6107CBC2EA27F2F6
              3817DBAD4F9D58D0AE043DE4B051159885DF6E038D64206A7FB0971DEC5CB6A5
              98D188661CF4E13E6CD02D7D78EDE8F36641BB14BD8F9E56EDB15D066BB4A7B1
              1D8C05ED52B0133CB7F5B108A8963FB809D5CD0E84F0E299DEFE465934F5F63B
              479A9CF73D6F6241BB167D780DB5168829CFAAC81F9E1CB16BB5B3E9241D46CD
              6041BB187425CD6E7C2C349AC44B4F8ADE3BCB82A6DDE763F780368305ED7290
              6041A245058BB0493FF6C12E589C6291AA82440A122AD38005ED72303BA3A1A3
              5ADB81DDDF858397B286C20EBE9B1ED4763282008FBBB8B6A3EDF6C6E316F67F
              99DA029505FD00C0274F61078B6A3DAC76F55741B9E7C2CDB6A922BA2E35EB96
              7ECF1F08517A75472B338588B14305E5ABD38205FD40186C5B2B3B957EB0D6DF
              4EBFB8B4BA98937DED84CD40B39B1E8A6C0E69D481D42958D00F0834254F649F
              685FA3232836BAA2A3A81976DAE902742AC5065B7DDBDEEAD9BE6CA63E6D58D0
              0F8C5EC2659DF43BB9314B57CF0E0C676BAB82C6AC9CC8AEF5152F415EF8FCC1
              BBEAA7C7827E80C013A79637873EEDAADDAC8A59B440EDF34ADF4E6E3341638F
              61289214B3FF2285E6137DC743ECBB7CFCCEB2E7760216F40305624EE4D64616
              1FA17AAE8DBD88AD065DB75B148EA7B5468B68F878512B912FD4FBD0A0D05C4C
              56E70D228B974E0FEEBC971E0BFA81138AC429B6F068C0268C0F6C4BBDF841CC
              F2F61ADB38050B9A91A0A3D17C42D88668D2F4935E4781705CBB51A166B560BB
              4393D3B0A0993E601F42C24AC067FB85AD406204A13755E4102F427E4890746E
              3E1AB92DAC89AD22FF699E3FB1A01917014123B11F9FF589308C033420E86FC4
              F87AD667C2300EF02D048D0D5EFF16C3994F2F6798D9806DE97F52D3453B62FC
              5D8CBF88111BFB900C73F7206BF34F31FE26C6CBFF03237782D48B0D4A130000
              000049454E44AE426082}
            Transparent = True
            ExplicitLeft = 2
            ExplicitTop = -2
          end
          object lblRemainingTargetTime: TLabel
            Left = 7
            Top = 25
            Width = 94
            Height = 27
            AutoSize = False
            Caption = '00:00:00'
            Font.Charset = ANSI_CHARSET
            Font.Color = 11239251
            Font.Height = -27
            Font.Name = 'Crystal'
            Font.Style = []
            ParentFont = False
          end
          object lblRemainingTargetEvent: TLabel
            Left = 7
            Top = 3
            Width = 96
            Height = 17
            Caption = 'Remainig target'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 11239251
            Font.Height = -12
            Font.Name = 'Century Gothic'
            Font.Style = []
            ParentFont = False
          end
          object lblTargetEventNo: TLabel
            Left = 114
            Top = 31
            Width = 58
            Height = 17
            Alignment = taRightJustify
            AutoSize = False
            Caption = '120'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = 11239251
            Font.Height = -12
            Font.Name = 'Century Gothic'
            Font.Style = []
            ParentFont = False
          end
        end
      end
      object WMPanel8: TWMPanel
        Left = 1
        Top = 552
        Width = 1388
        Height = 248
        ColorHighLight = 2497307
        ColorShadow = 2497307
        Align = alBottom
        DoubleBuffered = False
        ParentColor = True
        ParentDoubleBuffered = False
        TabOrder = 2
        DesignSize = (
          1388
          248)
        object wmibTimelineGotoCurrent: TWMImageSpeedButton
          Left = 1005
          Top = 8
          Width = 120
          Height = 30
          Action = frmSEC.actEditTimelineGotoCurrent
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D49484452000000780000
            001E0806000000E7C9DA95000002B74944415478DAED9B3D4C14411480DFDAF0
            73B90BA01567E8C899A821460B282440C49604A3959D367686821812628C8931
            16C6CE063A2B0824B460D4602185C446A3875606AE52815B116838E65D7CE763
            987507F62EBB37795F32B9CCDEEECDDBF976DECC5E76BDF35D3DF0EBC7F74E00
            78ACCA15553220D43345555EAA72AFED54C7572F9BCDE600BCA5E6544B4B4363
            0A3CEF44DC010A112895F66077670BFE6C6D6CA85A370A9E694EB50E3736A5E3
            8E4DA8223BDB45943C8B82375B4F9ECEC8C8750B1CC9EB3F577FA3E092CAD571
            C723D400B5B60211EC3022D87144B0E38860C711C18E5333C12377EF403AFDEF
            BE7AAD508089C917719F6F20B76FDD8495956FB0F876C97A9FFBE3A3F0FACDE2
            7F8F899BAA0BEEBDDC0DFD7DBDF0F94B1EA6A6E70E744EB6BD1D1E3C7C12F739
            1B11C196E049CFCCCEC1C74FF943DFDDB83E54FEE4E293C27104D70355158C02
            3399B4552A3675D6B9B339B83A38004F9F3D2FD769846046C0148F1980D7B11D
            3CE6DAF050E5377CDFAF1CCFDBB974F14265CAD0F7C17688A0A9C4B40F4E43F3
            0BAFCA1733B583B11134BA297B217A660B8B3F51828F7285DB0AE61DA5D7A973
            78C6A02982A602EA5C9E4A4D17E27146B02E586F07E34561EF973F1CD846F1DA
            C4EFBC60BDC3781D3B98771EC1A7026CA758F40F8D1ADE4EB504EBEDE036845F
            483C369BF8A392F814CDAF64BDAEAFD43994466DDA098A272C66538AE6C79BA4
            EB82C3E28F4A6C8B2C9B916523983AD8568AA99DA0FDC27ECB46B0BE4D171C16
            7F546A769BA4DF3EE8B749A6B9064F7875AD50911E263868BEE2A93CC9826DE2
            4F9C601E24C79472E804097D8519269864F15528C233489205DBC41F15F9ABD2
            7144B0E38860C711C18E23821D47043B0E0996C7661D843F362B0FBE3B087FF0
            FD0C80F74E5E5D7103EDD5951EEFEFCB6739F5DD2355065591A15CDFF8AA2CA8
            32A6D656F97D8B3623CBB7C883090000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D49484452000000780000
            001E0806000000E7C9DA95000002A34944415478DAED9BC16ED340104027C651
            7248AAA622958C9A8A4B48F903B8F5D01CB8F207F0035C111F80F819B872480F
            BDC127504542958AB0D4A0BA90A4C48A09EC448C35DDDA78A913D95DCD9356D1
            3A7676BCCF3BBB8EECCAB3E72F603078DF058037AA1CA8B201C26DE6872A87AA
            BCECF79F0C2B3B9D073DC7713E76BB7B9B9E770FAAD56AD1010A3998CFE7E0FB
            5F6138FC74B1582C1EA1E0B7BDDEC3A7BBBBF78B8E4D582127279F95E4E37728
            F8FBFEFEC1868C5CBBC0917C74743841C1BF55AE2E3A1E610DA8B51588608B11
            C19623822D47045B8E08B69CB509F63C0F5CF74E5C0FC310CECE46459F6F2ADB
            DB6DB8BCFC0993C9C4789F4E670782E0E29FC714CDCA05371A0D68B536613A9D
            C2F97970A5736AB51A9C9E7E29FA9C1311C186E0498F46DF60369B5DFB6E6BAB
            B5FCE4E2CBC24D04DF06562A1805BAAE6B948A933AAB5EAFABD1DF02DFF79775
            1A21981130C56306E0756C078F69B7EFC6BF1145BFE2E3793BCD66339E32F47D
            B01D226D2A49DA07A7A12008961733B583B11134BA297B217A66CB8ABF5482FF
            E70A3715CC3B4AAF53E7F08C4153044D05D4B93C95265D883719C1BA60BD1D8C
            17858DC7E32BDB285E93F8AD17AC7718AF6307F3CE23F85480ED4451746DD4F0
            765625586F07B721FC42E2B199C49F97D2A7687E25EB757DA5CEA1346AD24E5A
            3C593127A5687E7C92745D7056FC79296C916532B24C0453079B4A496A276DBF
            ACDF3211AC6FD30567C59F97B5DD26E9B70FFA6D52D25C83271C86B3587A96E0
            B4F98AA7F2320B3689BF748279909CA494432748E82BCC2CC1248BAF42119E41
            CA2CD824FEBCC85F959623822D47045B8E08B61C116C3922D87248B03C366B21
            FCB15979F0DD42F883EF7B8EE37C905757EC407B75E571E5EFCB673DF5DD6B55
            FAAA348B0E52C8C55895812AAFD4DAEAF80FB9752A8D151B67E8000000004945
            4E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D49484452000000780000
            001E0806000000E7C9DA95000002604944415478DAED9B3B4B034110C727ADA0
            D804B4D0CA5761A188A058D868935284F4D6B68AF80144626B6D1F7C9436A6D0
            42144454B0F1D12888421A1F216D9C312E37D9EC79ABEE7197657E30E4F67277
            33BBFFDDD9BDE32ED3B35D03A41F6D1D6D06AD038456E61DAD84B68276974181
            0771E314AD33E9C804A7BCA24D90C0BBB831977434422CEC91C06F2069D9572A
            24702DE92884F810813D4704F61C11D8734460CF11813D271681CF7300D9B6A0
            7C5D06C81D255DD570F6A7010E9F000AF7F6C73CCE036C5EFE7C4E1A702AF072
            1FC0E20840E90160E1ACB17186B300BD3B4957D78C086C09557AE918A0F8DCFC
            DFD678FD970B9F16FE2270ABE04C6012B0ABCD2E159B1A2BDF8D196014606CBF
            5E5623843202A578CA00BC4C7EE89C8DA9E01AE56A703EF7931F08A60CFD18F2
            A3089B4A4CC7D03454B8A87766E5876253A8D1ADB217A167B6A8F85DE04CE0DF
            F4705B817943E965D5383C63A829424D05AA71792A3575C4BF8C605D60DD0FC5
            4B82156F1BF7A9786DE2F75E60BDC178991A98379E824F05E4E7A5DA3C6AB81F
            DBD8A304D6FDD03E8277241E9B4DFC2E48758AE63D592FEB2B758E4AA3367EC2
            E2898AD994A2F9F926D17581A3E27741228B2C9B916523B06A605B514C7EC28E
            8BBA968DC0FA3E5DE0A8F85D10CB6D927EFBA0DF2699E61AAAF05539103D4AE0
            B0F98AA7F2340B6C13BF0B6279D0C1579D8429E5A80A2AF4156694C084BE0A25
            780649B3C036F1BB401E557A8E08EC3922B0E788C09E23027B8E08EC39F2DAAC
            DF54E4C577BFF97AF17D08374E403E5DF10DFA746532F3FDF1197D9FB486368B
            D69E7464C2BFF8403B405B45BBF90463F30708EF3AB8290000000049454E44AE
            426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D49484452000000780000
            001E0806000000E7C9DA95000002D34944415478DAED9B4D68D44014805F107F
            F690D873C5A3AC608B881EBA058BD6567B2B54F4B437BD281E44414414114544
            507A10BCE8CD534B0BBDD5B6A854D0228A072BBAD5A3DD73BB11D6B69475DEEA
            0BAFD3C48CCD2E4987F7C1B04C36D979992FF366B2244EFBFE022C5C7FB30700
            EEA9D2A38A07C266A6A2CA942A575BEE747E73760FD7F2B0BA32939B7AD4B26D
            F63938D54ADA010A09A8E53C586E3B01D59E0B0BB0656B070A1EC94D0C0E6C7F
            3794766C420359EA2C42B5FBFC280A5EDCF9A0CF93916B173892172F8FFF44C1
            3595ABD38E4768026A6D0522D86244B0E58860CB11C19623822DA769822F5D3C
            07AEEB06F5F972199E3C7D96F6F94672F64C11E6E6BEC3F4EB19E37D6EDEB802
            2F5F4DFFF398B469B8E0AEC31D70F448177CF95A82A1E1B1359DB3ABB5156EDD
            BE9FF6398722820DC1931E191D83D9CFA575DF9D3ED55FFFE4E2B3C246046F06
            1A2A18057A9E6B948AC33AAB6D5F1E8EF776C3C3C1C7F53A8D10CC0898E23103
            F03AB683C79C1CE80F7EC3F7FDE078DECEA18307822943DF07DB21A2A692B07D
            701A9A987C51BF98A91D8C8DA0D14DD90BD1335B5CFC9912FC3F57B8A960DE51
            7A9D3A87670C9A22682AA0CEE5A934EC42DCC808D605EBED60BC28ECFD878F6B
            B651BC26F15B2F58EF305EC70EE69D47F0A900DBA954FC75A386B7D328C17A3B
            B80DE117128FCD24FEA4643E45F32B59AFEB2B750EA5519376A2E2898B392C45
            F3E3C3A4EB82E3E24F4A6A8B2C93916522983AD8544A583B51FBC5FD9689607D
            9B2E382EFEA434ED3649BF7DD06F93C2E61A3CE11FF3E5407A9CE0A8F98AA7F2
            2C0B36893F738279909CB094432748E82BCC38C1248BAF42119E41B22CD824FE
            A4C85F959623822D47045B8E08B61C116C3922D87248B03C366B21FCB15979F0
            DD42960A45A81EFBF3E0FB5E585D795B7F75E5D33838BFFCB463131250DBE1C2
            727B1FBDBA5270FEBE7C9657DFDD55A5571537611B42BAE0089D54E59A5A5B95
            7E0326D24EB38C8E16900000000049454E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D49484452000000780000
            001E0806000000E7C9DA95000002B74944415478DAED9B3D4C14411480DFDAF0
            73B90BA01567E8C899A821460B282440C49604A3959D367686821812628C8931
            16C6CE063A2B0824B460D4602185C446A3875606AE52815B116838E65D7CE763
            987507F62EBB37795F32B9CCDEEECDDBF976DECC5E76BDF35D3DF0EBC7F74E00
            78ACCA15553220D43345555EAA72AFED54C7572F9BCDE600BCA5E6544B4B4363
            0A3CEF44DC010A112895F66077670BFE6C6D6CA85A370A9E694EB50E3736A5E3
            8E4DA8223BDB45943C8B82375B4F9ECEC8C8750B1CC9EB3F577FA3E092CAD571
            C723D400B5B60211EC3022D87144B0E38860C711C18E5333C12377EF403AFDEF
            BE7AAD508089C917719F6F20B76FDD8495956FB0F876C97A9FFBE3A3F0FACDE2
            7F8F899BAA0BEEBDDC0DFD7DBDF0F94B1EA6A6E70E744EB6BD1D1E3C7C12F739
            1B11C196E049CFCCCEC1C74FF943DFDDB83E54FEE4E293C27104D70355158C02
            3399B4552A3675D6B9B339B83A38004F9F3D2FD769846046C0148F1980D7B11D
            3CE6DAF050E5377CDFAF1CCFDBB974F14265CAD0F7C17688A0A9C4B40F4E43F3
            0BAFCA1733B583B11134BA297B217A660B8B3F51828F7285DB0AE61DA5D7A973
            78C6A02982A602EA5C9E4A4D17E27146B02E586F07E34561EF973F1CD846F1DA
            C4EFBC60BDC3781D3B98771EC1A7026CA758F40F8D1ADE4EB504EBEDE036845F
            483C369BF8A392F814CDAF64BDAEAFD43994466DDA098A272C66538AE6C79BA4
            EB82C3E28F4A6C8B2C9B916523983AD8568AA99DA0FDC27ECB46B0BE4D171C16
            7F546A769BA4DF3EE8B749A6B9064F7875AD50911E263868BEE2A93CC9826DE2
            4F9C601E24C79472E804097D8519269864F15528C233489205DBC41F15F9ABD2
            7144B0E38860C711C18E23821D47043B0E0996C7661D843F362B0FBE3B087FF0
            FD0C80F74E5E5D7103EDD5951EEFEFCB6739F5DD2355065591A15CDFF8AA2CA8
            32A6D656F97D8B3623CBB7C883090000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object wmibTimelineMoveLeft: TWMImageSpeedButton
          Left = 1128
          Top = 8
          Width = 30
          Height = 30
          Action = frmSEC.actEditTimelineMoveLeft
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000001494944415478DA63D4D5B76478F7E6912A03
            03430710BB00311F036DC02720DE03C415422272B719A5A5A5D51918184F7071
            0B08B0737033303232D1C4D6FFFFFF31FCFCF195E1DBD70F1F803C0B90C56BB9
            B8058338387969E45154F0E3FB2790E5EB40167F141496E1A3954FB1F9FCFDDB
            275F4016FF0786395D2C8501609A6218B5786859CCCACACA60616EC2A0ADADC9
            F0E0E123861D3BF6D0DE626969290647075B065E5E1E30FFF7EFDF0C73E62EA2
            9DC5205F5A5A98827D0937F0DD7B867DFB0F31BC7EFD863616A3FBF2FFFFFF0C
            E72F5C623873E63CC3DFBF7F09EA27D9624A7C49B6C51212E20C2ECE0E64FB92
            6C8B13E2A319383939C8F6E5D0B378C0821A040624712103BA67276AF97E6816
            99E8BEA77B25410E18B598EE160F58F376C01AF41AC02ECC713A77612C19A19D
            3660FF89A10D885D8198565EFF0CC4BB81B80A98A66E0200DDE325ADAF847F3C
            0000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA20000014B4944415478DA634C4CCA67D8BD7BBB2A03
            03430710BB00311F036DC02720DE03C415AEAE9EB7196564D5D49998984EA8AA
            6A08484A4A31B0B2B2D2C4D6DFBF7F333C7FFE8CE1F6ED1B1FFEFDFB6701B278
            ADBABA66909C9C028D3C8A0A1E3CB807B4FCE63A90C51F1D1C5CF868E5536C3E
            3F7060CF1790C5FF81614E174B610098A618462D1E5A1633313132F0F1F13370
            737331FCF8F183E1EDDB77B4B7989D9D9D4150508081858505CC07E6518667CF
            9ED3CE62982F7978B8E162A0ACF2FEFD07865FBF7ED1C662745FFEFFFF9FE1CB
            972F0C9F3E7D06B30901922DA6C497645BCCCECE06F4A520D9BE24DB62494949
            06666626B27D39F42C1EB0A0068101495CA8BEA77376A296EF87669189CDF790
            4AE227B092784B1F8B4905A316D3DDE2016BDE0E58835E03D885394EE72E8C25
            23B4D3A60E946B03625720E6A591673F03F16E20AE02A6A99B0003382C6F04B2
            B1E90000000049454E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000002334944415478DAD597BF4B1C4114C7BF3377
            2731E021A6308DE9420411515218B08CAD42C8FF904020692CD442ACC4C2C63F
            C0522C4C6C2D4C639A042B1111822172E95219CD1DA2DEEEF8C69DB97937BBC7
            EE49BC25036FE7DDDCDD7EDEF7CD6F31B0A940E529D932D94BB232528A48FB41
            4221CA39559FC9667FBDC6B120F033FAF08DAC372B4C88E4760671BE8AB5FF21
            1BD7E04FE4BCCA02B54061DBBC00944750AC4DA9A6AFB634F8AC557A138182B5
            7BEA2DC00726F8550D5669500EB4BEF4826828A447E805D0F099EA18B8954A0D
            9206745B8B78FF2AF3D0E05045E60793088EA9647E4138200FC206A818D0FA81
            07E7750C1C4BA90159300F2029CD8131EE8759C1D287CA0858F4E1E67F0F0AC0
            BB419A1A4F80DDDFC0873DA7B6AE9AB3100337A9652A39B4281DDC8247FB80B9
            61E071771444AD0E4C6C03D7613338CC02B6AAAC95A483164D200F49E51B5A7A
            A6075CBA4FAAC0E23E70701A29AE872CDD6960C9C01CD625A30074FBF347C0CC
            10D06F54EA976E9C006B3F800B527C15468AB5DA20CC00B683C6A6D6AA2D1970
            B904BC2595534C65A506AC1C02476791420BBDB576C052787D6AA063D4970B23
            AE2FF58B3E5680F59FC06510A9D3A6C15781F3DB076BA52CCDDBB46FF5763995
            AB47C0F1B99BBB75A3F03F06773AD5B90CAEDCA653AE0B48C797CC7FB5497CA1
            4DE27D3B9B44EEDB622E07011F9E34DAEFE5E8D32A80FB3AECE576BCCDED404F
            13015FD1D92BCC0B612E6DFAFEB4443649D6839472C74BDB5FAA76C8E6E9D2F6
            FD0660F06174A70E4A750000000049454E44AE426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000001664944415478DA63D4D5B764F850734C9581
            81A103885D80988F8136E01310EF01E20A8116ABDB8CB2ABFFAB33FCFD7D8273
            CF1401B62B3B1918BF7FA289ADFF39F9187EE9B8337C77C9F9C0C0CC6A01B278
            2DE7AE0941ECA756D1C8A3A8E0A7550CC377A7AC75208B3FF2F77AF0D1CAA7D8
            7CFEB178C71790C5FF81614E174B610098A618462D1E5A16B3B2B23258989B30
            686B6B323C78F88861C78E3DB4B7585A5A8AC1D1C19681979707CCFFFDFB37C3
            9CB98B686731C8979616A6605FC2C0BB77EF19F6ED3FC4F0FAF51BDA588CEECB
            FFFFFF339CBF7089E1CC99F30C7FFFFE25A89F648B29F125D9164B488833B838
            3B90ED4BB22D4E888F66E0E4E420DB9743CFE2010B6A101890C4850CE89E9DA8
            E5FBA15964A2FB9EEE95043960D462BA5B3C60CD5BFA36E82D810D7A6748835E
            03D885390EEEC25CDEC1C0F8E3336D7CCAC1CBF04BD703D685B1648476DAD481
            726D40EC0AC4BC34F22CC847BB81B80A98A66E020001935095E529E431000000
            0049454E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000001494944415478DA63D4D5B76478F7E6912A03
            03430710BB00311F036DC02720DE03C415422272B719A5A5A5D51918184F7071
            0B08B0737033303232D1C4D6FFFFFF31FCFCF195E1DBD70F1F803C0B90C56BB9
            B8058338387969E45154F0E3FB2790E5EB40167F141496E1A3954FB1F9FCFDDB
            275F4016FF0786395D2C8501609A6218B5786859CCCACACA60616EC2A0ADADC9
            F0E0E123861D3BF6D0DE626969290647075B065E5E1E30FFF7EFDF0C73E62EA2
            9DC5205F5A5A98827D0937F0DD7B867DFB0F31BC7EFD863616A3FBF2FFFFFF0C
            E72F5C623873E63CC3DFBF7F09EA27D9624A7C49B6C51212E20C2ECE0E64FB92
            6C8B13E2A319383939C8F6E5D0B378C0821A040624712103BA67276AF97E6816
            99E8BEA77B25410E18B598EE160F58F376C01AF41AC02ECC713A77612C19A19D
            3660FF89A10D885D8198565EFF0CC4BB81B80A98A66E0200DDE325ADAF847F3C
            0000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object wmibTimelineMoveRight: TWMImageSpeedButton
          Left = 1161
          Top = 8
          Width = 30
          Height = 30
          Action = frmSEC.actEditTimelineMoveRight
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA20000013C4944415478DA63D4D5B76478F7E6913A03
            03433B10BB00312F036DC02720DE03C4E54222727718A5A5A5351918188F7171
            0B08B0737033303232D1C4D6FFFFFF31FCFCF195E1DBD70FEF813C5390C51B81
            96FA7170F2D1C8A3A8E0C7F7CF40CBDFAF0159FC455058869B563EC5E6F3F76F
            9F7C0459FC1F18E674B1140680698A61D4E2C167B187870B8382BC1CC3D5ABD7
            194E9C3CC3F0FBF76FFA589C921CC7C0CACA0A667FFEFC8561FF81C30C4F9F3E
            A3BDC5A2A2220C4E8E760C424282703190EF8F9F384DB2EF498E636666660613
            13430643033D6009C748B6EFC94E5C94FA9EA2548DCBF77BF61E6078F1E225ED
            2CC6E5FBEFDF7F302C58B874185A3C20414DF7C43520D969C00A90012B3207AC
            92A02618B598EE160F58F376C01AF403D18531638476DA809633B432D0B6D3F6
            198877037105304DDD06008F591F9B6C07B2C10000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000001464944415478DA634C4CCA67D8BD7BBB3A03
            03433B10BB00312F036DC02720DE03C4E5AEAE9E77186564D5349998988EA9AA
            6A08484A4A31B0B2B2D2C4D6DFBF7F333C7FFE94E1F6ED9BEFFFFDFB670AB278
            A39A9A869FBCBC228D3C8A0A1E3E7CC070EBD6F535208BBF383ABA72B3B0B0D0
            C56290CF0F1CD8F31164F17F6098D3C5521800A62986518B079FC5C2C2C20C1C
            1CEC0C5FBF7E63F8F4E923C3BF7FFFE963B19494240330BF83D97FFEFC6178FF
            FE03C3CF9F3F696F311B1B1B83A0A0004A01F3E5CB57B27C4F721C33323232F0
            F1F132F0F0F080D9E4FA9EECC445A9EF294AD5B87DFF1EE8FB5FB4B31897EFFF
            FEFD07AC0C9E0F438B0724A8E99EB806243B0D5801326045A6B0B010B092E0A0
            7F25414D306A31DD2D1EB0E6ED8035E8C15D18A0E5E02E0C0B0B5DBA30668CD0
            4E9B2650AE9581B69DB6CF40BC1B882B8069EA3600B4C7265DDFD46E24000000
            0049454E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000002204944415478DAD597CF2B454114C7BF73DF
            0FC9EF95954424926C58B2B1791B2B1B59FA036C144B29D9D8508ABF81145B96
            B212498A58B161E1B790E7DE71A63BE39E3BDDE791BCC9D469CECC9B773FE77B
            E6DEF9211A5624A8B491CD920D905509FCBCC8E2431EC8B6C826CE87702608DC
            4E8D1DB25AF52B870AD6480A86C3A44CEE4F28B7643D0ABC4ECE207FB8010A0E
            647DB107CBB06DFA4C0045E0AB0AFC444EC5278843AD0084255B4A06B4DB5F07
            70AFC0D2562B78AD7D8FF5C182050CCA03F80A9E08E6104F68D383ECB9FE846B
            58600552487D0C6CAB55C0140BC4633E4F358729DF9796F2047841B002A43458
            88B82F0A28F679AD034A9A82443057971691EA855EA0AF1E583B07168F81573F
            0273E8BBF6FD200E0FBE03F62C70863AB773F4EAA7C33F5DBED06A7308ECDFB0
            546BE87B1087FBF6CBA7EB826093DAB417D65D75C05437D05419A579FD02583E
            019EFD1060A0AACE6BA8CFB2F17DB017AA568AB364E5A478B405186E0AC7A872
            45EAE78E80DDEB10A0806F413C08030E7E02CE68B081AB0C74D400E39D406345
            A47E83D42F91FA877C08CE33D526F5FCB32B0ACE6AD5D954E42B2BA3F6483330
            D418A957733F7D00ECDD30D56CCE83DF80337ADECD37DD5A0D8C7544EAEFDE80
            DC5604CEDBE9FE77E092A5DAD9CBE5E67372B18038593293368979DA24FAFF72
            9370B62D3A3908C04A77C98E3E004A7AD87376BC7576A0777185E915FAD2A6E0
            33F8DB4BDB23D926D9245DDA4E3F00BE697762CB144B160000000049454E44AE
            426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000001614944415478DA63D4D5B764F850734C9D81
            81A11D885D8098978136E01310EF01E2728116AB3B8CB2ABFF6B32FCFD7D8C73
            CF1401B62B3B1918BF7FA289ADFF39F9187EE97A327C77CE7ACFC0CC6A0AB278
            23D0523FF613CB68E45154F0D33C82E1BB6BDE1A90C55FF8BB5DB9197F7EA58B
            C5FF39F9193E166FFF08B2F83F30CCE962290C00D314C3A8C583CF620F0F1706
            05793986AB57AF339C387986E1F7EFDFF4B13825398E81959515CCFEFCF90BC3
            FE0387199E3E7D467B8B454545189C1CED18848404E16220DF1F3F719A64DF93
            1CC7CCCCCC0C2626860C86067A0C8C8C8C64FB9EECC445A9EF294AD5B87CBF67
            EF0186172F5ED2CE625CBEFFFEFD07C382854B87A1C50312D4744F5C03929D06
            AC0019B02273C02A096A82518BE96EF180356F810DFAC9C006FD72BA588CDCA0
            4774612E6F6760FCF185363EE5E065F8A507EAC26483BA30668CD04E9B2650AE
            9581B69DB6CF40BC1B882B8069EA3600B2E24A833AE1E2870000000049454E44
            AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA20000013C4944415478DA63D4D5B76478F7E6913A03
            03433B10BB00312F036DC02720DE03C4E54222727718A5A5A5351918188F7171
            0B08B0737033303232D1C4D6FFFFFF31FCFCF195E1DBD70FEF813C5390C51B81
            96FA7170F2D1C8A3A8E0C7F7CF40CBDFAF0159FC455058869B563EC5E6F3F76F
            9F7C0459FC1F18E674B1140680698A61D4E2C167B187870B8382BC1CC3D5ABD7
            194E9C3CC3F0FBF76FFA589C921CC7C0CACA0A667FFEFC8561FF81C30C4F9F3E
            A3BDC5A2A2220C4E8E760C424282703190EF8F9F384DB2EF498E636666660613
            13430643033D6009C748B6EFC94E5C94FA9EA2548DCBF77BF61E6078F1E225ED
            2CC6E5FBEFDF7F302C58B874185A3C20414DF7C43520D969C00A90012B3207AC
            92A02618B598EE160F58F376C01AF403D18531638476DA809633B432D0B6D3F6
            198877037105304DDD06008F591F9B6C07B2C10000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object wmibTimelineZoomIn: TWMImageSpeedButton
          Left = 1212
          Top = 8
          Width = 30
          Height = 30
          Action = frmSEC.actEditTimelineZoomIn
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000000F44944415478DA63D4D5B76478F7E6913A03
            03433B10BB00312F036DC02720DE0DC4154222727718A5A5A5351918188F7171
            0B08B0737033303232D1C4D6FFFFFF31FCFCF195E1DBD70FEF813C5390C51B81
            96FA7170F2D1C8A3A8E0C7F7CF40CBDFAF0559FC455058869B563EC5E6F3F76F
            9F7C0459FC1F18E674B1140680698A61D4E2518B472D1EB598B6166B6B6930C8
            CAC91065E8E3474F18AE5EBB411D8B33D29380B516235116FFFFFF9F61C6CC79
            D4B1D8C7DB9D415696481F3F7EC2B065EB4EEA584C6D306AF1A8C5A3160F1F8B
            07AC793B600DFA81E8C29831423B6D40CB195A1968DB69FB0CC4BB80B81298A6
            6E03002325B98E070978D10000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000000FD4944415478DA634C4CCA67D8BD7BBB3A03
            03433B10BB00312F036DC02720DE0DC415AEAE9E77186564D5349998988EA9AA
            6A08484A4A31B0B2B2D2C4D6DFBF7F333C7FFE94E1F6ED9BEFFFFDFB670AB278
            A39A9A869FBCBC228D3C8A0A1E3E7CC070EBD6F5B5208BBF383ABA72B3B0B0D0
            C56290CF0F1CD8F31164F17F6098D3C5521800A62986518B472D1EB578D462DA
            5ACCCDCDCDC0C1C14194A13F7EFC60F8FAF52B752C9696966260646424CAE2FF
            FFFF333C7DFA8C3A168B880893E4E3376FDE52C7626A83518B472D1EB578F858
            3C60CDDB016BD083BB3040CBC15D181616BA7461CC18A19D364DA05C2B036D3B
            6D9F817817105702D3D46D00A74FC0503D30AE6A0000000049454E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA20000020D4944415478DAC597BF4A043110C627D14E
            041BB112ECC45650502C6DB5F5197C037D00F1117C06110BB1D356B4F305B4B0
            150BF14FE9254EBCCCDD3793ECDD9E70778139E67697F9CD979964B36EF92212
            8F55B653B65DB6791ACFF864BB653B627B760C5E63E79E6D614C403BDED93612
            F88A9DFD0941655C26F0373B7313067F24709C30F46FB402BB1183B65152802D
            C4B9E67B8380568E4D468125B0C09C5C73706F0834E69F88D76209EF81116A7D
            977D5F99050C1CB22F50045AB8022BA519EA33D0BB52BD529B8121761308B10F
            44E50A5C28AC41C1C7199240024CC19BE0A8BA007B00CFB80CCF36E3CA3EC060
            1D8166EB0030348151AD2843602F09D307185860368160540F04236CD6EB046A
            3546E04F30490C038B1A84243B5821DA5982E5D530D5C9EE5E89CE5FFA504926
            9A04AA60A596ED71AFFBBFCD48A0F56B561DB5EA7F81CFB688B617DB81EFDF88
            0E1F46043B00F7EAECBB700FC9E09A9740016B9CFCD0AFAF805B35178254573B
            5D6BA96D88655777DA74359156E341354E3F6E34B89CA4BEA24ED486A867A571
            03293A1CE09298378D162030426345AD0213E9AD52F66484CBB689D38CEBB8B7
            911828BEA98ABD9A88AA6F240285CE245780A3AE77AC40E5D92AD8D61013C06E
            B603C104CA105A0513822001351BB5072B41ED01C03C56073725F19F31E8EC35
            D5E3EDD40EF4D3F884D974F9A32DC14F68BC1F6D5F6C376CC76C4FBF39C299E9
            61E3E6250000000049454E44AE426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000001194944415478DA63D4D5B764F850734C9D81
            81A11D885D8098978136E01310EF06E20A8116AB3B8CB2ABFF6B32FCFD7D8C73
            CF1401B62B3B1918BF7FA289ADFF39F9187EE97A327C77CE7ACFC0CC6A0AB278
            23D0523FF613CB68E45154F0D33C82E1BB6BEE5A90C55FF8BB5DB9197F7EA58B
            C5FF39F9193E166FFF08B2F83F30CCE962290C00D314C3A8C5A3168F5A3C6A31
            6D2DD6D6D26090959321CAD0C78F9E305CBD76833A1667A427313032321265F1
            FFFFFF1966CC9C471D8B7DBCDD19646589F4F1E3270C5BB6EEA48EC5D406A316
            8F5A3C6AF1F0B178C09AB7C006FD6460837E395D2C466ED023BA3097B73330FE
            F8421B9F72F032FCD2037561B2415D18334668A74D1328D7CA40DB4EDB6720DE
            05C495C034751B0038F5E474A0B8181C0000000049454E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000000F44944415478DA63D4D5B76478F7E6913A03
            03433B10BB00312F036DC02720DE0DC4154222727718A5A5A5351918188F7171
            0B08B0737033303232D1C4D6FFFFFF31FCFCF195E1DBD70FEF813C5390C51B81
            96FA7170F2D1C8A3A8E0C7F7CF40CBDFAF0559FC455058869B563EC5E6F3F76F
            9F7C0459FC1F18E674B1140680698A61D4E2518B472D1EB598B6166B6B6930C8
            CAC91065E8E3474F18AE5EBB411D8B33D29380B516235116FFFFFF9F61C6CC79
            D4B1D8C7DB9D415696481F3F7EC2B065EB4EEA584C6D306AF1A8C5A3160F1F8B
            07AC793B600DFA81E8C29831423B6D40CB195A1968DB69FB0CC4BB80B81298A6
            6E03002325B98E070978D10000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object wmibTimelineZoomOut: TWMImageSpeedButton
          Left = 1358
          Top = 8
          Width = 30
          Height = 30
          Action = frmSEC.actEditTimelineZoomOut
          Anchors = [akTop, akRight]
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA20000011C4944415478DA63D4D5B76478F7E6913A03
            03433B10BB00312F036DC02720DE0DC4154222727718A5A5A5351918188F7171
            0B08B0737033303232D1C4D6FFFFFF31FCFCF195E1DBD70FEF813C5390C51B80
            96FA7370F2D1C8A3A8E0C7F7CF40CBDFAF0559FC455058869B563EC5E6F3F76F
            9F7C0459FC1F18E674B1140680698A61D4E2C16D311B1B2B4368682098BD7AF5
            7A865FBF7ED3C76211116186D0900088C56B3630BC79F376D4E2216C312821F1
            F1A116A70202FC0CAE2E8E60F6EE3DFB193E7CF88822FFE9D327BC098E288BA3
            A3C318F87849ABB03E7DFECCB074E9AA216AF18005353630BC53F5A8C5203060
            9504A560D462BA5B3C60CDDB016BD00F4417C68C11DA69035ACED0CA40DB4EDB
            6720DE05C495C034751B000362079D6D6A06620000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000001324944415478DA634C4CCA67D8BD7BBB3A03
            03433B10BB00312F036DC02720DE0DC415AEAE9E77186564D5349998988EA9AA
            6A08484A4A31B0B2B2D2C4D6DFBF7F333C7FFE94E1F6ED9BEFFFFDFB670AB278
            03D0527F0505451A7914153C7CF880E1D6ADEB6B41167F717474E5666161A18B
            C5209F1F38B0E723C8E2FFC030A78BA530004C530CA3160F6E8B81D98F414C4C
            14CC7EF5EA3503307BD0C762505E17171703B35FBE7C054EA9A3160F5D8B4109
            89999919CD62160621212130FBDDBB77408BFFA0C8FFFDFB176F8223CA620909
            7106528BD33F7FFE30BC78F172885A3C60418D0D0CEF543D6A31080C58254129
            18B598EE160F58F376839A9A86BFBC3C7D1AF48F1E3D60B87913D2A007776180
            9683BB302C2C74E9C29831423B6D9A40B95606DA76DA3E03F12E20AE04A6A9DB
            0025940E5FE70001AC0000000049454E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA20000022E4944415478DACDD7BF4F14411407F0EFEC
            7192F043A5265C68008D35440D95B19686D6C68282C49204FF00424149626141
            4329268452132BA386DE04C5C2905093E35782703BBE7733EFF6DDECCD890537
            6EF298FDC1EDE7BD99B9BD5963AD456D0B530056299E520CA3C36670BDCDC62F
            1D537CA0583E98C74F33F6D6DEA783CF1477AF83994806B6831849E288629AE1
            6DDA99EB866A2C56B946741211FC1DC3A7B43318831935FA9C69BF6E9560D539
            C123709DE1D2B51015502720BD60035012F81B5E82359A2990233305AA2B163C
            5728FFC9BBE02D381CD34C5ABF1FA299FF400B53782EE7035C27D084BBA19500
            AFF87FD6306F0DDB8ECA710C6F838DAAAAA2506EFB24018AE12AB039EB3EF3FC
            137072E9518A2B85367C585B1EF316AC2792F150252BD0669B39F8DE1D60E3B1
            BBC90B7A02ECD53D9A3B48F046EEF66D30E19A704DC65855CBDD5AF568D8724C
            11FCE6918317BE00DFEB45759779B9CD51AEBA0DCE5437F7F90AAB3EE498DBC9
            DBC0EB87EE068B5F811FC7AEB22B5F21631C72DCE830D64DD8043356905B99C2
            4D91C00455BC3EE36EF07217D8AF2BD016E86F85DBA0EA12CC554B650C8DF403
            E3433E193FEE357ACE2D3D70F0DA37E0E0AC184F8E5FF42C3CBAF009F89EC8BB
            C1326B05E5D879028C0EE09FB6C373E0D947D5E5B698F5FF179CACAB7B3EB992
            7D9D923C40923D3293FF4840E13DFB5904122D04F497BFE74B9F187E938BBD64
            CBDB640BFA14AF3033C6BFB431BE829B7D693BA1784FF18A5EDAF6FF00E4FE47
            6E3C2C056A0000000049454E44AE426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA2000001494944415478DA63D4D5B764F850734C9D81
            81A11D885D8098978136E01310EF06E20A8116AB3B8CB2ABFF6B32FCFD7D8C73
            CF1401B62B3B1918BF7FA289ADFF39F9187EE97A327C77CE7ACFC0CC6A0AB278
            03E7DEA9FEECC797D2C8A3A8E0A77904C377D7BCB5208BBFF077BB7233FEFC4A
            178BFF73F2337C2CDEFE1164F17F6098D3C5521800A62986518B07B7C56C6CAC
            0CA1A18160F6EAD5EB197EFDFA4D1F8B454484194243022016AFD9C0F0E6CDDB
            518B87B0C5A084C4C7C787222620C0CFE0EAE20866EFDEB39FE1C3878F28F29F
            3E7DC29BE088B2383A3A8C818F97B40AEBD3E7CF0C4B97AE1AA2160F58506303
            C33B558F5A0C02035649500A462DA6BBC503D6BCDDC0B967B23FFB89E574B1F8
            A75938C377B7FCB5A85D98CBDB19187F7CA18D4F3978197EE981BA30D9A02E8C
            1923B4D3A609946B65A06DA7ED3310EF02E24A609ABA0D0026EB3285E9C58943
            0000000049454E44AE426082}
          EnableTextColor = 10262422
          DisableTextColor = 9076349
          DownTextColor = clWhite
          OverTextColor = 10262422
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000001E0000
            001E08060000003B30AEA20000011C4944415478DA63D4D5B76478F7E6913A03
            03433B10BB00312F036DC02720DE0DC4154222727718A5A5A5351918188F7171
            0B08B0737033303232D1C4D6FFFFFF31FCFCF195E1DBD70FEF813C5390C51B80
            96FA7370F2D1C8A3A8E0C7F7CF40CBDFAF0559FC455058869B563EC5E6F3F76F
            9F7C0459FC1F18E674B1140680698A61D4E2C16D311B1B2B4368682098BD7AF5
            7A865FBF7ED3C76211116186D0900088C56B3630BC79F376D4E2216C312821F1
            F1A116A70202FC0CAE2E8E60F6EE3DFB193E7CF88822FFE9D327BC098E288BA3
            A3C318F87849ABB03E7DFECCB074E9AA216AF18005353630BC53F5A8C5203060
            9504A560D462BA5B3C60CDDB016BD00F4417C68C11DA69035ACED0CA40DB4EDB
            6720DE05C495C034751B000362079D6D6A06620000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object Shape1: TShape
          Left = 1
          Top = 1
          Width = 1386
          Height = 1
          Align = alTop
          Brush.Style = bsClear
          Pen.Color = 8608301
        end
        object Image1: TImage
          Left = 1201
          Top = 11
          Width = 1
          Height = 25
          Anchors = [akTop, akRight]
          AutoSize = True
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D49484452000000010000
            00190806000000F090447F000000104944415478DA63F4F5F5FDCF3840040083
            B72F90D9E44A080000000049454E44AE426082}
        end
        object wmtbTimelineZoom: TWMTrackBar
          Left = 1250
          Top = 16
          Width = 100
          Height = 12
          Anchors = [akTop, akRight]
          TransParent = True
          AutoSize = True
          Max = 100
          Position = 100
          OnChange = wmtbTimelineZoomChange
          BarBorderEnableColor = 2497307
          BarBorderDisableColor = clGray
          BarBorderWidth = 0
          BarEnable.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D49484452000000640000
            000C080600000094451863000000984944415478DA63D4D5B7641805E483F7EF
            DEB202291E201680D2DC94F019475A84503B0081988D919191011B6680B14190
            094A238BC331488889E1C3FB770C833E4206730082E5A8EB57EA47C8480A406A
            037084C8C8AA8D06E02001B008F93F1A8083038C46C82003B0086163A06E91C5
            3A1AC11444080D2AF5D1081E4C114203478E98081E1211426D3098237848740C
            073BA0760403007EDF466AF0D108FA0000000049454E44AE426082}
          Rail.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000000C0000
            000C080600000056755CE7000000984944415478DA63DCB163070314B0017106
            104702B10E54EC0A102F07E21940FC0B24C008D5200DC45B80D880013BB800C4
            3E40FC14A40164F2493C8A61E022109B8334E402199318880379200DC7810C0B
            22359C0069F809F53031E027A91ABE90E5A43C206322911AF2411AD81920C1AA
            4F40F12520362336E24071E0CD008D389820C8E399401C03C49A50B1EB40BC04
            88A733409306002EEB364FB6B0CEF80000000049454E44AE426082}
          RailEndCenter = False
          RailEndVisible = False
          RailEndEnableCeolor = clWhite
          RailEndDisableCeolor = clSilver
          Direction = bdLeftToRight
          MouseClickEnabled = True
        end
        object wmibEventTimelineClose: TWMImageSpeedButton
          Left = 106
          Top = 20
          Width = 12
          Height = 12
          Action = frmSEC.actEditTimelineZoomIn
          AutoSize = True
          EnableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000000C0000
            000C080600000056755CE70000006F4944415478DA63640082FFFFFFB301296E
            4646C6F70C5800505E10487D05CAFF62842AAE02E2DF40BC0C28781F4DB12290
            8A026256206E6384EACE820AA06842530C929BC688436219D482287483187158
            FD1B2A8C612B231EF73260F317F91A487212499E263958498E38529306001EE1
            662F6D0045AE0000000049454E44AE426082}
          DisableBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000000C0000
            000C080600000056755CE70000006F4944415478DA63640082FFFFFF73002901
            4646C6170C5800505E02487D00CAFF60842A5E0DC4BF80B8112878094DB11E90
            AA076236200E6584EA9E0A1540D184A6182497CD8843A2116A413DBA418C38AC
            FE0515C6B095118F7B19B0F98B7C0D243989244F931CAC24471CA9490300608E
            66415633FB330000000049454E44AE426082}
          DownBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000000C0000
            000C080600000056755CE70000006F4944415478DA63640082FFFFFFF3002931
            4646C67B0C5800505E0948BD02CA7F61842AFE0CC4DF81D80728B80F4DB11390
            DA02C49C40CCCB08D57D052A80A2094D31484E871187840FD4822DE80631E2B0
            FA3B5418C356463CEE65C0E62FF23590E424923C4D72B0921C71A4260D00C334
            665D58AFCC3D0000000049454E44AE426082}
          OverBitmap.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000000C0000
            000C080600000056755CE70000006F4944415478DA63640082FFFFFF73002901
            4646C6170C5800505E02487D00CAFF60842A5E0DC4BF80B8112878094DB11E90
            AA076236200E6584EA9E0A1540D184A6182497CD8843A2116A413DBA418C38AC
            FE0515C6B095118F7B19B0F98B7C0D243989244F931CAC24471CA9490300608E
            66415633FB330000000049454E44AE426082}
          EnableTextColor = 16760237
          DisableTextColor = 6376003
          DownTextColor = clWhite
          OverTextColor = clWhite
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000000C0000
            000C080600000056755CE70000006F4944415478DA63640082FFFFFFB301296E
            4646C6F70C5800505E10487D05CAFF62842AAE02E2DF40BC0C28781F4DB12290
            8A026256206E6384EACE820AA06842530C929BC688436219D482287483187158
            FD1B2A8C612B231EF73260F317F91A487212499E263958498E38529306001EE1
            662F6D0045AE0000000049454E44AE426082}
          ParentFont = False
          Font.Charset = ANSI_CHARSET
          Font.Color = clSilver
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
        end
        object imgTimeline: TImage
          Left = 1
          Top = 12
          Width = 93
          Height = 24
          AutoSize = True
          Picture.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D494844520000005D0000
            0018080600000098E4F11C000004E74944415478DAED995B68154718C727862A
            05B168B5DE8AC6A42A1505438220F8209A63A322226814D40711120BE20DC389
            0A3E798B82F6C10B3D88A52A3E1845F1C10692EA8328AD26FA22B40F4D5A2A2A
            5E305E1FBC44FDFFCF7C1BBFB3EED9B3C793E3AE970FFECC996F6766777E333B
            FBCD9C82C6C64603FB02FACA587B003D3701ADB2B2B2EB37DA2A40D21B2A841E
            E3DA8BA0ED7C4A5600509390FE662C2CDA5FD09834E5393004FB107A49870BFA
            41248B25FB189A8FEBA7C3EE64D48CD05B9096499E30A7417F4AFE4B683EB400
            9A00F515FF53E8327402DA0FB01D74A2AD7148CE40FDA5DC5D5C1B107627A366
            84DE89B487791BF80C681F340CBA089D85FE3676867F074D84A61A3BA3D7417B
            01F89507F81EF487DDD12819A1B7190B48038F43DBA0F3D04AA8354DFD1268B3
            B16FC321680900772AF05F43859FA1A71AA1FF80F4AE7903B61AFA19DA0DAD82
            3AFD1AE09A8E3696E1E71E2881FC8FF4C33716C934E47786DDC9A85981442F8E
            8D84AE42C7A04550C619EA7C48D14E2D92EDD06CF84E85DDB1289B1B3A61339A
            1965EC1A4F635473C5D88FE626E54F9A82CEA8E61CC40FE7F7F0BFCC7473D4E1
            1B559DE6720C6A82EAD056FDFB8081E749BEE5B81FB954E4EBFE1AFA20E81AB4
            1EDAA1CA1441FFCAEF5BC67E347F35DE21E32C249CE515F0FF9E6587F956D5A0
            5E42F279EBB4CF33BC77E80BA1C3D0B7D07555A6C858E875D06472862E412BA0
            3F5CD0B929BA0D1D80BF36CB0EA7400FC334F47CDE4743E768721D1FEA2A5364
            2CF439D0496826B4CBD8F59F9BA17578C81BEAC1195A3EADD4A311ACC3EE995E
            8C8491550CBE66E41951114A3B5421D5B83FE04C4CE03AA3ADB8F8191454C1DF
            AEDAE71E83CBD93C75DB7A94A9F382EE71FF2629562C4AB9BFC7E0C555B95629
            D7EC86FE0B540A8DCF009DD613DA0AAD31F6D8E01B34F84C6EC81DE960E4B7E7
            01FA36E9287DADCAD7201D8C497304D4813231D55693D4AD91BA6532085C22CB
            0342E76037AB769D482FE60055CFD435A06A4270223468E8DC084D814667804E
            E0ABA10DC61E096C4443BBB2019C03F4B8035CD5BB273F4BD4CE98405BC4D72E
            1F6CCEF0728FD9DF2680EA0340EF2BF90ED5C65124ED04ACEAA4BC416A30A812
            0DBD5646A80FF4240D741E60113077A45C5AF8CADCCC157816D0ABF1BBC4558F
            655AE1AF523E775D0E4CC20D42CAB2CFFCF09707804EB8351EF5CB38FBD52CEF
            A707460D309FA34A43E7B6FE0234DBD808C40D9D1A61D4479417B35CBAF305BD
            C1B53677D535F61BD096E9FE02FA5DA0C765D062194260C7EA34742E15FFC843
            C654A121C646330C17D931CEF0AE183C22D0133AAC4B33D3EB33857EDD003DB9
            84E0773FBFFBB837474BA1FDD05CE8B8F2F35C8633FBA1ABFE16199474C653C6
            B51180CE1958EC7C005539BEF25CFBB94C367403F4943A1EE5E671197343E76C
            67617ECDA7CA03F9D9E900D02B4C00CB33F462E98B134A76C8E687EB2FC1978B
            2F27E892679BD54685926AD9496EB4DCD0693C7164AC3D1CE241D6113F581FC2
            F2A27C04A2E3F4848070A29E9CA12B1FDBD2717AC2E99B17741AD724EE4EA71B
            7BDCCBC23CAAFDCF558ED706FAB064AC5CDA2DA3F211593AE8C96BC6CE0A9EB5
            381BA63BC69EA13F92FC72F3E6CF0A2FBB0FE83F85DDC9A8991F746DCE3F453C
            71E4FAE4442F3CA7F1837E0FD0FF0FBB9351B3A0D0D319677C6F9FEBCF00BD57
            D89D8C9AE50A9DA78A7E2772AFF8F75DD89D8C9ABD06B1122B9BA8BB73660000
            000049454E44AE426082}
          Transparent = True
        end
        object wmtlPlaylist: TWMTimeLine
          Left = 0
          Top = 47
          Width = 1388
          Height = 202
          Anchors = [akLeft, akTop, akRight, akBottom]
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWhite
          Font.Height = -12
          Font.Name = 'Century Gothic'
          Font.Style = []
          ParentFont = False
          OnMouseWheel = wmtlPlaylistMouseWheel
          FrameNumber = 0
          TimeZoneProperty.Height = 46
          TimeZoneProperty.FrameStart = 0
          TimeZoneProperty.FrameCount = 180000
          TimeZoneProperty.FrameRate = 30.000000000000000000
          TimeZoneProperty.FrameStep = 15
          TimeZoneProperty.FrameGap = 7
          TimeZoneProperty.FrameSkip = 600
          TimeZoneProperty.FrameDayReset = False
          TimeZoneProperty.BarGap = 24
          TimeZoneProperty.BarLarge = 20
          TimeZoneProperty.BarSmall = 10
          TimeZoneProperty.BarColor = 7427652
          TimeZoneProperty.Color = 5060404
          TimeZoneProperty.ColorTo = 5060404
          TimeZoneProperty.MarkColor = clNavy
          TimeZoneProperty.MarkLineColor = 788230
          TimeZoneProperty.Font.Charset = ANSI_CHARSET
          TimeZoneProperty.Font.Color = 16765315
          TimeZoneProperty.Font.Height = -12
          TimeZoneProperty.Font.Name = 'Century Gothic'
          TimeZoneProperty.Font.Style = []
          TimeZoneProperty.RailImage.Data = {
            0954506E67496D61676589504E470D0A1A0A0000000D49484452000000090000
            000A0806000000660574BE0000004D4944415478DA63FCCFC00044F801E35056
            F40C484BE251F31CA4481FC83804C47C58147C02623B46A85D0E40BC0B885991
            14FC066237203EC088E4A050205E01C44C40FC0F8823807835480200C3641A10
            EB5741070000000049454E44AE426082}
          TimeZoneProperty.RailBarColor = clRed
          TimeZoneProperty.RailBarVisible = True
          TimeZoneProperty.Visible = True
          CompositionBarProperty.CaptionHeight = 2
          CompositionBarProperty.MinWidth = 120
          CompositionBarProperty.MaxWidth = 240
          CompositionBarProperty.MinHeight = 48
          CompositionBarProperty.MaxHeight = 480
          CompositionBarProperty.Width = 0
          CompositionBarProperty.Color = 2497307
          CompositionBarProperty.ColorSelected = clSilver
          CompositionBarProperty.Visible = False
          CompositionBarProperty.Font.Charset = DEFAULT_CHARSET
          CompositionBarProperty.Font.Color = clSilver
          CompositionBarProperty.Font.Height = -12
          CompositionBarProperty.Font.Name = 'Tahoma'
          CompositionBarProperty.Font.Style = []
          CompositionBarProperty.TimecodeBarFont.Charset = DEFAULT_CHARSET
          CompositionBarProperty.TimecodeBarFont.Color = clWindowText
          CompositionBarProperty.TimecodeBarFont.Height = -11
          CompositionBarProperty.TimecodeBarFont.Name = 'Tahoma'
          CompositionBarProperty.TimecodeBarFont.Style = []
          CompositionBarProperty.LineHorzColor = 788230
          CompositionBarProperty.LineVertColor = 788230
          ZoomBarProperty.Min = 0
          ZoomBarProperty.Max = 100
          ZoomBarProperty.Position = 0
          VideoGroupProperty.Count = 0
          VideoGroupProperty.HeightRate = 0.450000000000000000
          VideoGroupProperty.Visible = False
          AudioGroupProperty.Count = 0
          AudioGroupProperty.HeightRate = 0.450000000000000000
          AudioGroupProperty.Visible = False
          AudioGroupProperty.DefaultChannelStereo = True
          DataGroupProperty.Count = 4
          DataGroupProperty.HeightRate = 0.100000000000000000
          DataGroupProperty.Visible = True
          VideoTrackProperty.Font.Charset = DEFAULT_CHARSET
          VideoTrackProperty.Font.Color = clWindowText
          VideoTrackProperty.Font.Height = -11
          VideoTrackProperty.Font.Name = 'Tahoma'
          VideoTrackProperty.Font.Style = []
          VideoTrackProperty.Color = 13408882
          VideoTrackProperty.ColorCaption = 13408882
          VideoTrackProperty.ColorSelected = 13750737
          VideoTrackProperty.ColorSelectedCaption = 13750737
          AudioTrackProperty.Font.Charset = DEFAULT_CHARSET
          AudioTrackProperty.Font.Color = clWindowText
          AudioTrackProperty.Font.Height = -11
          AudioTrackProperty.Font.Name = 'Tahoma'
          AudioTrackProperty.Font.Style = []
          AudioTrackProperty.Color = 13408882
          AudioTrackProperty.ColorCaption = 13408882
          AudioTrackProperty.ColorSelected = 13750737
          AudioTrackProperty.ColorSelectedCaption = 13750737
          DataTrackProperty.Font.Charset = DEFAULT_CHARSET
          DataTrackProperty.Font.Color = clWhite
          DataTrackProperty.Font.Height = -20
          DataTrackProperty.Font.Name = 'Century Gothic'
          DataTrackProperty.Font.Style = []
          DataTrackProperty.Color = 13408882
          DataTrackProperty.ColorCaption = 13408882
          DataTrackProperty.ColorSelected = 13750737
          DataTrackProperty.ColorSelectedCaption = 13750737
          WorkAreaVisible = True
          ScrollType = ssMetro
          ScrollProperty.Color = 2497307
          ScrollProperty.ParentColor = False
          ScrollProperty.ThumbColor = 8608301
          ScrollProperty.Size = 4
          OnTrackCustomDrawEvent = wmtlPlaylistTrackCustomDrawEvent
          OnTrackHintEvent = wmtlPlaylistTrackHintEvent
        end
      end
    end
  end
end
