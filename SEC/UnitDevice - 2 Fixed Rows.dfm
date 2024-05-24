inherited frmDevice: TfrmDevice
  Caption = 'frmDevice'
  PixelsPerInch = 96
  TextHeight = 17
  inherited WMPanel: TWMPanel
    inherited pnlDesc: TWMPanel
      object acgDeviceList: TAdvColumnGrid
        Left = 1
        Top = 1
        Width = 313
        Height = 291
        Cursor = crDefault
        Align = alClient
        BevelWidth = 4
        BorderStyle = bsNone
        Color = 3812392
        ColCount = 2
        Ctl3D = False
        DrawingStyle = gdsClassic
        FixedColor = 6901572
        FixedCols = 0
        RowCount = 300
        Font.Charset = ANSI_CHARSET
        Font.Color = clSilver
        Font.Height = -12
        Font.Name = 'Century Gothic'
        Font.Style = []
        GradientEndColor = 1644825
        GradientStartColor = 1644825
        Options = [goFixedVertLine, goFixedHorzLine, goHorzLine, goColSizing, goRowSelect, goThumbTracking, goFixedHotTrack]
        ParentCtl3D = False
        ParentFont = False
        ScrollBars = ssBoth
        TabOrder = 0
        StyleElements = [seBorder]
        OnDrawCell = acgDeviceListDrawCell
        GridLineColor = 919047
        GridFixedLineColor = 1644825
        HoverRowCells = [hcNormal, hcSelected]
        OnGetCellColor = acgDeviceListGetCellColor
        HighlightColor = 8869383
        ActiveCellFont.Charset = DEFAULT_CHARSET
        ActiveCellFont.Color = clWindowText
        ActiveCellFont.Height = -11
        ActiveCellFont.Name = 'Tahoma'
        ActiveCellFont.Style = [fsBold]
        CellNode.NodeType = cnFlat
        CellNode.TreeColor = 919047
        ColumnHeaders.Strings = (
          'NO'
          'Name')
        ControlLook.FixedGradientHoverFrom = 13619409
        ControlLook.FixedGradientHoverTo = 12502728
        ControlLook.FixedGradientHoverMirrorFrom = 12502728
        ControlLook.FixedGradientHoverMirrorTo = 11254975
        ControlLook.FixedGradientDownFrom = 8816520
        ControlLook.FixedGradientDownTo = 7568510
        ControlLook.FixedGradientDownMirrorFrom = 7568510
        ControlLook.FixedGradientDownMirrorTo = 6452086
        ControlLook.FixedGradientDownBorder = 14007466
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
        EditWithTags = True
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
        FixedRowHeight = 22
        FixedFont.Charset = DEFAULT_CHARSET
        FixedFont.Color = 16760237
        FixedFont.Height = -12
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
        ScrollColor = clBlack
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
        SelectionColor = 3812392
        SelectionTextColor = clHighlightText
        ShowDesignHelper = False
        SortSettings.DefaultFormat = ssAutomatic
        SortSettings.HeaderColorTo = 16579058
        SortSettings.HeaderMirrorColor = 16380385
        SortSettings.HeaderMirrorColorTo = 16182488
        VAlignment = vtaCenter
        Version = '3.1.4.4'
        WordWrap = False
        Columns = <
          item
            AutoMinSize = 0
            AutoMaxSize = 0
            Alignment = taRightJustify
            Borders = []
            BorderPen.Color = clSilver
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 3812392
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clSilver
            Font.Height = -12
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'NO'
            HeaderAlignment = taCenter
            HeaderFont.Charset = DEFAULT_CHARSET
            HeaderFont.Color = 16760237
            HeaderFont.Height = -12
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
            CheckFalse = 'N'
            CheckTrue = 'Y'
            Color = 3812392
            ColumnPopupType = cpFixedCellsRClick
            DropDownCount = 8
            EditLength = 0
            Editor = edNormal
            FilterCaseSensitive = False
            Fixed = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clSilver
            Font.Height = -12
            Font.Name = 'Century Gothic'
            Font.Style = []
            Header = 'Name'
            HeaderAlignment = taCenter
            HeaderFont.Charset = DEFAULT_CHARSET
            HeaderFont.Color = 16760237
            HeaderFont.Height = -12
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
          end>
        ColWidths = (
          36
          68)
        RowHeights = (
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22
          22)
      end
    end
  end
end
