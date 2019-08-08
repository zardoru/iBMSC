Imports iBMSC.Editor

Public Class OpVisual
    Dim ReadOnly _
        lLeft() As Integer =
            {78, 110, 142, 174, 208, 240, 272, 304, 336, 368, 400, 432, 464, 498, 530, 562, 594, 626, 658, 690, 722, 754,
             788, 820, 852, 884, 918, 950}

    Structure ColumnOptionSet
        Public Width As NumericUpDown
        Public Title As TextBox
        Public SNote As Button
        Public SText As Button
        Public LNote As Button
        Public LText As Button
        Public BG As Button
    End Structure

    'Const Max As Integer = 99

    'Dim lWidth() As Integer
    'Dim lTitle() As String
    'Dim lColor() As Color
    'Dim lForeColor() As Color
    'Dim lColorL() As Color
    'Dim lForeColorL() As Color
    'Dim lBg() As Color

    'Dim WithEvents jWidth As New ArrayList
    'Dim WithEvents jTitle As New ArrayList
    'Dim WithEvents jColor As New ArrayList
    'Dim WithEvents jForeColor As New ArrayList
    'Dim WithEvents jColorL As New ArrayList
    'Dim WithEvents jForeColorL As New ArrayList
    'Dim WithEvents jBg As New ArrayList

    'Dim eColor(13) As Color
    'Dim eFont(3) As Font
    'Dim eI(4) As Integer

    'Dim WithEvents eaColor As New ArrayList
    'Dim WithEvents eaFont As New ArrayList
    'Dim WithEvents eaI As New ArrayList

    Dim ReadOnly VisualOptions As VisualSettings
    Dim ReadOnly TemporalColumns() As Column
    Dim ReadOnly ColumnOptions() As ColumnOptionSet


    Public Sub New(xvo As VisualSettings, xcol() As Column, monoFont As Font)
        InitializeComponent()

        'eColor = xeColor
        'eFont = xeFont
        'eI = xeI
        VisualOptions = xvo

        cButtonChange(cColumnTitle, VisualOptions.ColumnTitle.Color)
        cButtonChange(cBG, VisualOptions.Bg.Color)
        cButtonChange(cGrid, VisualOptions.pGrid.Color)
        cButtonChange(cSub, VisualOptions.pSub.Color)
        cButtonChange(cVerticalLine, VisualOptions.pVLine.Color)
        cButtonChange(cMeasureBarLine, VisualOptions.pMLine.Color)
        cButtonChange(cWaveForm, VisualOptions.pBGMWav.Color)
        cButtonChange(cMouseOver, VisualOptions.kMouseOver.Color)
        cButtonChange(cSelectedBorder, VisualOptions.kSelected.Color)
        cButtonChange(cAdjustLengthBorder, VisualOptions.kMouseOverE.Color)
        cButtonChange(cSelectionBox, VisualOptions.SelBox.Color)
        cButtonChange(cTSCursor, VisualOptions.PECursor.Color)
        cButtonChange(cTSSplitter, VisualOptions.PEHalf.Color)
        cButtonChange(cTSMouseOver, VisualOptions.PEMouseOver.Color)
        cButtonChange(cTSSelectionFill, VisualOptions.PESel.Color)
        cButtonChange(cTSBPM, VisualOptions.PEBPM.Color)

        fButtonChange(fColumnTitle, VisualOptions.ColumnTitleFont)
        fButtonChange(fNoteLabel, VisualOptions.kFont)
        fButtonChange(fMeasureLabel, VisualOptions.kMFont)
        fButtonChange(fTSBPM, VisualOptions.PEBPMFont)


        iNoteHeight.SetValClamped(VisualOptions.NoteHeight)
        iLabelVerticalShift.SetValClamped(VisualOptions.kLabelVShift)
        iLabelHorizShift.SetValClamped(VisualOptions.kLabelHShift)
        iLongLabelHorizShift.SetValClamped(VisualOptions.kLabelHShiftL)
        iHiddenNoteOpacity.SetValClamped(VisualOptions.kOpacity)
        iTSSensitivity.SetValClamped(VisualOptions.PEDeltaMouseOver)
        iMiddleSensitivity.SetValClamped(VisualOptions.MiddleDeltaRelease)

        TemporalColumns = xcol.Clone
        ReDim ColumnOptions(UBound(TemporalColumns))

        For i = 0 To UBound(TemporalColumns)
            Dim jw As New NumericUpDown
            With jw
                .BorderStyle = BorderStyle.FixedSingle
                .Location = New Point(lLeft(i), 12)
                .Maximum = 999
                .Size = New Size(33, 23)
                .Value = TemporalColumns(i).Width
            End With

            Dim jt As New TextBox
            With jt
                .BorderStyle = BorderStyle.FixedSingle
                .Location = New Point(lLeft(i), 34)
                .Size = New Size(33, 23)
                .Text = TemporalColumns(i).Title
            End With

            Dim js As New Button
            With js
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(i), 63)
                .Size = New Size(33, 66)
                .BackColor = Color.FromArgb(TemporalColumns(i).cNote)
                .ForeColor = TemporalColumns(i).cText
                .Text = To4Hex(TemporalColumns(i).cNote)
                .Name = "cNote"
            End With
            Dim jst As New Button
            With jst
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(i), 128)
                .Size = New Size(33, 66)
                .BackColor = Color.FromArgb(TemporalColumns(i).cNote)
                .ForeColor = TemporalColumns(i).cText
                .Text = To4Hex(TemporalColumns(i).cText.ToArgb)
                .Name = "cText"
            End With
            js.Tag = jst
            jst.Tag = js

            Dim jl As New Button
            With jl
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(i), 193)
                .Size = New Size(33, 66)
                .BackColor = Color.FromArgb(TemporalColumns(i).cLNote)
                .ForeColor = TemporalColumns(i).cLText
                .Text = To4Hex(TemporalColumns(i).cLNote)
                .Name = "cNote"
            End With
            Dim jlt As New Button
            With jlt
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(i), 258)
                .Size = New Size(33, 66)
                .BackColor = Color.FromArgb(TemporalColumns(i).cLNote)
                .ForeColor = TemporalColumns(i).cLText
                .Text = To4Hex(TemporalColumns(i).cLText.ToArgb)
                .Name = "cText"
            End With
            jl.Tag = jlt
            jlt.Tag = jl

            Dim jb As New Button
            With jb
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(i), 323)
                .Size = New Size(33, 66)
                .BackColor = TemporalColumns(i).cBG
                .ForeColor = IIf(CInt(TemporalColumns(i).cBG.GetBrightness*255) + 255 - TemporalColumns(i).cBG.A >= 128,
                                 Color.Black, Color.White)
                .Text = To4Hex(TemporalColumns(i).cBG.ToArgb)
                .Name = "cBG"
                .Tag = Nothing
            End With

            Panel1.Controls.Add(jw)
            Panel1.Controls.Add(jt)
            Panel1.Controls.Add(js)
            Panel1.Controls.Add(jst)
            Panel1.Controls.Add(jl)
            Panel1.Controls.Add(jlt)
            Panel1.Controls.Add(jb)
            ColumnOptions(i).Width = jw
            ColumnOptions(i).Title = jt
            ColumnOptions(i).SNote = js
            ColumnOptions(i).SText = jst
            ColumnOptions(i).LNote = jl
            ColumnOptions(i).LText = jlt
            ColumnOptions(i).BG = jb
            'AddHandler jw.ValueChanged, AddressOf WidthChanged
            'AddHandler jt.TextChanged, AddressOf TitleChanged
            AddHandler js.Click, AddressOf ButtonClick
            AddHandler jst.Click, AddressOf ButtonClick
            AddHandler jl.Click, AddressOf ButtonClick
            AddHandler jlt.Click, AddressOf ButtonClick
            AddHandler jb.Click, AddressOf ButtonClick
        Next
    End Sub

    Private Sub cButtonChange(xbutton As Button, c As Color)
        xbutton.Text = Hex(c.ToArgb)
        xbutton.BackColor = c
        xbutton.ForeColor = IIf(CInt(c.GetBrightness*255) + 255 - c.A >= 128, Color.Black, Color.White)
    End Sub

    Private Sub fButtonChange(xbutton As Button, f As Font)
        xbutton.Text = f.FontFamily.Name
        xbutton.Font = f
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        VisualOptions.ColumnTitle.Color = cColumnTitle.BackColor
        VisualOptions.Bg.Color = cBG.BackColor
        VisualOptions.pGrid.Color = cGrid.BackColor
        VisualOptions.pSub.Color = cSub.BackColor
        VisualOptions.pVLine.Color = cVerticalLine.BackColor
        VisualOptions.pMLine.Color = cMeasureBarLine.BackColor
        VisualOptions.pBGMWav.Color = cWaveForm.BackColor
        VisualOptions.kMouseOver.Color = cMouseOver.BackColor
        VisualOptions.kSelected.Color = cSelectedBorder.BackColor
        VisualOptions.kMouseOverE.Color = cAdjustLengthBorder.BackColor
        VisualOptions.SelBox.Color = cSelectionBox.BackColor
        VisualOptions.PECursor.Color = cTSCursor.BackColor
        VisualOptions.PEHalf.Color = cTSSplitter.BackColor
        VisualOptions.PEMouseOver.Color = cTSMouseOver.BackColor
        VisualOptions.PESel.Color = cTSSelectionFill.BackColor
        VisualOptions.PEBPM.Color = cTSBPM.BackColor

        VisualOptions.ColumnTitleFont = fColumnTitle.Font
        VisualOptions.kFont = fNoteLabel.Font
        VisualOptions.kMFont = fMeasureLabel.Font
        VisualOptions.PEBPMFont = fTSBPM.Font

        VisualOptions.NoteHeight = CInt(iNoteHeight.Value)
        VisualOptions.kLabelVShift = CInt(iLabelVerticalShift.Value)
        VisualOptions.kLabelHShift = CInt(iLabelHorizShift.Value)
        VisualOptions.kLabelHShiftL = CInt(iLongLabelHorizShift.Value)
        VisualOptions.kOpacity = CSng(iHiddenNoteOpacity.Value)
        VisualOptions.PEDeltaMouseOver = CInt(iTSSensitivity.Value)
        VisualOptions.MiddleDeltaRelease = CInt(iMiddleSensitivity.Value)

        MainWindow.setVO(VisualOptions)

        For i = 0 To ColumnOptions.Length - 1
            TemporalColumns(i).Title = ColumnOptions(i).Title.Text
            TemporalColumns(i).Width = ColumnOptions(i).Width.Value
            TemporalColumns(i).setNoteColor(ColumnOptions(i).SNote.BackColor.ToArgb)
            TemporalColumns(i).cText = ColumnOptions(i).SText.ForeColor
            TemporalColumns(i).setLNoteColor(ColumnOptions(i).LNote.BackColor.ToArgb)
            TemporalColumns(i).cLText = ColumnOptions(i).LText.ForeColor
            TemporalColumns(i).cBG = ColumnOptions(i).BG.BackColor
        Next

        MainWindow.Columns.column = TemporalColumns

        DialogResult = DialogResult.OK
        Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub OpVisual_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Font = MainWindow.Font

        'Language

        'Dim xSA() As String = Form1.lpvo

        'Me.Text = xSA(0)
        ''rb1.Text = xSA(1)
        ''rb2.Text = xSA(2)
        ''Label9.Text = xSA(3)
        ''Label10.Text = xSA(4)
        ''Label11.Text = "(" & xSA(5) & ")"
        ''Label20.Text = "(" & xSA(5) & ")"
        ''Label12.Text = xSA(6)
        ''Label13.Text = xSA(7)
        ''Label14.Text = xSA(8)
        ''Label17.Text = xSA(9)
        ''Label16.Text = xSA(10)
        ''Label15.Text = xSA(11)
        ''Label19.Text = xSA(12)
        ''Label18.Text = xSA(13)

        'Label1.Text = xSA(14)
        'Label2.Text = xSA(15)
        'Label4.Text = xSA(17)
        'Label6.Text = xSA(18)
        'Label5.Text = xSA(19)
        'Label7.Text = xSA(20)
        'Label8.Text = xSA(21)

        'Label37.Text = xSA(22)
        'Label30.Text = xSA(23)
        'Label26.Text = xSA(24)
        'Label27.Text = xSA(25)
        'Label29.Text = xSA(26)
        'Label40.Text = xSA(27)
        'Label22.Text = xSA(28)
        'Label31.Text = xSA(29)
        'Label98.Text = xSA(30)
        'Label86.Text = xSA(31)
        'Label88.Text = xSA(32)
        'Label21.Text = xSA(33)
        'Label24.Text = xSA(34)
        'Label28.Text = xSA(35)
        'Label25.Text = xSA(36)
        'Label38.Text = xSA(37)
        'Label39.Text = xSA(38)
        'Label34.Text = xSA(39)
        'Label35.Text = xSA(40)
        'Label23.Text = xSA(41)
        'OK_Button.Text = xSA(42)
        'Cancel_Button.Text = xSA(43)
        'Label3.Text = xSA(44)

        Me.Text = Strings.fopVisual.Title

        Label37.Text = Strings.fopVisual.ColumnCaption
        Label9.Text = Strings.fopVisual.ColumnCaptionFont
        Label30.Text = Strings.fopVisual.Background
        Label26.Text = Strings.fopVisual.Grid
        Label27.Text = Strings.fopVisual.SubGrid
        Label29.Text = Strings.fopVisual.VerticalLine
        Label40.Text = Strings.fopVisual.MeasureBarLine
        Label22.Text = Strings.fopVisual.BGMWaveform
        Label21.Text = Strings.fopVisual.NoteHeight
        Label24.Text = Strings.fopVisual.NoteLabel
        Label28.Text = Strings.fopVisual.MeasureLabel
        Label25.Text = Strings.fopVisual.LabelVerticalShift
        Label38.Text = Strings.fopVisual.LabelHorizontalShift
        Label39.Text = Strings.fopVisual.LongNoteLabelHorizontalShift
        Label33.Text = Strings.fopVisual.HiddenNoteOpacity
        Label34.Text = Strings.fopVisual.NoteBorderOnMouseOver
        Label35.Text = Strings.fopVisual.NoteBorderOnSelection
        Label23.Text = Strings.fopVisual.NoteBorderOnAdjustingLength
        Label31.Text = Strings.fopVisual.SelectionBoxBorder
        Label98.Text = Strings.fopVisual.TSCursor
        Label97.Text = Strings.fopVisual.TSSplitter
        Label96.Text = Strings.fopVisual.TSCursorSensitivity
        Label91.Text = Strings.fopVisual.TSMouseOverBorder
        Label86.Text = Strings.fopVisual.TSFill
        Label88.Text = Strings.fopVisual.TSBPM
        Label82.Text = Strings.fopVisual.TSBPMFont
        Label14.Text = Strings.fopVisual.MiddleSensitivity

        Label1.Text = Strings.fopVisual.Width
        Label2.Text = Strings.fopVisual.Caption
        Label4.Text = Strings.fopVisual.Note
        Label6.Text = Strings.fopVisual.Label
        Label5.Text = Strings.fopVisual.LongNote
        Label7.Text = Strings.fopVisual.LongNoteLabel
        Label8.Text = Strings.fopVisual.Bg

        '-----------------------------------------------
        'Page 2-----------------------------------------
        '-----------------------------------------------

        'With eaColor
        '    .Add(cColumnTitle)
        '    .Add(cBG)
        '    .Add(cGrid)
        '    .Add(cSub)
        '    .Add(cVerticalLine)
        '    .Add(cMeasureBarLine)
        '    .Add(cWaveForm)
        '    .Add(cSelectionBox)
        '    .Add(cTSCursor)
        '    .Add(cTSSelectionArea)
        '    .Add(cTSBPM)
        '    .Add(cMouseOver)
        '    .Add(cSelectedBorder)
        '    .Add(cAdjustLengthBorder)
        'End With
        'With eaFont
        '    .Add(fColumnTitle)
        '    .Add(fTSBPM)
        '    .Add(fNoteLabel)
        '    .Add(fMeasureLabel)
        'End With
        'With eaI
        '    .Add(iNoteHeight)
        '    .Add(iLabelVerticalShift)
        '    .Add(iLabelHorizShift)
        '    .Add(iLongNoteLabelHorizShift)
        '    .Add(iHiddenNoteOpacity)
        'End With

        'Dim i As Integer
        'For i = 0 To 13
        '    eaColor(i).Text = Hex(eColor(i).ToArgb)
        '    eaColor(i).BackColor = eColor(i)
        '    eaColor(i).ForeColor = IIf(eColor(i).GetBrightness + (255 - eColor(i).A) / 255 >= 0.5, Color.Black, Color.White)
        'Next
        'For i = 0 To 3
        '    eaFont(i).Text = eFont(i).FontFamily.Name
        '    eaFont(i).Font = eFont(i)
        'Next
        'For i = 0 To 4
        '    eaI(i).Value = eI(i)
        'Next

        'AddHandler cColumnTitle.Click, AddressOf BCClick
        'AddHandler cBG.Click, AddressOf BCClick
        'AddHandler cGrid.Click, AddressOf BCClick
        'AddHandler cSub.Click, AddressOf BCClick
        'AddHandler cVerticalLine.Click, AddressOf BCClick
        'AddHandler cMeasureBarLine.Click, AddressOf BCClick
        'AddHandler cWaveForm.Click, AddressOf BCClick
        'AddHandler cSelectionBox.Click, AddressOf BCClick
        'AddHandler cTSCursor.Click, AddressOf BCClick
        'AddHandler cTSSelectionArea.Click, AddressOf BCClick
        'AddHandler cTSBPM.Click, AddressOf BCClick
        'AddHandler cMouseOver.Click, AddressOf BCClick
        'AddHandler cSelectedBorder.Click, AddressOf BCClick
        'AddHandler cAdjustLengthBorder.Click, AddressOf BCClick
        'AddHandler fColumnTitle.Click, AddressOf BFClick
        'AddHandler fTSBPM.Click, AddressOf BFClick
        'AddHandler fNoteLabel.Click, AddressOf BFClick
        'AddHandler fMeasureLabel.Click, AddressOf BFClick
        'AddHandler iNoteHeight.ValueChanged, AddressOf BIValueChanged
        'AddHandler iLabelVerticalShift.ValueChanged, AddressOf BIValueChanged
        'AddHandler iLabelHorizShift.ValueChanged, AddressOf BIValueChanged
        'AddHandler iLongNoteLabelHorizShift.ValueChanged, AddressOf BIValueChanged
        'AddHandler iHiddenNoteOpacity.ValueChanged, AddressOf BIValueChanged

        '-----------------------------------------------
        'Page 1-----------------------------------------
        '-----------------------------------------------

        ''Width
        'For i = 0 To Columns.BGM
        '    Dim xjWidth As New NumericUpDown
        '    With xjWidth
        '        .BorderStyle = BorderStyle.FixedSingle
        '        .Location = New Point(lLeft(i), 12)
        '        .Maximum = 99
        '        .Size = New Size(33, 23)
        '        .Value = lWidth(i)
        '        .Tag = i
        '    End With
        '    Panel1.Controls.Add(xjWidth)
        '    jWidth.Add(xjWidth)
        '    AddHandler xjWidth.ValueChanged, AddressOf WidthChanged
        'Next
        '
        ''Title
        'For i = 0 To Columns.BGM
        '    Dim xjTitle As New TextBox
        '    With xjTitle
        '        .BorderStyle = BorderStyle.FixedSingle
        '        .Location = New Point(lLeft(i), 34)
        '        .Size = New Size(33, 23)
        '        .Text = lTitle(i)
        '        .Tag = i
        '    End With
        '    Panel1.Controls.Add(xjTitle)
        '    jTitle.Add(xjTitle)
        '    AddHandler xjTitle.TextChanged, AddressOf TitleChanged
        'Next
        '
        ''Color
        'For i = 0 To Columns.BGM
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(i), 63)
        '        .Size = New Size(33, 66)
        '        .BackColor = lColor(i)
        '        .ForeColor = lForeColor(i)
        '        .Text = To4Hex(lColor(i).ToArgb)
        '        .Tag = i
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jColor.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B1Click
        'Next
        '
        ''ForeColor
        'For i = 0 To Columns.BGM
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(i), 128)
        '        .Size = New Size(33, 66)
        '        .BackColor = lColor(i)
        '        .ForeColor = lForeColor(i)
        '        .Text = To4Hex(lForeColor(i).ToArgb)
        '        .Tag = i
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jForeColor.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B2Click
        'Next
        '
        ''ColorL
        'For i = 0 To Columns.BGM
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(i), 193)
        '        .Size = New Size(33, 66)
        '        .BackColor = lColorL(i)
        '        .ForeColor = lForeColorL(i)
        '        .Text = To4Hex(lColorL(i).ToArgb)
        '        .Tag = i
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jColorL.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B3Click
        'Next
        '
        ''ForeColorL
        'For i = 0 To Columns.BGM
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(i), 258)
        '        .Size = New Size(33, 66)
        '        .BackColor = lColorL(i)
        '        .ForeColor = lForeColorL(i)
        '        .Text = To4Hex(lForeColorL(i).ToArgb)
        '        .Tag = i
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jForeColorL.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B4Click
        'Next
        '
        ''BgColor
        'For i = 0 To Columns.BGM
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(i), 323)
        '        .Size = New Size(33, 66)
        '        .BackColor = lBg(i)
        '        .ForeColor = IIf(lBg(i).GetBrightness + (255 - lBg(i).A) / 255 >= 0.5, Color.Black, Color.White)
        '        .Text = To4Hex(lBg(i).ToArgb)
        '        .Tag = i
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jBg.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B5Click
        'Next
    End Sub

    'Private Sub rb1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Panel1.Visible = rb1.Checked
    'End Sub
    '
    'Private Sub rb2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Panel2.Visible = rb2.Checked
    'End Sub
    '
    Private Sub BCClick(sender As Object, e As EventArgs) _
        Handles cColumnTitle.Click,
                cBG.Click,
                cGrid.Click,
                cSub.Click,
                cVerticalLine.Click,
                cMeasureBarLine.Click,
                cWaveForm.Click,
                cMouseOver.Click,
                cSelectedBorder.Click,
                cAdjustLengthBorder.Click,
                cSelectionBox.Click,
                cTSCursor.Click,
                cTSSplitter.Click,
                cTSMouseOver.Click,
                cTSSelectionFill.Click,
                cTSBPM.Click

        'Dim xI As Integer = Val(sender.Tag)
        Dim s = CType(sender, Button)
        Dim xColorPicker As New ColorPicker
        xColorPicker.SetOrigColor(s.BackColor)
        If xColorPicker.ShowDialog(Me) = DialogResult.Cancel Then Exit Sub

        cButtonChange(s, xColorPicker.NewColor)
        'eColor(xI) = xColorPicker.NewColor
        'sender.Text = Hex(eColor(xI).ToArgb)
        'sender.BackColor = eColor(xI)
        'sender.ForeColor = IIf(eColor(xI).GetBrightness + (255 - eColor(xI).A) / 255 >= 0.5, Color.Black, Color.White)
    End Sub

    Private Sub BFClick(sender As Object, e As EventArgs) _
        Handles fColumnTitle.Click,
                fNoteLabel.Click,
                fMeasureLabel.Click,
                fTSBPM.Click

        'Dim xI As Integer = Val(sender.Tag)
        Dim s = CType(sender, Button)
        Dim xDFont As New FontDialog
        xDFont.Font = s.Font
        If xDFont.ShowDialog(Me) = DialogResult.Cancel Then Exit Sub

        fButtonChange(s, xDFont.Font)
        'eFont(xI) = xDFont.Font
        'sender.Text = eFont(xI).FontFamily.Name
        'sender.Font = eFont(xI)
    End Sub

    'Private Sub BIValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    eI(Val(sender.Tag)) = sender.Value
    'End Sub

    'Private Sub WidthChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    lWidth(Val(sender.Tag)) = sender.Value
    'End Sub
    'Private Sub TitleChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    lTitle(Val(sender.Tag)) = sender.Text
    'End Sub
    'Private Sub B1Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim xColorPicker As New ColorPicker
    '    Dim xI As Integer = Val(sender.Tag)
    '
    '    xColorPicker.SetOrigColor(lColor(xI))
    '    If xColorPicker.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub
    '    lColor(xI) = xColorPicker.NewColor
    '    sender.Text = To4Hex(lColor(xI).ToArgb)
    '    sender.BackColor = lColor(xI)
    '    jForeColor(xI).Backcolor = lColor(xI)
    'End Sub
    'Private Sub B2Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim xColorPicker As New ColorPicker
    '    Dim xI As Integer = Val(sender.Tag)
    '
    '    xColorPicker.SetOrigColor(lForeColor(xI))
    '    If xColorPicker.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub
    '    lForeColor(xI) = xColorPicker.NewColor
    '    sender.Text = To4Hex(lForeColor(xI).ToArgb)
    '    sender.ForeColor = lForeColor(xI)
    '    jColor(xI).ForeColor = lForeColor(xI)
    'End Sub
    'Private Sub B3Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim xColorPicker As New ColorPicker
    '    Dim xI As Integer = Val(sender.Tag)
    '
    '    xColorPicker.SetOrigColor(lColorL(xI))
    '    If xColorPicker.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub
    '    lColorL(xI) = xColorPicker.NewColor
    '    sender.Text = To4Hex(lColorL(xI).ToArgb)
    '    sender.BackColor = lColorL(xI)
    '    jForeColorL(xI).Backcolor = lColorL(xI)
    'End Sub
    'Private Sub B4Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim xColorPicker As New ColorPicker
    '    Dim xI As Integer = Val(sender.Tag)
    '
    '    xColorPicker.SetOrigColor(lForeColorL(xI))
    '    If xColorPicker.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub
    '    lForeColorL(xI) = xColorPicker.NewColor
    '    sender.Text = To4Hex(lForeColorL(xI).ToArgb)
    '    sender.ForeColor = lForeColorL(xI)
    '    jColorL(xI).ForeColor = lForeColorL(xI)
    'End Sub
    'Private Sub B5Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim xColorPicker As New ColorPicker
    '    Dim xI As Integer = Val(sender.Tag)
    '
    '    xColorPicker.SetOrigColor(lBg(xI))
    '    If xColorPicker.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub
    '    lBg(xI) = xColorPicker.NewColor
    '    sender.Text = To4Hex(lBg(xI).ToArgb)
    '    sender.BackColor = lBg(xI)
    '    sender.ForeColor = IIf(lBg(xI).GetBrightness + (255 - lBg(xI).A) / 255 >= 0.5, Color.Black, Color.White)
    'End Sub
    Private Sub ButtonClick(sender As Object, e As EventArgs)
        'Dim xI As Integer = Val(sender.Tag)
        Dim s = CType(sender, Button)
        Dim xColorPicker As New ColorPicker

        'xColorPicker.SetOrigColor(lColor(xI))
        If s.Name = "cText" Then xColorPicker.SetOrigColor(s.ForeColor) _
            Else xColorPicker.SetOrigColor(s.BackColor)

        If xColorPicker.ShowDialog(Me) = DialogResult.Cancel Then Exit Sub

        s.Text = To4Hex(xColorPicker.NewColor.ToArgb)
        Select Case s.Name
            Case "cNote"
                s.BackColor = xColorPicker.NewColor
                CType(s.Tag, Button).BackColor = xColorPicker.NewColor
            Case "cText"
                s.ForeColor = xColorPicker.NewColor
                CType(s.Tag, Button).ForeColor = xColorPicker.NewColor
            Case "cBG"
                s.BackColor = xColorPicker.NewColor
                s.ForeColor = IIf(CInt(xColorPicker.NewColor.GetBrightness*255) + 255 - xColorPicker.NewColor.A >= 128,
                                  Color.Black, Color.White)
        End Select
        'lColor(xI) = xColorPicker.NewColor
        'sender.Text = To4Hex(lColor(xI).ToArgb)
        'sender.BackColor = lColor(xI)
        'jForeColor(xI).Backcolor = lColor(xI)
    End Sub

    Private Function ColorArrayToIntArray(xC() As Color) As Integer()
        Dim xI(UBound(xC)) As Integer
        For i = 0 To UBound(xI)
            xI(i) = xC(i).ToArgb
        Next
        Return xI
    End Function

    Private Function To4Hex(xInt As Integer) As String
        Dim xCl As Color = Color.FromArgb(xInt)
        Return Hex(xCl.A) & vbCrLf & Hex(xCl.R) & vbCrLf & Hex(xCl.G) & vbCrLf & Hex(xCl.B)
    End Function
End Class
