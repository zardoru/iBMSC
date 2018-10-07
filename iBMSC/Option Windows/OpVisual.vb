Imports System.Runtime.CompilerServices
Imports System.Windows.Forms
Imports iBMSC.Editor

Public Class OpVisual
    Dim niB As Integer = MainWindow.niB
    Dim lLeft() As Integer = {78, 110, 142, 174, 208, 240, 272, 304, 336, 368, 400, 432, 464, 498, 530, 562, 594, 626, 658, 690, 722, 754, 788, 820, 852, 884, 918, 950}

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

    Dim vo As visualSettings
    Dim col() As Column
    Dim co() As ColumnOptionSet



    Public Sub New(ByVal xvo As visualSettings, ByVal xcol() As Column, ByVal monoFont As Font)
        InitializeComponent()

        'eColor = xeColor
        'eFont = xeFont
        'eI = xeI
        vo = xvo

        cButtonChange(Me.cColumnTitle, vo.ColumnTitle.Color)
        cButtonChange(Me.cBG, vo.Bg.Color)
        cButtonChange(Me.cGrid, vo.pGrid.Color)
        cButtonChange(Me.cSub, vo.pSub.Color)
        cButtonChange(Me.cVerticalLine, vo.pVLine.Color)
        cButtonChange(Me.cMeasureBarLine, vo.pMLine.Color)
        cButtonChange(Me.cWaveForm, vo.pBGMWav.Color)
        cButtonChange(Me.cMouseOver, vo.kMouseOver.Color)
        cButtonChange(Me.cSelectedBorder, vo.kSelected.Color)
        cButtonChange(Me.cAdjustLengthBorder, vo.kMouseOverE.Color)
        cButtonChange(Me.cSelectionBox, vo.SelBox.Color)
        cButtonChange(Me.cTSCursor, vo.PECursor.Color)
        cButtonChange(Me.cTSSplitter, vo.PEHalf.Color)
        cButtonChange(Me.cTSMouseOver, vo.PEMouseOver.Color)
        cButtonChange(Me.cTSSelectionFill, vo.PESel.Color)
        cButtonChange(Me.cTSBPM, vo.PEBPM.Color)

        fButtonChange(Me.fColumnTitle, vo.ColumnTitleFont)
        fButtonChange(Me.fNoteLabel, vo.kFont)
        fButtonChange(Me.fMeasureLabel, vo.kMFont)
        fButtonChange(Me.fTSBPM, vo.PEBPMFont)


        Me.iNoteHeight.SetValClamped(vo.kHeight)
        Me.iLabelVerticalShift.SetValClamped(vo.kLabelVShift)
        Me.iLabelHorizShift.SetValClamped(vo.kLabelHShift)
        Me.iLongLabelHorizShift.SetValClamped(vo.kLabelHShiftL)
        Me.iHiddenNoteOpacity.SetValClamped(vo.kOpacity)
        Me.iTSSensitivity.SetValClamped(vo.PEDeltaMouseOver)
        Me.iMiddleSensitivity.SetValClamped(vo.MiddleDeltaRelease)

        'lWidth = xlWidth
        'lTitle = xlTitle
        'lColor = xlColor
        'lForeColor = xlForeColor
        'lColorL = xlColorL
        'lForeColorL = xlForeColorL
        'lBg = xlBg

        col = xcol.Clone
        ReDim co(UBound(col))

        For xI1 As Integer = 0 To UBound(col)
            Dim jw As New NumericUpDown
            With jw
                .BorderStyle = BorderStyle.FixedSingle
                .Location = New Point(lLeft(xI1), 12)
                .Maximum = 999
                .Size = New Size(33, 23)
                .Value = col(xI1).Width
            End With

            Dim jt As New TextBox
            With jt
                .BorderStyle = BorderStyle.FixedSingle
                .Location = New Point(lLeft(xI1), 34)
                .Size = New Size(33, 23)
                .Text = col(xI1).Title
            End With

            Dim js As New Button
            With js
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(xI1), 63)
                .Size = New Size(33, 66)
                .BackColor = Color.FromArgb(col(xI1).cNote)
                .ForeColor = col(xI1).cText
                .Text = To4Hex(col(xI1).cNote)
                .Name = "cNote"
            End With
            Dim jst As New Button
            With jst
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(xI1), 128)
                .Size = New Size(33, 66)
                .BackColor = Color.FromArgb(col(xI1).cNote)
                .ForeColor = col(xI1).cText
                .Text = To4Hex(col(xI1).cText.ToArgb)
                .Name = "cText"
            End With
            js.Tag = jst
            jst.Tag = js

            Dim jl As New Button
            With jl
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(xI1), 193)
                .Size = New Size(33, 66)
                .BackColor = Color.FromArgb(col(xI1).cLNote)
                .ForeColor = col(xI1).cLText
                .Text = To4Hex(col(xI1).cLNote)
                .Name = "cNote"
            End With
            Dim jlt As New Button
            With jlt
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(xI1), 258)
                .Size = New Size(33, 66)
                .BackColor = Color.FromArgb(col(xI1).cLNote)
                .ForeColor = col(xI1).cLText
                .Text = To4Hex(col(xI1).cLText.ToArgb)
                .Name = "cText"
            End With
            jl.Tag = jlt
            jlt.Tag = jl

            Dim jb As New Button
            With jb
                .FlatStyle = FlatStyle.Popup
                .Font = monoFont
                .Location = New Point(lLeft(xI1), 323)
                .Size = New Size(33, 66)
                .BackColor = col(xI1).cBG
                .ForeColor = IIf(CInt(col(xI1).cBG.GetBrightness * 255) + 255 - col(xI1).cBG.A >= 128, Color.Black, Color.White)
                .Text = To4Hex(col(xI1).cBG.ToArgb)
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
            co(xI1).Width = jw
            co(xI1).Title = jt
            co(xI1).SNote = js
            co(xI1).SText = jst
            co(xI1).LNote = jl
            co(xI1).LText = jlt
            co(xI1).BG = jb
            'AddHandler jw.ValueChanged, AddressOf WidthChanged
            'AddHandler jt.TextChanged, AddressOf TitleChanged
            AddHandler js.Click, AddressOf ButtonClick
            AddHandler jst.Click, AddressOf ButtonClick
            AddHandler jl.Click, AddressOf ButtonClick
            AddHandler jlt.Click, AddressOf ButtonClick
            AddHandler jb.Click, AddressOf ButtonClick
        Next
    End Sub

    Private Sub cButtonChange(ByVal xbutton As Button, ByVal c As Color)
        xbutton.Text = Hex(c.ToArgb)
        xbutton.BackColor = c
        xbutton.ForeColor = IIf(CInt(c.GetBrightness * 255) + 255 - c.A >= 128, Color.Black, Color.White)
    End Sub

    Private Sub fButtonChange(ByVal xbutton As Button, ByVal f As Font)
        xbutton.Text = f.FontFamily.Name
        xbutton.Font = f
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        'With Form1

        '-----------------------------------------------
        'Page 2-----------------------------------------
        '-----------------------------------------------

        '.voTitle.Color = eColor(0)
        '.voBg.Color = eColor(1)
        '.voGrid.Color = eColor(2)
        '.voSub.Color = eColor(3)
        '.voVLine.Color = eColor(4)
        '.voMLine.Color = eColor(5)

        '.voBGMWav.Color = eColor(6)
        '.TWTransparency.Value = eColor(6).A
        '.TWTransparency2.Value = eColor(6).A
        '.TWSaturation.Value = eColor(6).GetSaturation * 1000
        '.TWSaturation2.Value = eColor(6).GetSaturation * 1000

        '.voSelBox.Color = eColor(7)
        '.voPECursor.Color = eColor(8)
        '.voPESel.Color = eColor(9)
        '.voPEBPM.Color = eColor(10)
        '.vKMouseOver.Color = eColor(11)
        '.vKSelected.Color = eColor(12)
        '.vKMouseOverE.Color = eColor(13)
        '.voTitleFont = eFont(0)
        '.voPEBPMFont = eFont(1)
        '.vKFont = eFont(2)
        '.vKMFont = eFont(3)
        '.vKHeight = eI(0)
        '.vKLabelVShift = eI(1)
        '.vKLabelHShift = eI(2)
        '.vKLabelHShiftL = eI(3)
        '.vKHidTransparency = CSng(eI(4)) / 100

        '-----------------------------------------------
        'Page 1-----------------------------------------
        '-----------------------------------------------

        '.kLength = lWidth
        '.kTitle = lTitle
        '.kColor = ColorArrayToIntArray(lColor)
        '.kForeColor = ColorArrayToIntArray(lForeColor)
        '.kColorL = ColorArrayToIntArray(lColorL)
        '.kForeColorL = ColorArrayToIntArray(lForeColorL)
        '.kBgColor = ColorArrayToIntArray(lBg)

        'End With

        vo.ColumnTitle.Color = Me.cColumnTitle.BackColor
        vo.Bg.Color = Me.cBG.BackColor
        vo.pGrid.Color = Me.cGrid.BackColor
        vo.pSub.Color = Me.cSub.BackColor
        vo.pVLine.Color = Me.cVerticalLine.BackColor
        vo.pMLine.Color = Me.cMeasureBarLine.BackColor
        vo.pBGMWav.Color = Me.cWaveForm.BackColor
        vo.kMouseOver.Color = Me.cMouseOver.BackColor
        vo.kSelected.Color = Me.cSelectedBorder.BackColor
        vo.kMouseOverE.Color = Me.cAdjustLengthBorder.BackColor
        vo.SelBox.Color = Me.cSelectionBox.BackColor
        vo.PECursor.Color = Me.cTSCursor.BackColor
        vo.PEHalf.Color = Me.cTSSplitter.BackColor
        vo.PEMouseOver.Color = Me.cTSMouseOver.BackColor
        vo.PESel.Color = Me.cTSSelectionFill.BackColor
        vo.PEBPM.Color = Me.cTSBPM.BackColor

        vo.ColumnTitleFont = Me.fColumnTitle.Font
        vo.kFont = Me.fNoteLabel.Font
        vo.kMFont = Me.fMeasureLabel.Font
        vo.PEBPMFont = Me.fTSBPM.Font

        vo.kHeight = CInt(Me.iNoteHeight.Value)
        vo.kLabelVShift = CInt(Me.iLabelVerticalShift.Value)
        vo.kLabelHShift = CInt(Me.iLabelHorizShift.Value)
        vo.kLabelHShiftL = CInt(Me.iLongLabelHorizShift.Value)
        vo.kOpacity = CSng(Me.iHiddenNoteOpacity.Value)
        vo.PEDeltaMouseOver = CInt(Me.iTSSensitivity.Value)
        vo.MiddleDeltaRelease = CInt(Me.iMiddleSensitivity.Value)

        MainWindow.setVO(vo)

        For xI1 As Integer = 0 To UBound(co)
            col(xI1).Title = co(xI1).Title.Text
            col(xI1).Width = co(xI1).Width.Value
            col(xI1).setNoteColor(co(xI1).SNote.BackColor.ToArgb)
            col(xI1).cText = co(xI1).SText.ForeColor
            col(xI1).setLNoteColor(co(xI1).LNote.BackColor.ToArgb)
            col(xI1).cLText = co(xI1).LText.ForeColor
            col(xI1).cBG = co(xI1).BG.BackColor
        Next

        MainWindow.column = col

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub OpVisual_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

        'Dim xI1 As Integer
        'For xI1 = 0 To 13
        '    eaColor(xI1).Text = Hex(eColor(xI1).ToArgb)
        '    eaColor(xI1).BackColor = eColor(xI1)
        '    eaColor(xI1).ForeColor = IIf(eColor(xI1).GetBrightness + (255 - eColor(xI1).A) / 255 >= 0.5, Color.Black, Color.White)
        'Next
        'For xI1 = 0 To 3
        '    eaFont(xI1).Text = eFont(xI1).FontFamily.Name
        '    eaFont(xI1).Font = eFont(xI1)
        'Next
        'For xI1 = 0 To 4
        '    eaI(xI1).Value = eI(xI1)
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
        'For xI1 = 0 To niB
        '    Dim xjWidth As New NumericUpDown
        '    With xjWidth
        '        .BorderStyle = BorderStyle.FixedSingle
        '        .Location = New Point(lLeft(xI1), 12)
        '        .Maximum = 99
        '        .Size = New Size(33, 23)
        '        .Value = lWidth(xI1)
        '        .Tag = xI1
        '    End With
        '    Panel1.Controls.Add(xjWidth)
        '    jWidth.Add(xjWidth)
        '    AddHandler xjWidth.ValueChanged, AddressOf WidthChanged
        'Next
        '
        ''Title
        'For xI1 = 0 To niB
        '    Dim xjTitle As New TextBox
        '    With xjTitle
        '        .BorderStyle = BorderStyle.FixedSingle
        '        .Location = New Point(lLeft(xI1), 34)
        '        .Size = New Size(33, 23)
        '        .Text = lTitle(xI1)
        '        .Tag = xI1
        '    End With
        '    Panel1.Controls.Add(xjTitle)
        '    jTitle.Add(xjTitle)
        '    AddHandler xjTitle.TextChanged, AddressOf TitleChanged
        'Next
        '
        ''Color
        'For xI1 = 0 To niB
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(xI1), 63)
        '        .Size = New Size(33, 66)
        '        .BackColor = lColor(xI1)
        '        .ForeColor = lForeColor(xI1)
        '        .Text = To4Hex(lColor(xI1).ToArgb)
        '        .Tag = xI1
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jColor.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B1Click
        'Next
        '
        ''ForeColor
        'For xI1 = 0 To niB
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(xI1), 128)
        '        .Size = New Size(33, 66)
        '        .BackColor = lColor(xI1)
        '        .ForeColor = lForeColor(xI1)
        '        .Text = To4Hex(lForeColor(xI1).ToArgb)
        '        .Tag = xI1
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jForeColor.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B2Click
        'Next
        '
        ''ColorL
        'For xI1 = 0 To niB
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(xI1), 193)
        '        .Size = New Size(33, 66)
        '        .BackColor = lColorL(xI1)
        '        .ForeColor = lForeColorL(xI1)
        '        .Text = To4Hex(lColorL(xI1).ToArgb)
        '        .Tag = xI1
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jColorL.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B3Click
        'Next
        '
        ''ForeColorL
        'For xI1 = 0 To niB
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(xI1), 258)
        '        .Size = New Size(33, 66)
        '        .BackColor = lColorL(xI1)
        '        .ForeColor = lForeColorL(xI1)
        '        .Text = To4Hex(lForeColorL(xI1).ToArgb)
        '        .Tag = xI1
        '    End With
        '    Panel1.Controls.Add(xjColor)
        '    jForeColorL.Add(xjColor)
        '    AddHandler xjColor.Click, AddressOf B4Click
        'Next
        '
        ''BgColor
        'For xI1 = 0 To niB
        '    Dim xjColor As New Button
        '    With xjColor
        '        .FlatStyle = FlatStyle.Popup
        '        .Font = New Font("Consolas", 9)
        '        .Location = New Point(lLeft(xI1), 323)
        '        .Size = New Size(33, 66)
        '        .BackColor = lBg(xI1)
        '        .ForeColor = IIf(lBg(xI1).GetBrightness + (255 - lBg(xI1).A) / 255 >= 0.5, Color.Black, Color.White)
        '        .Text = To4Hex(lBg(xI1).ToArgb)
        '        .Tag = xI1
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
    Private Sub BCClick(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles cColumnTitle.Click, _
            cBG.Click, _
            cGrid.Click, _
            cSub.Click, _
            cVerticalLine.Click, _
            cMeasureBarLine.Click, _
            cWaveForm.Click, _
            cMouseOver.Click, _
            cSelectedBorder.Click, _
            cAdjustLengthBorder.Click, _
            cSelectionBox.Click, _
            cTSCursor.Click, _
            cTSSplitter.Click, _
            cTSMouseOver.Click, _
            cTSSelectionFill.Click, _
            cTSBPM.Click

        'Dim xI As Integer = Val(sender.Tag)
        Dim s As Button = CType(sender, Button)
        Dim xColorPicker As New ColorPicker
        xColorPicker.SetOrigColor(s.BackColor)
        If xColorPicker.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub

        cButtonChange(s, xColorPicker.NewColor)
        'eColor(xI) = xColorPicker.NewColor
        'sender.Text = Hex(eColor(xI).ToArgb)
        'sender.BackColor = eColor(xI)
        'sender.ForeColor = IIf(eColor(xI).GetBrightness + (255 - eColor(xI).A) / 255 >= 0.5, Color.Black, Color.White)
    End Sub

    Private Sub BFClick(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles fColumnTitle.Click, _
            fNoteLabel.Click, _
            fMeasureLabel.Click, _
            fTSBPM.Click

        'Dim xI As Integer = Val(sender.Tag)
        Dim s As Button = CType(sender, Button)
        Dim xDFont As New FontDialog
        xDFont.Font = s.Font
        If xDFont.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub

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
    Private Sub ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim xI As Integer = Val(sender.Tag)
        Dim s As Button = CType(sender, Button)
        Dim xColorPicker As New ColorPicker

        'xColorPicker.SetOrigColor(lColor(xI))
        If s.Name = "cText" Then xColorPicker.SetOrigColor(s.ForeColor) _
        Else xColorPicker.SetOrigColor(s.BackColor)

        If xColorPicker.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub

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
                s.ForeColor = IIf(CInt(xColorPicker.NewColor.GetBrightness * 255) + 255 - xColorPicker.NewColor.A >= 128, Color.Black, Color.White)
        End Select
        'lColor(xI) = xColorPicker.NewColor
        'sender.Text = To4Hex(lColor(xI).ToArgb)
        'sender.BackColor = lColor(xI)
        'jForeColor(xI).Backcolor = lColor(xI)
    End Sub

    Private Function ColorArrayToIntArray(ByVal xC() As Color) As Integer()
        Dim xI(UBound(xC)) As Integer
        For xI1 As Integer = 0 To UBound(xI)
            xI(xI1) = xC(xI1).ToArgb
        Next
        Return xI
    End Function

    Private Function To4Hex(ByVal xInt As Integer) As String
        Dim xCl As Color = Color.FromArgb(xInt)
        Return Hex(xCl.A) & vbCrLf & Hex(xCl.R) & vbCrLf & Hex(xCl.G) & vbCrLf & Hex(xCl.B)
    End Function
End Class
