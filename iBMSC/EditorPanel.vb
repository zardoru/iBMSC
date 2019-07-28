Imports System.Linq
Imports iBMSC.Editor

Public Class EditorPanel
    Inherits Panel

    Public Shadows HorizontalScroll As Integer
    Public Shadows VerticalScroll As Integer
    Public DoubleBuffer As BufferedGraphics
    Public Buffer As Graphics
    Public Theme As visualSettings
    Public Editor As MainWindow
    Public WithEvents VerticalScrollBar As VScrollBar
    Public WithEvents HorizontalScrollBar As HScrollBar
    Public LastVerticalScroll As Integer
    Public LastHorizontalScroll As Integer

    Public Sub HsValueChanged(sender As Object, e As EventArgs) Handles HorizontalScrollBar.ValueChanged
        HorizontalScroll = sender.Value
        LastHorizontalScroll = sender.Value
        Refresh()
    End Sub

    Public Sub VsValueChanged(sender As Object, e As EventArgs) Handles VerticalScrollBar.ValueChanged
        VerticalScroll = sender.Value
        LastVerticalScroll = sender.Value
        Refresh()
    End Sub

    Public Sub Init(wnd As MainWindow, theme As visualSettings, vscroll As VScrollBar, hscroll As HScrollBar)
        Me.Theme = theme
        Parent = wnd
        Editor = wnd

        ' az: there's ought to be a Better Way, but for now we adapt what iBMSC originally had
        ' BringToFront()
        DoubleBuffered = True
    End Sub

    Private Sub DrawDragAndDrop()
        If Editor.DragDropFilename.Length > 0 Then
            Dim filenames = Join(Editor.DragDropFilename, vbCrLf)
            Dim brush As New SolidBrush(Color.FromArgb(&HC0FFFFFF))
            Dim format As New StringFormat With {
                .Alignment = StringAlignment.Center,
                .LineAlignment = StringAlignment.Center
            }
            Buffer.DrawString(filenames, Editor.Font, brush, DisplayRectangle, format)
        End If
    End Sub

    Function GetColumnHighlightColor(col As Color, Optional factor As Double = 2.0)
        Dim clamp = Function(x) IIf(x > 255, 255, x)
        Return Color.FromArgb(
                clamp(col.A * factor),
                clamp(col.R * factor),
                clamp(col.G * factor),
                clamp(col.B * factor))
    End Function

    Public Function GetColumnAtX(x As Integer) As Integer
        Dim i As Integer = 0
        'horizontal position of the mouse
        Dim mLeft As Integer = x / Editor.Grid.WidthScale + HorizontalScroll
        Dim col = 0
        If mLeft >= 0 Then
            Do
                'get the column where mouse is 
                If mLeft < Editor.Columns.GetColumnLeft(i + 1) Or i >= Editor.Columns.ColumnCount Then col = i : Exit Do
                i += 1
            Loop
        End If

        Dim eci = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(col)
        Return Editor.Columns.EnabledColumnIndexToColumnArrayIndex(eci)  'get the enabled column where mouse is 
    End Function

    Private Sub DrawBackgroundColor()
        Dim Grid = Editor.Grid
        Dim Columns = Editor.Columns

        For i = 0 To Columns.ColumnCount
            If HPositionToPanelX(Columns.GetColumnLeft(i + 1)) + 1 < 0 Then
                Continue For
            End If

            If (Columns.GetColumnLeft(i) - HorizontalScroll) * Grid.WidthScale + 1 > Size.Width Then
                Exit For
            End If

            If Not Columns.GetColumn(i).cBG.GetBrightness = 0 And Columns.GetColumnWidth(i) > 0 Then
                Dim col = Columns.GetColumn(i).cBG
                If i = GetColumnAtX(Editor.State.Mouse.MouseMoveStatus.X) Then
                    Dim bf = 1.2
                    col = GetColumnHighlightColor(col)
                End If
                Dim brush = New SolidBrush(col)

                Buffer.FillRectangle(brush,
                                          HPositionToPanelX(Columns.GetColumnLeft(i)) + 1,
                                          0,
                                          Columns.GetColumnWidth(i) * Editor.Grid.WidthScale,
                                          Size.Height)
            End If
        Next
    End Sub

    Private Function IsNoteVisible(note As Note) As Boolean
        Dim xUpperBorder As Single = Math.Abs(VerticalScroll) + Size.Height / Editor.Grid.HeightScale
        Dim xLowerBorder As Single = Math.Abs(VerticalScroll) - Theme.kHeight / Editor.Grid.HeightScale

        Dim AboveLower = note.VPosition >= xLowerBorder
        Dim HeadBelow = note.VPosition <= xLowerBorder
        Dim TailAbove = note.VPosition + note.Length >= xLowerBorder
        Dim IntersectsNT = HeadBelow And TailAbove
        Dim Intersects = (note.VPosition <= xLowerBorder And
                         Editor.Notes(note.LNPair).VPosition >= xLowerBorder)
        Dim AboveUpper = note.VPosition > xUpperBorder

        Dim NoteInside = (Not AboveUpper) And AboveLower

        Return NoteInside OrElse IntersectsNT OrElse Intersects
    End Function

    Private Function IsNoteVisible(noteindex As Integer) As Boolean
        Return IsNoteVisible(Editor.Notes(noteindex))
    End Function

    Private Sub DrawGridLines(measureIndex As Integer,
                              divisions As Integer, pen As Pen)
        Dim row = 0
        Dim xUpper As Double = Editor.MeasureUpper(measureIndex)
        Dim xCurr = Editor.MeasureBottom(measureIndex)
        Dim rowSize = 192 / divisions
        Do While xCurr < xUpper
            Dim y = VPositionToPanelY(xCurr)
            Buffer.DrawLine(pen,
                                     0, y,
                                     Size.Width, y)
            row += 1
            xCurr = Editor.MeasureBottom(measureIndex) + row * rowSize
        Loop
    End Sub

    Private Function DrawPanelLines(xVSu As Integer) As Integer
        'Vertical line
        If Editor.Grid.ShowVerticalLines Then
            For i = 0 To Editor.Columns.ColumnCount
                Dim xpos = (Editor.Columns.GetColumnLeft(i) - HorizontalScroll) * Editor.Grid.WidthScale
                If xpos + 1 < 0 Then Continue For
                If xpos + 1 > Size.Width Then Exit For
                If Editor.Columns.GetColumnWidth(i) > 0 Then
                    Buffer.DrawLine(Theme.pVLine,
                                         xpos, 0,
                                         xpos, Size.Height)
                End If
            Next
        End If

        'Grid, Sub, Measure
        Dim Measure
        For Measure = Editor.MeasureAtDisplacement(-VerticalScroll) To Editor.MeasureAtDisplacement(xVSu)
            'grid
            If Editor.Grid.ShowMainGrid Then
                DrawGridLines(Measure, Editor.Grid.Divider, Theme.pGrid)
            End If

            'sub
            If Editor.Grid.ShowSubGrid Then
                DrawGridLines(Measure, Editor.Grid.Subdivider, Theme.pSub)
            End If


            'measure and measurebar
            Dim xCurr = Editor.MeasureBottom(Measure)
            Dim pHeight = VPositionToPanelY(xCurr)
            If Editor.Grid.ShowMeasureBars Then
                Buffer.DrawLine(Theme.pMLine,
                                     0, pHeight,
                                     Size.Width, pHeight)
            End If
            If Editor.Grid.ShowMeasureNumber Then
                Dim brush = New SolidBrush(Editor.Columns.GetColumn(0).cText)
                Buffer.DrawString("[" & Add3Zeros(Measure).ToString & "]",
                                       Theme.kMFont, brush, -HorizontalScroll * Editor.Grid.WidthScale,
                                       pHeight - Theme.kMFont.Height)
            End If
        Next

        Dim vpos = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)
        Dim mouseLineHeight = VPositionToPanelY(vpos)
        Dim p = New Pen(Color.White)
        Buffer.DrawLine(p, 0, mouseLineHeight, Size.Width, mouseLineHeight)

        Return Measure
    End Function

    Private Sub DrawColumnCaptions()
        Dim Columns = Editor.Columns
        For i = 0 To Editor.Columns.ColumnCount
            Dim ColEnd = (Columns.GetColumnLeft(i + 1) - HorizontalScroll) * Editor.Grid.WidthScale
            If ColEnd + 1 < 0 Then
                Continue For
            End If

            Dim ColStart = (Columns.GetColumnLeft(i) - HorizontalScroll) * Editor.Grid.WidthScale
            If ColStart + 1 > Size.Width Then
                Exit For
            End If

            If Columns.GetColumnWidth(i) > 0 Then
                Buffer.DrawString(
                    Columns.nTitle(i),
                    Theme.ColumnTitleFont,
                    Theme.ColumnTitle,
                    ColStart, 0)
            End If
        Next
    End Sub

    Private Sub DrawSelectionBox()
        Dim pMouseMove = Editor.State.Mouse.pMouseMove
        Dim LastMouseDownLocation = Editor.State.Mouse.LastMouseDownLocation
        If Editor.IsSelectMode AndAlso
           (Editor.FocusedPanel Is Me) AndAlso
           Not (pMouseMove = New Point(-1, -1) Or
                LastMouseDownLocation = New Point(-1, -1)) Then
            Dim mdx = pMouseMove.X - LastMouseDownLocation.X
            Dim mdy = pMouseMove.Y - LastMouseDownLocation.Y
            Buffer.DrawRectangle(Theme.SelBox,
                                Math.Min(LastMouseDownLocation.X, pMouseMove.X),
                                Math.Min(LastMouseDownLocation.Y, pMouseMove.Y),
                                Math.Abs(mdx), Math.Abs(mdy))
        End If
    End Sub

    Private Sub DrawWaveform(xVSR As Integer)
        'If wWavL IsNot Nothing And wWavR IsNot Nothing And wPrecision > 0 Then
        '    If wLock Then
        '        For xI0 As Integer = 1 To UBound(Notes)
        '            If Notes(xI0).ColumnIndex >= ColumnType.BGM Then wPosition = Notes(xI0).VPosition : Exit For
        '        Next
        '    End If

        '    Dim xPtsL(xTHeight * wPrecision) As PointF
        '    Dim xPtsR(xTHeight * wPrecision) As PointF

        '    Dim xD1 As Double

        '    Dim bVPosition() As Double = {wPosition}
        '    Dim bBPM() As Decimal = {Notes(0).Value / 10000}
        '    Dim bWavDataIndex() As Decimal = {0}

        '    For Each Note In Editor.Notes.Skip(1).Where(Function(x) x.ColumnIndex = ColumnType.BPM)
        '        If Note.VPosition >= wPosition Then
        '            ReDim Preserve bVPosition(bVPosition.Length)
        '            ReDim Preserve bBPM(bBPM.Length)
        '            ReDim Preserve bWavDataIndex(bWavDataIndex.Length)
        '            bVPosition(UBound(bVPosition)) = Note.VPosition
        '            bBPM(UBound(bBPM)) = Note.Value / 10000
        '            bWavDataIndex(UBound(bWavDataIndex)) = (Note.VPosition - bVPosition(UBound(bVPosition) - 1)) * 1.25 * wSampleRate / bBPM(UBound(bBPM) - 1) + bWavDataIndex(UBound(bWavDataIndex) - 1)
        '        Else
        '            bBPM(0) = Note.Value / 10000
        '        End If
        '    Next


        '    For i = xTHeight * wPrecision To 0 Step -1
        '        Dim xI3 = (-i / wPrecision + xTHeight + xVSR * GxHeight - 1) / GxHeight
        '        For j = 1 To UBound(bVPosition)
        '            If bVPosition(j) >= xI3 Then Exit For
        '        Next

        '        xD1 = bWavDataIndex(-1) +
        '              (xI3 - bVPosition(-1)) * 1.25 * wSampleRate / bBPM(-1)

        '        If xD1 <= UBound(wWavL) And xD1 >= 0 Then
        '            xPtsL(i) = New PointF(wWavL(Int(xD1)) * wWidth + wLeft, i / wPrecision)
        '            xPtsR(i) = New PointF(wWavR(Int(xD1)) * wWidth + wLeft, i / wPrecision)
        '        Else
        '            xPtsL(i) = New PointF(wLeft, i / wPrecision)
        '            xPtsR(i) = New PointF(wLeft, i / wPrecision)
        '        End If
        '    Next
        '    Buffer.DrawLines(Theme.pBGMWav, xPtsL)
        '    Buffer.DrawLines(Theme.pBGMWav, xPtsR)
        'End If
    End Sub

    Private Sub DrawNotes()
        Dim upperBorder = Math.Abs(VerticalScroll) + Size.Height / Editor.Grid.HeightScale
        Dim lowerBorder = Math.Abs(VerticalScroll) - Theme.kHeight / Editor.Grid.HeightScale

        Dim renderNotes = Editor.Notes.
                                 Where(Function(x) x.VPosition <= upperBorder).
                                 Where(Function(x) IsNoteVisible(x)).
                                 Where(Function(x) Editor.Columns.nEnabled(x.ColumnIndex))

        For Each Note In renderNotes
            If Editor.NTInput Then
                DrawNoteNT(Note)
            Else
                DrawNote(Note)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Draw a note in a buffer.
    ''' </summary>
    ''' <param name="sNote">Note to be drawn.</param>

    Private Sub DrawNote(ByVal sNote As Note)
        Dim xAlpha As Single = 1.0F
        If sNote.Hidden Then xAlpha = Theme.kOpacity

        Dim xLabel As String = C10to36(sNote.Value \ 10000)
        If Editor.ShowFileName Then
            If Editor.Columns.IsColumnSound(sNote.ColumnIndex) Then
                If Editor.BmsWAV(C36to10(xLabel)) <> "" Then
                    xLabel = Path.GetFileNameWithoutExtension(Editor.BmsWAV(C36to10(xLabel)))
                End If
            Else
                If Editor.BmsBMP(C36to10(xLabel)) <> "" Then
                    xLabel = Path.GetFileNameWithoutExtension(Editor.BmsBMP(C36to10(xLabel)))
                End If
            End If
        End If

        Dim xPen As Pen
        Dim xBrush As Drawing2D.LinearGradientBrush
        Dim xBrush2 As SolidBrush

        Dim bright As Color
        Dim dark As Color

        Dim Grid = Editor.Grid
        Dim Columns = Editor.Columns
        Dim colX = Columns.GetColumnLeft(sNote.ColumnIndex)
        Dim colWidth = Columns.GetColumnWidth(sNote.ColumnIndex)
        Dim column = Columns.GetColumn(sNote.ColumnIndex)
        Dim startX = HPositionToPanelX(colX)
        Dim startY = VPositionToPanelY(sNote.VPosition) - Theme.kHeight

        Dim endX = HPositionToPanelX(colX + colWidth)
        Dim endY = VPositionToPanelY(sNote.VPosition)

        Dim p1 = New Point(startX,
                           startY - 10)
        Dim p2 = New Point(endX,
                           endY + 10)

        If Not sNote.LongNote Then
            xPen = New Pen(column.getBright(xAlpha))

            bright = column.getBright(xAlpha)
            dark = column.getDark(xAlpha)

            If sNote.Landmine Then
                bright = Color.Red
                dark = Color.Red
            End If

            xBrush2 = New SolidBrush(column.cText)
        Else
            bright = column.getLongBright(xAlpha)
            dark = column.getLongDark(xAlpha)

            xBrush2 = New SolidBrush(column.cLText)
        End If

        xPen = New Pen(bright)
        xBrush = New Drawing2D.LinearGradientBrush(p1, p2, bright, dark)

        Dim colDrawWidth = colWidth * Editor.Grid.WidthScale

        ' Fill
        Buffer.FillRectangle(xBrush,
                                      startX + 2,
                                      startY - Theme.kHeight + 1,
                                      colDrawWidth - 3,
                                      Theme.kHeight - 1)
        ' Outline
        Buffer.DrawRectangle(xPen,
                                     startX + 1,
                                     startY - Theme.kHeight,
                                     colDrawWidth - 2,
                                     Theme.kHeight)

        ' Label
        Dim label = IIf(Columns.IsColumnNumeric(sNote.ColumnIndex), sNote.Value / 10000, xLabel)
        Buffer.DrawString(label,
                                  Theme.kFont, xBrush2,
                                  startX + Theme.kLabelHShift,
                                  startY + Theme.kLabelVShift)

        If sNote.ColumnIndex < ColumnType.BGM Then
            If sNote.LNPair <> 0 Then
                DrawPairedLNBody(sNote, xAlpha)
            End If
        End If


        'e.Graphics.DrawString(sNote.TimeOffset.ToString("0.##"), New Font("Verdana", 9), Brushes.Cyan, _
        '                      New Point(HorizontalPositiontoDisplay(GetColumnLeft(sNote.ColumnIndex + 1), horizontalScroll), VerticalPositiontoDisplay(sNote.VPosition, verticalScroll, xHeight) - vo.kHeight - 2))

        'If ErrorCheck AndAlso (sNote.LongNote Xor sNote.PairWithI <> 0) Then e.Graphics.DrawImage(My.Resources.ImageError, _
        If Editor.ErrorCheck AndAlso sNote.HasError Then
            Dim ex = HPositionToPanelX(colX + colWidth / 2)
            Dim ey = VPositionToPanelY(sNote.VPosition) - Theme.kHeight / 2
            Buffer.DrawImage(My.Resources.ImageError,
                                CInt(ex - 12),
                                CInt(ey - 12),
                                24, 24)
        End If

        If sNote.Selected Then
            Buffer.DrawRectangle(
                Theme.kSelected,
                colX, VPositionToPanelY(sNote.VPosition) - Theme.kHeight - 1,
                colX + colDrawWidth, Theme.kHeight + 2)
        End If

    End Sub

    Private Sub DrawPairedLNBody(sNote As Note, xAlpha As Single)
        Dim column = Editor.Columns.GetColumn(sNote.ColumnIndex)

        Dim xPen2 As New Pen(column.getLongBright(xAlpha))

        Dim colX = Editor.Columns.GetColumnLeft(sNote.ColumnIndex)
        Dim colWidth = Editor.Columns.GetColumnWidth(sNote.ColumnIndex)

        Dim xBrush3 As New Drawing2D.LinearGradientBrush(
                    New Point(HPositionToPanelX(colX - 0.5 * colWidth),
                            VPositionToPanelY(Editor.Notes(sNote.LNPair).VPosition)),
                    New Point(HPositionToPanelX(colX + 1.5 * colWidth),
                            VPositionToPanelY(sNote.VPosition) + Theme.kHeight),
                    column.getLongBright(xAlpha),
                    column.getLongDark(xAlpha))

        ' assume ln pair column = our column
        Dim x1 = HPositionToPanelX(colX) ' HorizontalPositiontoDisplay(Editor.Columns.GetColumnLeft(Editor.Notes(sNote.LNPair).ColumnIndex))
        Dim y1 = VPositionToPanelY(Editor.Notes(sNote.LNPair).VPosition)
        Dim colDrawWidth = colWidth * Editor.Grid.WidthScale ' GetColumnWidth(Notes(sNote.LNPair).ColumnIndex) * GxWidth1
        Dim height = VPositionToPanelY(sNote.VPosition) - VPositionToPanelY(Editor.Notes(sNote.LNPair).VPosition)
        Buffer.FillRectangle(xBrush3,
                                      x1 + 3,
                                      y1 + 1,
                                      colDrawWidth - 5, ' az: ???
                                      height - Theme.kHeight - 1)
        Buffer.DrawRectangle(xPen2,
                                      x1 + 2,
                                      y1,
                                      colDrawWidth - 4,
                                      height - Theme.kHeight)
    End Sub

    ''' <summary>
    ''' Draw a note in a buffer under NT mode.
    ''' </summary>
    ''' <param name="sNote">Note to be drawn.</param>
    Private Sub DrawNoteNT(ByVal sNote As Note)

        Dim alpha As Single = 1.0F
        If sNote.Hidden Then alpha = Theme.kOpacity

        Dim Columns = Editor.Columns
        Dim column = Editor.Columns.GetColumn(sNote.ColumnIndex)
        Dim xLabel As String = C10to36(sNote.Value \ 10000)
        If Editor.ShowFileName Then
            If Columns.IsColumnSound(sNote.ColumnIndex) Then
                If Editor.BmsWAV(C36to10(xLabel)) <> "" Then
                    xLabel = Path.GetFileNameWithoutExtension(Editor.BmsWAV(C36to10(xLabel)))
                End If
            Else
                If Editor.BmsBMP(C36to10(xLabel)) <> "" Then
                    xLabel = Path.GetFileNameWithoutExtension(Editor.BmsBMP(C36to10(xLabel)))
                End If
            End If
        End If

        Dim xPen1 As Pen
        Dim xBrush As Drawing2D.LinearGradientBrush
        Dim xBrush2 As SolidBrush

        Dim p1 As Point
        Dim p2 As Point
        Dim bright As Color
        Dim dark As Color

        Dim colX = Columns.GetColumnLeft(sNote.ColumnIndex)
        Dim colWidth = Columns.GetColumnWidth(sNote.ColumnIndex)
        Dim startX = HPositionToPanelX(colX)
        Dim endX = HPositionToPanelX(colX + colWidth)
        Dim startY = VPositionToPanelY(sNote.VPosition)
        Dim endY = VPositionToPanelY(sNote.VPosition + sNote.Length)

        If sNote.Length = 0 Then
            p1 = New Point(startX,
                           startY - Theme.kHeight - 10)

            p2 = New Point(endX,
                           startY + 10)

            bright = column.getBright(alpha)
            dark = column.getDark(alpha)

            If sNote.Landmine Then
                bright = Color.Red
                dark = Color.Red
            End If

            xBrush2 = New SolidBrush(column.cText)
        Else
            p1 = New Point(HPositionToPanelX(colX - 0.5 * colWidth),
                           endY - Theme.kHeight)
            p2 = New Point(HPositionToPanelX(colX + 1.5 * colWidth),
                           startY)

            bright = column.getLongBright(alpha)
            dark = column.getLongDark(alpha)

            xBrush2 = New SolidBrush(column.cLText)
        End If

        xPen1 = New Pen(bright)
        xBrush = New Drawing2D.LinearGradientBrush(p1, p2, bright, dark)

        Dim colDrawWidth = Columns.GetColumnWidth(sNote.ColumnIndex) * Editor.Grid.WidthScale
        ' Note gradient
        Buffer.FillRectangle(xBrush,
                                     startX + 1,
                                     endY - Theme.kHeight + 1,
                                     colDrawWidth - 1,
                                     CInt(sNote.Length * Editor.Grid.HeightScale) + Theme.kHeight - 1)

        ' Outline
        Buffer.DrawRectangle(xPen1,
                                      startX + 1,
                                      endY - Theme.kHeight,
                                      colDrawWidth - 3,
                                      CInt(sNote.Length * Editor.Grid.HeightScale) + Theme.kHeight)

        ' Note B36
        Dim label = IIf(Columns.IsColumnNumeric(sNote.ColumnIndex), sNote.Value / 10000, xLabel)
        Buffer.DrawString(label,
                                  Theme.kFont, xBrush2,
                                  startX + Theme.kLabelHShiftL - 2,
                                  startY - Theme.kHeight + Theme.kLabelVShift)

        ' Draw paired body
        If sNote.ColumnIndex < ColumnType.BGM Then
            If sNote.Length = 0 And sNote.LNPair <> 0 Then
                DrawPairedLNBody(sNote, alpha)
            End If
        End If


        ' Select Box
        If sNote.Selected Then
            Buffer.DrawRectangle(Theme.kSelected,
                                    startX,
                                    endY - Theme.kHeight - 1,
                                    colWidth * Editor.Grid.WidthScale,
                                    CInt(sNote.Length * Editor.Grid.HeightScale) + Theme.kHeight + 2)
        End If

        ' Errors
        If Editor.ErrorCheck AndAlso sNote.HasError Then
            Dim sx = HPositionToPanelX(colX + colWidth / 2)
            Dim sy = CInt(VPositionToPanelY(sNote.VPosition) - Theme.kHeight / 2)
            Buffer.DrawImage(My.Resources.ImageError,
                                      sx - 12,
                                      sy - 12,
                                      24, 24)
        End If


    End Sub

    Private Function GetNoteRectangle(note As Note) As Rectangle
        Dim xDispX As Integer = HPositionToPanelX(Editor.Columns.GetColumnLeft(note.ColumnIndex))

        Dim xDispY As Integer = IIf(Not Editor.NTInput Or (Editor.State.NT.IsAdjustingNoteLength And Not Editor.State.NT.IsAdjustingUpperEnd),
                                    VPositionToPanelY(note.VPosition) - Theme.kHeight - 1,
                                    VPositionToPanelY(note.VPosition + note.Length) - Theme.kHeight - 1)

        Dim xDispW As Integer = Editor.Columns.GetColumnWidth(note.ColumnIndex) * Editor.Grid.WidthScale + 1
        Dim xDispH As Integer = IIf(Not Editor.NTInput Or Editor.State.NT.IsAdjustingNoteLength,
                                    Theme.kHeight + 3,
                                    note.Length * Editor.Grid.WidthScale + Theme.kHeight + 3)

        Return New Rectangle(xDispX, xDispY, xDispW, xDispH)
    End Function


    Private Sub DrawMouseOver()
        If Editor.NTInput Then
            If Not Editor.State.NT.IsAdjustingNoteLength Then
                DrawNoteNT(Editor.CurrentMouseoverNote)
            End If
        Else
            DrawNote(Editor.CurrentMouseoverNote)
        End If

        Dim rect = GetNoteRectangle(Editor.CurrentMouseoverNote)
        Dim pen = IIf(Editor.State.NT.IsAdjustingNoteLength, Theme.kMouseOverE, Theme.kMouseOver)
        Buffer.DrawRectangle(pen, rect.X, rect.Y, rect.Width - 1, rect.Height - 1)

        If ModifierMultiselectActive() Then
            Dim renderNotes = Editor.Notes.
                                     Where(Function(x) IsNoteVisible(x)).
                                     Where(Function(x) Editor.IsLabelMatch(x, Editor.CurrentMouseoverNote))
            For Each note In renderNotes
                Dim nrect = GetNoteRectangle(note)
                Buffer.DrawRectangle(pen, nrect.X, nrect.Y, nrect.Width - 1, nrect.Height - 1)
            Next
        End If

    End Sub

    Private Sub DrawTempNote()
        Dim xValue As Integer = (Editor.LWAV.SelectedIndex + 1) * 10000

        Dim alpha As Single = 1.0F
        If ModifierHiddenActive() Then
            alpha = Theme.kOpacity
        End If

        Dim column = Editor.Columns.GetColumn(Editor.State.Mouse.CurrentMouseColumn)
        Dim xText As String = C10to36(xValue \ 10000)
        If Editor.Columns.IsColumnNumeric(Editor.State.Mouse.CurrentMouseColumn) Then
            xText = column.Title
        End If

        Dim xPen As Pen
        Dim xBrush As Drawing2D.LinearGradientBrush
        Dim xBrush2 As SolidBrush

        Dim colX = Editor.Columns.GetColumnLeft(Editor.State.Mouse.CurrentMouseColumn)
        Dim colStartDrawX = HPositionToPanelX(colX)
        Dim colWidth = Editor.Columns.GetColumnWidth(Editor.State.Mouse.CurrentMouseColumn)
        Dim colEndDrawX = HPositionToPanelX(colX + colWidth)
        Dim colRightDrawX = HPositionToPanelX(colX + colWidth)
        Dim startY = VPositionToPanelY(Editor.State.Mouse.CurrentMouseRow)

        Dim point1 As New Point(colStartDrawX,
                                startY - Theme.kHeight - 10)
        Dim point2 As New Point(colEndDrawX,
                                startY + 10)

        Dim bright As Color
        Dim dark As Color
        If Editor.NTInput Or Not ModifierLongNoteActive() Then
            xPen = New Pen(column.getBright(alpha))
            bright = column.getBright(alpha)
            dark = column.getDark(alpha)

            xBrush2 = New SolidBrush(column.cText)
        Else
            xPen = New Pen(column.getLongBright(alpha))
            bright = column.getLongBright(alpha)
            dark = column.getLongDark(alpha)

            xBrush2 = New SolidBrush(column.cLText)
        End If

        ' Temp landmine
        If ModifierLandmineActive() Then
            bright = Color.Red
            dark = Color.Red
        End If

        xBrush = New Drawing2D.LinearGradientBrush(point1, point2, bright, dark)

        Dim colDrawWidth = Editor.Columns.GetColumnWidth(Editor.State.Mouse.CurrentMouseColumn) * Editor.Grid.WidthScale

        Buffer.FillRectangle(xBrush,
                             colStartDrawX + 2,
                             startY - Theme.kHeight + 1,
                             colDrawWidth - 3,
                             Theme.kHeight - 1)
        Buffer.DrawRectangle(xPen,
                                  colStartDrawX + 1,
                                  startY - Theme.kHeight,
                                  colDrawWidth - 2,
                                  Theme.kHeight)

        Buffer.DrawString(xText, Theme.kFont, xBrush2,
                          colX + Theme.kLabelHShiftL - 2,
                          startY - Theme.kHeight + Theme.kLabelVShift)
    End Sub

    Private Sub DrawTimeSelection()
        Dim xBPMStart = Editor.FirstBpm
        Dim xBPMHalf = Editor.FirstBpm
        Dim xBPMEnd = Editor.FirstBpm

        For Each note In Editor.Notes
            If note.ColumnIndex = ColumnType.BPM Then
                If note.VPosition <= Editor.State.TimeSelect.StartPoint Then
                    xBPMStart = note.Value
                End If
                If note.VPosition <= Editor.State.TimeSelect.HalfPoint Then
                    xBPMHalf = note.Value
                End If
                If note.VPosition <= Editor.State.TimeSelect.EndPoint Then
                    xBPMEnd = note.Value
                End If
            End If
            If note.VPosition > Editor.State.TimeSelect.EndPoint Then
                Exit For
            End If
        Next

        Dim bpmColStartX = Editor.Columns.GetColumnLeft(ColumnType.BPM)
        Dim startDrawX = HPositionToPanelX(bpmColStartX)
        Dim halfLineY = VPositionToPanelY(Editor.State.TimeSelect.StartPoint + Editor.State.TimeSelect.HalfPointLength)
        Dim endLineY = VPositionToPanelY(Editor.State.TimeSelect.StartPoint + Editor.State.TimeSelect.EndPointLength)
        Dim selectionDrawHeight = CInt(Math.Abs(Editor.State.TimeSelect.EndPointLength) * Editor.Grid.HeightScale)
        Dim startLineY = VPositionToPanelY(Editor.State.TimeSelect.StartPoint + Math.Max(Editor.State.TimeSelect.EndPointLength, 0))
        Dim startLineStringY = VPositionToPanelY(Editor.State.TimeSelect.StartPoint)

        'Selection area
        Buffer.FillRectangle(Theme.PESel,
                                  0, startLineY,
                                  Size.Width,
                                  selectionDrawHeight)
        'End Cursor
        Buffer.DrawLine(Theme.PECursor,
                        0, endLineY,
                        Size.Width, endLineY)
        'Half Cursor
        Buffer.DrawLine(Theme.PEHalf,
                        0, halfLineY,
                        Size.Width, halfLineY)
        'Start BPM
        Buffer.DrawString(xBPMStart / 10000,
                        Theme.PEBPMFont, Theme.PEBPM,
                        startDrawX,
                        startLineStringY - Theme.PEBPMFont.Height + 3)
        'Half BPM
        Buffer.DrawString(xBPMHalf / 10000,
                        Theme.PEBPMFont, Theme.PEBPM,
                        startDrawX,
                        halfLineY - Theme.PEBPMFont.Height + 3)
        'End BPM
        Buffer.DrawString(xBPMEnd / 10000,
                        Theme.PEBPMFont, Theme.PEBPM,
                        startDrawX,
                        endLineY - Theme.PEBPMFont.Height + 3)

        'SelLine
        If Editor.State.TimeSelect.MouseOverLine = 1 Then 'Start Cursor
            Buffer.DrawRectangle(Theme.PEMouseOver,
                                0, startLineStringY - 1,
                                Size.Width - 1, 2)
        ElseIf Editor.State.TimeSelect.MouseOverLine = 2 Then 'Half Cursor
            Buffer.DrawRectangle(Theme.PEMouseOver,
                                0, halfLineY - 1,
                                Size.Width - 1, 2)
        ElseIf Editor.State.TimeSelect.MouseOverLine = 3 Then 'End Cursor
            Buffer.DrawRectangle(Theme.PEMouseOver,
                                0, endLineY - 1,
                                Size.Width - 1, 2)
        End If
    End Sub

    Private Sub DrawClickAndScroll()
        Dim xDeltaLocation As Point = PointToScreen(New Point(0, 0))

        Dim xInitX As Single = Editor.State.Mouse.MiddleButtonLocation.X - xDeltaLocation.X
        Dim xInitY As Single = Editor.State.Mouse.MiddleButtonLocation.Y - xDeltaLocation.Y
        Dim xCurrX As Single = Cursor.Position.X - xDeltaLocation.X
        Dim xCurrY As Single = Cursor.Position.Y - xDeltaLocation.Y
        Dim xAngle As Double = Math.Atan2(xCurrY - xInitY, xCurrX - xInitX)
        Buffer.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        If Not (xInitX = xCurrX And xInitY = xCurrY) Then
            Dim xPointx() As PointF = {New PointF(xCurrX, xCurrY),
                                       New PointF(Math.Cos(xAngle + Math.PI / 2) * 10 + xInitX, Math.Sin(xAngle + Math.PI / 2) * 10 + xInitY),
                                       New PointF(Math.Cos(xAngle - Math.PI / 2) * 10 + xInitX, Math.Sin(xAngle - Math.PI / 2) * 10 + xInitY)}
            Buffer.FillPolygon(New Drawing2D.LinearGradientBrush(New Point(xInitX, xInitY), New Point(xCurrX, xCurrY), Color.FromArgb(0), Color.FromArgb(-1)), xPointx)
        End If

        Buffer.FillEllipse(Brushes.LightGray, xInitX - 10, xInitY - 10, 20, 20)
        Buffer.DrawEllipse(Pens.Black, xInitX - 8, xInitY - 8, 16, 16)

        Buffer.SmoothingMode = Drawing2D.SmoothingMode.Default

    End Sub


    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' TODO: move this
        ' If Me.WindowState = FormWindowState.Minimized Then Return
        If DisplayRectangle.Width <= 0 Or DisplayRectangle.Height <= 0 Then Return

        ' MyBase.OnPaint(e)
        If DoubleBuffer Is Nothing Then
            DoubleBuffer = BufferedGraphicsManager.Current.Allocate(e.Graphics, DisplayRectangle)
        End If

        Buffer = DoubleBuffer.Graphics

        If Theme Is Nothing Then
            e.Graphics.FillRectangle(New SolidBrush(BackColor), DisplayRectangle)
            e.Graphics.DrawString("Editor Pane", Font, Brushes.White, 0, 0)
            Exit Sub
        End If

        Buffer.Clear(Theme.Bg.Color)

        Dim xVSu As Integer = Math.Min(-VerticalScroll + Size.Height / Editor.Grid.HeightScale, Editor.GetMaxVPosition())

        'Bg color
        If Editor.Grid.ShowBackground Then
            DrawBackgroundColor()
        End If



        DrawPanelLines(xVSu)

        'Column Caption
        If Editor.Grid.ShowColumnCaptions Then
            DrawColumnCaptions()
        End If

        'WaveForm
        DrawWaveform(-VerticalScroll)

        'K
        'If Not K Is Nothing Then
        DrawNotes()

        'End If

        'Selection Box
        DrawSelectionBox()

        'Mouse Over
        If Editor.IsSelectMode AndAlso Editor.State.Mouse.CurrentHoveredNoteIndex <> -1 Then
            ' DrawNoteHoverHighlight(iI, HorizontalScroll, VerticalScroll, panelHeight, foundNoteIndex)
            DrawMouseOver()
        End If

        If Editor.ShouldDrawTempNote AndAlso
           Editor.State.Mouse.CurrentMouseColumn > -1 AndAlso
           Editor.State.Mouse.CurrentMouseRow > -1 Then
            DrawTempNote()
        End If

        'Time Selection
        If Editor.IsTimeSelectMode Then
            DrawTimeSelection()
        End If

        'Middle button: CLick and Scroll
        If Editor.State.Mouse.MiddleButtonClicked Then
            DrawClickAndScroll()
        End If

        'Drag/Drop
        DrawDragAndDrop()

        DoubleBuffer.Render(e.Graphics)
        ' MyBase.OnPaint(e)
    End Sub

    ''' <summary>
    ''' Transforms a position from absolute panel coordinates to screen panel coordinates
    ''' </summary>
    ''' <param name="xHPosition">Original horizontal position.</param>
    Private Function HPositionToPanelX(ByVal xHPosition As Integer) As Integer
        Return (xHPosition - HorizontalScroll) * Editor.Grid.WidthScale
    End Function

    ''' <summary>
    ''' Transform a position from absolute note row (192 parts per 4/4 measure) to a Y position in screen panel coordinates.
    ''' </summary>
    ''' <param name="xVPosition">Original vertical position.</param>
    Public Function VPositionToPanelY(ByVal xVPosition As Double) As Integer
        Return Size.Height - CInt((xVPosition + VerticalScroll) * Editor.Grid.HeightScale) - 1
    End Function

    Public Function SnapToGrid(ByVal xVPos As Double) As Double
        Dim xOffset As Double = Editor.MeasureBottom(Editor.MeasureAtDisplacement(xVPos))
        Dim xRatio As Double = 192.0R / Editor.Grid.Divider
        Return Math.Floor((xVPos - xOffset) / xRatio) * xRatio + xOffset
    End Function

End Class
