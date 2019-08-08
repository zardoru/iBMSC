Imports System.Drawing.Drawing2D
Imports System.Linq
Imports iBMSC.Editor

Public Class EditorPanel
    Inherits Panel

    Public ReadOnly Property HorizontalPosition As Integer
        Get
            Return HorizontalScrollBar.Value
        End Get
    End Property

    Public ReadOnly Property VerticalPosition As Integer
        Get
            Return VerticalScrollBar.Value
        End Get
    End Property

    Private _buffer As Graphics
    Private _theme As VisualSettings
    Private _editor As MainWindow
    Public WithEvents VerticalScrollBar As VScrollBar
    Public WithEvents HorizontalScrollBar As HScrollBar
    Public LastVerticalScroll As Integer
    Public LastHorizontalScroll As Integer

    Private Sub HsValueChanged(sender As Object, e As EventArgs) Handles HorizontalScrollBar.ValueChanged
        LastHorizontalScroll = sender.Value
        Refresh()
    End Sub

    Private Sub VsValueChanged(sender As Object, e As EventArgs) Handles VerticalScrollBar.ValueChanged
        LastVerticalScroll = sender.Value
        Refresh()
    End Sub

    Public Sub New()
        MyBase.New()
        HorizontalScrollBar = New HScrollBar()
        VerticalScrollBar = New VScrollBar()

        HorizontalScrollBar.Dock = DockStyle.Bottom
        VerticalScrollBar.Dock = DockStyle.Right

        Controls.Add(HorizontalScrollBar)
        Controls.Add(VerticalScrollBar)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

        VerticalScrollBar.Minimum = -10000
        VerticalScrollBar.Maximum = 500
    End Sub

    Public Sub New(paren As Control)
        Parent = paren
    End Sub

    Public Sub Init(wnd As MainWindow, theme As VisualSettings)

        _theme = theme
        _editor = wnd

    End Sub

    Private Sub DrawDragAndDrop()
        If _editor.DragDropFilename.Length > 0 Then
            Dim filenames = Join(_editor.DragDropFilename, vbCrLf)
            Dim brush As New SolidBrush(Color.FromArgb(&HC0FFFFFF))
            Dim format As New StringFormat With {
                    .Alignment = StringAlignment.Center,
                    .LineAlignment = StringAlignment.Center
            }
            _buffer.DrawString(filenames, _editor.Font, brush, DisplayRectangle, format)
        End If
    End Sub

    Private Function GetColumnHighlightColor(col As Color, Optional factor As Double = 2.0)
        Dim clamp = Function(x) IIf(x > 255, 255, x)
        Return Color.FromArgb(
            clamp(col.A * factor),
            clamp(col.R * factor),
            clamp(col.G * factor),
            clamp(col.B * factor))
    End Function

    Public Function GetColumnAtX(x As Integer) As Integer
        Dim i = 0
        'horizontal position of the mouse
        Dim mLeft As Integer = x / _editor.Grid.WidthScale + HorizontalPosition
        Dim col = 0
        If mLeft >= 0 Then
            Do
                'get the column where mouse is 
                If mLeft < _editor.Columns.GetColumnLeft(i + 1) Or i >= _editor.Columns.ColumnCount Then col = i : Exit Do
                i += 1
            Loop
        End If

        'get the enabled column where mouse is 
        Return _editor.Columns.NormalizeIndex(col)
    End Function

    Private Sub DrawBackgroundColor()
        Dim grid = _editor.Grid
        Dim columns = _editor.Columns

        For i = 0 To columns.ColumnCount
            If HPositionToPanelX(columns.GetColumnLeft(i + 1)) + 1 < 0 Then
                Continue For
            End If

            If (columns.GetColumnLeft(i) - HorizontalPosition) * grid.WidthScale + 1 > Size.Width Then
                Exit For
            End If

            If Not columns.GetColumn(i).cBG.GetBrightness = 0 And columns.GetWidth(i) > 0 Then
                Dim col = columns.GetColumn(i).cBG
                If i = GetColumnAtX(_editor.State.Mouse.MouseMoveStatus.X) Then
                    Dim bf = 1.2
                    col = GetColumnHighlightColor(col)
                End If
                Dim brush = New SolidBrush(col)

                _buffer.FillRectangle(brush,
                                     HPositionToPanelX(columns.GetColumnLeft(i)) + 1,
                                     0,
                                     columns.GetWidth(i) * _editor.Grid.WidthScale,
                                     Size.Height)
            End If
        Next
    End Sub

    Private Function IsNoteVisible(note As Note) As Boolean
        Dim xUpperBorder As Single = Math.Abs(VerticalPosition) + Size.Height / _editor.Grid.HeightScale
        Dim xLowerBorder As Single = Math.Abs(VerticalPosition) - _theme.NoteHeight / _editor.Grid.HeightScale

        Dim aboveLower = note.VPosition >= xLowerBorder
        Dim headBelow = note.VPosition <= xLowerBorder
        Dim tailAbove = note.VPosition + note.Length >= xLowerBorder
        Dim intersectsNt = headBelow And tailAbove
        Dim intersects = (note.VPosition <= xLowerBorder And
                          _editor.Notes(note.LNPair).VPosition >= xLowerBorder)
        Dim aboveUpper = note.VPosition > xUpperBorder

        Dim noteInside = (Not aboveUpper) And aboveLower

        Return noteInside OrElse intersectsNt OrElse intersects
    End Function

    Private Function IsNoteVisible(noteIndex As Integer) As Boolean
        Return IsNoteVisible(_editor.Notes(noteIndex))
    End Function

    Private Sub DrawGridLines(measureIndex As Integer,
                              divisions As Integer, pen As Pen)
        Dim row = 0
        Dim xUpper As Double = _editor.MeasureUpper(measureIndex)
        Dim xCurr = _editor.MeasureBottom(measureIndex)
        Dim rowSize = 192 / divisions
        Do While xCurr < xUpper
            Dim y = VPositionToPanelY(xCurr)
            _buffer.DrawLine(pen,
                            0, y,
                            Size.Width, y)
            row += 1
            xCurr = _editor.MeasureBottom(measureIndex) + row * rowSize
        Loop
    End Sub

    Private Sub DrawPanelLines(xVSu As Integer)
        'Vertical line
        If _editor.Grid.ShowVerticalLines Then
            For i = 0 To _editor.Columns.ColumnCount
                Dim xpos = (_editor.Columns.GetColumnLeft(i) - HorizontalPosition) * _editor.Grid.WidthScale
                If xpos + 1 < 0 Then Continue For
                If xpos + 1 > Size.Width Then Exit For
                If _editor.Columns.GetWidth(i) > 0 Then
                    _buffer.DrawLine(_theme.pVLine,
                                    xpos, 0,
                                    xpos, Size.Height)
                End If
            Next
        End If

        'Grid, Sub, Measure
        For Measure = _editor.MeasureAtDisplacement(-VerticalPosition) To _editor.MeasureAtDisplacement(xVSu)
            'grid
            If _editor.Grid.ShowMainGrid Then
                DrawGridLines(Measure, _editor.Grid.Divider, _theme.pGrid)
            End If

            'sub
            If _editor.Grid.ShowSubGrid Then
                DrawGridLines(Measure, _editor.Grid.Subdivider, _theme.pSub)
            End If


            'measure and measurebar
            Dim xCurr = _editor.MeasureBottom(Measure)
            Dim pHeight = VPositionToPanelY(xCurr)
            If _editor.Grid.ShowMeasureBars Then
                _buffer.DrawLine(_theme.pMLine,
                                0, pHeight,
                                Size.Width, pHeight)
            End If
            If _editor.Grid.ShowMeasureNumber Then
                Dim brush = New SolidBrush(_editor.Columns.GetColumn(0).cText)
                _buffer.DrawString("[" & Add3Zeros(Measure).ToString & "]",
                                  _theme.kMFont, brush, -HorizontalPosition * _editor.Grid.WidthScale,
                                  pHeight - _theme.kMFont.Height)
            End If
        Next

        Dim vpos = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)
        Dim mouseLineHeight = VPositionToPanelY(vpos)
        Dim p = New Pen(Color.White)
        _buffer.DrawLine(p, 0, mouseLineHeight, Size.Width, mouseLineHeight)
    End Sub

    Private Sub DrawColumnCaptions()
        Dim columns = _editor.Columns
        For i = 0 To _editor.Columns.ColumnCount
            Dim colEnd = (columns.GetColumnLeft(i + 1) - HorizontalPosition) * _editor.Grid.WidthScale
            If colEnd + 1 < 0 Then
                Continue For
            End If

            Dim colStart = (columns.GetColumnLeft(i) - HorizontalPosition) * _editor.Grid.WidthScale
            If colStart + 1 > Size.Width Then
                Exit For
            End If

            If columns.GetWidth(i) > 0 Then
                _buffer.DrawString(
                    columns.GetName(i),
                    _theme.ColumnTitleFont,
                    _theme.ColumnTitle,
                    colStart, 0)
            End If
        Next
    End Sub

    Private Sub DrawSelectionBox()
        Dim pMouseMove = _editor.State.Mouse.pMouseMove
        Dim lastMouseDownLocation = _editor.State.Mouse.LastMouseDownLocation
        If _editor.IsSelectMode AndAlso
           (_editor.FocusedPanel Is Me) AndAlso
           Not (pMouseMove = New Point(-1, -1) Or
                lastMouseDownLocation = New Point(-1, -1)) Then
            Dim mdx = pMouseMove.X - lastMouseDownLocation.X
            Dim mdy = pMouseMove.Y - lastMouseDownLocation.Y
            _buffer.DrawRectangle(_theme.SelBox,
                                 Math.Min(lastMouseDownLocation.X, pMouseMove.X),
                                 Math.Min(lastMouseDownLocation.Y, pMouseMove.Y),
                                 Math.Abs(mdx), Math.Abs(mdy))
        End If
    End Sub

    Private Sub DrawWaveform(xVsr As Integer)
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
        Dim upperBorder = Math.Abs(VerticalPosition) + Size.Height / _editor.Grid.HeightScale
        Dim lowerBorder = Math.Abs(VerticalPosition) - _theme.NoteHeight / _editor.Grid.HeightScale

        Dim renderNotes = _editor.Notes.
                Where(Function(x) x.VPosition <= upperBorder).
                Where(Function(x) IsNoteVisible(x)).
                Where(Function(x) _editor.Columns.IsEnabled(x.ColumnIndex))

        For Each Note In renderNotes
            If _editor.NtInput Then
                DrawNoteNt(Note)
            Else
                DrawNote(Note)
            End If
        Next
    End Sub

    ''' <summary>
    '''     Draw a note in a buffer.
    ''' </summary>
    ''' <param name="sNote">Note to be drawn.</param>

    Private Sub DrawNote(sNote As Note)
        Dim alpha = 1.0F
        If sNote.Hidden Then alpha = _theme.kOpacity

        Dim xLabel As String = C10to36(sNote.Value \ 10000)
        If _editor.ShowFileName Then
            If _editor.Columns.IsColumnSound(sNote.ColumnIndex) Then
                If _editor.BmsWAV(C36to10(xLabel)) <> "" Then
                    xLabel = Path.GetFileNameWithoutExtension(_editor.BmsWAV(C36to10(xLabel)))
                End If
            Else
                If _editor.BmsBMP(C36to10(xLabel)) <> "" Then
                    xLabel = Path.GetFileNameWithoutExtension(_editor.BmsBMP(C36to10(xLabel)))
                End If
            End If
        End If

        Dim xPen As Pen
        Dim xBrush As LinearGradientBrush
        Dim xBrush2 As SolidBrush

        Dim bright As Color
        Dim dark As Color

        Dim grid = _editor.Grid
        Dim columns = _editor.Columns
        Dim colX = columns.GetColumnLeft(sNote.ColumnIndex)
        Dim colWidth = columns.GetWidth(sNote.ColumnIndex)
        Dim column = columns.GetColumn(sNote.ColumnIndex)
        Dim startX = HPositionToPanelX(colX)
        Dim startY = VPositionToPanelY(sNote.VPosition) - _theme.NoteHeight

        Dim endX = HPositionToPanelX(colX + colWidth)
        Dim endY = VPositionToPanelY(sNote.VPosition)

        Dim p1 = New Point(startX,
                           startY - 10)
        Dim p2 = New Point(endX,
                           endY + 10)

        If Not sNote.LongNote Then
            xPen = New Pen(column.GetBright(alpha))

            bright = column.GetBright(alpha)
            dark = column.GetDark(alpha)

            If sNote.Landmine Then
                bright = Color.Red
                dark = Color.Red
            End If

            xBrush2 = New SolidBrush(column.cText)
        Else
            bright = column.GetLongBright(alpha)
            dark = column.GetLongDark(alpha)

            xBrush2 = New SolidBrush(column.cLText)
        End If

        xPen = New Pen(bright)
        xBrush = New LinearGradientBrush(p1, p2, bright, dark)

        Dim colDrawWidth = colWidth * _editor.Grid.WidthScale

        ' Fill
        _buffer.FillRectangle(xBrush,
                             startX + 2,
                             startY + 1,
                             colDrawWidth - 3,
                             _theme.NoteHeight - 1)
        ' Outline
        _buffer.DrawRectangle(xPen,
                             startX + 1,
                             startY,
                             colDrawWidth - 2,
                             _theme.NoteHeight)

        ' Label
        Dim label = IIf(columns.IsColumnNumeric(sNote.ColumnIndex), sNote.Value / 10000, xLabel)
        _buffer.DrawString(label,
                          _theme.kFont, xBrush2,
                          startX + _theme.kLabelHShift,
                          startY + _theme.kLabelVShift)

        If sNote.ColumnIndex < ColumnType.BGM Then
            If sNote.LNPair <> 0 Then
                DrawPairedLnBody(sNote, alpha)
            End If
        End If


        If _editor.ErrorCheck AndAlso sNote.HasError Then
            Dim ex = HPositionToPanelX(colX + colWidth / 2)
            Dim ey = VPositionToPanelY(sNote.VPosition) - _theme.NoteHeight / 2
            _buffer.DrawImage(My.Resources.ImageError,
                             CInt(ex - 12),
                             CInt(ey - 12),
                             24, 24)
        End If

        If sNote.Selected Then
            _buffer.DrawRectangle(
                _theme.kSelected,
                colX, VPositionToPanelY(sNote.VPosition) - _theme.NoteHeight - 1,
                colDrawWidth, _theme.NoteHeight + 2)
        End If
    End Sub

    Private Sub DrawPairedLnBody(sNote As Note, xAlpha As Single)
        Dim column = _editor.Columns.GetColumn(sNote.ColumnIndex)

        Dim xPen2 As New Pen(column.GetLongBright(xAlpha))

        Dim colX = _editor.Columns.GetColumnLeft(sNote.ColumnIndex)
        Dim colWidth = _editor.Columns.GetWidth(sNote.ColumnIndex)

        Dim xBrush3 As New LinearGradientBrush(
            New Point(HPositionToPanelX(colX - 0.5 * colWidth),
                      VPositionToPanelY(_editor.Notes(sNote.LNPair).VPosition)),
            New Point(HPositionToPanelX(colX + 1.5 * colWidth),
                      VPositionToPanelY(sNote.VPosition) + _theme.NoteHeight),
            column.GetLongBright(xAlpha),
            column.GetLongDark(xAlpha))

        ' assume ln pair column = our column
        Dim x1 = HPositionToPanelX(colX)
        Dim y1 = VPositionToPanelY(_editor.Notes(sNote.LNPair).VPosition)
        Dim colDrawWidth = colWidth * _editor.Grid.WidthScale ' GetColumnWidth(Notes(sNote.LNPair).ColumnIndex) * GxWidth1
        Dim bodyHeight = VPositionToPanelY(sNote.VPosition) - VPositionToPanelY(_editor.Notes(sNote.LNPair).VPosition)
        _buffer.FillRectangle(xBrush3,
                             x1 + 3,
                             y1 + 1,
                             colDrawWidth - 5, ' az: ???
                             bodyHeight - _theme.NoteHeight - 1)
        _buffer.DrawRectangle(xPen2,
                             x1 + 2,
                             y1,
                             colDrawWidth - 4,
                             bodyHeight - _theme.NoteHeight)
    End Sub

    ''' <summary>
    '''     Draw a note in a buffer under NT mode.
    ''' </summary>
    ''' <param name="sNote">Note to be drawn.</param>
    Private Sub DrawNoteNt(sNote As Note)

        Dim alpha = 1.0F
        If sNote.Hidden Then alpha = _theme.kOpacity

        Dim columns = _editor.Columns
        Dim column = _editor.Columns.GetColumn(sNote.ColumnIndex)
        Dim xLabel As String = C10to36(sNote.Value \ 10000)
        If _editor.ShowFileName Then
            If columns.IsColumnSound(sNote.ColumnIndex) Then
                If _editor.BmsWAV(C36to10(xLabel)) <> "" Then
                    xLabel = Path.GetFileNameWithoutExtension(_editor.BmsWAV(C36to10(xLabel)))
                End If
            Else
                If _editor.BmsBMP(C36to10(xLabel)) <> "" Then
                    xLabel = Path.GetFileNameWithoutExtension(_editor.BmsBMP(C36to10(xLabel)))
                End If
            End If
        End If

        Dim xPen1 As Pen
        Dim xBrush As LinearGradientBrush
        Dim xBrush2 As SolidBrush

        Dim p1 As Point
        Dim p2 As Point
        Dim bright As Color
        Dim dark As Color

        Dim colX = columns.GetColumnLeft(sNote.ColumnIndex)
        Dim colWidth = columns.GetWidth(sNote.ColumnIndex)
        Dim startX = HPositionToPanelX(colX)
        Dim endX = HPositionToPanelX(colX + colWidth)
        Dim startY = VPositionToPanelY(sNote.VPosition)
        Dim endY = VPositionToPanelY(sNote.VPosition + sNote.Length)

        If sNote.Length = 0 Then
            p1 = New Point(startX,
                           startY - _theme.NoteHeight - 10)

            p2 = New Point(endX,
                           startY + 10)

            bright = column.GetBright(alpha)
            dark = column.GetDark(alpha)

            If sNote.Landmine Then
                bright = Color.Red
                dark = Color.Red
            End If

            xBrush2 = New SolidBrush(column.cText)
        Else
            p1 = New Point(HPositionToPanelX(colX - 0.5 * colWidth),
                           endY - _theme.NoteHeight)
            p2 = New Point(HPositionToPanelX(colX + 1.5 * colWidth),
                           startY)

            bright = column.GetLongBright(alpha)
            dark = column.GetLongDark(alpha)

            xBrush2 = New SolidBrush(column.cLText)
        End If

        xPen1 = New Pen(bright)
        xBrush = New LinearGradientBrush(p1, p2, bright, dark)

        Dim colDrawWidth = columns.GetWidth(sNote.ColumnIndex) * _editor.Grid.WidthScale
        ' Note gradient
        _buffer.FillRectangle(xBrush,
                             startX + 1,
                             endY - _theme.NoteHeight + 1,
                             colDrawWidth - 1,
                             CInt(sNote.Length * _editor.Grid.HeightScale) + _theme.NoteHeight - 1)

        ' Outline
        _buffer.DrawRectangle(xPen1,
                             startX + 1,
                             endY - _theme.NoteHeight,
                             colDrawWidth - 3,
                             CInt(sNote.Length * _editor.Grid.HeightScale) + _theme.NoteHeight)

        ' Note B36
        Dim label = IIf(columns.IsColumnNumeric(sNote.ColumnIndex), sNote.Value / 10000, xLabel)
        _buffer.DrawString(label,
                          _theme.kFont, xBrush2,
                          startX + _theme.kLabelHShiftL - 2,
                          startY - _theme.NoteHeight + _theme.kLabelVShift)

        ' Draw paired body
        If sNote.ColumnIndex < ColumnType.BGM Then
            If sNote.Length = 0 And sNote.LNPair <> 0 Then
                DrawPairedLnBody(sNote, alpha)
            End If
        End If


        ' Select Box
        If sNote.Selected Then
            _buffer.DrawRectangle(_theme.kSelected,
                                 startX,
                                 endY - _theme.NoteHeight - 1,
                                 colWidth * _editor.Grid.WidthScale,
                                 CInt(sNote.Length * _editor.Grid.HeightScale) + _theme.NoteHeight + 2)
        End If

        ' Errors
        If _editor.ErrorCheck AndAlso sNote.HasError Then
            Dim sx = HPositionToPanelX(colX + colWidth / 2)
            Dim sy = CInt(VPositionToPanelY(sNote.VPosition) - _theme.NoteHeight / 2)
            _buffer.DrawImage(My.Resources.ImageError,
                             sx - 12,
                             sy - 12,
                             24, 24)
        End If
    End Sub

    Private Function GetNoteRectangle(note As Note) As Rectangle
        Dim colLeft As Integer = HPositionToPanelX(_editor.Columns.GetColumnLeft(note.ColumnIndex))

        Dim noteY As Integer =
                IIf(
                    Not _editor.NtInput Or
                    (_editor.State.NT.IsAdjustingNoteLength And Not _editor.State.NT.IsAdjustingUpperEnd),
                    VPositionToPanelY(note.VPosition) - _theme.NoteHeight - 1,
                    VPositionToPanelY(note.VPosition + note.Length) - _theme.NoteHeight - 1)

        Dim colWidth As Integer = _editor.Columns.GetWidth(note.ColumnIndex) * _editor.Grid.WidthScale + 1
        Dim noteHeight As Integer = IIf(Not _editor.NtInput Or _editor.State.NT.IsAdjustingNoteLength,
                                    _theme.NoteHeight + 3,
                                    note.Length * _editor.Grid.WidthScale + _theme.NoteHeight + 3)

        Return New Rectangle(colLeft, noteY, colWidth, noteHeight)
    End Function


    Private Sub DrawMouseOver()
        If _editor.NtInput Then
            If Not _editor.State.NT.IsAdjustingNoteLength Then
                DrawNoteNt(_editor.CurrentMouseoverNote)
            End If
        Else
            DrawNote(_editor.CurrentMouseoverNote)
        End If

        Dim rect = GetNoteRectangle(_editor.CurrentMouseoverNote)
        Dim pen = IIf(_editor.State.NT.IsAdjustingNoteLength, _theme.kMouseOverE, _theme.kMouseOver)
        _buffer.DrawRectangle(pen, rect.X, rect.Y, rect.Width - 1, rect.Height - 1)

        If ModifierMultiselectActive() Then
            Dim renderNotes = _editor.Notes.
                    Where(Function(x) IsNoteVisible(x)).
                    Where(Function(x) _editor.IsLabelMatch(x, _editor.CurrentMouseoverNote))
            For Each note In renderNotes
                Dim nrect = GetNoteRectangle(note)
                _buffer.DrawRectangle(pen, nrect.X, nrect.Y, nrect.Width - 1, nrect.Height - 1)
            Next
        End If
    End Sub

    Private Sub DrawTempNote()
        Dim xValue As Integer = (_editor.LWAV.SelectedIndex + 1) * 10000

        Dim alpha = 1.0F
        If ModifierHiddenActive() Then
            alpha = _theme.kOpacity
        End If

        Dim column = _editor.Columns.GetColumn(_editor.State.Mouse.CurrentMouseColumn)
        Dim xText As String = C10to36(xValue \ 10000)
        If _editor.Columns.IsColumnNumeric(_editor.State.Mouse.CurrentMouseColumn) Then
            xText = column.Title
        End If

        Dim xPen As Pen
        Dim xBrush As LinearGradientBrush
        Dim xBrush2 As SolidBrush

        Dim colX = _editor.Columns.GetColumnLeft(_editor.State.Mouse.CurrentMouseColumn)
        Dim colStartDrawX = HPositionToPanelX(colX)
        Dim colWidth = _editor.Columns.GetWidth(_editor.State.Mouse.CurrentMouseColumn)
        Dim colEndDrawX = HPositionToPanelX(colX + colWidth)
        Dim colRightDrawX = HPositionToPanelX(colX + colWidth)
        Dim startY = VPositionToPanelY(_editor.State.Mouse.CurrentMouseRow)

        Dim point1 As New Point(colStartDrawX,
                                startY - _theme.NoteHeight - 10)
        Dim point2 As New Point(colEndDrawX,
                                startY + 10)

        Dim bright As Color
        Dim dark As Color
        If _editor.NtInput Or Not ModifierLongNoteActive() Then
            xPen = New Pen(column.GetBright(alpha))
            bright = column.GetBright(alpha)
            dark = column.GetDark(alpha)

            xBrush2 = New SolidBrush(column.cText)
        Else
            xPen = New Pen(column.GetLongBright(alpha))
            bright = column.GetLongBright(alpha)
            dark = column.GetLongDark(alpha)

            xBrush2 = New SolidBrush(column.cLText)
        End If

        ' Temp landmine
        If ModifierLandmineActive() Then
            bright = Color.Red
            dark = Color.Red
        End If

        xBrush = New LinearGradientBrush(point1, point2, bright, dark)

        Dim colDrawWidth = _editor.Columns.GetWidth(_editor.State.Mouse.CurrentMouseColumn) * _editor.Grid.WidthScale

        _buffer.FillRectangle(xBrush,
                             colStartDrawX + 2,
                             startY - _theme.NoteHeight + 1,
                             colDrawWidth - 3,
                             _theme.NoteHeight - 1)
        _buffer.DrawRectangle(xPen,
                             colStartDrawX + 1,
                             startY - _theme.NoteHeight,
                             colDrawWidth - 2,
                             _theme.NoteHeight)

        _buffer.DrawString(xText, _theme.kFont, xBrush2,
                          colStartDrawX + _theme.kLabelHShiftL - 2,
                          startY - _theme.NoteHeight + _theme.kLabelVShift)
    End Sub

    Private Sub DrawTimeSelection()
        Dim xBpmStart = _editor.FirstBpm
        Dim xBpmHalf = _editor.FirstBpm
        Dim xBpmEnd = _editor.FirstBpm

        For Each note In _editor.Notes
            If note.ColumnIndex = ColumnType.BPM Then
                If note.VPosition <= _editor.State.TimeSelect.StartPoint Then
                    xBpmStart = note.Value
                End If
                If note.VPosition <= _editor.State.TimeSelect.HalfPoint Then
                    xBpmHalf = note.Value
                End If
                If note.VPosition <= _editor.State.TimeSelect.EndPoint Then
                    xBpmEnd = note.Value
                End If
            End If
            If note.VPosition > _editor.State.TimeSelect.EndPoint Then
                Exit For
            End If
        Next

        Dim bpmColStartX = _editor.Columns.GetColumnLeft(ColumnType.BPM)
        Dim startDrawX = HPositionToPanelX(bpmColStartX)
        Dim halfLineY = VPositionToPanelY(_editor.State.TimeSelect.StartPoint + _editor.State.TimeSelect.HalfPointLength)
        Dim endLineY = VPositionToPanelY(_editor.State.TimeSelect.StartPoint + _editor.State.TimeSelect.EndPointLength)
        Dim selectionDrawHeight = CInt(Math.Abs(_editor.State.TimeSelect.EndPointLength) * _editor.Grid.HeightScale)
        Dim startLineY =
                VPositionToPanelY(
                    _editor.State.TimeSelect.StartPoint + Math.Max(_editor.State.TimeSelect.EndPointLength, 0))
        Dim startLineStringY = VPositionToPanelY(_editor.State.TimeSelect.StartPoint)

        'Selection area
        _buffer.FillRectangle(_theme.PESel,
                             0, startLineY,
                             Size.Width,
                             selectionDrawHeight)
        'End Cursor
        _buffer.DrawLine(_theme.PECursor,
                        0, endLineY,
                        Size.Width, endLineY)
        'Half Cursor
        _buffer.DrawLine(_theme.PEHalf,
                        0, halfLineY,
                        Size.Width, halfLineY)
        'Start BPM
        _buffer.DrawString(xBpmStart / 10000,
                          _theme.PEBPMFont, _theme.PEBPM,
                          startDrawX,
                          startLineStringY - _theme.PEBPMFont.Height + 3)
        'Half BPM
        _buffer.DrawString(xBpmHalf / 10000,
                          _theme.PEBPMFont, _theme.PEBPM,
                          startDrawX,
                          halfLineY - _theme.PEBPMFont.Height + 3)
        'End BPM
        _buffer.DrawString(xBpmEnd / 10000,
                          _theme.PEBPMFont, _theme.PEBPM,
                          startDrawX,
                          endLineY - _theme.PEBPMFont.Height + 3)

        'SelLine
        If _editor.State.TimeSelect.MouseOverLine = 1 Then 'Start Cursor
            _buffer.DrawRectangle(_theme.PEMouseOver,
                                 0, startLineStringY - 1,
                                 Size.Width - 1, 2)
        ElseIf _editor.State.TimeSelect.MouseOverLine = 2 Then 'Half Cursor
            _buffer.DrawRectangle(_theme.PEMouseOver,
                                 0, halfLineY - 1,
                                 Size.Width - 1, 2)
        ElseIf _editor.State.TimeSelect.MouseOverLine = 3 Then 'End Cursor
            _buffer.DrawRectangle(_theme.PEMouseOver,
                                 0, endLineY - 1,
                                 Size.Width - 1, 2)
        End If
    End Sub

    Private Sub DrawClickAndScroll()
        Dim xDeltaLocation As Point = PointToScreen(New Point(0, 0))

        Dim startX As Single = _editor.State.Mouse.MiddleButtonLocation.X - xDeltaLocation.X
        Dim startY As Single = _editor.State.Mouse.MiddleButtonLocation.Y - xDeltaLocation.Y
        Dim currentX As Single = Cursor.Position.X - xDeltaLocation.X
        Dim currentY As Single = Cursor.Position.Y - xDeltaLocation.Y
        Dim angle As Double = Math.Atan2(currentY - startY, currentX - startX)
        _buffer.SmoothingMode = SmoothingMode.HighQuality

        If Not (startX = currentX And startY = currentY) Then
            Dim xPointx() As PointF = {New PointF(currentX, currentY),
                                       New PointF(Math.Cos(angle + Math.PI / 2) * 10 + startX,
                                                  Math.Sin(angle + Math.PI / 2) * 10 + startY),
                                       New PointF(Math.Cos(angle - Math.PI / 2) * 10 + startX,
                                                  Math.Sin(angle - Math.PI / 2) * 10 + startY)}
            _buffer.FillPolygon(
                New LinearGradientBrush(New Point(startX, startY), New Point(currentX, currentY), Color.FromArgb(0),
                                        Color.FromArgb(-1)), xPointx)
        End If

        _buffer.FillEllipse(Brushes.LightGray, startX - 10, startY - 10, 20, 20)
        _buffer.DrawEllipse(Pens.Black, startX - 8, startY - 8, 16, 16)

        _buffer.SmoothingMode = SmoothingMode.Default
    End Sub


    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' TODO: move this
        ' If Me.WindowState = FormWindowState.Minimized Then Return
        ' If DisplayRectangle.Width <= 0 Or DisplayRectangle.Height <= 0 Then Return

        MyBase.OnPaint(e)

        _buffer = e.Graphics

        If _theme Is Nothing Then
            e.Graphics.FillRectangle(New SolidBrush(BackColor), DisplayRectangle)
            e.Graphics.DrawString("Editor Pane", Font, Brushes.White, 0, 0)
            Exit Sub
        End If

        _buffer.Clear(_theme.Bg.Color)

        Dim xVSu As Integer = Math.Min(-VerticalPosition + Size.Height / _editor.Grid.HeightScale, _editor.GetMaxVPosition())

        'Bg color
        If _editor.Grid.ShowBackground Then
            DrawBackgroundColor()
        End If


        DrawPanelLines(xVSu)

        'Column Caption
        If _editor.Grid.ShowColumnCaptions Then
            DrawColumnCaptions()
        End If

        'WaveForm
        DrawWaveform(-VerticalPosition)

        'K
        'If Not K Is Nothing Then
        DrawNotes()

        'End If

        'Selection Box
        DrawSelectionBox()

        'Mouse Over
        If _editor.IsSelectMode AndAlso _editor.State.Mouse.CurrentHoveredNoteIndex <> -1 Then
            ' DrawNoteHoverHighlight(iI, HorizontalScroll, VerticalScroll, panelHeight, foundNoteIndex)
            DrawMouseOver()
        End If

        If _editor.ShouldDrawTempNote AndAlso
           _editor.State.Mouse.CurrentMouseColumn > -1 AndAlso
           _editor.State.Mouse.CurrentMouseRow > -1 Then
            DrawTempNote()
        End If

        'Time Selection
        If _editor.IsTimeSelectMode Then
            DrawTimeSelection()
        End If

        'Middle button: CLick and Scroll
        If _editor.State.Mouse.MiddleButtonClicked Then
            DrawClickAndScroll()
        End If

        'Drag/Drop
        DrawDragAndDrop()
    End Sub

    ''' <summary>
    '''     Transforms a position from absolute panel coordinates to screen panel coordinates
    ''' </summary>
    ''' <param name="xHPosition">Original horizontal position.</param>
    Private Function HPositionToPanelX(xHPosition As Integer) As Integer
        Return (xHPosition - HorizontalPosition) * _editor.Grid.WidthScale
    End Function

    ''' <summary>
    '''     Transform a position from absolute note row (192 parts per 4/4 measure) to a Y position in screen panel
    '''     coordinates.
    ''' </summary>
    ''' <param name="xVPosition">Original vertical position.</param>
    Private Function VPositionToPanelY(xVPosition As Double) As Integer
        Return Size.Height - CInt((xVPosition + VerticalPosition) * _editor.Grid.HeightScale) - 1
    End Function

    Public Function SnapToGrid(vpos As Double) As Double
        Dim offset As Double = _editor.MeasureBottom(_editor.MeasureAtDisplacement(vpos))
        Dim divisor As Double = 192.0R / _editor.Grid.Divider
        Return Math.Floor((vpos - offset) / divisor) * divisor + offset
    End Function

    Friend Sub SetWidthScale(widthScale As Single)
        HorizontalScrollBar.LargeChange = Width / widthScale
        With HorizontalScrollBar
            If .Value > .Maximum - .LargeChange + 1 Then
                .Value = .Maximum - .LargeChange + 1
            End If
        End With
    End Sub

    Friend Sub ScrollChart(xAmount As Double, yAmount As Double)
        With VerticalScrollBar
            Dim i = .Value + yAmount
            If i > 0 Then i = 0
            If i < .Minimum Then i = .Minimum
            .Value = i
        End With
        With HorizontalScrollBar
            Dim dx = .Value + xAmount
            If dx > .Maximum - .LargeChange + 1 Then
                dx = .Maximum - .LargeChange + 1
            End If
            If dx < .Minimum Then dx = .Minimum
            .Value = dx
        End With
    End Sub

    Friend Sub ScrollTo(v As Double)
        VerticalScrollBar.Minimum = Math.Min(v, VerticalScrollBar.Minimum)
        VerticalScrollBar.Maximum = Math.Max(v, VerticalScrollBar.Maximum)
        VerticalScrollBar.Value = v
    End Sub
End Class
