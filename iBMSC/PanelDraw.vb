Imports iBMSC.Editor

Partial Public Class MainWindow

    Private Sub RefreshPanelAll()
        If IsInitializing Then Exit Sub
        RefreshPanel(0, PMainInL.DisplayRectangle)
        RefreshPanel(1, PMainIn.DisplayRectangle)
        RefreshPanel(2, PMainInR.DisplayRectangle)
    End Sub

    Dim bufferlist As Dictionary(Of Integer, BufferedGraphics) = New Dictionary(Of Integer, BufferedGraphics)
    Dim rectList As Dictionary(Of Integer, Rectangle) = New Dictionary(Of Integer, Rectangle)
    Private Function GetBuffer(xIndex As Integer, DisplayRect As Rectangle)
        If bufferlist.ContainsKey(xIndex) AndAlso rectList.Item(xIndex) = DisplayRect Then
            Return bufferlist.Item(xIndex)
        Else
            If bufferlist.ContainsKey(xIndex) Then
                bufferlist.Item(xIndex).Dispose()
                bufferlist.Remove(xIndex)
                rectList.Remove(xIndex)
            End If

            Dim gfx = BufferedGraphicsManager.Current.Allocate(spMain(xIndex).CreateGraphics, DisplayRect)
            bufferlist.Add(xIndex, gfx)
            rectList.Add(xIndex, DisplayRect)
            Return gfx
        End If
    End Function

    Private Sub RefreshPanel(ByVal xIndex As Integer, ByVal DisplayRect As Rectangle)
        If Me.WindowState = FormWindowState.Minimized Then Return
        If DisplayRect.Width <= 0 Or DisplayRect.Height <= 0 Then Return
        'If spMain.Count = 0 Then Return
        'Dim currentContext As BufferedGraphicsContext = BufferedGraphicsManager.Current
        Dim e1 As BufferedGraphics = GetBuffer(xIndex, DisplayRect)
        e1.Graphics.FillRectangle(vo.Bg, DisplayRect)

        Dim xTHeight As Integer = spMain(xIndex).Height
        Dim xTWidth As Integer = spMain(xIndex).Width
        Dim xPanelHScroll As Integer = PanelHScroll(xIndex)
        Dim xPanelDisplacement As Integer = PanelVScroll(xIndex)
        Dim xVSR As Integer = -PanelVScroll(xIndex)
        Dim xVSu As Integer = IIf(xVSR + xTHeight / gxHeight > GetMaxVPosition(), GetMaxVPosition(), xVSR + xTHeight / gxHeight)

        'e1.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        Dim xI1 As Integer

        'Bg color
        DrawBackgroundColor(e1, xTHeight, xTWidth, xPanelHScroll, xI1)

        xI1 = DrawPanelLines(e1, xTHeight, xTWidth, xPanelHScroll, xPanelDisplacement, xVSu)

        'Column Caption
        xI1 = DrawColumnCaptions(e1, xTWidth, xPanelHScroll, xI1)

        'WaveForm
        DrawWaveform(e1, xTHeight, xVSR, xI1)

        'K
        'If Not K Is Nothing Then
        DrawNotes(e1, xTHeight, xPanelHScroll, xPanelDisplacement)

        'End If

        'Selection Box
        DrawSelectionBox(xIndex, e1)

        'Mouse Over
        If TBSelect.Checked AndAlso Not KMouseOver = -1 Then
            DrawMouseOver(e1, xTHeight, xPanelHScroll, xPanelDisplacement)
        End If

        If ShouldDrawTempNote AndAlso (SelectedColumn > -1 And TempVPosition > -1) Then
            DrawTempNote(e1, xTHeight, xPanelHScroll, xPanelDisplacement)
        End If

        'Time Selection
        If TBTimeSelect.Checked Then
            DrawTimeSelection(e1, xTHeight, xTWidth, xPanelHScroll, xPanelDisplacement)
        End If

        'Middle button: CLick and Scroll
        If MiddleButtonClicked Then
            e1 = DrawClickAndScroll(xIndex, e1)
        End If

        'Drag/Drop
        DrawDragAndDrop(xIndex, e1)

        e1.Render(spMain(xIndex).CreateGraphics)
        'e1.Dispose()
    End Sub

    Private Sub DrawTempNote(e1 As BufferedGraphics, xTHeight As Integer, xHS As Integer, xVS As Integer)
        Dim xValue As Integer = (LWAV.SelectedIndex + 1) * 10000

        Dim xAlpha As Single = 1.0F
        If ModifierHiddenActive() Then
            xAlpha = vo.kOpacity
        End If

        Dim xText As String = C10to36(xValue \ 10000)
        If IsColumnNumeric(SelectedColumn) Then
            xText = GetColumn(SelectedColumn).Title
        End If

        Dim xPen As Pen
        Dim xBrush As Drawing2D.LinearGradientBrush
        Dim xBrush2 As SolidBrush
        Dim point1 As New Point(HorizontalPositiontoDisplay(nLeft(SelectedColumn), xHS),
                                NoteRowToPanelHeight(TempVPosition, xVS, xTHeight) - vo.kHeight - 10)
        Dim point2 As New Point(HorizontalPositiontoDisplay(nLeft(SelectedColumn) + GetColumnWidth(SelectedColumn), xHS),
                                NoteRowToPanelHeight(TempVPosition, xVS, xTHeight) + 10)

        Dim bright As Color
        Dim dark As Color
        If NTInput Or Not ModifierLongNoteActive() Then
            xPen = New Pen(GetColumn(SelectedColumn).getBright(xAlpha))
            bright = GetColumn(SelectedColumn).getBright(xAlpha)
            dark = GetColumn(SelectedColumn).getDark(xAlpha)

            xBrush2 = New SolidBrush(GetColumn(SelectedColumn).cText)
        Else
            xPen = New Pen(GetColumn(SelectedColumn).getLongBright(xAlpha))
            bright = GetColumn(SelectedColumn).getLongBright(xAlpha)
            dark = GetColumn(SelectedColumn).getLongDark(xAlpha)

            xBrush2 = New SolidBrush(GetColumn(SelectedColumn).cLText)
        End If

        ' Temp landmine
        If ModifierLandmineActive() Then
            bright = Color.Red
            dark = Color.Red
        End If

        xBrush = New Drawing2D.LinearGradientBrush(point1, point2, bright, dark)

        e1.Graphics.FillRectangle(xBrush, HorizontalPositiontoDisplay(nLeft(SelectedColumn), xHS) + 2,
                                  NoteRowToPanelHeight(TempVPosition, xVS, xTHeight) - vo.kHeight + 1,
                                  GetColumnWidth(SelectedColumn) * gxWidth - 3,
                                  vo.kHeight - 1)
        e1.Graphics.DrawRectangle(xPen,
                                  HorizontalPositiontoDisplay(nLeft(SelectedColumn), xHS) + 1,
                                  NoteRowToPanelHeight(TempVPosition, xVS, xTHeight) - vo.kHeight,
                                  GetColumnWidth(SelectedColumn) * gxWidth - 2,
                                  vo.kHeight)

        e1.Graphics.DrawString(xText, vo.kFont, xBrush2,
                        HorizontalPositiontoDisplay(nLeft(SelectedColumn), xHS) + vo.kLabelHShiftL - 2,
                        NoteRowToPanelHeight(TempVPosition, xVS, xTHeight) - vo.kHeight + vo.kLabelVShift)
    End Sub

    Private Sub DrawDragAndDrop(xIndex As Integer, e1 As BufferedGraphics)
        If UBound(DDFileName) > -1 Then
            'Dim xFont As New Font("Cambria", 12)
            Dim xBrush As New SolidBrush(Color.FromArgb(&HC0FFFFFF))
            Dim xCenterX As Single = spMain(xIndex).DisplayRectangle.Width / 2
            Dim xCenterY As Single = spMain(xIndex).DisplayRectangle.Height / 2
            Dim xFormat As New System.Drawing.StringFormat
            xFormat.Alignment = StringAlignment.Center
            xFormat.LineAlignment = StringAlignment.Center
            e1.Graphics.DrawString(Join(DDFileName, vbCrLf), Me.Font, xBrush, spMain(xIndex).DisplayRectangle, xFormat)
        End If
    End Sub

    Private Sub DrawSelectionBox(xIndex As Integer, e1 As BufferedGraphics)
        If TBSelect.Checked AndAlso xIndex = PanelFocus AndAlso Not (pMouseMove = New Point(-1, -1) Or LastMouseDownLocation = New Point(-1, -1)) Then
            e1.Graphics.DrawRectangle(vo.SelBox, IIf(pMouseMove.X > LastMouseDownLocation.X, LastMouseDownLocation.X, pMouseMove.X),
                                                IIf(pMouseMove.Y > LastMouseDownLocation.Y, LastMouseDownLocation.Y, pMouseMove.Y),
                                                Math.Abs(pMouseMove.X - LastMouseDownLocation.X), Math.Abs(pMouseMove.Y - LastMouseDownLocation.Y))
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

    Private Sub DrawBackgroundColor(e1 As BufferedGraphics, xTHeight As Integer, xTWidth As Integer, xHS As Integer, xI1 As Integer)
        If gShowBG Then
            For xI1 = 0 To gColumns
                If nLeft(xI1 + 1) * gxWidth - xHS * gxWidth + 1 < 0 Then Continue For
                If nLeft(xI1) * gxWidth - xHS * gxWidth + 1 > xTWidth Then Exit For
                If Not GetColumn(xI1).cBG.GetBrightness = 0 And GetColumnWidth(xI1) > 0 Then
                    Dim col = GetColumn(xI1).cBG
                    If xI1 = GetColumnAtX(MouseMoveStatus.X, xHS) Then
                        Dim bf = 1.2
                        col = GetColumnHighlightColor(col)
                    End If
                    Dim brush = New SolidBrush(col)

                    e1.Graphics.FillRectangle(brush,
                                              nLeft(xI1) * gxWidth - xHS * gxWidth + 1,
                                              0,
                                              GetColumnWidth(xI1) * gxWidth,
                                              xTHeight)
                End If
            Next
        End If
    End Sub

    Private Function DrawColumnCaptions(e1 As BufferedGraphics, xTWidth As Integer, xHS As Integer, xI1 As Integer) As Integer
        If gShowC Then
            For xI1 = 0 To gColumns
                If nLeft(xI1 + 1) * gxWidth - xHS * gxWidth + 1 < 0 Then Continue For
                If nLeft(xI1) * gxWidth - xHS * gxWidth + 1 > xTWidth Then Exit For
                If GetColumnWidth(xI1) > 0 Then e1.Graphics.DrawString(nTitle(xI1), vo.ColumnTitleFont, vo.ColumnTitle, nLeft(xI1) * gxWidth - xHS * gxWidth, 0)
            Next
        End If

        Return xI1
    End Function

    Private Function DrawPanelLines(e1 As BufferedGraphics,
                               xTHeight As Integer, xTWidth As Integer,
                               xHS As Integer, xVS As Integer,
                               xVSu As Integer) As Integer
        'Vertical line
        If gShowVerticalLine Then
            For xI1 = 0 To gColumns
                Dim xpos = nLeft(xI1) * gxWidth - xHS * gxWidth
                If xpos + 1 < 0 Then Continue For
                If xpos + 1 > xTWidth Then Exit For
                If GetColumnWidth(xI1) > 0 Then e1.Graphics.DrawLine(vo.pVLine,
                                                              xpos, 0,
                                                              xpos, xTHeight)
            Next
        End If

        'Grid, Sub, Measure
        Dim Measure
        For Measure = MeasureAtDisplacement(-xVS) To MeasureAtDisplacement(xVSu)
            'grid
            If gShowGrid Then DrawGridLines(e1,
                                        xTHeight, xTWidth,
                                        xVS, Measure,
                                        gDivide, vo.pGrid)

            'sub
            If gShowSubGrid Then DrawGridLines(e1,
                                         xTHeight, xTWidth,
                                         xVS, Measure,
                                         gSub, vo.pSub)


            'measure and measurebar
            Dim xCurr = MeasureBottom(Measure)
            Dim Height = NoteRowToPanelHeight(xCurr, xVS, xTHeight)
            If gShowMeasureBar Then e1.Graphics.DrawLine(vo.pMLine, 0, Height,
                                                 xTWidth, Height)
            If gShowMeasureNumber Then e1.Graphics.DrawString("[" & Add3Zeros(Measure).ToString & "]", vo.kMFont,
                                                  New SolidBrush(GetColumn(0).cText), -xHS * gxWidth,
                                                  Height - vo.kMFont.Height)
        Next

        Dim vpos = GetMouseVPosition(gSnap)
        Dim mouseLineHeight = NoteRowToPanelHeight(vpos, xVS, xTHeight)
        Dim p = New Pen(Color.White)
        e1.Graphics.DrawLine(p, 0, mouseLineHeight, xTWidth, mouseLineHeight)

        Return Measure
    End Function

    Private Sub DrawGridLines(e1 As BufferedGraphics,
                              xTHeight As Integer, xTWidth As Integer,
                              xVS As Integer, measureIndex As Integer,
                              divisions As Integer, pen As Pen)
        Dim Line = 0
        Dim xUpper As Double = MeasureUpper(measureIndex)
        Dim xCurr = MeasureBottom(measureIndex)
        Dim xDiff = 192 / divisions
        Do While xCurr < xUpper
            Dim Height = NoteRowToPanelHeight(xCurr, xVS, xTHeight)
            e1.Graphics.DrawLine(pen, 0, Height,
                                      xTWidth, Height)
            Line += 1
            xCurr = MeasureBottom(measureIndex) + Line * xDiff
        Loop
    End Sub

    Private Function IsNoteVisible(note As Note, xTHeight As Integer, xVS As Integer) As Boolean
        Dim xUpperBorder As Single = Math.Abs(xVS) + xTHeight / gxHeight
        Dim xLowerBorder As Single = Math.Abs(xVS) - vo.kHeight / gxHeight

        Dim AboveLower = note.VPosition >= xLowerBorder
        Dim HeadBelow = note.VPosition <= xLowerBorder
        Dim TailAbove = note.VPosition + note.Length >= xLowerBorder
        Dim IntersectsNT = HeadBelow And TailAbove
        Dim Intersecs = (note.VPosition <= xLowerBorder And Notes(note.LNPair).VPosition >= xLowerBorder)
        Dim AboveUpper = note.VPosition > xUpperBorder

        Dim NoteInside = (Not AboveUpper) And AboveLower

        Return NoteInside OrElse IntersectsNT OrElse IntersectsNT
    End Function

    Private Function IsNoteVisible(noteindex As Integer, xTHeight As Integer, xVS As Integer) As Boolean
        Return IsNoteVisible(Notes(noteindex), xTHeight, xVS)
    End Function

    Private Sub DrawNotes(e1 As BufferedGraphics, xTHeight As Integer, xHS As Integer, xVS As Integer)
        Dim xI1 As Integer
        Dim xUpperBorder As Single = Math.Abs(xVS) + xTHeight / gxHeight
        Dim xLowerBorder As Single = Math.Abs(xVS) - vo.kHeight / gxHeight

        For xI1 = 0 To UBound(Notes)
            If Notes(xI1).VPosition > xUpperBorder Then Exit For
            If Not IsNoteVisible(xI1, xTHeight, xVS) Then Continue For
            If NTInput Then
                DrawNoteNT(Notes(xI1), e1, xHS, xVS, xTHeight)
            Else
                DrawNote(Notes(xI1), e1, xHS, xVS, xTHeight)
            End If
        Next
    End Sub

    Private Function GetNoteRectangle(note As Note, xTHeight As Integer, xHS As Integer, xVS As Integer) As Rectangle
        Dim xDispX As Integer = HorizontalPositiontoDisplay(nLeft(note.ColumnIndex), xHS)

        Dim xDispY As Integer = IIf(Not NTInput Or (bAdjustLength And Not bAdjustUpper),
                                    NoteRowToPanelHeight(note.VPosition, xVS, xTHeight) - vo.kHeight - 1,
                                    NoteRowToPanelHeight(note.VPosition +
                                    note.Length, xVS, xTHeight) -
                                    vo.kHeight - 1)

        Dim xDispW As Integer = GetColumnWidth(note.ColumnIndex) * gxWidth + 1
        Dim xDispH As Integer = IIf(Not NTInput Or bAdjustLength,
                                    vo.kHeight + 3,
                                    note.Length * gxHeight + vo.kHeight + 3)

        Return New Rectangle(xDispX, xDispY, xDispW, xDispH)
    End Function

    Private Function GetNoteRectangle(noteIndex As Integer, xTHeight As Integer, xHS As Integer, xVS As Integer) As Rectangle
        Return GetNoteRectangle(Notes(noteIndex), xTHeight, xHS, xVS)
    End Function


    Private Sub DrawMouseOver(e1 As BufferedGraphics, xTHeight As Integer, xHS As Integer, xVS As Integer)
        If NTInput Then
            If Not bAdjustLength Then DrawNoteNT(Notes(KMouseOver), e1, xHS, xVS, xTHeight)
        Else
            DrawNote(Notes(KMouseOver), e1, xHS, xVS, xTHeight)
        End If

        Dim rect = GetNoteRectangle(KMouseOver, xTHeight, xHS, xVS)
        Dim pen = IIf(bAdjustLength, vo.kMouseOverE, vo.kMouseOver)
        e1.Graphics.DrawRectangle(pen, rect.X, rect.Y, rect.Width - 1, rect.Height - 1)

        If ModifierMultiselectActive() Then
            For Each note In Notes
                If IsNoteVisible(note, xTHeight, xVS) AndAlso IsLabelMatch(note, KMouseOver) Then
                    Dim nrect = GetNoteRectangle(note, xTHeight, xHS, xVS)
                    e1.Graphics.DrawRectangle(pen, nrect.X, nrect.Y, nrect.Width - 1, nrect.Height - 1)
                End If
            Next
        End If

    End Sub

    Private Sub DrawTimeSelection(e1 As BufferedGraphics, xTHeight As Integer, xTWidth As Integer, xHS As Integer, xVS As Integer)
        Dim xI1 As Integer
        Dim xBPMStart = Notes(0).Value
        Dim xBPMHalf = Notes(0).Value
        Dim xBPMEnd = Notes(0).Value

        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).ColumnIndex = niBPM Then
                If Notes(xI1).VPosition <= vSelStart Then xBPMStart = Notes(xI1).Value
                If Notes(xI1).VPosition <= vSelStart + vSelHalf Then xBPMHalf = Notes(xI1).Value
                If Notes(xI1).VPosition <= vSelStart + vSelLength Then xBPMEnd = Notes(xI1).Value
            End If
            If Notes(xI1).VPosition > vSelStart + vSelLength Then Exit For
        Next

        'Selection area
        e1.Graphics.FillRectangle(vo.PESel,
                                  0,
                                  NoteRowToPanelHeight(vSelStart + IIf(vSelLength > 0, vSelLength, 0), xVS, xTHeight) + Math.Abs(CInt(vSelLength <> 0)),
                                  xTWidth,
                                  CInt(Math.Abs(vSelLength) * gxHeight))
        'End Cursor
        e1.Graphics.DrawLine(vo.PECursor,
                             0,
                             NoteRowToPanelHeight(vSelStart + vSelLength, xVS, xTHeight),
                             xTWidth,
                             NoteRowToPanelHeight(vSelStart + vSelLength, xVS, xTHeight))
        'Half Cursor
        e1.Graphics.DrawLine(vo.PEHalf,
                             0,
                             NoteRowToPanelHeight(vSelStart + vSelHalf, xVS, xTHeight),
                             xTWidth,
                             NoteRowToPanelHeight(vSelStart + vSelHalf, xVS, xTHeight))
        'Start BPM
        e1.Graphics.DrawString(xBPMStart / 10000,
                               vo.PEBPMFont, vo.PEBPM,
                               (-xHS + nLeft(niBPM)) * gxWidth,
                               NoteRowToPanelHeight(vSelStart, xVS, xTHeight) - vo.PEBPMFont.Height + 3)
        'Half BPM
        e1.Graphics.DrawString(xBPMHalf / 10000,
                               vo.PEBPMFont, vo.PEBPM,
                               (-xHS + nLeft(niBPM)) * gxWidth,
                               NoteRowToPanelHeight(vSelStart + vSelHalf, xVS, xTHeight) - vo.PEBPMFont.Height + 3)
        'End BPM
        e1.Graphics.DrawString(xBPMEnd / 10000,
                               vo.PEBPMFont, vo.PEBPM,
                               (-xHS + nLeft(niBPM)) * gxWidth,
                               NoteRowToPanelHeight(vSelStart + vSelLength, xVS, xTHeight) - vo.PEBPMFont.Height + 3)

        'SelLine
        If vSelMouseOverLine = 1 Then 'Start Cursor
            e1.Graphics.DrawRectangle(vo.PEMouseOver,
                                      0, NoteRowToPanelHeight(vSelStart, xVS, xTHeight) - 1,
                                      xTWidth - 1, 2)
        ElseIf vSelMouseOverLine = 2 Then 'Half Cursor
            e1.Graphics.DrawRectangle(vo.PEMouseOver,
                                      0, NoteRowToPanelHeight(vSelStart + vSelHalf, xVS, xTHeight) - 1,
                                      xTWidth - 1, 2)
        ElseIf vSelMouseOverLine = 3 Then 'End Cursor
            e1.Graphics.DrawRectangle(vo.PEMouseOver,
                                      0, NoteRowToPanelHeight(vSelStart + vSelLength, xVS, xTHeight) - 1,
                                      xTWidth - 1, 2)
        End If
    End Sub

    Private Function DrawClickAndScroll(xIndex As Integer, e1 As BufferedGraphics) As BufferedGraphics
        Dim xDeltaLocation As Point = spMain(xIndex).PointToScreen(New Point(0, 0))

        Dim xInitX As Single = MiddleButtonLocation.X - xDeltaLocation.X
        Dim xInitY As Single = MiddleButtonLocation.Y - xDeltaLocation.Y
        Dim xCurrX As Single = Cursor.Position.X - xDeltaLocation.X
        Dim xCurrY As Single = Cursor.Position.Y - xDeltaLocation.Y
        Dim xAngle As Double = Math.Atan2(xCurrY - xInitY, xCurrX - xInitX)
        e1.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality

        If Not (xInitX = xCurrX And xInitY = xCurrY) Then
            Dim xPointx() As PointF = {New PointF(xCurrX, xCurrY),
                                       New PointF(Math.Cos(xAngle + Math.PI / 2) * 10 + xInitX, Math.Sin(xAngle + Math.PI / 2) * 10 + xInitY),
                                       New PointF(Math.Cos(xAngle - Math.PI / 2) * 10 + xInitX, Math.Sin(xAngle - Math.PI / 2) * 10 + xInitY)}
            e1.Graphics.FillPolygon(New Drawing2D.LinearGradientBrush(New Point(xInitX, xInitY), New Point(xCurrX, xCurrY), Color.FromArgb(0), Color.FromArgb(-1)), xPointx)
        End If

        e1.Graphics.FillEllipse(Brushes.LightGray, xInitX - 10, xInitY - 10, 20, 20)
        e1.Graphics.DrawEllipse(Pens.Black, xInitX - 8, xInitY - 8, 16, 16)

        e1.Graphics.SmoothingMode = Drawing2D.SmoothingMode.Default
        Return e1
    End Function

    Private Sub DrawWaveform(e1 As BufferedGraphics, xTHeight As Integer, xVSR As Integer, xI1 As Integer)
        If wWavL IsNot Nothing And wWavR IsNot Nothing And wPrecision > 0 Then
            If wLock Then
                For xI0 As Integer = 1 To UBound(Notes)
                    If Notes(xI0).ColumnIndex >= niB Then wPosition = Notes(xI0).VPosition : Exit For
                Next
            End If

            Dim xPtsL(xTHeight * wPrecision) As PointF
            Dim xPtsR(xTHeight * wPrecision) As PointF

            Dim xD1 As Double

            Dim bVPosition() As Double = {wPosition}
            Dim bBPM() As Decimal = {Notes(0).Value / 10000}
            Dim bWavDataIndex() As Decimal = {0}

            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).ColumnIndex = niBPM Then
                    If Notes(xI1).VPosition >= wPosition Then
                        ReDim Preserve bVPosition(UBound(bVPosition) + 1)
                        ReDim Preserve bBPM(UBound(bBPM) + 1)
                        ReDim Preserve bWavDataIndex(UBound(bWavDataIndex) + 1)
                        bVPosition(UBound(bVPosition)) = Notes(xI1).VPosition
                        bBPM(UBound(bBPM)) = Notes(xI1).Value / 10000
                        bWavDataIndex(UBound(bWavDataIndex)) = (Notes(xI1).VPosition - bVPosition(UBound(bVPosition) - 1)) * 1.25 * wSampleRate / bBPM(UBound(bBPM) - 1) + bWavDataIndex(UBound(bWavDataIndex) - 1)
                    Else
                        bBPM(0) = Notes(xI1).Value / 10000
                    End If
                End If
            Next

            Dim xI2 As Integer = 0
            Dim xI3 As Double

            For xI1 = xTHeight * wPrecision To 0 Step -1
                xI3 = (-xI1 / wPrecision + xTHeight + xVSR * gxHeight - 1) / gxHeight
                For xI2 = 1 To UBound(bVPosition)
                    If bVPosition(xI2) >= xI3 Then Exit For
                Next
                xI2 -= 1
                xD1 = bWavDataIndex(xI2) + (xI3 - bVPosition(xI2)) * 1.25 * wSampleRate / bBPM(xI2)

                If xD1 <= UBound(wWavL) And xD1 >= 0 Then
                    xPtsL(xI1) = New PointF(wWavL(Int(xD1)) * wWidth + wLeft, xI1 / wPrecision)
                    xPtsR(xI1) = New PointF(wWavR(Int(xD1)) * wWidth + wLeft, xI1 / wPrecision)
                Else
                    xPtsL(xI1) = New PointF(wLeft, xI1 / wPrecision)
                    xPtsR(xI1) = New PointF(wLeft, xI1 / wPrecision)
                End If
            Next
            e1.Graphics.DrawLines(vo.pBGMWav, xPtsL)
            e1.Graphics.DrawLines(vo.pBGMWav, xPtsR)
        End If
    End Sub

    ''' <summary>
    ''' Draw a note in a buffer.
    ''' </summary>
    ''' <param name="sNote">Note to be drawn.</param>
    ''' <param name="e">Buffer.</param>
    ''' <param name="xHS">HS.Value.</param>
    ''' <param name="xVS">VS.Value.</param>
    ''' <param name="xHeight">Display height of the panel. (not ClipRectangle.Height)</param>

    Private Sub DrawNote(ByVal sNote As Note, ByVal e As BufferedGraphics, ByVal xHS As Long, ByVal xVS As Long, ByVal xHeight As Integer) ', Optional ByVal CheckError As Boolean = True) ', Optional ByVal ConnectToIndex As Long = 0)
        If Not nEnabled(sNote.ColumnIndex) Then Exit Sub
        Dim xAlpha As Single = 1.0F
        If sNote.Hidden Then xAlpha = vo.kOpacity

        Dim xLabel As String = C10to36(sNote.Value \ 10000)
        If ShowFileName Then
            If IsColumnSound(sNote.ColumnIndex) Then
                If hWAV(C36to10(xLabel)) <> "" Then xLabel = Path.GetFileNameWithoutExtension(hWAV(C36to10(xLabel)))
            Else
                If hBMP(C36to10(xLabel)) <> "" Then xLabel = Path.GetFileNameWithoutExtension(hBMP(C36to10(xLabel)))
            End If
        End If

        Dim xPen As Pen
        Dim xBrush As Drawing2D.LinearGradientBrush
        Dim xBrush2 As SolidBrush

        Dim bright As Color
        Dim dark As Color
        Dim p1 = New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS),
                           NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight - 10)
        Dim p2 = New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex) + GetColumnWidth(sNote.ColumnIndex), xHS),
                           NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) + 10)

        If Not sNote.LongNote Then
            xPen = New Pen(GetColumn(sNote.ColumnIndex).getBright(xAlpha))

            bright = GetColumn(sNote.ColumnIndex).getBright(xAlpha)
            dark = GetColumn(sNote.ColumnIndex).getDark(xAlpha)

            If sNote.Landmine Then
                bright = Color.Red
                dark = Color.Red
            End If

            xBrush2 = New SolidBrush(GetColumn(sNote.ColumnIndex).cText)
        Else
            bright = GetColumn(sNote.ColumnIndex).getLongBright(xAlpha)
            dark = GetColumn(sNote.ColumnIndex).getLongDark(xAlpha)

            xBrush2 = New SolidBrush(GetColumn(sNote.ColumnIndex).cLText)
        End If

        xPen = New Pen(bright)
        xBrush = New Drawing2D.LinearGradientBrush(p1, p2, bright, dark)

        ' Fill
        e.Graphics.FillRectangle(xBrush, HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS) + 2,
                                 NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight + 1,
                                 GetColumnWidth(sNote.ColumnIndex) * gxWidth - 3,
                                 vo.kHeight - 1)
        ' Outline
        e.Graphics.DrawRectangle(xPen,
                                 HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS) + 1,
                                 NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight,
                                 GetColumnWidth(sNote.ColumnIndex) * gxWidth - 2,
                                 vo.kHeight)

        ' Label
        e.Graphics.DrawString(IIf(IsColumnNumeric(sNote.ColumnIndex), sNote.Value / 10000, xLabel),
                              vo.kFont, xBrush2,
                              HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS) + vo.kLabelHShift,
                              NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight + vo.kLabelVShift)

        If sNote.ColumnIndex < niB Then
            If sNote.LNPair <> 0 Then
                DrawPairedLNBody(sNote, e, xHS, xVS, xHeight, xAlpha)
            End If
        End If


        'e.Graphics.DrawString(sNote.TimeOffset.ToString("0.##"), New Font("Verdana", 9), Brushes.Cyan, _
        '                      New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex + 1), xHS), VerticalPositiontoDisplay(sNote.VPosition, xVS, xHeight) - vo.kHeight - 2))

        'If ErrorCheck AndAlso (sNote.LongNote Xor sNote.PairWithI <> 0) Then e.Graphics.DrawImage(My.Resources.ImageError, _
        If ErrorCheck AndAlso sNote.HasError Then e.Graphics.DrawImage(My.Resources.ImageError,
                                                            CInt(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex) + GetColumnWidth(sNote.ColumnIndex) / 2, xHS) - 12),
                                                            CInt(NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight / 2 - 12),
                                                            24, 24)

        If sNote.Selected Then e.Graphics.DrawRectangle(vo.kSelected, HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS), NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight - 1, GetColumnWidth(sNote.ColumnIndex) * gxWidth, vo.kHeight + 2)

    End Sub

    Private Sub DrawPairedLNBody(sNote As Note, e As BufferedGraphics, xHS As Long, xVS As Long, xHeight As Integer, xAlpha As Single)
        Dim xPen2 As New Pen(GetColumn(sNote.ColumnIndex).getLongBright(xAlpha))
        Dim xBrush3 As New Drawing2D.LinearGradientBrush(
                    New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex) - 0.5 * GetColumnWidth(sNote.ColumnIndex), xHS),
                            NoteRowToPanelHeight(Notes(sNote.LNPair).VPosition, xVS, xHeight)),
                    New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex) + 1.5 * GetColumnWidth(sNote.ColumnIndex), xHS),
                            NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) + vo.kHeight),
                    GetColumn(sNote.ColumnIndex).getLongBright(xAlpha),
                    GetColumn(sNote.ColumnIndex).getLongDark(xAlpha))
        e.Graphics.FillRectangle(xBrush3, HorizontalPositiontoDisplay(nLeft(Notes(sNote.LNPair).ColumnIndex), xHS) + 3, NoteRowToPanelHeight(Notes(sNote.LNPair).VPosition, xVS, xHeight) + 1,
                                        GetColumnWidth(Notes(sNote.LNPair).ColumnIndex) * gxWidth - 5, NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - NoteRowToPanelHeight(Notes(sNote.LNPair).VPosition, xVS, xHeight) - vo.kHeight - 1)
        e.Graphics.DrawRectangle(xPen2, HorizontalPositiontoDisplay(nLeft(Notes(sNote.LNPair).ColumnIndex), xHS) + 2, NoteRowToPanelHeight(Notes(sNote.LNPair).VPosition, xVS, xHeight),
                                        GetColumnWidth(Notes(sNote.LNPair).ColumnIndex) * gxWidth - 4, NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - NoteRowToPanelHeight(Notes(sNote.LNPair).VPosition, xVS, xHeight) - vo.kHeight)
    End Sub

    ''' <summary>
    ''' Draw a note in a buffer under NT mode.
    ''' </summary>
    ''' <param name="sNote">Note to be drawn.</param>
    ''' <param name="e">Buffer.</param>
    ''' <param name="xHS">HS.Value.</param>
    ''' <param name="xVS">VS.Value.</param>
    ''' <param name="xHeight">Display height of the panel. (not ClipRectangle.Height)</param>

    Private Sub DrawNoteNT(ByVal sNote As Note, ByVal e As BufferedGraphics, ByVal xHS As Long, ByVal xVS As Long, ByVal xHeight As Integer) ', Optional ByVal CheckError As Boolean = True)
        If Not nEnabled(sNote.ColumnIndex) Then Exit Sub
        Dim xAlpha As Single = 1.0F
        If sNote.Hidden Then xAlpha = vo.kOpacity

        Dim xLabel As String = C10to36(sNote.Value \ 10000)
        If ShowFileName Then
            If IsColumnSound(sNote.ColumnIndex) Then
                If hWAV(C36to10(xLabel)) <> "" Then xLabel = Path.GetFileNameWithoutExtension(hWAV(C36to10(xLabel)))
            Else
                If hBMP(C36to10(xLabel)) <> "" Then xLabel = Path.GetFileNameWithoutExtension(hBMP(C36to10(xLabel)))
            End If
        End If

        Dim xPen1 As Pen
        Dim xBrush As Drawing2D.LinearGradientBrush
        Dim xBrush2 As SolidBrush

        Dim p1 As Point
        Dim p2 As Point
        Dim bright As Color
        Dim dark As Color

        If sNote.Length = 0 Then
            p1 = New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS),
                           NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight - 10)

            p2 = New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex) + GetColumnWidth(sNote.ColumnIndex), xHS),
                           NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) + 10)

            bright = GetColumn(sNote.ColumnIndex).getBright(xAlpha)
            dark = GetColumn(sNote.ColumnIndex).getDark(xAlpha)

            If sNote.Landmine Then
                bright = Color.Red
                dark = Color.Red
            End If

            xBrush2 = New SolidBrush(GetColumn(sNote.ColumnIndex).cText)
        Else
            p1 = New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex) - 0.5 * GetColumnWidth(sNote.ColumnIndex), xHS),
                           NoteRowToPanelHeight(sNote.VPosition + sNote.Length, xVS, xHeight) - vo.kHeight)
            p2 = New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex) + 1.5 * GetColumnWidth(sNote.ColumnIndex), xHS),
                                      NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight))

            bright = GetColumn(sNote.ColumnIndex).getLongBright(xAlpha)
            dark = GetColumn(sNote.ColumnIndex).getLongDark(xAlpha)

            xBrush2 = New SolidBrush(GetColumn(sNote.ColumnIndex).cLText)
        End If

        xPen1 = New Pen(bright)
        xBrush = New Drawing2D.LinearGradientBrush(p1, p2, bright, dark)

        ' Note gradient
        e.Graphics.FillRectangle(xBrush,
                                     HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS) + 1,
                                     NoteRowToPanelHeight(sNote.VPosition + sNote.Length, xVS, xHeight) - vo.kHeight + 1,
                                     GetColumnWidth(sNote.ColumnIndex) * gxWidth - 1,
                                     CInt(sNote.Length * gxHeight) + vo.kHeight - 1)

        ' Outline
        e.Graphics.DrawRectangle(xPen1, HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS) + 1,
                                     NoteRowToPanelHeight(sNote.VPosition + sNote.Length, xVS, xHeight) - vo.kHeight,
                                            GetColumnWidth(sNote.ColumnIndex) * gxWidth - 3, CInt(sNote.Length * gxHeight) + vo.kHeight)

        ' Note B36
        e.Graphics.DrawString(IIf(IsColumnNumeric(sNote.ColumnIndex), sNote.Value / 10000, xLabel),
                              vo.kFont, xBrush2,
                              HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS) + vo.kLabelHShiftL - 2,
                              NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight + vo.kLabelVShift)

        ' Draw paired body
        If sNote.ColumnIndex < niB Then
            If sNote.Length = 0 And sNote.LNPair <> 0 Then
                DrawPairedLNBody(sNote, e, xHS, xVS, xHeight, xAlpha)
            End If
        End If


        ' Select Box
        If sNote.Selected Then
            e.Graphics.DrawRectangle(vo.kSelected,
                                    HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex), xHS),
                                    NoteRowToPanelHeight(sNote.VPosition + sNote.Length, xVS, xHeight) - vo.kHeight - 1,
                                    GetColumnWidth(sNote.ColumnIndex) * gxWidth,
                                    CInt(sNote.Length * gxHeight) + vo.kHeight + 2)
        End If

        ' Errors
        If ErrorCheck AndAlso sNote.HasError Then
            e.Graphics.DrawImage(My.Resources.ImageError,
                                 CInt(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex) + GetColumnWidth(sNote.ColumnIndex) / 2, xHS) - 12),
                                 CInt(NoteRowToPanelHeight(sNote.VPosition, xVS, xHeight) - vo.kHeight / 2 - 12),
                                 24, 24)
        End If

        'e.Graphics.DrawString(sNote.TimeOffset.ToString("0.##"), New Font("Verdana", 9), Brushes.Cyan, _
        '                      New Point(HorizontalPositiontoDisplay(nLeft(sNote.ColumnIndex + 1), xHS), VerticalPositiontoDisplay(sNote.VPosition, xVS, xHeight) - vo.kHeight - 2))

    End Sub
End Class
