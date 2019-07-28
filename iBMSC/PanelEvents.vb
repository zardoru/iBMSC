Imports System.Linq
Imports iBMSC.Editor

Partial Public Class EditorPanel
    Public Sub PreviewKeyDownEvent(ByVal sender As Object, ByVal e As PreviewKeyDownEventArgs) Handles Me.PreviewKeyDown
        If e.KeyCode = Keys.ShiftKey Or e.KeyCode = Keys.ControlKey Then
            Editor.RefreshPanelAll()
            Editor.POStatusRefresh()
            Exit Sub
        End If

        If e.KeyCode = 18 Then Exit Sub

        Dim iI As Integer = sender.Tag
        Dim xTargetColumn As Integer = -1
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void

        Select Case e.KeyCode
            Case Keys.Home
                VerticalScrollBar.Value = 0

            Case Keys.End
                VerticalScrollBar.Value = VerticalScrollBar.Minimum

            Case Keys.PageUp
                VerticalScrollBar.Value = Math.Max(VerticalScrollBar.Value - Editor.Grid.PageUpDnScroll, VerticalScrollBar.Minimum)

            Case Keys.PageDown
                VerticalScrollBar.Value = Math.Min(VerticalScrollBar.Value + Editor.Grid.PageUpDnScroll, 0)
        End Select

        PMainInMouseMove(sender)
    End Sub

    Private Sub ResizeEvent(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Resize
        If Not Created Then Exit Sub

        If VerticalScrollBar Is Nothing OrElse
           HorizontalScrollBar Is Nothing OrElse
           Editor Is Nothing Then
            Exit Sub
        End If

        VerticalScrollBar.LargeChange = sender.Height * 0.9
        VerticalScrollBar.Maximum = VerticalScrollBar.LargeChange - 1
        HorizontalScrollBar.LargeChange = sender.Width / Editor.Grid.WidthScale

        If HorizontalScrollBar.Value > HorizontalScrollBar.Maximum - HorizontalScrollBar.LargeChange + 1 Then
            HorizontalScrollBar.Value = HorizontalScrollBar.Maximum - HorizontalScrollBar.LargeChange + 1
        End If

        Refresh()
    End Sub

    Private Sub LostFocusEvent(ByVal sender As Object, ByVal e As EventArgs) Handles Me.LostFocus
        Editor.RefreshPanelAll()
    End Sub

    Private Sub MouseDownEvent(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseDown
        ' az: the hell is this for?
        Editor.tempFirstMouseDown = Editor.FirstClickDisabled And Not sender.Focused

        Editor.PanelFocus = sender.Tag
        sender.Focus()
        Editor.State.Mouse.LastMouseDownLocation = New Point(-1, -1)
        LastVerticalScroll = VerticalScroll

        If Editor.NTInput Then
            Editor.State.NT.IsAdjustingUpperEnd = False
            Editor.State.NT.IsAdjustingNoteLength = False
        End If
        Editor.State.IsDuplicatingSelectedNotes = False
        Editor.State.SelectedNotesWereDuplicated = False

        If Editor.State.Mouse.MiddleButtonClicked Then
            Editor.State.Mouse.MiddleButtonClicked = False
            Exit Sub
        End If


        Select Case e.Button
            Case MouseButtons.Left
                If Editor.tempFirstMouseDown And Not Editor.IsTimeSelectMode Then
                    Editor.RefreshPanelAll()
                    Exit Select
                End If

                Editor.State.Mouse.CurrentHoveredNoteIndex = -1
                'If K Is Nothing Then pMouseDown = e.Location : Exit Select

                'Find the clicked K
                Dim NoteIndex As Integer = GetClickedNote(e)

                Editor.PanelPreviewNoteIndex(NoteIndex)

                For Each note In Editor.Notes
                    note.TempMouseDown = False
                Next

                HandleCurrentModeOnClick(e, NoteIndex)
                Editor.RefreshPanelAll()
                Editor.POStatusRefresh()

            Case MouseButtons.Middle
                If Editor.MiddleButtonMoveMethod = 1 Then
                    Editor.State.Mouse.tempX = e.X
                    Editor.State.Mouse.tempY = e.Y
                    Editor.State.Mouse.tempV = VerticalScroll
                    Editor.State.Mouse.tempH = HorizontalScroll
                Else
                    Editor.State.Mouse.MiddleButtonLocation = Cursor.Position
                    Editor.State.Mouse.MiddleButtonClicked = True
                    Editor.TimerMiddle.Enabled = True
                End If

            Case MouseButtons.Right
                DeselectOrRemove(e)
        End Select
    End Sub

    Private Sub DeselectOrRemove(e As MouseEventArgs)
        Editor.State.Mouse.CurrentHoveredNoteIndex = -1

        Editor.ClearSelectionArray()

        If Not Editor.tempFirstMouseDown Then
            Dim i As Integer
            For i = Editor.Notes.Length - 1 To 1 Step -1
                Dim note = Editor.Notes(i)
                'If mouse is clicking on a K
                If MouseInNote(e, note) Then

                    If My.Computer.Keyboard.ShiftKeyDown Then
                        Editor.SelectWavFromNote(note)
                    Else
                        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                        Dim xRedo As UndoRedo.LinkedURCmd = Nothing

                        RedoRemoveNote(note, xUndo, xRedo)
                        Editor.RemoveNote(i)

                        Editor.AddUndoChain(xUndo, xRedo)
                        Editor.RefreshPanelAll()
                    End If

                    Exit For
                End If
            Next

            Editor.CalculateTotalPlayableNotes()
        End If
    End Sub

    Private Function GetClickedNote(e As MouseEventArgs) As Integer
        Dim NoteIndex As Integer = -1
        For i = UBound(Editor.Notes) To 0 Step -1
            Dim note = Editor.Notes(i)

            If MouseInNote(e, note) Then
                NoteIndex = i

                If Editor.NTInput And My.Computer.Keyboard.ShiftKeyDown Then
                    Editor.State.NT.IsAdjustingUpperEnd = e.Y <= VPositionToPanelY(note.VPosition + note.Length)
                    Editor.State.NT.IsAdjustingNoteLength = e.Y >= VPositionToPanelY(note.VPosition) - Theme.kHeight Or
                                                            Editor.State.NT.IsAdjustingUpperEnd
                End If

                Exit For

            End If
        Next

        Return NoteIndex
    End Function



    Private Sub HandleCurrentModeOnClick(e As MouseEventArgs, ByRef clickedNoteIndex As Integer)
        Dim Notes = Editor.Notes
        Dim mouseVPos = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)
        Dim clickedNote As Note = Nothing
        If clickedNoteIndex > 0 Then
            clickedNote = Editor.Notes(clickedNoteIndex)
        End If

        If Editor.IsSelectMode Then
            OnSelectModeLeftClick(e, clickedNoteIndex)
        ElseIf Editor.NTInput And Editor.IsWriteMode Then
            Editor.State.Mouse.CurrentMouseRow = -1
            Editor.State.Mouse.CurrentMouseColumn = -1

            If mouseVPos < 0 Or mouseVPos >= Editor.GetMaxVPosition() Then Exit Sub

            Dim col = GetColumnAtEvent(e)

            For j As Integer = UBound(Editor.Notes) To 1 Step -1
                If Editor.Notes(j).VPosition = mouseVPos And
                   Editor.Notes(j).ColumnIndex = col Then
                    clickedNoteIndex = j
                    Exit For
                End If
            Next



            Dim Hidden As Boolean = ModifierHiddenActive()

            If clickedNoteIndex > 0 Then
                ' Editor.SelectSingleNote(clickedNote)

                'KMouseDown = xITemp
                clickedNote.TempMouseDown = True
                clickedNote.Length = mouseVPos - clickedNote.VPosition

                Editor.State.NT.IsAdjustingUpperEnd = True

                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = Nothing


                RedoLongNoteModify(clickedNote, clickedNote.VPosition, clickedNote.Length, xUndo, xRedo)
                Editor.AddUndoChain(xUndo, xRedo)

            ElseIf Editor.Columns.IsColumnNumeric(col) Then

                Dim xMessage As String = Strings.Messages.PromptEnterNumeric
                If col = ColumnType.BPM Then xMessage = Strings.Messages.PromptEnterBPM
                If col = ColumnType.STOPS Then xMessage = Strings.Messages.PromptEnterSTOP
                If col = ColumnType.SCROLLS Then xMessage = Strings.Messages.PromptEnterSCROLL

                Dim valstr As String = InputBox(xMessage, Text)
                Dim value As Double = Val(valstr) * 10000

                If (col = ColumnType.SCROLLS And valstr = "0") Or value <> 0 Then
                    If col <> ColumnType.SCROLLS And value <= 0 Then value = 1

                    Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                    Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                    Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                    For Each note In Editor.Notes
                        If note.VPosition = mouseVPos AndAlso note.ColumnIndex = col Then
                            Command.RedoRemoveNote(note, xUndo, xRedo)
                        End If
                    Next

                    Dim n = New Note(col, mouseVPos, value, 0, Hidden)
                    Command.RedoAddNote(n, xUndo, xRedo)

                    Editor.AddNote(n)
                    Editor.AddUndoChain(xUndo, xBaseRedo.Next)
                End If

                ' ShouldDrawTempNote = True (az: Why?)

            Else
                Dim xLbl As Integer = Editor.CurrentWavSelectedIndex * 10000

                Dim Landmine As Boolean = ModifierLandmineActive()

                Dim note = New Note With {
                    .VPosition = mouseVPos,
                    .ColumnIndex = col,
                    .Value = xLbl,
                    .Hidden = Hidden,
                    .Landmine = Landmine,
                    .TempMouseDown = True,
                    .LNPair = -1
                }
                Editor.AppendNote(note)
            End If

            Editor.ValidateNotesArray()

        ElseIf Editor.IsTimeSelectMode Then

            If clickedNoteIndex >= 0 Then
                mouseVPos = clickedNote.VPosition
            End If

            UpdateTimeSelectLineOver(e)

            If Not Editor.State.TimeSelect.Adjust Then
                If Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine Then

                    Editor.State.TimeSelect.EndPointLength += Editor.State.TimeSelect.StartPoint - mouseVPos
                    Editor.State.TimeSelect.HalfPointLength += Editor.State.TimeSelect.StartPoint - mouseVPos
                    Editor.State.TimeSelect.StartPoint = mouseVPos

                ElseIf Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine Then
                    Editor.State.TimeSelect.HalfPointLength = mouseVPos
                    If Editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 Then
                        Editor.State.TimeSelect.HalfPointLength = SnapToGrid(Editor.State.TimeSelect.HalfPointLength)
                    End If
                    Editor.State.TimeSelect.HalfPointLength -= Editor.State.TimeSelect.StartPoint

                ElseIf Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine Then
                    Editor.State.TimeSelect.EndPointLength = mouseVPos
                    If Editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 Then
                        Editor.State.TimeSelect.EndPointLength = SnapToGrid(Editor.State.TimeSelect.EndPointLength)
                    End If
                    Editor.State.TimeSelect.EndPointLength -= Editor.State.TimeSelect.StartPoint

                Else
                    Editor.State.TimeSelect.EndPointLength = 0
                    Editor.State.TimeSelect.StartPoint = mouseVPos
                    If Editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 Then
                        Editor.State.TimeSelect.StartPoint = SnapToGrid(Editor.State.TimeSelect.StartPoint)
                    End If
                End If
                Editor.State.TimeSelect.ValidateSelection(Editor.GetMaxVPosition())

            Else
                If Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine Then
                    Editor.SortByVPositionInsertion()
                    Editor.State.TimeSelect.PStart = Editor.State.TimeSelect.StartPoint
                    Editor.State.TimeSelect.PLength = Editor.State.TimeSelect.EndPointLength
                    Editor.State.TimeSelect.PHalf = Editor.State.TimeSelect.HalfPointLength
                    Editor.State.TimeSelect.K = Notes
                    ReDim Preserve Editor.State.TimeSelect.K(UBound(Editor.State.TimeSelect.K))

                    If Editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 And Not My.Computer.Keyboard.CtrlKeyDown Then mouseVPos = SnapToGrid(mouseVPos)
                    Editor.AddUndoChain(New UndoRedo.Void, New UndoRedo.Void)
                    Editor.BPMChangeHalf(mouseVPos - Editor.State.TimeSelect.HalfPointLength - Editor.State.TimeSelect.StartPoint, , True)


                ElseIf Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine Or Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine Then
                    Editor.SortByVPositionInsertion()
                    Editor.State.TimeSelect.PStart = Editor.State.TimeSelect.StartPoint
                    Editor.State.TimeSelect.PLength = Editor.State.TimeSelect.EndPointLength
                    Editor.State.TimeSelect.PHalf = Editor.State.TimeSelect.HalfPointLength
                    Editor.State.TimeSelect.K = Notes.Clone()
                    ReDim Preserve Editor.State.TimeSelect.K(UBound(Editor.State.TimeSelect.K))

                    If Editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 And Not My.Computer.Keyboard.CtrlKeyDown Then mouseVPos = SnapToGrid(mouseVPos)
                    Editor.AddUndoChain(New UndoRedo.Void, New UndoRedo.Void)
                    Dim v = IIf(Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine,
                                            mouseVPos - Editor.State.TimeSelect.StartPoint,
                                            Editor.State.TimeSelect.EndPoint - mouseVPos)

                    Editor.BPMChangeTop(v / Editor.State.TimeSelect.EndPointLength, , True)
                Else
                    Editor.State.TimeSelect.EndPointLength = mouseVPos
                    If Editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 And Not My.Computer.Keyboard.CtrlKeyDown Then Editor.State.TimeSelect.EndPointLength = SnapToGrid(Editor.State.TimeSelect.EndPointLength)
                    Editor.State.TimeSelect.EndPointLength -= Editor.State.TimeSelect.StartPoint
                End If

            End If

            If Editor.State.TimeSelect.EndPointLength Then
                Dim xVLower As Double = Math.Min(Editor.State.TimeSelect.StartPoint, Editor.State.TimeSelect.EndPoint)
                Dim xVUpper As Double = Math.Max(Editor.State.TimeSelect.StartPoint, Editor.State.TimeSelect.EndPoint)
                If Editor.NTInput Then
                    For Each note In Notes.Skip(1)
                        note.Selected = Not note.VPosition >= xVUpper And
                                        Not note.VPosition + note.Length < xVLower And
                                        Editor.Columns.nEnabled(note.ColumnIndex)
                    Next
                Else
                    For Each note In Notes.Skip(1)
                        note.Selected = note.VPosition >= xVLower And
                                        note.VPosition < xVUpper And
                                        Editor.Columns.nEnabled(note.ColumnIndex)
                    Next
                End If
            Else
                Editor.DeselectAllNotes()
            End If

        End If
    End Sub

    Private Sub UpdateTimeSelectLineOver(e As MouseEventArgs)
        Editor.State.TimeSelect.Adjust = ModifierLongNoteActive()

        Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.None
        If Math.Abs(e.Y - VPositionToPanelY(Editor.State.TimeSelect.EndPoint)) <= Theme.PEDeltaMouseOver Then
            Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine
        ElseIf Math.Abs(e.Y - VPositionToPanelY(Editor.State.TimeSelect.HalfPoint)) <= Theme.PEDeltaMouseOver Then
            Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine
        ElseIf Math.Abs(e.Y - VPositionToPanelY(Editor.State.TimeSelect.StartPoint)) <= Theme.PEDeltaMouseOver Then
            Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine
        End If
    End Sub

    Private Sub OnSelectModeLeftClick(e As MouseEventArgs, clickedNoteIndex As Integer)
        Dim Notes = Editor.Notes

        If clickedNoteIndex >= 0 And e.Clicks = 2 Then
            DoubleClickNoteIndex(clickedNoteIndex)
        ElseIf clickedNoteIndex > 0 Then
            'KMouseDown = -1
            Editor.ClearSelectionArray()


            'KMouseDown = xITemp
            Notes(clickedNoteIndex).TempMouseDown = True

            If My.Computer.Keyboard.CtrlKeyDown And Not ModifierMultiselectActive() Then
                Editor.State.IsDuplicatingSelectedNotes = True
            ElseIf ModifierMultiselectActive() Then
                For i = 0 To UBound(Notes)
                    If IsNoteVisible(i) Then
                        If Editor.IsLabelMatch(Notes(i), clickedNoteIndex) Then
                            Notes(i).Selected = Not Notes(i).Selected
                        End If
                    End If
                Next
            Else
                ' az description: If the clicked note is not selected, select only this one.
                'Otherwise, we clicked an already selected note
                'and we should rebuild the selected note array.
                If Not Notes(clickedNoteIndex).Selected Then
                    For i = 0 To UBound(Notes)
                        If Notes(i).Selected Then Notes(i).Selected = False
                    Next
                    Notes(clickedNoteIndex).Selected = True
                End If

                Dim SelectedCount As Integer = Editor.GetSelectedNotes().Count()

                ' adjustsingle if selectedcount is 1
                Editor.State.NT.IsAdjustingSingleNote = SelectedCount = 1
                ' Editor.RegenerateSelectedNotesArray()

                Editor.State.uAdded = False
            End If
        Else ' NoteIndex <= 0
            Editor.ClearSelectionArray()
            Editor.State.Mouse.LastMouseDownLocation = e.Location

            If Not My.Computer.Keyboard.CtrlKeyDown Then
                For i = 0 To UBound(Notes)
                    Notes(i).Selected = False
                    Notes(i).TempSelected = False
                Next
            Else
                For i = 0 To UBound(Notes)
                    Notes(i).TempSelected = Notes(i).Selected
                Next
            End If
        End If
    End Sub

    ' Handles a double click on a note in select mode.
    Private Sub DoubleClickNoteIndex(clickedNoteIndex As Integer)
        Dim note As Note = Editor.Notes(clickedNoteIndex)
        Dim noteColumn As Integer = note.ColumnIndex

        If Editor.Columns.IsColumnNumeric(noteColumn) Then
            'BPM/Stop prompt
            Dim xMessage As String = Strings.Messages.PromptEnterNumeric
            If noteColumn = ColumnType.BPM Then xMessage = Strings.Messages.PromptEnterBPM
            If noteColumn = ColumnType.STOPS Then xMessage = Strings.Messages.PromptEnterSTOP
            If noteColumn = ColumnType.SCROLLS Then xMessage = Strings.Messages.PromptEnterSCROLL


            Dim valstr As String = InputBox(xMessage, Me.Text)
            Dim PromptValue As Double = Val(valstr) * 10000
            If (noteColumn = ColumnType.SCROLLS And valstr = "0") Or PromptValue <> 0 Then

                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = Nothing
                RedoRelabelNote(note, PromptValue, xUndo, xRedo)
                If clickedNoteIndex = 0 Then
                    Editor.THBPM.Value = PromptValue / 10000
                Else
                    Editor.Notes(clickedNoteIndex).Value = PromptValue
                End If
                Editor.AddUndoChain(xUndo, xRedo)
            End If
        Else
            'Label prompt
            Dim xStr As String = UCase(Trim(InputBox(Strings.Messages.PromptEnter, Me.Text)))

            If Len(xStr) = 0 Then Return

            If IsBase36(xStr) And Not (xStr = "00" Or xStr = "0") Then
                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = Nothing
                RedoRelabelNote(note, C36to10(xStr) * 10000, xUndo, xRedo)
                Editor.Notes(clickedNoteIndex).Value = C36to10(xStr) * 10000
                Editor.AddUndoChain(xUndo, xRedo)
                Return
            Else
                MsgBox(Strings.Messages.InvalidLabel, MsgBoxStyle.Critical, Strings.Messages.Err)
            End If

        End If
    End Sub

    Private Function MouseInNote(e As MouseEventArgs, note As Note) As Boolean
        Return e.X >= HPositionToPanelX(Editor.Columns.GetColumnLeft(note.ColumnIndex)) + 1 And
               e.X <= HPositionToPanelX(Editor.Columns.GetColumnRight(note.ColumnIndex)) - 1 And
               e.Y >= VPositionToPanelY(note.VPosition + IIf(Editor.NTInput, note.Length, 0)) - Theme.kHeight And
               e.Y <= VPositionToPanelY(note.VPosition)
    End Function


    Public Sub PMainInMouseMove(ByVal sender As Panel)
        Dim p As Point = sender.PointToClient(Cursor.Position)
        PMainInMouseMove(sender, New MouseEventArgs(MouseButtons.None, 0, p.X, p.Y, 0))
    End Sub

    Public Sub PMainInMouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseMove
        Editor.State.Mouse.MouseMoveStatus = e.Location

        Dim Notes = Editor.Notes

        Select Case e.Button
            Case MouseButtons.None
                'If K Is Nothing Then Exit Select
                If Editor.State.Mouse.MiddleButtonClicked Then Exit Select

                If Editor.IsFullscreen Then
                    Editor.SetToolstripVisible(e.Y > 5)
                End If

                Dim mouseRemainInSameRegion = False
                Dim foundNoteIndex = UpdateNTInputState(e, Notes, mouseRemainInSameRegion)

                If Editor.IsSelectMode Then

                    If mouseRemainInSameRegion Then Exit Select
                    Editor.State.Mouse.CurrentHoveredNoteIndex = -1

                    UpdateTimeSelectLineOver(e)

                    Editor.State.Mouse.CurrentHoveredNoteIndex = foundNoteIndex

                ElseIf Editor.IsWriteMode Then
                    Editor.UpdateMouseRowAndColumn()

                    Editor.LNDisplayLength = 0
                    If foundNoteIndex > -1 Then
                        Editor.LNDisplayLength = Notes(foundNoteIndex).Length
                    End If

                    Editor.RefreshPanelAll()
                End If

            Case MouseButtons.Left
                If Editor.tempFirstMouseDown And Not Editor.IsTimeSelectMode Then Exit Select

                Editor.State.Mouse.tempX = 0
                Editor.State.Mouse.tempY = 0
                If e.X < 0 Or e.X > Width Or e.Y < 0 Or e.Y > Height Then
                    If e.X < 0 Then Editor.State.Mouse.tempX = e.X
                    If e.X > Width Then Editor.State.Mouse.tempX = e.X - Width
                    If e.Y < 0 Then Editor.State.Mouse.tempY = e.Y
                    If e.Y > Height Then
                    Else
                        Editor.Timer1.Enabled = False
                    End If

                    If Editor.IsSelectMode Then

                        Editor.State.Mouse.pMouseMove = e.Location

                        'If K Is Nothing Then RefreshPanelAll() : Exit Select

                        If Not Editor.State.Mouse.LastMouseDownLocation = New Point(-1, -1) Then
                            UpdateSelectionBox()

                            'ElseIf Not KMouseDown = -1 Then
                        ElseIf Editor.GetSelectedNotes().Count() <> 0 Then
                            UpdateSelectedNotes(e)
                        ElseIf Editor.State.IsDuplicatingSelectedNotes Then
                            OnDuplicateSelectedNotes(e)
                        End If

                    ElseIf Editor.IsWriteMode Then

                        If Editor.NTInput Then
                            OnWriteModeMouseMove()
                        Else
                            Editor.State.Mouse.CurrentMouseRow = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)
                            Editor.State.Mouse.CurrentMouseColumn = GetColumnAtEvent(e)
                        End If

                    ElseIf Editor.IsTimeSelectMode Then
                        OnTimeSelectClick(e)
                    End If
                End If


            Case MouseButtons.Middle
                OnPanelMousePan(e)
        End Select

        Dim col = GetColumnAtEvent(e)
        Dim vps = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)
        If vps <> lastVPos Or col <> lastColumn Then
            lastVPos = vps
            lastColumn = col
            Editor.POStatusRefresh()
            Editor.RefreshPanelAll() 'az: refreshing the line is important now...
        End If

    End Sub

    ''' <summary>
    ''' Updates NT Input state and calculates hovered note.
    ''' </summary>
    ''' <param name="e">Mouse event</param>
    ''' <param name="Notes">Notes array to check</param>
    ''' <returns>The index of the note that the event's coordinates hovers above.</returns>
    Private Function UpdateNTInputState(e As MouseEventArgs, Notes() As Note, ByRef xMouseRemainInSameRegion As Boolean) As Integer
        Dim foundNoteIndex = -1

        For noteIndex = UBound(Notes) To 0 Step -1
            If MouseInNote(e, Notes(noteIndex)) Then
                foundNoteIndex = noteIndex

                xMouseRemainInSameRegion = foundNoteIndex = Editor.State.Mouse.CurrentHoveredNoteIndex
                If Editor.NTInput Then
                    Dim vy = VPositionToPanelY(Notes(noteIndex).VPosition + Notes(noteIndex).Length)

                    Dim xbAdjustUpper As Boolean = (e.Y <= vy) And ModifierLongNoteActive()
                    Dim xbAdjustLength As Boolean = (e.Y >= vy - Theme.kHeight Or xbAdjustUpper) And ModifierLongNoteActive()

                    xMouseRemainInSameRegion = xMouseRemainInSameRegion And
                                               xbAdjustUpper = Editor.State.NT.IsAdjustingUpperEnd And
                                               xbAdjustLength = Editor.State.NT.IsAdjustingNoteLength

                    Editor.State.NT.IsAdjustingUpperEnd = xbAdjustUpper
                    Editor.State.NT.IsAdjustingNoteLength = xbAdjustLength
                End If

                Exit For
            End If
        Next

        Return foundNoteIndex
    End Function

    Dim lastVPos = -1
    Dim lastColumn = -1

    Private Sub UpdateSelectedNotes(e As MouseEventArgs)
        Dim currentClickedNoteIndex As Integer

        For i = 1 To Editor.Notes.Length - 1
            If Editor.Notes(i).TempMouseDown Then
                currentClickedNoteIndex = i
                Exit For
            End If
        Next

        Dim clickedNote = Editor.Notes(currentClickedNoteIndex)

        Dim mouseVPosition = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)

        If Editor.State.NT.IsAdjustingNoteLength And
            Editor.State.NT.IsAdjustingSingleNote Then

            If Editor.State.NT.IsAdjustingUpperEnd AndAlso
                mouseVPosition < clickedNote.VPosition Then

                Editor.State.NT.IsAdjustingUpperEnd = False
                clickedNote.VPosition += clickedNote.Length
                clickedNote.Length *= -1

            ElseIf Not Editor.State.NT.IsAdjustingUpperEnd AndAlso
                   mouseVPosition > clickedNote.VPosition + clickedNote.Length Then

                Editor.State.NT.IsAdjustingUpperEnd = True
                clickedNote.VPosition += clickedNote.Length
                clickedNote.Length *= -1
            End If

        End If

        'If moving
        If Not Editor.State.NT.IsAdjustingNoteLength Then
            OnSelectModeMoveNotes(e, clickedNote)

        ElseIf Editor.State.NT.IsAdjustingUpperEnd Then
            Dim dVPosition = mouseVPosition - clickedNote.VPosition - clickedNote.Length  'delta Length
            '< 0 means shorten, > 0 means lengthen

            OnAdjustUpperEnd(dVPosition)

        Else    'If adjusting lower end
            Dim dVPosition = mouseVPosition - clickedNote.VPosition  'delta VPosition
            '> 0 means shorten, < 0 means lengthen

            OnAdjustLowerEnd(dVPosition)
        End If

        Editor.ValidateNotesArray()
    End Sub

    Private Sub OnPanelMousePan(e As MouseEventArgs)
        If Editor.MiddleButtonMoveMethod = 1 Then
            Dim Mouse = Editor.State.Mouse
            Dim i As Integer = Mouse.tempV + (Mouse.tempY - e.Y) / Editor.Grid.HeightScale
            Dim j As Integer = Mouse.tempH + (Mouse.tempX - e.X) / Editor.Grid.WidthScale
            If i > 0 Then i = 0
            If j < 0 Then j = 0

            If i < VerticalScrollBar.Minimum Then
                i = VerticalScrollBar.Minimum
            End If
            VerticalScrollBar.Value = i

            With HorizontalScrollBar
                If j > .Maximum - .LargeChange + 1 Then
                    j = .Maximum - .LargeChange + 1
                End If

                .Value = j
            End With


        End If
    End Sub

    Private Sub OnTimeSelectClick(e As MouseEventArgs)
        Dim i As Integer
        Dim hoverNoteIndex As Integer = -1
        Dim note As Note = Nothing
        If Editor.Notes IsNot Nothing Then
            For i = Editor.Notes.Length - 1 To 0 Step -1
                If MouseInNote(e, Editor.Notes(i)) Then
                    hoverNoteIndex = i
                    note = Editor.Notes(i)
                    Exit For
                End If
            Next
        End If

        Dim snap = Editor.Grid.IsSnapEnabled And Not My.Computer.Keyboard.CtrlKeyDown
        If Not Editor.State.TimeSelect.Adjust Then
            Dim mouseRow As Double = Editor.GetMouseVPosition(snap)
            If hoverNoteIndex >= 0 Then mouseRow = note.VPosition

            If Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine Then
                Dim StartPointDY = Editor.State.TimeSelect.StartPoint - mouseRow
                Editor.State.TimeSelect.EndPointLength += StartPointDY
                Editor.State.TimeSelect.HalfPointLength += StartPointDY
                Editor.State.TimeSelect.StartPoint = mouseRow

            ElseIf Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine Then
                Editor.State.TimeSelect.HalfPoint = mouseRow

            ElseIf Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine Then
                Editor.State.TimeSelect.EndPoint = mouseRow
            Else
                Editor.State.TimeSelect.EndPoint = mouseRow
                Editor.State.TimeSelect.HalfPointLength = Editor.State.TimeSelect.EndPointLength / 2
            End If

            Editor.State.TimeSelect.ValidateSelection(Editor.GetMaxVPosition())

        Else
            Dim xL1 As Double = Editor.GetMouseVPosition(snap)

            If Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine Then
                Editor.State.TimeSelect.StartPoint = Editor.State.TimeSelect.PStart
                Editor.State.TimeSelect.EndPointLength = Editor.State.TimeSelect.PLength
                Editor.State.TimeSelect.HalfPointLength = Editor.State.TimeSelect.PHalf
                Editor.Notes = Editor.State.TimeSelect.K
                ReDim Preserve Editor.Notes(Editor.Notes.Length - 1)

                Editor.BPMChangeHalf(xL1 - Editor.State.TimeSelect.HalfPointLength - Editor.State.TimeSelect.StartPoint, , True)


            ElseIf Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine Or Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine Then
                Editor.State.TimeSelect.StartPoint = Editor.State.TimeSelect.PStart
                Editor.State.TimeSelect.EndPointLength = Editor.State.TimeSelect.PLength
                Editor.State.TimeSelect.HalfPointLength = Editor.State.TimeSelect.PHalf
                Editor.Notes = Editor.State.TimeSelect.K
                ReDim Preserve Editor.Notes(UBound(Editor.Notes))

                Editor.BPMChangeTop(
                    IIf(Editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine,
                        xL1 - Editor.State.TimeSelect.StartPoint,
                        Editor.State.TimeSelect.EndPoint - xL1) / Editor.State.TimeSelect.EndPointLength, , True)


            Else
                Editor.State.TimeSelect.EndPointLength = xL1
                Editor.State.TimeSelect.ValidateSelection(Editor.GetMaxVPosition())
            End If
        End If

        Dim Notes = Editor.Notes

        If Editor.State.TimeSelect.EndPointLength Then
            Dim xVLower As Double = Math.Min(Editor.State.TimeSelect.StartPoint, Editor.State.TimeSelect.EndPoint)
            Dim xVUpper As Double = Math.Max(Editor.State.TimeSelect.StartPoint, Editor.State.TimeSelect.EndPoint)

            If Editor.NTInput Then
                For j As Integer = 1 To UBound(Notes)
                    Notes(j).Selected = Notes(j).VPosition < xVUpper And
                                        Notes(j).VPosition + Notes(j).Length >= xVLower And
                                        Editor.Columns.nEnabled(Notes(j).ColumnIndex)
                Next
            Else
                For j As Integer = 1 To UBound(Notes)
                    Notes(j).Selected = Notes(j).VPosition >= xVLower And
                                        Notes(j).VPosition < xVUpper And
                                        Editor.Columns.nEnabled(Notes(j).ColumnIndex)
                Next
            End If
        Else
            Editor.DeselectAllNotes()
        End If

    End Sub

    Private Sub OnAdjustUpperEnd(dVPosition As Double)
        Dim minLength As Double = 0
        Dim maxHeight As Double = 191999
        Dim Notes = Editor.Notes
        For Each note In Notes.Skip(1).Where(Function(x) x.Selected)

            If note.Length + dVPosition < minLength Then
                minLength = note.Length + dVPosition
            End If

            If note.Length + note.VPosition + dVPosition > maxHeight Then
                maxHeight = note.Length + note.VPosition + dVPosition
            End If
        Next
        maxHeight -= 191999 ' az: what. WHAT WHY

        'declare undo variables
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        'start moving
        ' az refactoring: this... was looking at the global notes and addressing the selected note array.
        ' I get why ideally you'd use it (as a sort of optimization) but that was not in use here.
        ' So... Maybe in the future use that array?
        For Each note In Editor.Notes.Skip(1).Where(Function(x) x.Selected)
            Dim xLen = note.Length + dVPosition - minLength - maxHeight
            RedoLongNoteModify(note, note.VPosition, xLen, xUndo, xRedo)

            note.Length = xLen
        Next

        'Add undo
        If dVPosition - minLength - maxHeight <> 0 Then
            Editor.AddUndoChain(xUndo, xBaseRedo.Next, Editor.State.uAdded)
            If Not Editor.State.uAdded Then Editor.State.uAdded = True
        End If
    End Sub


    Private Sub OnAdjustLowerEnd(dVPosition As Double)
        Dim minLength As Double = 0
        Dim minVPosition As Double = 0
        Dim Notes = Editor.Notes
        For Each note In Notes.Skip(1).Where(Function(x) x.Selected)
            If note.Length - dVPosition < minLength Then
                minLength = note.Length - dVPosition
            End If
            If note.VPosition + dVPosition < minVPosition Then
                minVPosition = note.VPosition + dVPosition
            End If
        Next

        'declare undo variables
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        'start moving
        ' az refactoring: same as OnAdjustUpperEnd. Why use SelectedNotes here.

        For Each note In Notes.Where(Function(x) x.Selected)
            Dim newVPos = note.VPosition + dVPosition + minLength - minVPosition
            Dim newLen = note.Length - dVPosition - minLength + minVPosition
            RedoLongNoteModify(note, newVPos, newLen, xUndo, xRedo)

            note.VPosition = newVPos
            note.Length = newLen
        Next

        'Add undo
        If dVPosition + minLength - minVPosition <> 0 Then
            Editor.AddUndoChain(xUndo, xBaseRedo.Next, Editor.State.uAdded)
            If Not Editor.State.uAdded Then Editor.State.uAdded = True
        End If
    End Sub

    Private Sub OnDuplicateSelectedNotes(e As MouseEventArgs)
        Dim Notes = Editor.Notes
        Dim tempNoteIndex As Integer
        For tempNoteIndex = 1 To UBound(Notes)
            If Notes(tempNoteIndex).TempMouseDown Then Exit For
        Next

        Dim mouseVPosition = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)
        If Editor.DisableVerticalMove Then
            mouseVPosition = Notes(tempNoteIndex).VPosition
        End If

        Dim dVPosition As Double = mouseVPosition - Notes(tempNoteIndex).VPosition  'delta VPosition

        Dim currCol = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(GetColumnAtEvent(e))
        Dim noteCol = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(Notes(tempNoteIndex).ColumnIndex)
        Dim colChange As Integer = currCol - noteCol 'delta Column

        'Ks cannot be beyond the left, the upper and the lower boundary
        Dim dstColumn As Integer = 0
        Dim mVPosition As Double = 0
        Dim muVPosition As Double = 191999
        For Each note In Notes.Skip(1).Where(Function(x) x.Selected)

            If Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + colChange < dstColumn Then
                dstColumn = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + colChange
            End If

            If note.VPosition + dVPosition < mVPosition Then
                mVPosition = note.VPosition + dVPosition
            End If

            Dim noteTailPos = note.VPosition + IIf(Editor.NTInput, note.Length, 0)
            If noteTailPos + dVPosition > muVPosition Then
                muVPosition = noteTailPos + dVPosition
            End If

        Next
        muVPosition -= 191999

        'If not moving then exit
        If (Not Editor.State.SelectedNotesWereDuplicated) And
            colChange - dstColumn = 0 And
            dVPosition - mVPosition - muVPosition = 0 Then
            Return
        End If

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If Not Editor.State.SelectedNotesWereDuplicated Then     'If Editor.State.uAdded = False
            DuplicateSelectedNotes(tempNoteIndex, dVPosition, colChange, dstColumn, mVPosition, muVPosition)
            Editor.State.SelectedNotesWereDuplicated = True

        Else
            For Each note In Editor.GetSelectedNotes()
                Dim idx = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + colChange - dstColumn
                note.ColumnIndex = Editor.Columns.EnabledColumnIndexToColumnArrayIndex(idx)
                note.VPosition = note.VPosition + dVPosition - mVPosition - muVPosition
                RedoAddNote(note, xUndo, xRedo)
            Next

            Editor.AddUndoChain(xUndo, xBaseRedo.Next, True)
        End If

        Editor.ValidateNotesArray()
    End Sub


    Private Sub OnWriteModeMouseMove()
        'If Not KMouseDown = -1 Then
        Dim selectedNotes = Editor.GetSelectedNotes().ToArray()

        If selectedNotes.Count() <> 0 Then
            Dim note As Note = Editor.Notes.
                                      Skip(1).
                                      Where(Function(x) x.TempMouseDown).
                                      First()

            Dim mouseVPosition = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)
            Dim maxVPos = Editor.GetMaxVPosition()
            Dim NT = Editor.State.NT

            With note
                If NT.IsAdjustingUpperEnd AndAlso
                   mouseVPosition < .VPosition Then

                    NT.IsAdjustingUpperEnd = False
                    .VPosition += .Length
                    .Length *= -1

                ElseIf Not NT.IsAdjustingUpperEnd AndAlso
                       mouseVPosition > .VPosition + .Length Then

                    NT.IsAdjustingUpperEnd = True
                    .VPosition += .Length
                    .Length *= -1

                End If

                If NT.IsAdjustingUpperEnd Then
                    .Length = mouseVPosition - .VPosition
                Else
                    .Length = .VPosition + .Length - mouseVPosition
                    .VPosition = mouseVPosition
                End If

                If .VPosition < 0 Then .Length += .VPosition : .VPosition = 0
                If .VPosition + .Length >= maxVPos Then .Length = maxVPos - 1 - .VPosition

                If selectedNotes(0).LNPair = -1 Then 'If new note
                    Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                    Dim xRedo As UndoRedo.LinkedURCmd = Nothing
                    RedoAddNote(note, xUndo, xRedo)
                    Editor.AddUndoChain(xUndo, xRedo, True)

                Else 'If existing note
                    Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                    Dim xRedo As UndoRedo.LinkedURCmd = Nothing
                    RedoLongNoteModify(selectedNotes.First(), .VPosition, .Length, xUndo, xRedo)
                    Editor.AddUndoChain(xUndo, xRedo, True)
                End If

                Editor.State.Mouse.CurrentMouseColumn = .ColumnIndex
                Editor.State.Mouse.CurrentMouseRow = mouseVPosition
                Editor.LNDisplayLength = .Length

            End With

            Editor.ValidateNotesArray()

        End If
    End Sub

    Private Sub OnSelectModeMoveNotes(e As MouseEventArgs, currentNote As Note)
        Dim mouseVPosition = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)

        If Editor.DisableVerticalMove Then
            mouseVPosition = Editor.GetSelectedNotes().First().VPosition
        End If

        Dim dVPosition = mouseVPosition - currentNote.VPosition  'delta VPosition

        Dim mouseColumn As Integer
        Dim i = 0
        Dim mLeft As Integer = e.X / Editor.Grid.WidthScale + HorizontalScroll 'horizontal position of the mouse
        If mLeft >= 0 Then
            Do
                If mLeft < Editor.Columns.GetColumnLeft(i + 1) Or i >= Editor.Columns.ColumnCount Then
                    mouseColumn = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(i)
                    Exit Do 'get the column where mouse is 
                End If
                i += 1
            Loop
        End If

        'get the enabled delta column where mouse is 
        Dim dColumn = mouseColumn - Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(currentNote.ColumnIndex)

        'Ks cannot be beyond the left, the upper and the lower boundary
        mLeft = 0
        Dim mVPosition As Double = 0
        Dim muVPosition As Double = 191999
        For Each note In Editor.GetSelectedNotes()
            Dim idx = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex)
            mLeft = Math.Min(idx + dColumn, mLeft)
            mVPosition = Math.Min(note.VPosition + dVPosition, mVPosition)
            muVPosition = Math.Max(note.VPosition + IIf(Editor.NTInput, note.Length, 0) + dVPosition, muVPosition)
        Next
        muVPosition -= 191999 ' az: this magic number again. why?

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        'start moving
        For Each note In Editor.GetSelectedNotes()
            Dim idx = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + dColumn - mLeft
            Dim newColumn = Editor.Columns.EnabledColumnIndexToColumnArrayIndex(idx)
            Dim newVPosition = note.VPosition + dVPosition - mVPosition - muVPosition
            RedoMoveNote(note, newColumn, newVPosition, xUndo, xRedo)

            note.ColumnIndex = newColumn
            note.VPosition = newVPosition
        Next

        Editor.AddUndoChain(xUndo, xBaseRedo.Next, Editor.State.uAdded)
        If Not Editor.State.uAdded Then Editor.State.uAdded = True

        'End If
    End Sub

    Private Sub UpdateSelectionBox()
        Dim pMouseMove = Editor.State.Mouse.pMouseMove
        Dim LastMouseDownLocation = Editor.State.Mouse.LastMouseDownLocation
        Dim SelectionBox As New Rectangle(Math.Min(pMouseMove.X, LastMouseDownLocation.X),
                                          Math.Min(pMouseMove.Y, LastMouseDownLocation.Y),
                                          Math.Abs(pMouseMove.X - LastMouseDownLocation.X),
                                          Math.Abs(pMouseMove.Y - LastMouseDownLocation.Y))
        Dim NoteRect As Rectangle

        Dim Columns = Editor.Columns
        For Each note In Editor.Notes.Skip(1)
            Dim noteStartX = Columns.GetColumnLeft(note.ColumnIndex)
            Dim noteTailPos = note.VPosition + IIf(Editor.NTInput, note.Length, 0)
            Dim drawWidth = Columns.GetColumnWidth(note.ColumnIndex) * Editor.Grid.WidthScale - 2
            Dim drawHeight = Theme.kHeight + IIf(Editor.NTInput, note.Length * Editor.Grid.HeightScale, 0)
            NoteRect = New Rectangle(HPositionToPanelX(noteStartX) + 1,
                                    VPositionToPanelY(noteTailPos) - Theme.kHeight,
                                    drawWidth,
                                    drawHeight)


            Dim columnEnabled = Columns.nEnabled(note.ColumnIndex)

            If NoteRect.IntersectsWith(SelectionBox) Then
                note.Selected = Not note.TempSelected And
                                columnEnabled
            Else
                note.Selected = note.TempSelected And
                                columnEnabled
            End If
        Next
    End Sub

    Private Sub DuplicateSelectedNotes(tempNoteIndex As Integer, dVPosition As Double, dColumn As Integer, mLeft As Integer, mVPosition As Double, muVPosition As Double)
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        Dim Notes = Editor.Notes

        Notes(tempNoteIndex).Selected = True

        Dim selectedNotes = Editor.GetSelectedNotes().ToArray()
        Dim selectedNotesCount = selectedNotes.Length

        Dim duplicatedNotes(selectedNotesCount - 1) As Note
        Dim j As Integer = 0
        For Each note In selectedNotes
            Dim idx = Editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + dColumn - mLeft
            duplicatedNotes(j) = note.Clone
            duplicatedNotes(j).ColumnIndex = Editor.Columns.EnabledColumnIndexToColumnArrayIndex(idx)
            duplicatedNotes(j).VPosition = note.VPosition + dVPosition - mVPosition - muVPosition
            RedoAddNote(duplicatedNotes(j), xUndo, xRedo)

            note.Selected = False
            j += 1
        Next
        Notes(tempNoteIndex).TempMouseDown = False

        'copy to K
        Editor.Notes = Editor.Notes.Concat(duplicatedNotes).ToArray()

        Editor.AddUndoChain(xUndo, xBaseRedo.Next)
    End Sub

    'Private Sub DrawNoteHoverHighlight(foundNoteIndex As Integer)
    '    Dim xDispX As Integer = HPositionToPanelX(GetColumnLeft(Notes(foundNoteIndex).ColumnIndex))
    '    Dim xDispY As Integer = IIf(Not NTInput Or (IsAdjustingNoteLength And Not IsAdjustingUpperEnd),
    '                                VPositionToPanelY(Notes(foundNoteIndex).VPosition, VerticalScroll, xHeight) - vo.kHeight - 1,
    '                                VPositionToPanelY(Notes(foundNoteIndex).VPosition + Notes(foundNoteIndex).Length, VerticalScroll, xHeight) - vo.kHeight - 1)
    '    Dim xDispW As Integer = GetColumnWidth(Notes(foundNoteIndex).ColumnIndex) * Editor.Grid.WidthScale + 1
    '    Dim xDispH As Integer = IIf(Not NTInput Or IsAdjustingNoteLength,
    '                                vo.kHeight + 3,
    '                                Notes(foundNoteIndex).Length * Editor.Grid.HeightScale + vo.kHeight + 3)

    '    Dim e1 As BufferedGraphics = BufferedGraphicsManager.Current.Allocate(spMain(iI).CreateGraphics, New Rectangle(xDispX, xDispY, xDispW, xDispH))
    '    e1.Graphics.FillRectangle(vo.Bg, New Rectangle(xDispX, xDispY, xDispW, xDispH))

    '    If NTInput Then DrawNoteNT(Notes(foundNoteIndex), e1, HorizontalScroll, VerticalScroll, xHeight) Else DrawNote(Notes(foundNoteIndex), e1, HorizontalScroll, VerticalScroll, xHeight)

    '    e1.Graphics.DrawRectangle(IIf(IsAdjustingNoteLength, vo.kMouseOverE, vo.kMouseOver), xDispX, xDispY, xDispW - 1, xDispH - 1)

    'End Sub

    Private Function GetColumnAtEvent(e As MouseEventArgs)
        Return GetColumnAtX(e.X)
    End Function

    Private Sub PMainInMouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseUp
        Editor.State.Mouse.tempX = 0
        Editor.State.Mouse.tempY = 0
        Editor.State.Mouse.tempV = 0
        Editor.State.Mouse.tempH = 0
        LastVerticalScroll = -1
        LastHorizontalScroll = -1
        Editor.Timer1.Enabled = False

        Dim iI As Integer = sender.Tag

        Dim dv = (Editor.State.Mouse.MiddleButtonLocation - Cursor.Position)
        If Editor.State.Mouse.MiddleButtonClicked AndAlso
            e.Button = Windows.Forms.MouseButtons.Middle AndAlso
            dv.X ^ 2 + dv.Y ^ 2 >= Theme.MiddleDeltaRelease Then
            Editor.State.Mouse.MiddleButtonClicked = False
        End If

        If Editor.IsSelectMode Then
            Editor.State.Mouse.LastMouseDownLocation = New Point(-1, -1)
            Editor.State.Mouse.pMouseMove = New Point(-1, -1)

            If Editor.State.IsDuplicatingSelectedNotes And
                Not Editor.State.SelectedNotesWereDuplicated And
                Not ModifierMultiselectActive() Then

                For Each note In Editor.Notes
                    If note.TempMouseDown Then
                        note.Selected = Not note.Selected
                        Exit For
                    End If
                Next
            End If

            Editor.State.IsDuplicatingSelectedNotes = False
            Editor.State.SelectedNotesWereDuplicated = False

        ElseIf Editor.IsWriteMode Then

            If Not Editor.NTInput And Not Editor.tempFirstMouseDown Then
                Dim xVPosition = Editor.GetMouseVPosition(Editor.Grid.IsSnapEnabled)

                Dim xColumn = GetColumnAtEvent(e)
                Dim Notes = Editor.Notes

                If e.Button = Windows.Forms.MouseButtons.Left Then
                    Dim HiddenNote As Boolean = ModifierHiddenActive()
                    Dim LongNote As Boolean = ModifierLongNoteActive()
                    Dim Landmine As Boolean = ModifierLandmineActive()
                    Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                    Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                    Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                    If Editor.Columns.IsColumnNumeric(xColumn) Then
                        Dim xMessage As String = Strings.Messages.PromptEnterNumeric
                        If xColumn = ColumnType.BPM Then xMessage = Strings.Messages.PromptEnterBPM
                        If xColumn = ColumnType.STOPS Then xMessage = Strings.Messages.PromptEnterSTOP
                        If xColumn = ColumnType.SCROLLS Then xMessage = Strings.Messages.PromptEnterSCROLL

                        Dim valstr As String = InputBox(xMessage, Me.Text)
                        Dim value As Long = Val(valstr) * 10000

                        If (xColumn = ColumnType.SCROLLS And valstr = "0") Or value <> 0 Then
                            For i = 1 To UBound(Notes)
                                If Notes(i).VPosition = xVPosition AndAlso Notes(i).ColumnIndex = xColumn Then _
                            RedoRemoveNote(Notes(i), xUndo, xRedo)
                            Next

                            Dim n = New Note(xColumn, xVPosition, value, LongNote, HiddenNote)
                            RedoAddNote(n, xUndo, xRedo)
                            Editor.AddNote(n)

                            Editor.AddUndoChain(xUndo, xBaseRedo.Next)
                        End If

                    Else
                        Dim xValue As Integer = Editor.CurrentWavSelectedIndex * 10000

                        For i = 1 To UBound(Notes)
                            If Notes(i).VPosition = xVPosition AndAlso Notes(i).ColumnIndex = xColumn Then _
                            RedoRemoveNote(Notes(i), xUndo, xRedo)
                        Next

                        Dim n = New Note(xColumn, xVPosition, xValue,
                                         LongNote, HiddenNote, True, Landmine)

                        RedoAddNote(n, xUndo, xRedo)
                        Editor.AddNote(n)

                        Editor.AddUndoChain(xUndo, xRedo)
                    End If
                End If
            End If

            Editor.State.Mouse.CurrentMouseRow = -1
            Editor.State.Mouse.CurrentMouseColumn = -1
        End If

        ' az refactoring: Not a full note refresh?
        Editor.CalculateGreatestVPosition()
        Editor.RefreshPanelAll()
    End Sub

    Private Sub PMainInMouseWheel(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseWheel
        If Editor.State.Mouse.MiddleButtonClicked Then
            Editor.State.Mouse.MiddleButtonClicked = False
        End If

        Dim i = VerticalScroll - Math.Sign(e.Delta) * Editor.Grid.WheelScroll
        VerticalScrollBar.Value = Clamp(i, VerticalScrollBar.Minimum, 0)
    End Sub
End Class
