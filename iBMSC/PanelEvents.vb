Imports System.Linq
Imports iBMSC.Editor

Partial Public Class EditorPanel
    Private Sub PreviewKeyDownEvent(sender As Object, e As PreviewKeyDownEventArgs) Handles Me.PreviewKeyDown
        If e.KeyCode = Keys.ShiftKey Or e.KeyCode = Keys.ControlKey Then
            _editor.RefreshPanelAll()
            _editor.PoStatusRefresh()
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
                VerticalScrollBar.Value = Math.Max(VerticalScrollBar.Value - _editor.Grid.PageUpDnScroll,
                                                   VerticalScrollBar.Minimum)

            Case Keys.PageDown
                VerticalScrollBar.Value = Math.Min(VerticalScrollBar.Value + _editor.Grid.PageUpDnScroll, 0)
        End Select

        OnMouseMove(sender)
    End Sub

    Private Sub ResizeEvent(sender As Object, e As EventArgs) Handles Me.Resize
        If Not Created Then Exit Sub

        If VerticalScrollBar Is Nothing OrElse
           HorizontalScrollBar Is Nothing OrElse
           _editor Is Nothing Then
            Exit Sub
        End If

        VerticalScrollBar.LargeChange = sender.Height * 0.9
        VerticalScrollBar.Maximum = VerticalScrollBar.LargeChange - 1
        HorizontalScrollBar.LargeChange = sender.Width / _editor.Grid.WidthScale

        If HorizontalScrollBar.Value > HorizontalScrollBar.Maximum - HorizontalScrollBar.LargeChange + 1 Then
            HorizontalScrollBar.Value = HorizontalScrollBar.Maximum - HorizontalScrollBar.LargeChange + 1
        End If

        Refresh()
    End Sub

    Private Sub LostFocusEvent(sender As Object, e As EventArgs) Handles Me.LostFocus
        _editor.RefreshPanelAll()
    End Sub

    Private Sub MouseDownEvent(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        ' az: the hell is this for?
        _editor.TempFirstMouseDown = _editor.FirstClickDisabled And Not sender.Focused

        _editor.PanelFocus = sender.Tag
        sender.Focus()
        _editor.State.Mouse.LastMouseDownLocation = New Point(-1, -1)
        LastVerticalScroll = VerticalPosition

        If _editor.NtInput Then
            _editor.State.NT.IsAdjustingUpperEnd = False
            _editor.State.NT.IsAdjustingNoteLength = False
        End If
        _editor.State.IsDuplicatingSelectedNotes = False
        _editor.State.SelectedNotesWereDuplicated = False

        If _editor.State.Mouse.MiddleButtonClicked Then
            _editor.State.Mouse.MiddleButtonClicked = False
            Exit Sub
        End If


        Select Case e.Button
            Case MouseButtons.Left
                If _editor.TempFirstMouseDown And Not _editor.IsTimeSelectMode Then
                    _editor.RefreshPanelAll()
                    Exit Select
                End If

                _editor.State.Mouse.CurrentHoveredNoteIndex = -1
                'If K Is Nothing Then pMouseDown = e.Location : Exit Select

                'Find the clicked K
                Dim noteIndex As Integer = GetClickedNote(e)

                _editor.PanelPreviewNoteIndex(noteIndex)

                For Each note In _editor.Notes
                    note.TempMouseDown = False
                Next

                HandleCurrentModeOnClick(e, noteIndex)
                _editor.RefreshPanelAll()
                _editor.PoStatusRefresh()

            Case MouseButtons.Middle
                If _editor.MiddleButtonMoveMethod = 1 Then
                    _editor.State.Mouse.tempX = e.X
                    _editor.State.Mouse.tempY = e.Y
                    _editor.State.Mouse.tempV = VerticalPosition
                    _editor.State.Mouse.tempH = HorizontalPosition
                Else
                    _editor.State.Mouse.MiddleButtonLocation = Cursor.Position
                    _editor.State.Mouse.MiddleButtonClicked = True
                    _editor.TimerMiddle.Enabled = True
                End If

            Case MouseButtons.Right
                DeselectOrRemove(e)
        End Select
    End Sub

    Private Sub DeselectOrRemove(e As MouseEventArgs)
        _editor.State.Mouse.CurrentHoveredNoteIndex = -1

        _editor.ClearSelectionArray()

        If Not _editor.TempFirstMouseDown Then
            Dim i As Integer
            For i = _editor.Notes.Length - 1 To 1 Step -1
                Dim note = _editor.Notes(i)
                'If mouse is clicking on a K
                If MouseInNote(e, note) Then

                    If My.Computer.Keyboard.ShiftKeyDown Then
                        _editor.SelectWavFromNote(note)
                    Else
                        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                        Dim xRedo As UndoRedo.LinkedURCmd = Nothing

                        RedoRemoveNote(note, xUndo, xRedo)
                        _editor.RemoveNote(i)

                        _editor.AddUndoChain(xUndo, xRedo)
                        _editor.RefreshPanelAll()
                    End If

                    Exit For
                End If
            Next

            _editor.CalculateTotalPlayableNotes()
        End If
    End Sub

    Private Function GetClickedNote(e As MouseEventArgs) As Integer
        Dim noteIndex As Integer = -1
        For i = UBound(_editor.Notes) To 0 Step -1
            Dim note = _editor.Notes(i)

            If MouseInNote(e, note) Then
                noteIndex = i

                If _editor.NtInput And My.Computer.Keyboard.ShiftKeyDown Then
                    _editor.State.NT.IsAdjustingUpperEnd = e.Y <= VPositionToPanelY(note.VPosition + note.Length)
                    _editor.State.NT.IsAdjustingNoteLength = e.Y >= VPositionToPanelY(note.VPosition) - _theme.NoteHeight Or
                                                            _editor.State.NT.IsAdjustingUpperEnd
                End If

                Exit For

            End If
        Next

        Return noteIndex
    End Function


    Private Sub HandleCurrentModeOnClick(e As MouseEventArgs, ByRef clickedNoteIndex As Integer)
        Dim notes = _editor.Notes
        Dim mouseVPos = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)
        Dim clickedNote As Note = Nothing
        If clickedNoteIndex > 0 Then
            clickedNote = _editor.Notes(clickedNoteIndex)
        End If

        If _editor.IsSelectMode Then
            OnSelectModeLeftClick(e, clickedNoteIndex)
        ElseIf _editor.NtInput And _editor.IsWriteMode Then
            _editor.State.Mouse.CurrentMouseRow = -1
            _editor.State.Mouse.CurrentMouseColumn = -1

            If mouseVPos < 0 Or mouseVPos >= _editor.GetMaxVPosition() Then Exit Sub

            Dim col = GetColumnAtEvent(e)

            For j As Integer = UBound(_editor.Notes) To 1 Step -1
                If _editor.Notes(j).VPosition = mouseVPos And
                   _editor.Notes(j).ColumnIndex = col Then
                    clickedNoteIndex = j
                    Exit For
                End If
            Next


            Dim hidden As Boolean = ModifierHiddenActive()

            If clickedNoteIndex > 0 Then
                ' Editor.SelectSingleNote(clickedNote)

                'KMouseDown = xITemp
                clickedNote.TempMouseDown = True
                clickedNote.Length = mouseVPos - clickedNote.VPosition

                _editor.State.NT.IsAdjustingUpperEnd = True

                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = Nothing


                RedoLongNoteModify(clickedNote, clickedNote.VPosition, clickedNote.Length, xUndo, xRedo)
                _editor.AddUndoChain(xUndo, xRedo)

            ElseIf _editor.Columns.IsColumnNumeric(col) Then

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

                    For Each note In _editor.Notes
                        If note.VPosition = mouseVPos AndAlso note.ColumnIndex = col Then
                            RedoRemoveNote(note, xUndo, xRedo)
                        End If
                    Next

                    Dim n = New Note(col, mouseVPos, value, 0, hidden)
                    RedoAddNote(n, xUndo, xRedo)

                    _editor.AddNote(n)
                    _editor.AddUndoChain(xUndo, xBaseRedo.Next)
                End If

                ' ShouldDrawTempNote = True (az: Why?)

            Else
                Dim xLbl As Integer = _editor.CurrentWavSelectedIndex * 10000

                Dim landmine As Boolean = ModifierLandmineActive()

                Dim note = New Note With {
                        .VPosition = mouseVPos,
                        .ColumnIndex = col,
                        .Value = xLbl,
                        .Hidden = hidden,
                        .Landmine = landmine,
                        .TempMouseDown = True,
                        .LNPair = -1
                        }
                _editor.AppendNote(note)
            End If

            _editor.ValidateNotesArray()

        ElseIf _editor.IsTimeSelectMode Then

            If clickedNoteIndex >= 0 Then
                mouseVPos = clickedNote.VPosition
            End If

            UpdateTimeSelectLineOver(e)

            If Not _editor.State.TimeSelect.Adjust Then
                If _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine Then

                    _editor.State.TimeSelect.EndPointLength += _editor.State.TimeSelect.StartPoint - mouseVPos
                    _editor.State.TimeSelect.HalfPointLength += _editor.State.TimeSelect.StartPoint - mouseVPos
                    _editor.State.TimeSelect.StartPoint = mouseVPos

                ElseIf _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine Then
                    _editor.State.TimeSelect.HalfPointLength = mouseVPos
                    If _editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 Then
                        _editor.State.TimeSelect.HalfPointLength = SnapToGrid(_editor.State.TimeSelect.HalfPointLength)
                    End If
                    _editor.State.TimeSelect.HalfPointLength -= _editor.State.TimeSelect.StartPoint

                ElseIf _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine Then
                    _editor.State.TimeSelect.EndPointLength = mouseVPos
                    If _editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 Then
                        _editor.State.TimeSelect.EndPointLength = SnapToGrid(_editor.State.TimeSelect.EndPointLength)
                    End If
                    _editor.State.TimeSelect.EndPointLength -= _editor.State.TimeSelect.StartPoint

                Else
                    _editor.State.TimeSelect.EndPointLength = 0
                    _editor.State.TimeSelect.StartPoint = mouseVPos
                    If _editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 Then
                        _editor.State.TimeSelect.StartPoint = SnapToGrid(_editor.State.TimeSelect.StartPoint)
                    End If
                End If
                _editor.State.TimeSelect.ValidateSelection(_editor.GetMaxVPosition())

            Else
                If _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine Then
                    _editor.ValidateNotesArray()
                    _editor.State.TimeSelect.PStart = _editor.State.TimeSelect.StartPoint
                    _editor.State.TimeSelect.PLength = _editor.State.TimeSelect.EndPointLength
                    _editor.State.TimeSelect.PHalf = _editor.State.TimeSelect.HalfPointLength
                    _editor.State.TimeSelect.Notes = notes.Clone()

                    If _editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 And Not My.Computer.Keyboard.CtrlKeyDown Then _
                        mouseVPos = SnapToGrid(mouseVPos)
                    _editor.AddUndoChain(New UndoRedo.Void, New UndoRedo.Void)
                    _editor.BPMChangeHalf(
                        mouseVPos - _editor.State.TimeSelect.HalfPointLength - _editor.State.TimeSelect.StartPoint, , True)


                ElseIf _
                    _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine Or
                    _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine Then
                    _editor.ValidateNotesArray()
                    _editor.State.TimeSelect.PStart = _editor.State.TimeSelect.StartPoint
                    _editor.State.TimeSelect.PLength = _editor.State.TimeSelect.EndPointLength
                    _editor.State.TimeSelect.PHalf = _editor.State.TimeSelect.HalfPointLength
                    _editor.State.TimeSelect.Notes = notes.Clone()
                    ReDim Preserve _editor.State.TimeSelect.Notes(UBound(_editor.State.TimeSelect.Notes))

                    If _editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 And Not My.Computer.Keyboard.CtrlKeyDown Then _
                        mouseVPos = SnapToGrid(mouseVPos)
                    _editor.AddUndoChain(New UndoRedo.Void, New UndoRedo.Void)
                    Dim v = IIf(_editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine,
                                mouseVPos - _editor.State.TimeSelect.StartPoint,
                                _editor.State.TimeSelect.EndPoint - mouseVPos)

                    _editor.BPMChangeTop(v / _editor.State.TimeSelect.EndPointLength, , True)
                Else
                    _editor.State.TimeSelect.EndPointLength = mouseVPos
                    If _editor.Grid.IsSnapEnabled And clickedNoteIndex <= 0 And Not My.Computer.Keyboard.CtrlKeyDown Then _
                        _editor.State.TimeSelect.EndPointLength = SnapToGrid(_editor.State.TimeSelect.EndPointLength)
                    _editor.State.TimeSelect.EndPointLength -= _editor.State.TimeSelect.StartPoint
                End If

            End If

            If _editor.State.TimeSelect.EndPointLength Then
                Dim xVLower As Double = Math.Min(_editor.State.TimeSelect.StartPoint, _editor.State.TimeSelect.EndPoint)
                Dim xVUpper As Double = Math.Max(_editor.State.TimeSelect.StartPoint, _editor.State.TimeSelect.EndPoint)
                If _editor.NtInput Then
                    For Each note In notes.Skip(1)
                        note.Selected = Not note.VPosition >= xVUpper And
                                        Not note.VPosition + note.Length < xVLower And
                                        _editor.Columns.IsEnabled(note.ColumnIndex)
                    Next
                Else
                    For Each note In notes.Skip(1)
                        note.Selected = note.VPosition >= xVLower And
                                        note.VPosition < xVUpper And
                                        _editor.Columns.IsEnabled(note.ColumnIndex)
                    Next
                End If
            Else
                _editor.DeselectAllNotes()
            End If

        End If
    End Sub

    Private Sub UpdateTimeSelectLineOver(e As MouseEventArgs)
        _editor.State.TimeSelect.Adjust = ModifierLongNoteActive()

        _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.None
        If Math.Abs(e.Y - VPositionToPanelY(_editor.State.TimeSelect.EndPoint)) <= _theme.PEDeltaMouseOver Then
            _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine
        ElseIf Math.Abs(e.Y - VPositionToPanelY(_editor.State.TimeSelect.HalfPoint)) <= _theme.PEDeltaMouseOver Then
            _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine
        ElseIf Math.Abs(e.Y - VPositionToPanelY(_editor.State.TimeSelect.StartPoint)) <= _theme.PEDeltaMouseOver Then
            _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine
        End If
    End Sub

    Private Sub OnSelectModeLeftClick(e As MouseEventArgs, clickedNoteIndex As Integer)
        Dim notes = _editor.Notes

        If clickedNoteIndex >= 0 And e.Clicks = 2 Then
            DoubleClickNoteIndex(clickedNoteIndex)
        ElseIf clickedNoteIndex > 0 Then
            'KMouseDown = -1
            _editor.ClearSelectionArray()


            'KMouseDown = xITemp
            notes(clickedNoteIndex).TempMouseDown = True

            If My.Computer.Keyboard.CtrlKeyDown And Not ModifierMultiselectActive() Then
                _editor.State.IsDuplicatingSelectedNotes = True
            ElseIf ModifierMultiselectActive() Then
                For i = 0 To UBound(notes)
                    If IsNoteVisible(i) Then
                        If _editor.IsLabelMatch(notes(i), clickedNoteIndex) Then
                            notes(i).Selected = Not notes(i).Selected
                        End If
                    End If
                Next
            Else
                ' az description: If the clicked note is not selected, select only this one.
                'Otherwise, we clicked an already selected note
                'and we should rebuild the selected note array.
                If Not notes(clickedNoteIndex).Selected Then
                    For i = 0 To UBound(notes)
                        If notes(i).Selected Then notes(i).Selected = False
                    Next
                    notes(clickedNoteIndex).Selected = True
                End If

                Dim selectedCount As Integer = _editor.GetSelectedNotes().Count()

                ' adjustsingle if selectedcount is 1
                _editor.State.NT.IsAdjustingSingleNote = selectedCount = 1
                ' Editor.RegenerateSelectedNotesArray()

                _editor.State.OverwriteLastUndoRedoCommand = False
            End If
        Else ' NoteIndex <= 0
            _editor.ClearSelectionArray()
            _editor.State.Mouse.LastMouseDownLocation = e.Location

            If Not My.Computer.Keyboard.CtrlKeyDown Then
                For i = 0 To UBound(notes)
                    notes(i).Selected = False
                    notes(i).TempSelected = False
                Next
            Else
                For i = 0 To UBound(notes)
                    notes(i).TempSelected = notes(i).Selected
                Next
            End If
        End If
    End Sub

    ' Handles a double click on a note in select mode.
    Private Sub DoubleClickNoteIndex(clickedNoteIndex As Integer)
        Dim note As Note = _editor.Notes(clickedNoteIndex)
        Dim noteColumn As Integer = note.ColumnIndex

        If _editor.Columns.IsColumnNumeric(noteColumn) Then
            'BPM/Stop prompt
            Dim xMessage As String = Strings.Messages.PromptEnterNumeric
            If noteColumn = ColumnType.BPM Then xMessage = Strings.Messages.PromptEnterBPM
            If noteColumn = ColumnType.STOPS Then xMessage = Strings.Messages.PromptEnterSTOP
            If noteColumn = ColumnType.SCROLLS Then xMessage = Strings.Messages.PromptEnterSCROLL


            Dim valstr As String = InputBox(xMessage, Me.Text)
            Dim promptValue As Double = Val(valstr) * 10000
            If (noteColumn = ColumnType.SCROLLS And valstr = "0") Or promptValue <> 0 Then

                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = Nothing
                RedoRelabelNote(note, promptValue, xUndo, xRedo)
                If clickedNoteIndex = 0 Then
                    _editor.THBPM.Value = promptValue / 10000
                Else
                    _editor.Notes(clickedNoteIndex).Value = promptValue
                End If
                _editor.AddUndoChain(xUndo, xRedo)
            End If
        Else
            'Label prompt
            Dim xStr As String = UCase(Trim(InputBox(Strings.Messages.PromptEnter, Me.Text)))

            If Len(xStr) = 0 Then Return

            If IsBase36(xStr) And Not (xStr = "00" Or xStr = "0") Then
                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = Nothing
                RedoRelabelNote(note, C36to10(xStr) * 10000, xUndo, xRedo)
                _editor.Notes(clickedNoteIndex).Value = C36to10(xStr) * 10000
                _editor.AddUndoChain(xUndo, xRedo)
                Return
            Else
                MsgBox(Strings.Messages.InvalidLabel, MsgBoxStyle.Critical, Strings.Messages.Err)
            End If

        End If
    End Sub

    Private Function MouseInNote(e As MouseEventArgs, note As Note) As Boolean
        Return e.X >= HPositionToPanelX(_editor.Columns.GetColumnLeft(note.ColumnIndex)) + 1 And
               e.X <= HPositionToPanelX(_editor.Columns.GetColumnRight(note.ColumnIndex)) - 1 And
               e.Y >= VPositionToPanelY(note.VPosition + IIf(_editor.NtInput, note.Length, 0)) - _theme.NoteHeight And
               e.Y <= VPositionToPanelY(note.VPosition)
    End Function


    Public Sub _MouseMoveEvent(sender As Panel)
        Dim p As Point = sender.PointToClient(Cursor.Position)
        MouseMoveEvent(sender, New MouseEventArgs(MouseButtons.None, 0, p.X, p.Y, 0))
    End Sub

    Public Sub MouseMoveEvent(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        _editor.State.Mouse.MouseMoveStatus = e.Location

        Dim notes = _editor.Notes

        Select Case e.Button
            Case MouseButtons.None
                'If K Is Nothing Then Exit Select
                If _editor.State.Mouse.MiddleButtonClicked Then Exit Select

                If _editor.IsFullscreen Then
                    _editor.SetToolstripVisible(e.Y > 5)
                End If

                Dim mouseRemainInSameRegion = False
                Dim foundNoteIndex = UpdateNtInputState(e, notes, mouseRemainInSameRegion)

                If _editor.IsSelectMode Then

                    If mouseRemainInSameRegion Then Exit Select
                    _editor.State.Mouse.CurrentHoveredNoteIndex = -1

                    UpdateTimeSelectLineOver(e)

                    _editor.State.Mouse.CurrentHoveredNoteIndex = foundNoteIndex

                ElseIf _editor.IsWriteMode Then
                    _editor.UpdateMouseRowAndColumn()

                    _editor.LnDisplayLength = 0
                    If foundNoteIndex > -1 Then
                        _editor.LnDisplayLength = notes(foundNoteIndex).Length
                    End If

                    _editor.RefreshPanelAll()
                End If

            Case MouseButtons.Left
                If _editor.TempFirstMouseDown And Not _editor.IsTimeSelectMode Then Exit Select

                _editor.State.Mouse.tempX = 0
                _editor.State.Mouse.tempY = 0
                If Not (e.X < 0 Or e.X > Width Or e.Y < 0 Or e.Y > Height) Then
                    If e.X < 0 Then _editor.State.Mouse.tempX = e.X
                    If e.X > Width Then _editor.State.Mouse.tempX = e.X - Width
                    If e.Y < 0 Then _editor.State.Mouse.tempY = e.Y
                    If e.Y > Height Then
                    Else
                        ' _editor.Timer1.Enabled = False
                    End If

                    If _editor.IsSelectMode Then

                        _editor.State.Mouse.pMouseMove = e.Location

                        'If K Is Nothing Then RefreshPanelAll() : Exit Select

                        If Not _editor.State.Mouse.LastMouseDownLocation = New Point(-1, -1) Then
                            UpdateSelectionBox()

                            'ElseIf Not KMouseDown = -1 Then
                        ElseIf _editor.GetSelectedNotes().Count() <> 0 Then
                            UpdateSelectedNotes(e)
                        ElseIf _editor.State.IsDuplicatingSelectedNotes Then
                            OnDuplicateSelectedNotes(e)
                        End If

                    ElseIf _editor.IsWriteMode Then

                        If _editor.NtInput Then
                            OnWriteModeMouseMove()
                        Else
                            _editor.State.Mouse.CurrentMouseRow = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)
                            _editor.State.Mouse.CurrentMouseColumn = GetColumnAtEvent(e)
                        End If

                    ElseIf _editor.IsTimeSelectMode Then
                        OnTimeSelectClick(e)
                    End If
                End If


            Case MouseButtons.Middle
                OnPanelMousePan(e)
        End Select

        Dim col = GetColumnAtEvent(e)
        Dim vps = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)
        If vps <> _lastVPos Or col <> _lastColumn Then
            _lastVPos = vps
            _lastColumn = col
            _editor.RefreshPanelAll() 'az: refreshing the line is important now...
        End If

        _editor.PoStatusRefresh()
        Refresh()
    End Sub

    ''' <summary>
    '''     Updates NT Input state and calculates hovered note.
    ''' </summary>
    ''' <param name="e">Mouse event</param>
    ''' <param name="notes">Notes array to check</param>
    ''' <returns>The index of the note that the event's coordinates hovers above.</returns>
    Private Function UpdateNtInputState(e As MouseEventArgs, notes() As Note, ByRef xMouseRemainInSameRegion As Boolean) _
        As Integer
        Dim foundNoteIndex = -1

        For noteIndex = UBound(notes) To 0 Step -1
            If MouseInNote(e, notes(noteIndex)) Then
                foundNoteIndex = noteIndex

                xMouseRemainInSameRegion = foundNoteIndex = _editor.State.Mouse.CurrentHoveredNoteIndex
                If _editor.NtInput Then
                    Dim vy = VPositionToPanelY(notes(noteIndex).VPosition + notes(noteIndex).Length)

                    Dim xbAdjustUpper As Boolean = (e.Y <= vy) And ModifierLongNoteActive()
                    Dim xbAdjustLength As Boolean = (e.Y >= vy - _theme.NoteHeight Or xbAdjustUpper) And
                                                    ModifierLongNoteActive()

                    xMouseRemainInSameRegion = xMouseRemainInSameRegion And
                                               xbAdjustUpper = _editor.State.NT.IsAdjustingUpperEnd And
                                               xbAdjustLength = _editor.State.NT.IsAdjustingNoteLength

                    _editor.State.NT.IsAdjustingUpperEnd = xbAdjustUpper
                    _editor.State.NT.IsAdjustingNoteLength = xbAdjustLength
                End If

                Exit For
            End If
        Next

        Return foundNoteIndex
    End Function

    Dim _lastVPos = -1
    Dim _lastColumn = -1

    Private Sub UpdateSelectedNotes(e As MouseEventArgs)
        Dim currentClickedNoteIndex As Integer

        For i = 1 To _editor.Notes.Length - 1
            If _editor.Notes(i).TempMouseDown Then
                currentClickedNoteIndex = i
                Exit For
            End If
        Next

        Dim clickedNote = _editor.Notes(currentClickedNoteIndex)

        Dim mouseVPosition = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)

        If _editor.State.NT.IsAdjustingNoteLength And
           _editor.State.NT.IsAdjustingSingleNote Then

            If _editor.State.NT.IsAdjustingUpperEnd AndAlso
               mouseVPosition < clickedNote.VPosition Then

                _editor.State.NT.IsAdjustingUpperEnd = False
                clickedNote.VPosition += clickedNote.Length
                clickedNote.Length *= -1

            ElseIf Not _editor.State.NT.IsAdjustingUpperEnd AndAlso
                   mouseVPosition > clickedNote.VPosition + clickedNote.Length Then

                _editor.State.NT.IsAdjustingUpperEnd = True
                clickedNote.VPosition += clickedNote.Length
                clickedNote.Length *= -1
            End If

        End If

        'If moving
        If Not _editor.State.NT.IsAdjustingNoteLength Then
            OnSelectModeMoveNotes(e, clickedNote)

        ElseIf _editor.State.NT.IsAdjustingUpperEnd Then
            Dim dVPosition = mouseVPosition - clickedNote.VPosition - clickedNote.Length  'delta Length
            '< 0 means shorten, > 0 means lengthen

            OnAdjustUpperEnd(dVPosition)

        Else 'If adjusting lower end
            Dim dVPosition = mouseVPosition - clickedNote.VPosition  'delta VPosition
            '> 0 means shorten, < 0 means lengthen

            OnAdjustLowerEnd(dVPosition)
        End If

        _editor.ValidateNotesArray()
    End Sub

    Private Sub OnPanelMousePan(e As MouseEventArgs)
        If _editor.MiddleButtonMoveMethod = 1 Then
            Dim mouse = _editor.State.Mouse
            Dim i As Integer = mouse.tempV + (mouse.tempY - e.Y) / _editor.Grid.HeightScale
            Dim j As Integer = mouse.tempH + (mouse.tempX - e.X) / _editor.Grid.WidthScale
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
        If _editor.Notes IsNot Nothing Then
            For i = _editor.Notes.Length - 1 To 0 Step -1
                If MouseInNote(e, _editor.Notes(i)) Then
                    hoverNoteIndex = i
                    note = _editor.Notes(i)
                    Exit For
                End If
            Next
        End If

        Dim snap = _editor.Grid.IsSnapEnabled And Not My.Computer.Keyboard.CtrlKeyDown
        If Not _editor.State.TimeSelect.Adjust Then
            Dim mouseRow As Double = _editor.GetMouseVPosition(snap)
            If hoverNoteIndex >= 0 Then mouseRow = note.VPosition

            If _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine Then
                Dim startPointDy = _editor.State.TimeSelect.StartPoint - mouseRow
                _editor.State.TimeSelect.EndPointLength += startPointDy
                _editor.State.TimeSelect.HalfPointLength += startPointDy
                _editor.State.TimeSelect.StartPoint = mouseRow

            ElseIf _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine Then
                _editor.State.TimeSelect.HalfPoint = mouseRow

            ElseIf _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine Then
                _editor.State.TimeSelect.EndPoint = mouseRow
            Else
                _editor.State.TimeSelect.EndPoint = mouseRow
                _editor.State.TimeSelect.HalfPointLength = _editor.State.TimeSelect.EndPointLength / 2
            End If

            _editor.State.TimeSelect.ValidateSelection(_editor.GetMaxVPosition())

        Else
            Dim xL1 As Double = _editor.GetMouseVPosition(snap)

            If _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.HalfLine Then
                _editor.State.TimeSelect.StartPoint = _editor.State.TimeSelect.PStart
                _editor.State.TimeSelect.EndPointLength = _editor.State.TimeSelect.PLength
                _editor.State.TimeSelect.HalfPointLength = _editor.State.TimeSelect.PHalf
                _editor.Notes = _editor.State.TimeSelect.Notes
                ReDim Preserve _editor.Notes(_editor.Notes.Length - 1)

                _editor.BPMChangeHalf(xL1 - _editor.State.TimeSelect.HalfPointLength - _editor.State.TimeSelect.StartPoint,
                                     , True)


            ElseIf _
                _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine Or
                _editor.State.TimeSelect.MouseOverLine = TimeSelectLine.StartLine Then
                _editor.State.TimeSelect.StartPoint = _editor.State.TimeSelect.PStart
                _editor.State.TimeSelect.EndPointLength = _editor.State.TimeSelect.PLength
                _editor.State.TimeSelect.HalfPointLength = _editor.State.TimeSelect.PHalf
                _editor.Notes = _editor.State.TimeSelect.Notes
                ReDim Preserve _editor.Notes(UBound(_editor.Notes))

                _editor.BPMChangeTop(
                    IIf(_editor.State.TimeSelect.MouseOverLine = TimeSelectLine.EndLine,
                        xL1 - _editor.State.TimeSelect.StartPoint,
                        _editor.State.TimeSelect.EndPoint - xL1) / _editor.State.TimeSelect.EndPointLength, , True)


            Else
                _editor.State.TimeSelect.EndPointLength = xL1
                _editor.State.TimeSelect.ValidateSelection(_editor.GetMaxVPosition())
            End If
        End If

        Dim notes = _editor.Notes

        If _editor.State.TimeSelect.EndPointLength Then
            Dim xVLower As Double = Math.Min(_editor.State.TimeSelect.StartPoint, _editor.State.TimeSelect.EndPoint)
            Dim xVUpper As Double = Math.Max(_editor.State.TimeSelect.StartPoint, _editor.State.TimeSelect.EndPoint)

            If _editor.NtInput Then
                For j = 1 To UBound(notes)
                    notes(j).Selected = notes(j).VPosition < xVUpper And
                                        notes(j).VPosition + notes(j).Length >= xVLower And
                                        _editor.Columns.IsEnabled(notes(j).ColumnIndex)
                Next
            Else
                For j = 1 To UBound(notes)
                    notes(j).Selected = notes(j).VPosition >= xVLower And
                                        notes(j).VPosition < xVUpper And
                                        _editor.Columns.IsEnabled(notes(j).ColumnIndex)
                Next
            End If
        Else
            _editor.DeselectAllNotes()
        End If
    End Sub

    Private Sub OnAdjustUpperEnd(dVPosition As Double)
        Dim minLength As Double = 0
        Dim maxHeight As Double = 191999
        Dim notes = _editor.Notes
        For Each note In notes.Skip(1).Where(Function(x) x.Selected)

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
        For Each note In _editor.Notes.Skip(1).Where(Function(x) x.Selected)
            Dim xLen = note.Length + dVPosition - minLength - maxHeight
            RedoLongNoteModify(note, note.VPosition, xLen, xUndo, xRedo)

            note.Length = xLen
        Next

        'Add undo
        If dVPosition - minLength - maxHeight <> 0 Then
            _editor.AddUndoChain(xUndo, xBaseRedo.Next, _editor.State.OverwriteLastUndoRedoCommand)
            If Not _editor.State.OverwriteLastUndoRedoCommand Then _editor.State.OverwriteLastUndoRedoCommand = True
        End If
    End Sub


    Private Sub OnAdjustLowerEnd(dVPosition As Double)
        Dim minLength As Double = 0
        Dim minVPosition As Double = 0
        Dim notes = _editor.Notes
        For Each note In notes.Skip(1).Where(Function(x) x.Selected)
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

        For Each note In notes.Where(Function(x) x.Selected)
            Dim newVPos = note.VPosition + dVPosition + minLength - minVPosition
            Dim newLen = note.Length - dVPosition - minLength + minVPosition
            RedoLongNoteModify(note, newVPos, newLen, xUndo, xRedo)

            note.VPosition = newVPos
            note.Length = newLen
        Next

        'Add undo
        If dVPosition + minLength - minVPosition <> 0 Then
            _editor.AddUndoChain(xUndo, xBaseRedo.Next, _editor.State.OverwriteLastUndoRedoCommand)
            If Not _editor.State.OverwriteLastUndoRedoCommand Then _editor.State.OverwriteLastUndoRedoCommand = True
        End If
    End Sub

    Private Sub OnDuplicateSelectedNotes(e As MouseEventArgs)
        Dim notes = _editor.Notes
        Dim tempNoteIndex As Integer
        For tempNoteIndex = 1 To UBound(notes)
            If notes(tempNoteIndex).TempMouseDown Then Exit For
        Next

        Dim mouseVPosition = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)
        If _editor.DisableVerticalMove Then
            mouseVPosition = notes(tempNoteIndex).VPosition
        End If

        Dim dVPosition As Double = mouseVPosition - notes(tempNoteIndex).VPosition  'delta VPosition

        Dim currCol = _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(GetColumnAtEvent(e))
        Dim noteCol = _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(notes(tempNoteIndex).ColumnIndex)
        Dim colChange As Integer = currCol - noteCol 'delta Column

        'Ks cannot be beyond the left, the upper and the lower boundary
        Dim dstColumn = 0
        Dim mVPosition As Double = 0
        Dim muVPosition As Double = 191999
        For Each note In notes.Skip(1).Where(Function(x) x.Selected)

            If _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + colChange < dstColumn Then
                dstColumn = _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + colChange
            End If

            If note.VPosition + dVPosition < mVPosition Then
                mVPosition = note.VPosition + dVPosition
            End If

            Dim noteTailPos = note.VPosition + IIf(_editor.NtInput, note.Length, 0)
            If noteTailPos + dVPosition > muVPosition Then
                muVPosition = noteTailPos + dVPosition
            End If

        Next
        muVPosition -= 191999

        'If not moving then exit
        If (Not _editor.State.SelectedNotesWereDuplicated) And
           colChange - dstColumn = 0 And
           dVPosition - mVPosition - muVPosition = 0 Then
            Return
        End If

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If Not _editor.State.SelectedNotesWereDuplicated Then 'If Editor.State.uAdded = False
            DuplicateSelectedNotes(tempNoteIndex, dVPosition, colChange, dstColumn, mVPosition, muVPosition)
            _editor.State.SelectedNotesWereDuplicated = True

        Else
            For Each note In _editor.GetSelectedNotes()
                Dim idx = _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + colChange - dstColumn
                note.ColumnIndex = _editor.Columns.EnabledColumnIndexToColumnArrayIndex(idx)
                note.VPosition = note.VPosition + dVPosition - mVPosition - muVPosition
                RedoAddNote(note, xUndo, xRedo)
            Next

            _editor.AddUndoChain(xUndo, xBaseRedo.Next, True)
        End If

        _editor.ValidateNotesArray()
    End Sub


    Private Sub OnWriteModeMouseMove()
        'If Not KMouseDown = -1 Then
        Dim selectedNotes = _editor.GetSelectedNotes().ToArray()

        If selectedNotes.Count() <> 0 Then
            Dim note As Note = _editor.Notes.
                    Skip(1).First(Function(x) x.TempMouseDown)

            Dim mouseVPosition = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)
            Dim maxVPos = _editor.GetMaxVPosition()
            Dim nt = _editor.State.NT

            With note
                If nt.IsAdjustingUpperEnd AndAlso
                   mouseVPosition < .VPosition Then

                    nt.IsAdjustingUpperEnd = False
                    .VPosition += .Length
                    .Length *= -1

                ElseIf Not nt.IsAdjustingUpperEnd AndAlso
                       mouseVPosition > .VPosition + .Length Then

                    nt.IsAdjustingUpperEnd = True
                    .VPosition += .Length
                    .Length *= -1

                End If

                If nt.IsAdjustingUpperEnd Then
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
                    _editor.AddUndoChain(xUndo, xRedo, True)

                Else 'If existing note
                    Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                    Dim xRedo As UndoRedo.LinkedURCmd = Nothing
                    RedoLongNoteModify(selectedNotes.First(), .VPosition, .Length, xUndo, xRedo)
                    _editor.AddUndoChain(xUndo, xRedo, True)
                End If

                _editor.State.Mouse.CurrentMouseColumn = .ColumnIndex
                _editor.State.Mouse.CurrentMouseRow = mouseVPosition
                _editor.LnDisplayLength = .Length

            End With

            _editor.ValidateNotesArray()

        End If
    End Sub

    Private Sub OnSelectModeMoveNotes(e As MouseEventArgs, currentNote As Note)
        ' delta vpos
        Dim dVPosition As Single
        If _editor.DisableVerticalMove Then
            dVPosition = 0
        Else
            Dim mouseVPosition = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)
            dVPosition = mouseVPosition - currentNote.VPosition
        End If

        ' delta column
        Dim mouseColumn = _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(GetColumnAtX(e.X))

        'get the enabled delta column where the mouse is 
        Dim dColumn = mouseColumn - _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(currentNote.ColumnIndex)

        ' Don't do anything if there's no change
        If dColumn = 0 AndAlso dVPosition = 0 Then
            Return
        End If

        Debug.Indent()
        Debug.WriteLine(String.Format("dCol/dV: {0} $ {1} | noteCol/VPos: {2} $ {3}",
                                      dColumn, dVPosition,
                                      currentNote.ColumnIndex, currentNote.VPosition))

        'Ks cannot be beyond the left, the upper and the lower boundary
        Dim mLeft = 0
        Dim mVPosition As Double = 0
        Dim muVPosition As Double = 191999
        For Each note In _editor.GetSelectedNotes()
            Dim idx = _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex)
            mLeft = Math.Min(idx + dColumn, mLeft)
            mVPosition = Math.Min(note.VPosition + dVPosition, mVPosition)
            muVPosition = Math.Max(note.VPosition + IIf(_editor.NtInput, note.Length, 0) + dVPosition, muVPosition)
        Next
        muVPosition -= 191999 ' az: this magic number again. why?

        Debug.WriteLine(String.Format("mVPos/muVPos/mLeft: {0} and {1} and {2}", mVPosition, muVPosition, mLeft))

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        ' az: Okay so uAdded (now OverwriteLastUndoRedoCommand, to be more descriptive!) 
        ' does try to solve a real problem, which Is to 
        ' try to overwrite a move if there already is one, so the Undo/Redo history isn't spammed to hell and back.
        ' Thing is, you want to work with the columnindex and vpos of when the first uAdded was used
        ' otherwise you end up with a nasty bug where only the last move is saved...
        ' So this is where this MoveStartCol and MoveStartVPos crap comes in.
        ' Those will store the position where the note originally was, 
        ' so that this overwritten undo/redo works properly.

        'start moving
        ' az: yeah, start moving I guess. 
        For Each note In _editor.GetSelectedNotes()

            ' This is it. We Save MoveStart variables here
            ' to perform the move - this is a new move command, not an old one we're overwriting.
            If Not _editor.State.OverwriteLastUndoRedoCommand Then
                note.MoveStartVPos = note.VPosition
                note.MoveStartColumnIndex = note.ColumnIndex
            End If

            Dim idx = _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + dColumn - mLeft
            Dim newColumn = _editor.Columns.EnabledColumnIndexToColumnArrayIndex(idx)
            Dim newVPosition = note.VPosition + dVPosition - mVPosition - muVPosition

            ' We could've done a, "if not uAdded note, else note.MoveStartClone" but effectively
            ' it doesn't matter, because the vposition and columnindex is the same for either case
            ' especially with the initialization of MoveStartVPos and MoveStartColumnIndex above.
            RedoMoveNote(note.MoveStartClone(), newColumn, newVPosition, xUndo, xRedo)

            note.ColumnIndex = newColumn
            note.VPosition = newVPosition
        Next


        Debug.WriteLine(String.Format("uAdded: {0}", _editor.State.OverwriteLastUndoRedoCommand))
        _editor.AddUndoChain(xUndo, xBaseRedo.Next, _editor.State.OverwriteLastUndoRedoCommand)

        ' az documentation: We set it to true so that on the moves that follow 
        ' it overwrites the last UR until we release the mouse button.
        If Not _editor.State.OverwriteLastUndoRedoCommand Then
            _editor.State.OverwriteLastUndoRedoCommand = True
        End If

        Debug.Unindent()
        'End If
    End Sub

    Private Sub UpdateSelectionBox()
        Dim pMouseMove = _editor.State.Mouse.pMouseMove
        Dim lastMouseDownLocation = _editor.State.Mouse.LastMouseDownLocation
        Dim selectionBox As New Rectangle(Math.Min(pMouseMove.X, lastMouseDownLocation.X),
                                          Math.Min(pMouseMove.Y, lastMouseDownLocation.Y),
                                          Math.Abs(pMouseMove.X - lastMouseDownLocation.X),
                                          Math.Abs(pMouseMove.Y - lastMouseDownLocation.Y))
        Dim noteRect As Rectangle

        Dim columns = _editor.Columns
        For Each note In _editor.Notes.Skip(1)
            Dim noteStartX = columns.GetColumnLeft(note.ColumnIndex)
            Dim noteTailPos = note.VPosition + IIf(_editor.NtInput, note.Length, 0)
            Dim drawWidth = columns.GetWidth(note.ColumnIndex) * _editor.Grid.WidthScale - 2
            Dim drawHeight = _theme.NoteHeight + IIf(_editor.NtInput, note.Length * _editor.Grid.HeightScale, 0)
            noteRect = New Rectangle(HPositionToPanelX(noteStartX) + 1,
                                     VPositionToPanelY(noteTailPos) - _theme.NoteHeight,
                                     drawWidth,
                                     drawHeight)


            Dim columnEnabled = columns.IsEnabled(note.ColumnIndex)

            If noteRect.IntersectsWith(selectionBox) Then
                note.Selected = Not note.TempSelected And
                                columnEnabled
            Else
                note.Selected = note.TempSelected And
                                columnEnabled
            End If
        Next
    End Sub

    Private Sub DuplicateSelectedNotes(tempNoteIndex As Integer, dVPosition As Double, dColumn As Integer,
                                       mLeft As Integer, mVPosition As Double, muVPosition As Double)
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        Dim notes = _editor.Notes

        notes(tempNoteIndex).Selected = True

        Dim selectedNotes = _editor.GetSelectedNotes().ToArray()
        Dim selectedNotesCount = selectedNotes.Length

        Dim duplicatedNotes(selectedNotesCount - 1) As Note
        Dim j = 0
        For Each note In selectedNotes
            Dim idx = _editor.Columns.ColumnArrayIndexToEnabledColumnIndex(note.ColumnIndex) + dColumn - mLeft
            duplicatedNotes(j) = note.Clone
            duplicatedNotes(j).ColumnIndex = _editor.Columns.EnabledColumnIndexToColumnArrayIndex(idx)
            duplicatedNotes(j).VPosition = note.VPosition + dVPosition - mVPosition - muVPosition
            RedoAddNote(duplicatedNotes(j), xUndo, xRedo)

            note.Selected = False
            j += 1
        Next
        notes(tempNoteIndex).TempMouseDown = False

        'copy to K
        _editor.Notes = _editor.Notes.Concat(duplicatedNotes).ToArray()

        _editor.AddUndoChain(xUndo, xBaseRedo.Next)
    End Sub

    Private Function GetColumnAtEvent(e As MouseEventArgs)
        Return GetColumnAtX(e.X)
    End Function

    Private Sub PMainInMouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        _editor.State.Mouse.tempX = 0
        _editor.State.Mouse.tempY = 0
        _editor.State.Mouse.tempV = 0
        _editor.State.Mouse.tempH = 0
        LastVerticalScroll = -1
        LastHorizontalScroll = -1

        Dim iI As Integer = sender.Tag

        Dim dv = (_editor.State.Mouse.MiddleButtonLocation - Cursor.Position)
        If _editor.State.Mouse.MiddleButtonClicked AndAlso
           e.Button = MouseButtons.Middle AndAlso
           dv.X ^ 2 + dv.Y ^ 2 >= _theme.MiddleDeltaRelease Then
            _editor.State.Mouse.MiddleButtonClicked = False
        End If

        If _editor.IsSelectMode Then
            _editor.State.Mouse.LastMouseDownLocation = New Point(-1, -1)
            _editor.State.Mouse.pMouseMove = New Point(-1, -1)

            If _editor.State.IsDuplicatingSelectedNotes And
               Not _editor.State.SelectedNotesWereDuplicated And
               Not ModifierMultiselectActive() Then

                For Each note In _editor.Notes
                    If note.TempMouseDown Then
                        note.Selected = Not note.Selected
                        Exit For
                    End If
                Next
            End If

            _editor.State.IsDuplicatingSelectedNotes = False
            _editor.State.SelectedNotesWereDuplicated = False

        ElseIf _editor.IsWriteMode Then

            If Not _editor.NtInput And Not _editor.TempFirstMouseDown Then
                Dim xVPosition = _editor.GetMouseVPosition(_editor.Grid.IsSnapEnabled)

                Dim xColumn = GetColumnAtEvent(e)
                Dim notes = _editor.Notes

                If e.Button = MouseButtons.Left Then
                    Dim hiddenNote As Boolean = ModifierHiddenActive()
                    Dim longNote As Boolean = ModifierLongNoteActive()
                    Dim landmine As Boolean = ModifierLandmineActive()
                    Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                    Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                    Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                    If _editor.Columns.IsColumnNumeric(xColumn) Then
                        Dim xMessage As String = Strings.Messages.PromptEnterNumeric
                        If xColumn = ColumnType.BPM Then xMessage = Strings.Messages.PromptEnterBPM
                        If xColumn = ColumnType.STOPS Then xMessage = Strings.Messages.PromptEnterSTOP
                        If xColumn = ColumnType.SCROLLS Then xMessage = Strings.Messages.PromptEnterSCROLL

                        Dim valstr As String = InputBox(xMessage, Me.Text)
                        Dim value As Long = Val(valstr) * 10000

                        If (xColumn = ColumnType.SCROLLS And valstr = "0") Or value <> 0 Then
                            For i = 1 To UBound(notes)
                                If notes(i).VPosition = xVPosition AndAlso notes(i).ColumnIndex = xColumn Then _
                                    RedoRemoveNote(notes(i), xUndo, xRedo)
                            Next

                            Dim n = New Note(xColumn, xVPosition, value, longNote, hiddenNote)
                            RedoAddNote(n, xUndo, xRedo)
                            _editor.AddNote(n)

                            _editor.AddUndoChain(xUndo, xBaseRedo.Next)
                        End If

                    Else
                        Dim xValue As Integer = _editor.CurrentWavSelectedIndex * 10000

                        For i = 1 To UBound(notes)
                            If notes(i).VPosition = xVPosition AndAlso notes(i).ColumnIndex = xColumn Then _
                                RedoRemoveNote(notes(i), xUndo, xRedo)
                        Next

                        Dim n = New Note(xColumn, xVPosition, xValue,
                                         longNote, hiddenNote, True, landmine)

                        RedoAddNote(n, xUndo, xRedo)
                        _editor.AddNote(n)

                        _editor.AddUndoChain(xUndo, xRedo)
                    End If
                End If
            End If

            _editor.State.Mouse.CurrentMouseRow = -1
            _editor.State.Mouse.CurrentMouseColumn = -1
        End If

        ' az refactoring: Not a full note refresh?
        _editor.RefreshPanelAll()
    End Sub

    Private Sub PMainInMouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        If _editor.State.Mouse.MiddleButtonClicked Then
            _editor.State.Mouse.MiddleButtonClicked = False
        End If

        Dim i = VerticalPosition - Math.Sign(e.Delta) * _editor.Grid.WheelScroll + HorizontalScrollBar.Height
        VerticalScrollBar.Value = Clamp(i, VerticalScrollBar.Minimum, 0)
    End Sub

    Public Sub OnUpdateScroll(newMin As Integer)
        VerticalScrollBar.Minimum = newMin
    End Sub
End Class
