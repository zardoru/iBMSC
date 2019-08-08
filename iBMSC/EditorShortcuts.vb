Imports System.Linq
Imports iBMSC.Editor

Partial Public Class MainWindow
    Private xUndo As UndoRedo.LinkedURCmd
    Private xRedo As Object


    Private Sub KeyDownEvent(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If Not _panelList.Any(Function(x) x.Focused) Then
            Return
        End If

        Select Case e.KeyCode
            Case Keys.Up
                Dim xVPosition As Double = 192 / Grid.Divider
                If My.Computer.Keyboard.CtrlKeyDown Then xVPosition = 1

                'Ks cannot be beyond the upper boundary
                Dim muVPosition As Double = GetMaxVPosition() - 1
                For i = 1 To UBound(Notes)
                    If Notes(i).Selected Then
                        'K(i).VPosition = Math.Floor(K(i).VPosition / (192 / Grid.Divider)) * 192 / Grid.Divider
                        muVPosition =
                            IIf(Notes(i).VPosition + IIf(NtInput, Notes(i).Length, 0) + xVPosition > muVPosition,
                                Notes(i).VPosition + IIf(NtInput, Notes(i).Length, 0) + xVPosition,
                                muVPosition)
                    End If
                Next
                muVPosition -= 191999

                'xRedo = sCmdKMs(0, xVPosition - muVPosition, True)
                Dim xVPos As Double
                For i = 1 To UBound(Notes)
                    If Not Notes(i).Selected Then Continue For

                    xVPos = Notes(i).VPosition + xVPosition - muVPosition
                    RedoMoveNote(Notes(i), Notes(i).ColumnIndex, xVPos, xUndo, xRedo)
                    Notes(i).VPosition = xVPos
                Next
                'xUndo = sCmdKMs(0, -xVPosition + muVPosition, True)

                If xVPosition - muVPosition <> 0 Then
                    AddUndoChain(xUndo, xBaseRedo.Next)
                End If

                ValidateNotesArray()
                RefreshPanelAll()

            Case Keys.Down
                Dim xVPosition As Double = -192 / Grid.Divider
                If My.Computer.Keyboard.CtrlKeyDown Then xVPosition = -1

                'Ks cannot be beyond the lower boundary
                Dim mVPosition As Double = 0
                For i = 1 To UBound(Notes)
                    If Notes(i).Selected Then
                        'K(i).VPosition = Math.Ceiling(K(i).VPosition / (192 / Grid.Divider)) * 192 / Grid.Divider
                        mVPosition = IIf(Notes(i).VPosition + xVPosition < mVPosition,
                                         Notes(i).VPosition + xVPosition,
                                         mVPosition)
                    End If
                Next

                'xRedo = sCmdKMs(0, xVPosition - mVPosition, True)
                Dim xVPos As Double
                For i = 1 To UBound(Notes)
                    If Not Notes(i).Selected Then Continue For

                    xVPos = Notes(i).VPosition + xVPosition - mVPosition
                    RedoMoveNote(Notes(i), Notes(i).ColumnIndex, xVPos, xUndo, xRedo)
                    Notes(i).VPosition = xVPos
                Next
                'xUndo = sCmdKMs(0, -xVPosition + mVPosition, True)

                If xVPosition - mVPosition <> 0 Then
                    AddUndoChain(xUndo, xBaseRedo.Next)
                End If
                ValidateNotesArray()
                RefreshPanelAll()

            Case Keys.Left

                'Ks cannot be beyond the left boundary
                Dim mLeft = 0
                For i = 1 To UBound(Notes)
                    If Notes(i).Selected Then
                        mLeft = Math.Min(Columns.ColumnArrayIndexToEnabledColumnIndex(Notes(i).ColumnIndex) - 1, mLeft)
                    End If
                Next

                Dim xCol As Integer
                For i = 1 To UBound(Notes)
                    If Not Notes(i).Selected Then Continue For

                    xCol = Columns.NormalizeIndex(Notes(i).ColumnIndex - 1 - mLeft)
                    RedoMoveNote(Notes(i), xCol, Notes(i).VPosition, xUndo, xRedo)
                    Notes(i).ColumnIndex = xCol
                Next
                'xUndo = sCmdKMs(1 + mLeft, 0, True)

                If -1 - mLeft <> 0 Then AddUndoChain(xUndo, xBaseRedo.Next)
                UpdatePairing()
                CalculateTotalPlayableNotes()
                RefreshPanelAll()

            Case Keys.Right
                'xRedo = sCmdKMs(1, 0, True)
                Dim xCol As Integer
                For i = 1 To UBound(Notes)
                    If Not Notes(i).Selected Then Continue For

                    xCol = Columns.NormalizeIndex(Notes(i).ColumnIndex + 1)
                    RedoMoveNote(Notes(i), xCol, Notes(i).VPosition, xUndo, xRedo)
                    Notes(i).ColumnIndex = xCol
                Next
                'xUndo = sCmdKMs(-1, 0, True)

                AddUndoChain(xUndo, xBaseRedo.Next)
                UpdatePairing()
                CalculateTotalPlayableNotes()
                RefreshPanelAll()

            Case Keys.Delete
                mnDelete_Click(mnDelete, New EventArgs)


            Case Keys.Oemcomma
                If Grid.Divider * 2 <= CGDivide.Maximum Then CGDivide.Value = Grid.Divider * 2

            Case Keys.OemPeriod
                If Grid.Divider \ 2 >= CGDivide.Minimum Then CGDivide.Value = Grid.Divider \ 2

            Case Keys.OemQuestion
                'Dim xTempSwap As Integer = Grid.Slash
                'Grid.Slash = CGDivide.Value
                'CGDivide.Value = xTempSwap
                CGDivide.Value = Grid.Slash

            Case Keys.Oemplus
                With CGHeight
                    .Value += IIf(.Value > .Maximum - .Increment, .Maximum - .Value, .Increment)
                End With

            Case Keys.OemMinus
                With CGHeight
                    .Value -= IIf(.Value < .Minimum + .Increment, .Value - .Minimum, .Increment)
                End With

            Case Keys.Add
                IncreaseCurrentWav()
            Case Keys.Subtract
                DecreaseCurrentWav()

            Case Keys.G
                'az: don't trigger when we use Go To Measure
                If Not My.Computer.Keyboard.CtrlKeyDown Then CGSnap.Checked = Not Grid.IsSnapEnabled

            Case Keys.L
                If Not My.Computer.Keyboard.CtrlKeyDown Then POBLong_Click(Nothing, Nothing)

            Case Keys.S
                If Not My.Computer.Keyboard.CtrlKeyDown Then POBNormal_Click(Nothing, Nothing)

            Case Keys.D
                CGDisableVertical.Checked = Not CGDisableVertical.Checked

            Case Keys.NumPad0, Keys.D0
                MoveToBGM(xUndo, xRedo)

            Case Keys.Oem1, Keys.NumPad1, Keys.D1 : MoveToColumn(ColumnType.A1, xUndo, xRedo)
            Case Keys.Oem2, Keys.NumPad2, Keys.D2 : MoveToColumn(ColumnType.A2, xUndo, xRedo)
            Case Keys.Oem3, Keys.NumPad3, Keys.D3 : MoveToColumn(ColumnType.A3, xUndo, xRedo)
            Case Keys.Oem4, Keys.NumPad4, Keys.D4 : MoveToColumn(ColumnType.A4, xUndo, xRedo)
            Case Keys.Oem5, Keys.NumPad5, Keys.D5 : MoveToColumn(ColumnType.A5, xUndo, xRedo)
            Case Keys.Oem6, Keys.NumPad6, Keys.D6 : MoveToColumn(ColumnType.A6, xUndo, xRedo)
            Case Keys.Oem7, Keys.NumPad7, Keys.D7 : MoveToColumn(ColumnType.A7, xUndo, xRedo)
            Case Keys.Oem8, Keys.NumPad8, Keys.D8 : MoveToColumn(ColumnType.A8, xUndo, xRedo)
        End Select

        If My.Computer.Keyboard.CtrlKeyDown And (Not My.Computer.Keyboard.AltKeyDown) And
            (Not My.Computer.Keyboard.ShiftKeyDown) Then
            Select Case e.KeyCode
                Case Keys.Z : TBUndo_Click(TBUndo, New EventArgs)
                Case Keys.Y : TBRedo_Click(TBRedo, New EventArgs)
                Case Keys.X : TBCut_Click(TBCut, New EventArgs)
                Case Keys.C : TBCopy_Click(TBCopy, New EventArgs)
                Case Keys.V : TBPaste_Click(TBPaste, New EventArgs)
                Case Keys.A : mnSelectAll_Click(mnSelectAll, New EventArgs)
                Case Keys.F : TBFind_Click(TBFind, New EventArgs)
                Case Keys.T : TBStatistics_Click(TBStatistics, New EventArgs)
            End Select
        End If

        If ModifierMultiselectActive() Then
            If e.KeyCode = Keys.A And State.Mouse.CurrentHoveredNoteIndex <> -1 Then
                SelectAllWithHoveredNoteLabel()
            End If
        End If

        PoStatusRefresh()
    End Sub

    Private Sub MoveToBGM(xUndo As UndoRedo.LinkedURCmd, xRedo As UndoRedo.LinkedURCmd)
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For j = 1 To UBound(Notes)
            If Not Notes(j).Selected Then Continue For

            With Notes(j)
                Dim currentBGMColumn As Integer = ColumnType.BGM

                'TODO: optimize the for loops below
                If NtInput Then
                    For i = 1 To UBound(Notes)
                        Dim IntersectA = Notes(i).VPosition <= Notes(j).VPosition + Notes(j).Length
                        Dim IntersectB = Notes(i).VPosition + Notes(i).Length >= Notes(j).VPosition
                        If Notes(i).ColumnIndex = currentBGMColumn AndAlso
                            IntersectA And
                            IntersectB Then
                            currentBGMColumn += 1
                            i = 1
                        End If
                    Next
                Else
                    For i = 1 To UBound(Notes)
                        If Notes(i).ColumnIndex = currentBGMColumn AndAlso
                            Notes(i).VPosition = Notes(j).VPosition Then
                            currentBGMColumn += 1
                            i = 1
                        End If
                    Next
                End If

                RedoMoveNote(Notes(j), currentBGMColumn, .VPosition, xUndo, xRedo)
                .ColumnIndex = currentBGMColumn
            End With
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        UpdatePairing()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
    End Sub

    Private Sub MoveToColumn(targetColumn As Integer, xUndo As UndoRedo.LinkedURCmd, xRedo As UndoRedo.LinkedURCmd)
        If targetColumn = -1 Then Return
        If Not Columns.IsEnabled(targetColumn) Then Return

        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        Dim bMoveAndDeselectFirstNote = My.Computer.Keyboard.ShiftKeyDown

        For j = 1 To UBound(Notes)
            If Not Notes(j).Selected Then Continue For

            RedoMoveNote(Notes(j), targetColumn, Notes(j).VPosition, xUndo, xRedo)
            Notes(j).ColumnIndex = targetColumn

            If bMoveAndDeselectFirstNote Then
                Notes(j).Selected = False
                PanelPreviewNoteIndex(j)

                ' az: Add selected notes to undo
                ' to preserve selection status
                ' this works because the note find
                ' does not account for selection status
                ' when checking equality! (equalsBMSE, equalsNT)
                For i = 1 To UBound(Notes)
                    If i = j Then Continue For
                    If Notes(i).Selected Then
                        RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition, xUndo, xRedo)
                    End If
                Next

                Exit For
            End If
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)

        ValidateNotesArray()

        RefreshPanelAll()
    End Sub

    Private Sub SelectAllWithHoveredNoteLabel()
        For Each note In Notes
            note.Selected = IsLabelMatch(note, CurrentHoveredNote) Or note.Selected
        Next
    End Sub
End Class
