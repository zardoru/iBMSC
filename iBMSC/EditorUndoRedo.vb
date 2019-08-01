Imports iBMSC.Editor


'Undo appends before, Redo appends after.
'After a sequence of Commands, 
'   Undo will be the first one to execute, 
'   Redo will be the last one to execute.
'Remember to save the first Redo.

'In case where undo is Nothing: Dont worry.
'In case where redo is Nothing: 
'   If only one redo is in a sequence, put Nothing.
'   If several redo are in a sequence, 
'       Create Void first. 
'       Record its reference into a seperate copy. (xBaseRedo = xRedo)
'       Use this xRedo as the BaseRedo.
'       When calling AddUndo subroutine, use xBaseRedo.Next as cRedo.

'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
'Dim xRedo As UndoRedo.LinkedURCmd = Nothing
'... 'Command.RedoRemoveNote(K(i), True, xUndo, xRedo)
'AddUndo(xUndo, xRedo)

'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
'Dim xRedo As New UndoRedo.Void
'Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
'... 'Command.RedoRemoveNote(K(i), True, xUndo, xRedo)
'AddUndo(xUndo, xBaseRedo.Next)


Module Command
    Public Sub RedoRemoveNote(xN As Note, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.AddNote(xN.Clone)
        Dim xRedo As New UndoRedo.RemoveNote(xN.Clone)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Public Sub RedoAddNote(note As Note,
                           ByRef BaseUndo As UndoRedo.LinkedURCmd,
                           ByRef BaseRedo As UndoRedo.LinkedURCmd,
                           Optional autoinc As Boolean = False)
        Dim xUndo As New UndoRedo.RemoveNote(note.Clone)
        Dim xRedo As New UndoRedo.AddNote(note.Clone)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Public Sub RedoAddNote(xIndices() As Note, xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                           ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For Each note In xIndices
            Dim xUndo As New UndoRedo.RemoveNote(note.Clone)
            Dim xRedo As New UndoRedo.AddNote(note.Clone)
            xUndo.Next = BaseUndo
            BaseUndo = xUndo
            If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
            BaseRedo = xRedo
        Next
    End Sub

    Public Sub RedoAddNoteSelected(Notes() As Note, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                                   ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For i = 1 To Notes.Length - 1
            If Not Notes(i).Selected Then Continue For

            Dim xUndo As New UndoRedo.RemoveNote(Notes(i).Clone)
            Dim xRedo As New UndoRedo.AddNote(Notes(i).Clone)
            xUndo.Next = BaseUndo
            BaseUndo = xUndo
            If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
            BaseRedo = xRedo

        Next
    End Sub

    Public Sub RedoAddNoteAll(Notes() As Note, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                              ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For Each note In Notes

            Dim xRedo As New UndoRedo.AddNote(note.Clone)
            If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
            BaseRedo = xRedo

        Next
        Dim xUndo As New UndoRedo.RemoveAllNotes
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
    End Sub

    Public Sub RedoRemoveNote(xIndices() As Note, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                              ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For Each note In xIndices
            Dim xUndo As New UndoRedo.AddNote(note.Clone)
            Dim xRedo As New UndoRedo.RemoveNote(note.Clone)
            xUndo.Next = BaseUndo
            BaseUndo = xUndo
            If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
            BaseRedo = xRedo
        Next
    End Sub

    Public Sub RedoRemoveNoteSelected(Notes() As Note, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                                      ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For i = 1 To Notes.Length - 1
            If Not Notes(i).Selected Then Continue For

            Dim xUndo As New UndoRedo.AddNote(Notes(i).Clone)
            Dim xRedo As New UndoRedo.RemoveNote(Notes(i).Clone)
            xUndo.Next = BaseUndo
            BaseUndo = xUndo
            If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
            BaseRedo = xRedo
        Next
    End Sub

    Public Sub RedoRemoveNoteAll(Notes() As Note, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                                 ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For i = 1 To Notes.Length - 1
            Dim xUndo As New UndoRedo.AddNote(Notes(i).Clone)
            xUndo.Next = BaseUndo
            BaseUndo = xUndo

        Next
        Dim xRedo As New UndoRedo.RemoveAllNotes
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Public Sub RedoChangeNote(startNote As Note, endNote As Note, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                              ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.ChangeNote(endNote.Clone, startNote.Clone)
        Dim xRedo As New UndoRedo.ChangeNote(startNote.Clone, endNote.Clone)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub


    Public Sub RedoMoveNote(note As Note, nCol As Integer, nVPos As Double, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                            ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim noteAfterModification = note.Clone
        noteAfterModification.ColumnIndex = nCol
        noteAfterModification.VPosition = nVPos
        Dim xUndo As New UndoRedo.MoveNote(noteAfterModification, note.ColumnIndex, note.VPosition)
        Dim xRedo As New UndoRedo.MoveNote(note, nCol, nVPos)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub


    Public Sub RedoLongNoteModify(note As Note, nVPos As Double, nLong As Double, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                                  ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim n = note.Clone
        n.VPosition = nVPos
        n.Length = nLong

        Dim xUndo As New UndoRedo.LongNoteModify(n, note.VPosition, note.Length)
        Dim xRedo As New UndoRedo.LongNoteModify(note, nVPos, n.Length)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Public Sub RedoHiddenNoteModify(xN As Note, nHide As Boolean, xSel As Boolean,
                                    ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim noteAfterModification = xN.Clone
        noteAfterModification.Hidden = nHide
        Dim xUndo As New UndoRedo.HiddenNoteModify(noteAfterModification, xN.Hidden)
        Dim xRedo As New UndoRedo.HiddenNoteModify(xN, nHide)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Public Sub RedoRelabelNote(xN As Note, nVal As Long, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                               ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim noteAfterModification = xN.Clone
        noteAfterModification.Value = nVal
        Dim xUndo As New UndoRedo.RelabelNote(noteAfterModification, xN.Value)
        Dim xRedo As New UndoRedo.RelabelNote(xN, nVal)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Public Sub RedoChangeTimeSelection(pStart As Double, pLen As Double, pHalf As Double,
                                       nStart As Double, nLen As Double, nHalf As Double, xSel As Boolean,
                                       ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.ChangeTimeSelection(pStart, pLen, pHalf, xSel)
        Dim xRedo As New UndoRedo.ChangeTimeSelection(nStart, nLen, nHalf, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    'Public Sub RedoChangeVisibleColumns(ByVal pBLP As Boolean, ByVal pSTOP As Boolean, ByVal pPlayer As Integer, _
    '                                     ByVal nBLP As Boolean, ByVal nSTOP As Boolean, ByVal nPlayer As Integer, _
    'ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
    '    Dim xUndo As New UndoRedo.ChangeVisibleColumns(pBLP, pSTOP, pPlayer)
    '    Dim xRedo As New UndoRedo.ChangeVisibleColumns(nBLP, nSTOP, nPlayer)
    '    xUndo.Next = BaseUndo
    '    BaseUndo = xUndo
    '    If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
    '    BaseRedo = xRedo
    'End Sub

    Public Sub RedoNT(becomeNT As Boolean, autoConvert As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                      ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.NT(Not becomeNT, autoConvert)
        Dim xRedo As New UndoRedo.NT(becomeNT, autoConvert)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Public Sub RedoWavIncrease(wavinc As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                               ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.WavAutoincFlag(Not wavinc)
        Dim xRedo As New UndoRedo.WavAutoincFlag(wavinc)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseUndo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub
End Module

Partial Public Class MainWindow
    'Variables for undo/redo
    Dim sUndo(99) As UndoRedo.LinkedURCmd
    Dim sRedo(99) As UndoRedo.LinkedURCmd
    Dim sI As Integer = 0

    ' az: this belongs here and I suspect this is the "next undo slot"
    ' but no confirmation yet.
    Private Function sIA() As Integer
        Return IIf(sI > 98, 0, sI + 1)
    End Function

    Private Function sIM() As Integer
        Return IIf(sI < 1, 99, sI - 1)
    End Function

    Public Sub RedoChangeMeasureLengthSelected(nVal As Double, ByRef BaseUndo As UndoRedo.LinkedURCmd,
                                               ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xIndices(LBeat.SelectedIndices.Count - 1) As Integer
        LBeat.SelectedIndices.CopyTo(xIndices, 0)
        If xIndices.Length = 0 Then Exit Sub

        Dim xmLen(- 1) As Double
        Dim xUndo(- 1) As UndoRedo.ChangeMeasureLength
        For Each i As Integer In xIndices
            Dim xI As Integer = Array.IndexOf(xmLen, MeasureLength(i))
            If xI = - 1 Then
                ReDim Preserve xmLen(UBound(xmLen) + 1)
                ReDim Preserve xUndo(UBound(xUndo) + 1)
                xmLen(UBound(xmLen)) = MeasureLength(i)
                xUndo(UBound(xUndo)) = New UndoRedo.ChangeMeasureLength(MeasureLength(i), New Integer() {i})
            Else
                With xUndo(xI)
                    ReDim Preserve .Indices(UBound(.Indices) + 1)
                    .Indices(UBound(.Indices)) = i
                End With
            End If
        Next
        For i = 0 To UBound(xUndo) - 1
            xUndo(i).Next = xUndo(i + 1)
        Next
        xUndo(UBound(xUndo)).Next = BaseUndo
        BaseUndo = xUndo(0)

        Dim xRedo As New UndoRedo.ChangeMeasureLength(nVal, xIndices.Clone)
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub PerformCommand(sCmd As UndoRedo.LinkedURCmd)
        For j = 1 To UBound(Notes)
            Notes(j).Selected = False
        Next
        LBeat.SelectedIndices.Clear()

        Do While sCmd IsNot Nothing
            Dim xType As Byte = sCmd.ofType

            Select Case xType
                Case UndoRedo.opAddNote
                    Dim xCmd As UndoRedo.AddNote = sCmd

                    ReDim Preserve Notes(UBound(Notes) + 1)
                    Notes(UBound(Notes)) = xCmd.note

                    If TBWavIncrease.Checked Then
                        IncreaseCurrentWav()
                    End If
                Case UndoRedo.opRemoveNote
                    Dim xCmd As UndoRedo.RemoveNote = sCmd
                    Dim j As Integer = FindNoteIndex(xCmd.note)

                    If j < Notes.Length Then
                        For xI3 As Integer = j + 1 To UBound(Notes)
                            Notes(xI3 - 1) = Notes(xI3)
                        Next
                        ReDim Preserve Notes(UBound(Notes) - 1)
                    End If

                    If TBWavIncrease.Checked Then
                        DecreaseCurrentWav()
                    End If

                Case UndoRedo.opChangeNote
                    Dim xCmd As UndoRedo.ChangeNote = sCmd
                    Dim j As Integer = FindNoteIndex(xCmd.note)

                    If j < Notes.Length Then
                        Notes(j) = xCmd.note
                    End If

                Case UndoRedo.opMoveNote
                    Dim xCmd As UndoRedo.MoveNote = sCmd
                    Dim j As Integer = FindNoteIndex(xCmd.note)

                    If j < Notes.Length Then
                        With Notes(j)
                            .ColumnIndex = xCmd.NColumnIndex
                            .VPosition = xCmd.NVPosition
                            .Selected = xCmd.note.Selected And Columns.IsEnabled(.ColumnIndex)
                        End With
                    End If

                Case UndoRedo.opLongNoteModify
                    Dim xCmd As UndoRedo.LongNoteModify = sCmd
                    Dim j As Integer = FindNoteIndex(xCmd.note)

                    If j < Notes.Length Then
                        With Notes(j)
                            If NTInput Then
                                .VPosition = xCmd.NVPosition
                                .Length = xCmd.NLongNote
                            Else
                                .LongNote = xCmd.NLongNote
                            End If
                            .Selected = xCmd.note.Selected And Columns.IsEnabled(.ColumnIndex)
                        End With
                    End If

                Case UndoRedo.opHiddenNoteModify
                    Dim xCmd As UndoRedo.HiddenNoteModify = sCmd
                    Dim j As Integer = FindNoteIndex(xCmd.note)

                    If j < Notes.Length Then
                        Notes(j).Hidden = xCmd.NHidden
                        Notes(j).Selected = xCmd.note.Selected And Columns.IsEnabled(Notes(j).ColumnIndex)
                    End If

                Case UndoRedo.opRelabelNote
                    Dim xCmd As UndoRedo.RelabelNote = sCmd
                    Dim j As Integer = FindNoteIndex(xCmd.note)

                    If j < Notes.Length Then
                        Notes(j).Value = xCmd.NValue
                        Notes(j).Selected = xCmd.note.Selected And Columns.IsEnabled(Notes(j).ColumnIndex)
                    End If

                Case UndoRedo.opRemoveAllNotes
                    ReDim Preserve Notes(0)

                Case UndoRedo.opChangeMeasureLength
                    Dim xCmd As UndoRedo.ChangeMeasureLength = sCmd
                    Dim xxD As Long = GetDenominator(xCmd.Value/192)
                    'Dim xDenom As Integer = 192 / GCD(xCmd.Value, 192.0R)
                    'If xDenom < 4 Then xDenom = 4
                    For Each xM As Integer In xCmd.Indices
                        MeasureLength(xM) = xCmd.Value
                        LBeat.Items(xM) = Add3Zeros(xM) & ": " & (xCmd.Value/192) &
                                          IIf(xxD > 10000, "", " ( " & CLng(xCmd.Value/192*xxD) & " / " & xxD & " ) ")
                        LBeat.SelectedIndices.Add(xM)
                    Next
                    UpdateMeasureBottom()

                Case UndoRedo.opChangeTimeSelection
                    Dim xCmd As UndoRedo.ChangeTimeSelection = sCmd
                    State.TimeSelect.StartPoint = xCmd.SelStart
                    State.TimeSelect.EndPointLength = xCmd.SelLength
                    State.TimeSelect.HalfPointLength = xCmd.SelHalf
                    If xCmd.Selected Then
                        Dim xSelLo As Double = State.TimeSelect.StartPoint +
                                               Math.Min(State.TimeSelect.EndPointLength, 0)
                        Dim xSelHi As Double = State.TimeSelect.StartPoint +
                                               Math.Max(State.TimeSelect.EndPointLength, 0)
                        For j = 1 To UBound(Notes)
                            Notes(j).Selected = Notes(j).VPosition >= xSelLo AndAlso
                                                Notes(j).VPosition < xSelHi AndAlso
                                                Columns.IsEnabled(Notes(j).ColumnIndex)
                        Next
                    End If

                Case UndoRedo.opNT
                    Dim xCmd As UndoRedo.NT = sCmd
                    NTInput = xCmd.BecomeNT
                    TBNTInput.Checked = NTInput
                    mnNTInput.Checked = NTInput

                    POBLong.Enabled = Not NTInput
                    POBLongShort.Enabled = Not NTInput
                    State.NT.IsAdjustingNoteLength = False
                    State.NT.IsAdjustingUpperEnd = False

                    If xCmd.AutoConvert Then
                        If NTInput Then ConvertBMSE2NT() Else ConvertNT2BMSE()
                    End If
                Case UndoRedo.opWavAutoincFlag
                    Dim xcmd As UndoRedo.WavAutoincFlag = sCmd
                    TBWavIncrease.Checked = xcmd.Checked

                Case UndoRedo.opVoid

                Case UndoRedo.opNoOperation
                    'Exit Do

            End Select

            sCmd = sCmd.Next
        Loop

        THBPM.Value = Notes(0).Value/10000
        If _isSaved Then SetIsSaved(False)

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Public Sub AddUndoChain(sCUndo As UndoRedo.LinkedURCmd,
                            sCRedo As UndoRedo.LinkedURCmd,
                            Optional ByVal OverWrite As Boolean = False)
        If sCUndo Is Nothing And sCRedo Is Nothing Then Exit Sub
        If _isSaved Then SetIsSaved(False)
        If Not OverWrite Then sI = sIA()

        sUndo(sI) = sCUndo
        sRedo(sI) = sCRedo
        sUndo(sIA) = New UndoRedo.NoOperation
        sRedo(sIA) = New UndoRedo.NoOperation
        TBUndo.Enabled = True
        TBRedo.Enabled = False
        mnUndo.Enabled = True
        mnRedo.Enabled = False
    End Sub

    Private Sub ClearUndo()

        ReDim sUndo(99)
        ReDim sRedo(99)
        sUndo(0) = New UndoRedo.NoOperation
        sUndo(1) = New UndoRedo.NoOperation
        sRedo(0) = New UndoRedo.NoOperation
        sRedo(1) = New UndoRedo.NoOperation
        sI = 0
        TBUndo.Enabled = False
        TBRedo.Enabled = False
        mnUndo.Enabled = False
        mnRedo.Enabled = False
    End Sub
End Class
