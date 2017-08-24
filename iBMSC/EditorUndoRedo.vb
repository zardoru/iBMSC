Imports iBMSC.Editor

Partial Public Class MainWindow
    Private Sub Operate(ByVal sCmd As UndoRedo.LinkedURCmd)
        For xI2 As Integer = 1 To UBound(Notes)
            Notes(xI2).Selected = False
        Next
        LBeat.SelectedIndices.Clear()

        Do While sCmd IsNot Nothing
            Dim xType As Byte = sCmd.ofType

            Select Case xType
                Case UndoRedo.opAddNote
                    Dim xCmd As UndoRedo.AddNote = sCmd

                    ReDim Preserve Notes(UBound(Notes) + 1)
                    With Notes(UBound(Notes))
                        .ColumnIndex = xCmd.ColumnIndex
                        .VPosition = xCmd.VPosition
                        .Value = xCmd.Value
                        If NTInput Then .Length = xCmd.LongNote Else .LongNote = xCmd.LongNote
                        .Hidden = xCmd.Hidden
                        .Selected = xCmd.Selected And nEnabled(.ColumnIndex)
                    End With

                Case UndoRedo.opRemoveNote
                    Dim xCmd As UndoRedo.RemoveNote = sCmd
                    Dim xI2 As Integer = FindNoteIndex(xCmd.ColumnIndex, xCmd.VPosition, xCmd.Value, xCmd.LongNote, xCmd.Hidden)

                    If xI2 < Notes.Length Then
                        For xI3 As Integer = xI2 + 1 To UBound(Notes)
                            Notes(xI3 - 1) = Notes(xI3)
                        Next
                        ReDim Preserve Notes(UBound(Notes) - 1)
                    End If

                Case UndoRedo.opChangeNote
                    Dim xCmd As UndoRedo.ChangeNote = sCmd
                    Dim xI2 As Integer = FindNoteIndex(xCmd.ColumnIndex, xCmd.VPosition, xCmd.Value, xCmd.LongNote, xCmd.Hidden)

                    If xI2 < Notes.Length Then
                        With Notes(xI2)
                            .ColumnIndex = xCmd.NColumnIndex
                            .VPosition = xCmd.NVPosition
                            .Value = xCmd.NValue
                            If NTInput Then .Length = xCmd.NLongNote Else .LongNote = xCmd.LongNote
                            .Hidden = xCmd.NHidden
                            .Selected = xCmd.Selected And nEnabled(.ColumnIndex)
                        End With
                    End If

                Case UndoRedo.opMoveNote
                    Dim xCmd As UndoRedo.MoveNote = sCmd
                    Dim xI2 As Integer = FindNoteIndex(xCmd.ColumnIndex, xCmd.VPosition, xCmd.Value, xCmd.LongNote, xCmd.Hidden)

                    If xI2 < Notes.Length Then
                        With Notes(xI2)
                            .ColumnIndex = xCmd.NColumnIndex
                            .VPosition = xCmd.NVPosition
                            .Selected = xCmd.Selected And nEnabled(.ColumnIndex)
                        End With
                    End If

                Case UndoRedo.opLongNoteModify
                    Dim xCmd As UndoRedo.LongNoteModify = sCmd
                    Dim xI2 As Integer = FindNoteIndex(xCmd.ColumnIndex, xCmd.VPosition, xCmd.Value, xCmd.LongNote, xCmd.Hidden)

                    If xI2 < Notes.Length Then
                        With Notes(xI2)
                            If NTInput Then
                                .VPosition = xCmd.NVPosition
                                .Length = xCmd.NLongNote
                            Else
                                .LongNote = xCmd.NLongNote
                            End If
                            .Selected = xCmd.Selected And nEnabled(.ColumnIndex)
                        End With
                    End If

                Case UndoRedo.opHiddenNoteModify
                    Dim xCmd As UndoRedo.HiddenNoteModify = sCmd
                    Dim xI2 As Integer = FindNoteIndex(xCmd.ColumnIndex, xCmd.VPosition, xCmd.Value, xCmd.LongNote, xCmd.Hidden)

                    If xI2 < Notes.Length Then
                        Notes(xI2).Hidden = xCmd.NHidden
                        Notes(xI2).Selected = xCmd.Selected And nEnabled(Notes(xI2).ColumnIndex)
                    End If

                Case UndoRedo.opRelabelNote
                    Dim xCmd As UndoRedo.RelabelNote = sCmd
                    Dim xI2 As Integer = FindNoteIndex(xCmd.ColumnIndex, xCmd.VPosition, xCmd.Value, xCmd.LongNote, xCmd.Hidden)

                    If xI2 < Notes.Length Then
                        Notes(xI2).Value = xCmd.NValue
                        Notes(xI2).Selected = xCmd.Selected And nEnabled(Notes(xI2).ColumnIndex)
                    End If

                Case UndoRedo.opRemoveAllNotes
                    ReDim Preserve Notes(0)

                Case UndoRedo.opChangeMeasureLength
                    Dim xCmd As UndoRedo.ChangeMeasureLength = sCmd
                    Dim xxD As Long = GetDenominator(xCmd.Value / 192)
                    'Dim xDenom As Integer = 192 / GCD(xCmd.Value, 192.0R)
                    'If xDenom < 4 Then xDenom = 4
                    For Each xM As Integer In xCmd.Indices
                        MeasureLength(xM) = xCmd.Value
                        LBeat.Items(xM) = Add3Zeros(xM) & ": " & (xCmd.Value / 192) & IIf(xxD > 10000, "", " ( " & CLng(xCmd.Value / 192 * xxD) & " / " & xxD & " ) ")
                        LBeat.SelectedIndices.Add(xM)
                    Next
                    UpdateMeasureBottom()

                Case UndoRedo.opChangeTimeSelection
                    Dim xCmd As UndoRedo.ChangeTimeSelection = sCmd
                    vSelStart = xCmd.SelStart
                    vSelLength = xCmd.SelLength
                    vSelHalf = xCmd.SelHalf
                    If xCmd.Selected Then
                        Dim xSelLo As Double = vSelStart + IIf(vSelLength < 0, vSelLength, 0)
                        Dim xSelHi As Double = vSelStart + IIf(vSelLength > 0, vSelLength, 0)
                        For xI2 As Integer = 1 To UBound(Notes)
                            Notes(xI2).Selected = Notes(xI2).VPosition >= xSelLo AndAlso
                                              Notes(xI2).VPosition < xSelHi AndAlso
                                              nEnabled(Notes(xI2).ColumnIndex)
                        Next
                    End If

                Case UndoRedo.opNT
                    Dim xCmd As UndoRedo.NT = sCmd
                    NTInput = xCmd.BecomeNT
                    TBNTInput.Checked = NTInput
                    mnNTInput.Checked = NTInput

                    POBLong.Enabled = Not NTInput
                    POBLongShort.Enabled = Not NTInput
                    bAdjustLength = False
                    bAdjustUpper = False

                    If xCmd.AutoConvert Then
                        If NTInput Then ConvertBMSE2NT() Else ConvertNT2BMSE()
                    End If

                Case UndoRedo.opVoid

                Case UndoRedo.opNoOperation
                    'Exit Do

            End Select

            sCmd = sCmd.Next
        Loop

        THBPM.Value = Notes(0).Value / 10000
        If IsSaved Then SetIsSaved(False)

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub AddUndo(ByVal sCUndo As UndoRedo.LinkedURCmd, ByVal sCRedo As UndoRedo.LinkedURCmd, Optional ByVal OverWrite As Boolean = False)
        If sCUndo Is Nothing And sCRedo Is Nothing Then Exit Sub
        If IsSaved Then SetIsSaved(False)
        If Not OverWrite Then sI = sIA()

        'ClearURReference(sUndo(sI))
        'ClearURReference(sRedo(sI))
        'ClearURReference(sUndo(sIA))
        'ClearURReference(sRedo(sIA))
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
        'For xI1 As Integer = 0 To 99
        '    ClearURReference(sUndo(xI1))
        '    ClearURReference(sRedo(xI1))
        'Next

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

    Private Sub RedoAddNote(ByVal xCol As Integer, ByVal xVPos As Double, ByVal xVal As Long, ByVal xLong As Double, ByVal xHide As Boolean,
    ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.RemoveNote(xCol, xVPos, xVal, xLong, xHide)
        Dim xRedo As New UndoRedo.AddNote(xCol, xVPos, xVal, xLong, xHide, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoAddNote(ByVal xN As Note, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.RemoveNote(xN.ColumnIndex,
                                             xN.VPosition,
                                             xN.Value,
                                IIf(NTInput, xN.Length, xN.LongNote),
                                             xN.Hidden)
        Dim xRedo As New UndoRedo.AddNote(xN.ColumnIndex,
                                          xN.VPosition,
                                          xN.Value,
                             IIf(NTInput, xN.Length, xN.LongNote),
                                          xN.Hidden,
                                          xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoAddNote(ByVal xIndices() As Integer, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For xI1 As Integer = 0 To UBound(xIndices)
            With Notes(xIndices(xI1))
                Dim xUndo As New UndoRedo.RemoveNote(.ColumnIndex,
                                                     .VPosition,
                                                     .Value,
                                        IIf(NTInput, .Length, .LongNote),
                                                     .Hidden)
                Dim xRedo As New UndoRedo.AddNote(.ColumnIndex,
                                                  .VPosition,
                                                  .Value,
                                     IIf(NTInput, .Length, .LongNote),
                                                  .Hidden,
                                                  xSel)
                xUndo.Next = BaseUndo
                BaseUndo = xUndo
                If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
                BaseRedo = xRedo
            End With
        Next
    End Sub

    Private Sub RedoAddNoteSelected(ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For
            With Notes(xI1)
                Dim xUndo As New UndoRedo.RemoveNote(.ColumnIndex,
                                                     .VPosition,
                                                     .Value,
                                        IIf(NTInput, .Length, .LongNote),
                                                     .Hidden)
                Dim xRedo As New UndoRedo.AddNote(.ColumnIndex,
                                                  .VPosition,
                                                  .Value,
                                     IIf(NTInput, .Length, .LongNote),
                                                  .Hidden,
                                                  xSel)
                xUndo.Next = BaseUndo
                BaseUndo = xUndo
                If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
                BaseRedo = xRedo
            End With
        Next
    End Sub

    Private Sub RedoAddNoteAll(ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For xI1 As Integer = 1 To UBound(Notes)
            With Notes(xI1)
                Dim xRedo As New UndoRedo.AddNote(.ColumnIndex,
                                                  .VPosition,
                                                  .Value,
                                     IIf(NTInput, .Length, .LongNote),
                                                  .Hidden,
                                                  xSel)
                If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
                BaseRedo = xRedo
            End With
        Next
        Dim xUndo As New UndoRedo.RemoveAllNotes
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
    End Sub

    Private Sub RedoRemoveNote(ByVal xCol As Integer, ByVal xVPos As Double, ByVal xVal As Integer, ByVal xLong As Double, ByVal xHide As Boolean,
    ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.AddNote(xCol, xVPos, xVal, xLong, xHide, xSel)
        Dim xRedo As New UndoRedo.RemoveNote(xCol, xVPos, xVal, xLong, xHide)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoRemoveNote(ByVal xN As Note, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.AddNote(xN.ColumnIndex,
                                          xN.VPosition,
                                          xN.Value,
                             IIf(NTInput, xN.Length, xN.LongNote),
                                          xN.Hidden,
                                          xSel)
        Dim xRedo As New UndoRedo.RemoveNote(xN.ColumnIndex,
                                             xN.VPosition,
                                             xN.Value,
                                IIf(NTInput, xN.Length, xN.LongNote),
                                             xN.Hidden)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoRemoveNote(ByVal xIndices() As Integer, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For xI1 As Integer = 0 To UBound(xIndices)
            With Notes(xIndices(xI1))
                Dim xUndo As New UndoRedo.AddNote(.ColumnIndex,
                                                  .VPosition,
                                                  .Value,
                                     IIf(NTInput, .Length, .LongNote),
                                                  .Hidden,
                                                  xSel)
                Dim xRedo As New UndoRedo.RemoveNote(.ColumnIndex,
                                                     .VPosition,
                                                     .Value,
                                        IIf(NTInput, .Length, .LongNote),
                                                     .Hidden)
                xUndo.Next = BaseUndo
                BaseUndo = xUndo
                If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
                BaseRedo = xRedo
            End With
        Next
    End Sub

    Private Sub RedoRemoveNoteSelected(ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For
            With Notes(xI1)
                Dim xUndo As New UndoRedo.AddNote(.ColumnIndex,
                                                  .VPosition,
                                                  .Value,
                                     IIf(NTInput, .Length, .LongNote),
                                                  .Hidden,
                                                  xSel)
                Dim xRedo As New UndoRedo.RemoveNote(.ColumnIndex,
                                                     .VPosition,
                                                     .Value,
                                        IIf(NTInput, .Length, .LongNote),
                                                     .Hidden)
                xUndo.Next = BaseUndo
                BaseUndo = xUndo
                If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
                BaseRedo = xRedo
            End With
        Next
    End Sub

    Private Sub RedoRemoveNoteAll(ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        For xI1 As Integer = 1 To UBound(Notes)
            With Notes(xI1)
                Dim xUndo As New UndoRedo.AddNote(.ColumnIndex,
                                                  .VPosition,
                                                  .Value,
                                     IIf(NTInput, .Length, .LongNote),
                                                  .Hidden,
                                                  xSel)
                xUndo.Next = BaseUndo
                BaseUndo = xUndo
            End With
        Next
        Dim xRedo As New UndoRedo.RemoveAllNotes
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoChangeNote(ByVal xCol As Integer, ByVal xVPos As Double, ByVal xVal As Integer, ByVal xLong As Double, ByVal xHide As Boolean,
                               ByVal nCol As Integer, ByVal nVPos As Double, ByVal nVal As Integer, ByVal nLong As Double, ByVal nHide As Boolean,
    ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.ChangeNote(nCol, nVPos, nVal, nLong, nHide, xCol, xVPos, xVal, xLong, xHide, xSel)
        Dim xRedo As New UndoRedo.ChangeNote(xCol, xVPos, xVal, xLong, xHide, nCol, nVPos, nVal, nLong, nHide, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoChangeNote(ByVal xN As Note, ByVal nCol As Integer, ByVal nVPos As Double, ByVal nVal As Integer, ByVal nLong As Double, ByVal nHide As Boolean,
    ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.ChangeNote(nCol, nVPos, nVal, nLong, nHide,
                                             xN.ColumnIndex,
                                             xN.VPosition,
                                             xN.Value,
                                IIf(NTInput, xN.Length, xN.LongNote),
                                             xN.Hidden,
                                             xSel)
        Dim xRedo As New UndoRedo.ChangeNote(xN.ColumnIndex,
                                             xN.VPosition,
                                             xN.Value,
                                IIf(NTInput, xN.Length, xN.LongNote),
                                             xN.Hidden,
                                             nCol, nVPos, nVal, nLong, nHide, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoChangeNote(ByVal xN As Note, ByVal nN As Note, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.ChangeNote(nN.ColumnIndex,
                                             nN.VPosition,
                                             nN.Value,
                                IIf(NTInput, nN.Length, nN.LongNote),
                                             nN.Hidden,
                                             xN.ColumnIndex,
                                             xN.VPosition,
                                             xN.Value,
                                IIf(NTInput, xN.Length, xN.LongNote),
                                             xN.Hidden,
                                             xSel)
        Dim xRedo As New UndoRedo.ChangeNote(xN.ColumnIndex,
                                             xN.VPosition,
                                             xN.Value,
                                IIf(NTInput, xN.Length, xN.LongNote),
                                             xN.Hidden,
                                             nN.ColumnIndex,
                                             nN.VPosition,
                                             nN.Value,
                                IIf(NTInput, nN.Length, nN.LongNote),
                                             nN.Hidden,
                                             xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoMoveNote(ByVal xCol As Integer, ByVal xVPos As Double, ByVal xVal As Integer, ByVal xLong As Double, ByVal xHide As Boolean,
                             ByVal nCol As Integer, ByVal nVPos As Double, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.MoveNote(nCol, nVPos, xVal, xLong, xHide, xCol, xVPos, xSel)
        Dim xRedo As New UndoRedo.MoveNote(xCol, xVPos, xVal, xLong, xHide, nCol, nVPos, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoMoveNote(ByVal xN As Note, ByVal nCol As Integer, ByVal nVPos As Double,
    ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.MoveNote(nCol, nVPos,
                                           xN.Value,
                              IIf(NTInput, xN.Length, xN.LongNote),
                                           xN.Hidden,
                                           xN.ColumnIndex,
                                           xN.VPosition,
                                           xSel)
        Dim xRedo As New UndoRedo.MoveNote(xN.ColumnIndex,
                                           xN.VPosition,
                                           xN.Value,
                              IIf(NTInput, xN.Length, xN.LongNote),
                                           xN.Hidden,
                                           nCol, nVPos, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoLongNoteModify(ByVal xCol As Integer, ByVal xVPos As Double, ByVal xVal As Integer, ByVal xLong As Double, ByVal xHide As Boolean,
                                   ByVal nVPos As Double, ByVal nLong As Double, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.LongNoteModify(xCol, nVPos, xVal, nLong, xHide, xVPos, xLong, xSel)
        Dim xRedo As New UndoRedo.LongNoteModify(xCol, xVPos, xVal, xLong, xHide, nVPos, nLong, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoLongNoteModify(ByVal xN As Note, ByVal nVPos As Double, ByVal nLong As Double,
    ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.LongNoteModify(xN.ColumnIndex, nVPos, xN.Value, nLong, xN.Hidden,
                                                                 xN.VPosition, IIf(NTInput, xN.Length, xN.LongNote), xSel)
        Dim xRedo As New UndoRedo.LongNoteModify(xN.ColumnIndex,
                                                 xN.VPosition,
                                                 xN.Value,
                                    IIf(NTInput, xN.Length, xN.LongNote),
                                                 xN.Hidden,
                                                 nVPos, nLong, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoHiddenNoteModify(ByVal xCol As Integer, ByVal xVPos As Double, ByVal xVal As Integer, ByVal xLong As Double, ByVal xHide As Boolean,
    ByVal nHide As Boolean, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.HiddenNoteModify(xCol, xVPos, xVal, xLong, nHide, xHide, xSel)
        Dim xRedo As New UndoRedo.HiddenNoteModify(xCol, xVPos, xVal, xLong, xHide, nHide, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoHiddenNoteModify(ByVal xN As Note, ByVal nHide As Boolean,
    ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.HiddenNoteModify(xN.ColumnIndex,
                                                   xN.VPosition,
                                                   xN.Value,
                                      IIf(NTInput, xN.Length, xN.LongNote),
                                                   nHide,
                                                   xN.Hidden,
                                                   xSel)
        Dim xRedo As New UndoRedo.HiddenNoteModify(xN.ColumnIndex,
                                                   xN.VPosition,
                                                   xN.Value,
                                      IIf(NTInput, xN.Length, xN.LongNote),
                                                   xN.Hidden,
                                                   nHide,
                                                   xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoRelabelNote(ByVal xCol As Integer, ByVal xVPos As Double, ByVal xVal As Integer, ByVal xLong As Double, ByVal xHide As Boolean,
    ByVal nVal As Integer, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.RelabelNote(xCol, xVPos, nVal, xLong, xHide, xVal, xSel)
        Dim xRedo As New UndoRedo.RelabelNote(xCol, xVPos, xVal, xLong, xHide, nVal, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoRelabelNote(ByVal xN As Note, ByVal nVal As Integer, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.RelabelNote(xN.ColumnIndex,
                                              xN.VPosition,
                                              nVal,
                                 IIf(NTInput, xN.Length, xN.LongNote),
                                              xN.Hidden,
                                              xN.Value,
                                              xSel)
        Dim xRedo As New UndoRedo.RelabelNote(xN.ColumnIndex,
                                              xN.VPosition,
                                              xN.Value,
                                 IIf(NTInput, xN.Length, xN.LongNote),
                                              xN.Hidden,
                                              nVal,
                                              xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoChangeMeasureLengthSelected(ByVal nVal As Double, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xIndices(LBeat.SelectedIndices.Count - 1) As Integer
        LBeat.SelectedIndices.CopyTo(xIndices, 0)
        If xIndices.Length = 0 Then Exit Sub

        Dim xmLen(-1) As Double
        Dim xUndo(-1) As UndoRedo.ChangeMeasureLength
        For Each xI1 As Integer In xIndices
            Dim xI As Integer = Array.IndexOf(xmLen, MeasureLength(xI1))
            If xI = -1 Then
                ReDim Preserve xmLen(UBound(xmLen) + 1)
                ReDim Preserve xUndo(UBound(xUndo) + 1)
                xmLen(UBound(xmLen)) = MeasureLength(xI1)
                xUndo(UBound(xUndo)) = New UndoRedo.ChangeMeasureLength(MeasureLength(xI1), New Integer() {xI1})
            Else
                With xUndo(xI)
                    ReDim Preserve .Indices(UBound(.Indices) + 1)
                    .Indices(UBound(.Indices)) = xI1
                End With
            End If
        Next
        For xI1 As Integer = 0 To UBound(xUndo) - 1
            xUndo(xI1).Next = xUndo(xI1 + 1)
        Next
        xUndo(UBound(xUndo)).Next = BaseUndo
        BaseUndo = xUndo(0)

        Dim xRedo As New UndoRedo.ChangeMeasureLength(nVal, xIndices.Clone)
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    Private Sub RedoChangeTimeSelection(ByVal pStart As Double, ByVal pLen As Double, ByVal pHalf As Double,
    ByVal nStart As Double, ByVal nLen As Double, ByVal nHalf As Double, ByVal xSel As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.ChangeTimeSelection(pStart, pLen, pHalf, xSel)
        Dim xRedo As New UndoRedo.ChangeTimeSelection(nStart, nLen, nHalf, xSel)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub

    'Private Sub RedoChangeVisibleColumns(ByVal pBLP As Boolean, ByVal pSTOP As Boolean, ByVal pPlayer As Integer, _
    '                                     ByVal nBLP As Boolean, ByVal nSTOP As Boolean, ByVal nPlayer As Integer, _
    'ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
    '    Dim xUndo As New UndoRedo.ChangeVisibleColumns(pBLP, pSTOP, pPlayer)
    '    Dim xRedo As New UndoRedo.ChangeVisibleColumns(nBLP, nSTOP, nPlayer)
    '    xUndo.Next = BaseUndo
    '    BaseUndo = xUndo
    '    If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
    '    BaseRedo = xRedo
    'End Sub

    Private Sub RedoNT(ByVal becomeNT As Boolean, ByVal autoConvert As Boolean, ByRef BaseUndo As UndoRedo.LinkedURCmd, ByRef BaseRedo As UndoRedo.LinkedURCmd)
        Dim xUndo As New UndoRedo.NT(Not becomeNT, autoConvert)
        Dim xRedo As New UndoRedo.NT(becomeNT, autoConvert)
        xUndo.Next = BaseUndo
        BaseUndo = xUndo
        If BaseRedo IsNot Nothing Then BaseRedo.Next = xRedo
        BaseRedo = xRedo
    End Sub
End Class
