Imports System.Linq
Imports iBMSC.Editor

Partial Public Class MainWindow

    Private Sub BVCCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BVCCalculate.Click
        If Not TBTimeSelect.Checked Then Exit Sub

        SortByVPositionInsertion()
        BPMChangeByValue(Val(TVCBPM.Text) * 10000)

        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
        POStatusRefresh()

        Beep()
        TVCBPM.Focus()
    End Sub

    Private Sub BVCApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BVCApply.Click
        If Not TBTimeSelect.Checked Then Exit Sub

        SortByVPositionInsertion()
        BPMChangeTop(Val(TVCM.Text) / Val(TVCD.Text))

        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
        POStatusRefresh()
        CalculateGreatestVPosition()

        Beep()
        TVCM.Focus()
        'Select Case spFocus
        '    Case 0 : PMainInL.Focus()
        '    Case 1 : PMainIn.Focus()
        '    Case 2 : PMainInR.Focus()
        'End Select
    End Sub

    Private Sub BPMChangeTop(ByVal xRatio As Double, Optional ByVal bAddUndo As Boolean = True, Optional ByVal bOverWriteUndo As Boolean = False)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If vSelLength = 0 Then GoTo EndofSub
        If xRatio = 1 Or xRatio <= 0 Then GoTo EndofSub

        Dim xVLower As Double = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        Dim xVUpper As Double = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If xVLower < 0 Then xVLower = 0
        If xVUpper >= GetMaxVPosition() Then xVUpper = GetMaxVPosition() - 1

        Dim xBPM As Integer = Notes(0).Value
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        Dim xValueL As Integer = xBPM
        Dim xValueU As Integer = xBPM

        'Save undo
        'For xI3 = 1 To UBound(K)
        '    K(xI3).Selected = True
        'Next
        'xUndo = "KZ" & vbCrLf & _
        '        sCmdKs(False) & vbCrLf & _
        '        "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"

        Me.RedoRemoveNoteAll(False, xUndo, xRedo)

        'Start
        If Not NTInput Then
            'Below Selection
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition > xVLower Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then xBPM = Notes(xI1).Value
            Next
            xValueL = xBPM
            xI2 = xI1

            'Within Selection
            For xI1 = xI2 To UBound(Notes)
                If Notes(xI1).VPosition > xVUpper Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then
                    xBPM = Notes(xI1).Value
                    Notes(xI1).Value = Notes(xI1).Value * xRatio
                End If
                Notes(xI1).VPosition = (Notes(xI1).VPosition - xVLower) * xRatio + xVLower
            Next
            xValueU = xBPM
            xI2 = xI1

            'Above Selection
            For xI1 = xI2 To UBound(Notes)
                Notes(xI1).VPosition += (xRatio - 1) * (xVUpper - xVLower)
            Next

            'Add BPMs
            AddNote(New Note(niBPM, xVLower, xValueL * xRatio), False, True, False)
            AddNote(New Note(niBPM, xVUpper + (xRatio - 1) * (xVUpper - xVLower), xValueU), False, True, False)

        Else
            Dim xAddBPML As Boolean = True
            Dim xAddBPMU As Boolean = True

            For xI1 = 1 To UBound(Notes)
                'Modify notes
                If Notes(xI1).VPosition <= xVLower Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueL = Notes(xI1).Value
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVLower Then xAddBPML = False : Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio <= 655359999, Notes(xI1).Value * xRatio, 655359999)
                    End If

                    'If longnote then adjust length
                    If Notes(xI1).VPosition + Notes(xI1).Length > xVLower Then
                        Notes(xI1).Length += (IIf(xVUpper < Notes(xI1).VPosition + Notes(xI1).Length, xVUpper, Notes(xI1).VPosition + Notes(xI1).Length) - xVLower) * (xRatio - 1)
                    End If

                ElseIf Notes(xI1).VPosition <= xVUpper Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVUpper Then xAddBPMU = False Else Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio <= 655359999, Notes(xI1).Value * xRatio, 655359999)
                    End If

                    'Adjust Length
                    Notes(xI1).Length += (IIf(xVUpper < Notes(xI1).Length + Notes(xI1).VPosition, xVUpper, Notes(xI1).Length + Notes(xI1).VPosition) - Notes(xI1).VPosition) * (xRatio - 1)

                    'Adjust VPosition
                    Notes(xI1).VPosition = (Notes(xI1).VPosition - xVLower) * xRatio + xVLower

                Else
                    Notes(xI1).VPosition += (xVUpper - xVLower) * (xRatio - 1)
                End If
            Next

            'Add BPMs
            If xAddBPML Then AddNote(New Note(niBPM, xVLower, xValueL * xRatio), False, True, False)
            If xAddBPMU Then AddNote(New Note(niBPM, (xVUpper - xVLower) * xRatio + xVLower, xValueU), False, True, False)
        End If

        'Check BPM Overflow
        For xI3 = 1 To UBound(Notes)
            If Notes(xI3).ColumnIndex = niBPM AndAlso Notes(xI3).Value < 1 Then Notes(xI3).Value = 1
        Next

        Me.RedoAddNoteAll(False, xUndo, xRedo)

        'Restore selection
        Dim pSelStart As Double = vSelStart
        Dim pSelLength As Double = vSelLength
        Dim pSelHalf As Double = vSelHalf
        If vSelLength < 0 Then vSelStart += (xRatio - 1) * (xVUpper - xVLower)
        vSelLength = vSelLength * xRatio
        vSelHalf = vSelHalf * xRatio
        ValidateSelection()
        Me.RedoChangeTimeSelection(pSelStart, pSelLength, pSelHalf, vSelStart, vSelLength, vSelHalf, True, xUndo, xRedo)

        'Save redo
        'For xI3 = 1 To UBound(K)
        '    K(xI3).Selected = True
        'Next
        'xRedo = "KZ" & vbCrLf & _
        '                      sCmdKs(False) & vbCrLf & _
        '                      "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"

        'Restore note selection
        xVLower = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        xVUpper = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If Not NTInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And Notes(xI3).VPosition < xVUpper And nEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And nEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

EndofSub:
        If bAddUndo Then AddUndo(xUndo, xBaseRedo.Next, bOverWriteUndo)
    End Sub

    Private Sub BPMChangeHalf(ByVal dVPosition As Double, Optional ByVal bAddUndo As Boolean = True, Optional ByVal bOverWriteUndo As Boolean = False)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If vSelLength = 0 Then GoTo EndofSub
        If dVPosition = 0 Then GoTo EndofSub

        Dim xVLower As Double = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        Dim xVHalf As Double = vSelStart + vSelHalf
        Dim xVUpper As Double = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If dVPosition + xVHalf <= xVLower Or dVPosition + xVHalf >= xVUpper Then GoTo EndofSub

        If xVLower < 0 Then xVLower = 0
        If xVUpper >= GetMaxVPosition() Then xVUpper = GetMaxVPosition() - 1
        If xVHalf > xVUpper Then xVHalf = xVUpper
        If xVHalf < xVLower Then xVHalf = xVLower

        Dim xBPM As Integer = Notes(0).Value
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        Dim xValueL As Integer = xBPM
        Dim xValueM As Integer = xBPM
        Dim xValueU As Integer = xBPM

        Dim xRatio1 As Double = (xVHalf - xVLower + dVPosition) / (xVHalf - xVLower)
        Dim xRatio2 As Double = (xVUpper - xVHalf - dVPosition) / (xVUpper - xVHalf)

        'Save undo
        'For xI3 = 1 To UBound(K)
        'K(xI3).Selected = True
        'Next
        'xUndo = "KZ" & vbCrLf & _
        '        sCmdKs(False) & vbCrLf & _
        '        "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"

        Me.RedoRemoveNoteAll(False, xUndo, xRedo)

        If Not NTInput Then
            'Below Selection
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition > xVLower Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then xBPM = Notes(xI1).Value
            Next
            xValueL = xBPM
            xI2 = xI1

            'Below Half
            For xI1 = xI2 To UBound(Notes)
                If Notes(xI1).VPosition > xVHalf Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then
                    xBPM = Notes(xI1).Value
                    Notes(xI1).Value = Notes(xI1).Value * xRatio1
                End If
                Notes(xI1).VPosition = (Notes(xI1).VPosition - xVLower) * xRatio1 + xVLower
            Next
            xValueM = xBPM
            xI2 = xI1

            'Above Half
            For xI1 = xI2 To UBound(Notes)
                If Notes(xI1).VPosition > xVUpper Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then
                    xBPM = Notes(xI1).Value
                    Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio2 <= 655359999, Notes(xI1).Value * xRatio2, 655359999)
                End If
                Notes(xI1).VPosition = (Notes(xI1).VPosition - xVHalf) * xRatio2 + xVHalf + dVPosition
            Next
            xValueU = xBPM
            xI2 = xI1

            'Above Selection
            'For xI1 = xI2 To UBound(K)
            '    K(xI1).VPosition += (xRatio - 1) * (xVUpper - xVLower)
            'Next

            'Add BPMs
            ' az: cond. removed; 
            ' IIf(xVHalf <> xVLower AndAlso xValueL * xRatio1 <= 655359999, xValueL * xRatio1, 655359999)
            AddNote(New Note(niBPM, xVLower, xValueL * xRatio1), False, True, False)
            ' az: cond removed;
            ' IIf(xVHalf <> xVUpper AndAlso xValueM * xRatio2 <= 655359999, xValueM * xRatio2, 655359999)
            AddNote(New Note(niBPM, xVHalf + dVPosition, xValueM * xRatio2), False, True, False)
            AddNote(New Note(niBPM, xVUpper, xValueU), False, True, False)

        Else
            Dim xAddBPML As Boolean = True
            Dim xAddBPMM As Boolean = True
            Dim xAddBPMU As Boolean = True

            'Modify notes
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition <= xVLower Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueL = Notes(xI1).Value
                        xValueM = Notes(xI1).Value
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVLower Then
                            xAddBPML = False

                            ' az: condition removed;
                            ' IIf(xVHalf <> xVLower AndAlso Notes(xI1).Value * xRatio1 <= 655359999, Notes(xI1).Value * xRatio1, 655359999)
                            Notes(xI1).Value = Notes(xI1).Value * xRatio1
                        End If
                    End If

                    'If longnote then adjust length
                    Dim xEnd As Double = Notes(xI1).VPosition + Notes(xI1).Length
                    If xEnd > xVUpper Then
                    ElseIf xEnd > xVHalf Then
                        Notes(xI1).Length = (xEnd - xVHalf) * xRatio2 + xVHalf + dVPosition - Notes(xI1).VPosition
                    ElseIf xEnd > xVLower Then
                        Notes(xI1).Length = (xEnd - xVLower) * xRatio1 + xVLower - Notes(xI1).VPosition
                    End If

                ElseIf Notes(xI1).VPosition <= xVHalf Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueM = Notes(xI1).Value
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVHalf Then
                            xAddBPMM = False
                            ' az: cond. remove
                            ' IIf(xVHalf <> xVUpper AndAlso Notes(xI1).Value * xRatio2 <= 655359999, Notes(xI1).Value * xRatio2, 655359999)
                            Notes(xI1).Value = Notes(xI1).Value * xRatio2
                        Else
                            ' az: cond. remove
                            ' IIf(Notes(xI1).Value * xRatio1 <= 655359999, Notes(xI1).Value * xRatio1, 655359999)
                            Notes(xI1).Value = Notes(xI1).Value * xRatio1
                        End If
                    End If

                    'Adjust Length
                    Dim xEnd As Double = Notes(xI1).VPosition + Notes(xI1).Length
                    If xEnd > xVUpper Then
                        Notes(xI1).Length = xEnd - xVLower - (Notes(xI1).VPosition - xVLower) * xRatio1
                    ElseIf xEnd > xVHalf Then
                        Notes(xI1).Length = (xVHalf - Notes(xI1).VPosition) * xRatio1 + (xEnd - xVHalf) * xRatio2
                    Else
                        Notes(xI1).Length *= xRatio1
                    End If

                    'Adjust VPosition
                    Notes(xI1).VPosition = (Notes(xI1).VPosition - xVLower) * xRatio1 + xVLower

                ElseIf Notes(xI1).VPosition <= xVUpper Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVUpper Then xAddBPMU = False Else Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio2 <= 655359999, Notes(xI1).Value * xRatio2, 655359999)
                    End If

                    'Adjust Length
                    Dim xEnd As Double = Notes(xI1).VPosition + Notes(xI1).Length
                    If xEnd > xVUpper Then
                        Notes(xI1).Length = (xVUpper - Notes(xI1).VPosition) * xRatio2 + xEnd - xVUpper
                    Else
                        Notes(xI1).Length *= xRatio2
                    End If

                    'Adjust VPosition
                    Notes(xI1).VPosition = (Notes(xI1).VPosition - xVHalf) * xRatio2 + xVHalf + dVPosition

                    'Else
                    '    K(xI1).VPosition += (xVUpper - xVLower) * (xRatio - 1)
                End If
            Next

            'Add BPMs
            ' IIf(xVHalf <> xVLower AndAlso xValueL * xRatio1 <= 655359999, xValueL * xRatio1, 655359999)
            If xAddBPML Then AddNote(New Note(niBPM, xVLower, xValueL * xRatio1), False, True, False)
            ' IIf(xVHalf <> xVUpper AndAlso xValueM * xRatio2 <= 655359999, xValueM * xRatio2, 655359999)
            If xAddBPMM Then AddNote(New Note(niBPM, xVHalf + dVPosition, xValueM * xRatio2), False, True, False)
            If xAddBPMU Then AddNote(New Note(niBPM, xVUpper, xValueU), False, True, False)
        End If

        'Check BPM Overflow
        'For xI3 = 1 To UBound(Notes)
        '    If Notes(xI3).ColumnIndex = niBPM Then
        '        If Notes(xI3).Value > 655359999 Then Notes(xI3).Value = 655359999
        '        If Notes(xI3).Value < 1 Then Notes(xI3).Value = 1
        '    End If
        'Next

        'Restore selection
        'If vSelLength < 0 Then vSelStart += (xRatio - 1) * (xVUpper - xVLower)
        'vSelLength = vSelLength * xRatio
        Dim pSelHalf As Double = vSelHalf
        vSelHalf += dVPosition
        ValidateSelection()
        Me.RedoChangeTimeSelection(vSelStart, vSelLength, pSelHalf, vSelStart, vSelStart, vSelHalf, True, xUndo, xRedo)

        Me.RedoAddNoteAll(False, xUndo, xRedo)


        'Restore note selection
        xVLower = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        xVUpper = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If Not NTInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And Notes(xI3).VPosition < xVUpper And nEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And nEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

EndofSub:
        If bAddUndo Then AddUndo(xUndo, xBaseRedo.Next, bOverWriteUndo)
    End Sub

    Private Sub BPMChangeByValue(ByVal xValue As Integer)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If vSelLength = 0 Then Return

        Dim xVLower As Double = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        Dim xVHalf As Double = vSelStart + vSelHalf
        Dim xVUpper As Double = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If xVHalf = xVUpper Then xVHalf = xVLower
        'If dVPosition + xVHalf <= xVLower Or dVPosition + xVHalf >= xVUpper Then GoTo EndofSub

        If xVLower < 0 Then xVLower = 0
        If xVUpper >= GetMaxVPosition() Then xVUpper = GetMaxVPosition() - 1
        If xVHalf > xVUpper Then xVHalf = xVUpper
        If xVHalf < xVLower Then xVHalf = xVLower

        Dim xBPM = Notes(0).Value
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        Dim xConstBPM As Double = 0

        'Below Selection
        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).VPosition > xVLower Then Exit For
            If Notes(xI1).ColumnIndex = niBPM Then xBPM = Notes(xI1).Value
        Next
        xI2 = xI1
        Dim xVPos() As Double = {xVLower}
        Dim xVal() = {xBPM}

        'Within Selection
        Dim xU As Integer = 0
        For xI1 = xI2 To UBound(Notes)
            If Notes(xI1).VPosition > xVUpper Then Exit For

            If Notes(xI1).ColumnIndex = niBPM Then
                xU = UBound(xVPos) + 1
                ReDim Preserve xVPos(xU)
                ReDim Preserve xVal(xU)
                xVPos(xU) = Notes(xI1).VPosition
                xVal(xU) = Notes(xI1).Value
            End If
        Next
        ReDim Preserve xVPos(xU + 1)
        xVPos(xU + 1) = xVUpper

        'Calculate Constant BPM
        For xI1 = 0 To xU
            xConstBPM += (xVPos(xI1 + 1) - xVPos(xI1)) / xVal(xI1)
        Next
        xConstBPM = (xVUpper - xVLower) / xConstBPM

        'Compare BPM        '(xVHalf - xVLower) / xValue + (xVUpper - xVHalf) / xResult = (xVUpper - xVLower) / xConstBPM
        If (xVUpper - xVLower) / xConstBPM <= (xVHalf - xVLower) / xValue Then
            Dim Limit = ((xVHalf - xVLower) * xConstBPM / (xVUpper - xVLower) / 10000)
            MsgBox("Please enter a value that is greater than " & Limit & ".", MsgBoxStyle.Critical, Strings.Messages.Err)
            Return
        End If
        Dim xTempDivider As Double = xConstBPM * (xVHalf - xVLower) - xValue * (xVUpper - xVLower)

        ' az: I want to allow negative values, maybe...
        If xTempDivider = 0 Then
            Return ' nullop this
        End If

        ' apply div. by 10k to nullify mult. caused by divider being divided by 10k
        Dim xResult = (xVHalf - xVUpper) * xValue / xTempDivider * xConstBPM ' order here is important to avoid an overflow

        RedoRemoveNoteAll(False, xUndo, xRedo)

        'Adjust note
        If Not NTInput Then
            'Below Selection
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition > xVLower Then Exit For
            Next
            xI2 = xI1
            If xI2 > UBound(Notes) Then GoTo EndOfAdjustment

            'Within Selection
            Dim xTempTime As Double
            Dim xTempVPos As Double
            For xI1 = xI2 To UBound(Notes)
                If Notes(xI1).VPosition >= xVUpper Then Exit For
                xTempTime = 0

                xTempVPos = Notes(xI1).VPosition
                For xI3 = 0 To xU
                    If xTempVPos < xVPos(xI3 + 1) Then Exit For
                    xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                Next
                xTempTime += (xTempVPos - xVPos(xI3)) / xVal(xI3)

                If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                    Notes(xI1).VPosition = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf
                Else
                    Notes(xI1).VPosition = xTempTime * xValue + xVLower
                End If
            Next

        Else
            Dim xTempTime As Double
            Dim xTempVPos As Double
            Dim xTempEnd As Double

            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).Length Then xTempEnd = Notes(xI1).VPosition + Notes(xI1).Length

                If Notes(xI1).VPosition > xVLower And Notes(xI1).VPosition < xVUpper Then
                    xTempTime = 0

                    xTempVPos = Notes(xI1).VPosition
                    For xI3 = 0 To xU
                        If xTempVPos < xVPos(xI3 + 1) Then Exit For
                        xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                    Next
                    xTempTime += (xTempVPos - xVPos(xI3)) / xVal(xI3)

                    If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                        Notes(xI1).VPosition = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf
                    Else
                        Notes(xI1).VPosition = xTempTime * xValue + xVLower
                    End If
                End If

                If Notes(xI1).Length Then
                    If xTempEnd > xVLower And xTempEnd < xVUpper Then
                        xTempTime = 0

                        For xI3 = 0 To xU
                            If xTempEnd < xVPos(xI3 + 1) Then Exit For
                            xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                        Next
                        xTempTime += (xTempEnd - xVPos(xI3)) / xVal(xI3)

                        If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                            Notes(xI1).Length = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf - Notes(xI1).VPosition
                        Else
                            Notes(xI1).Length = xTempTime * xValue + xVLower - Notes(xI1).VPosition
                        End If

                    Else
                        Notes(xI1).Length = xTempEnd - Notes(xI1).VPosition
                    End If
                End If

            Next
        End If

EndOfAdjustment:

        'Delete BPMs
        xI1 = 1
        Do While xI1 <= UBound(Notes)
            If Notes(xI1).VPosition > xVUpper Then Exit Do
            If Notes(xI1).VPosition >= xVLower And Notes(xI1).ColumnIndex = niBPM Then
                For xI3 = xI1 + 1 To UBound(Notes)
                    Notes(xI3 - 1) = Notes(xI3)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
            Else
                xI1 += 1
            End If
        Loop

        'Add BPMs
        ReDim Preserve Notes(UBound(Notes) + 2)
        With Notes(UBound(Notes) - 1)
            .ColumnIndex = niBPM
            .VPosition = xVHalf
            .Value = xResult
        End With
        With Notes(UBound(Notes))
            .ColumnIndex = niBPM
            .VPosition = xVUpper
            .Value = xVal(xU)
        End With
        If xVLower <> xVHalf Then
            ReDim Preserve Notes(UBound(Notes) + 1)
            With Notes(UBound(Notes))
                .ColumnIndex = niBPM
                .VPosition = xVLower
                .Value = xValue
            End With
        End If

        'Save redo
        'For xI3 = 1 To UBound(K)
        '    K(xI3).Selected = True
        'Next
        'xRedo = "KZ" & vbCrLf & _
        '                      sCmdKs(False) & vbCrLf & _
        '                      "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"
        Me.RedoAddNoteAll(False, xUndo, xRedo)

        'Restore note selection
        xVLower = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        xVUpper = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If Not NTInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And Notes(xI3).VPosition < xVUpper And nEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And nEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

        'EndofSub:
        AddUndo(xUndo, xBaseRedo.Next)
    End Sub

    Private Sub ConvertAreaToStop()
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If vSelLength = 0 Then Return

        Dim xVLower As Double = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        Dim xVUpper As Double = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)

        Dim notesInRange = From note In Notes
                           Where note.VPosition > xVLower And note.VPosition <= xVUpper
                           Select note

        If notesInRange.Count() > 0 Then
            MessageBox.Show("The selected area can't have notes anywhere but at the start.")
            Return
        End If

        RedoRemoveNoteAll(False, xUndo, xRedo)

        ' Translate notes
        For I = 1 To UBound(Notes)
            If Notes(I).VPosition > xVUpper Then
                Notes(I).VPosition -= vSelLength
            End If
        Next

        ' Add Stop
        ReDim Preserve Notes(UBound(Notes) + 1)
        With Notes(UBound(Notes))
            .ColumnIndex = niSTOP
            .VPosition = xVLower
            .Value = vSelLength * 10000
        End With

        Me.RedoAddNoteAll(False, xUndo, xRedo)

        AddUndo(xUndo, xBaseRedo.Next)

    End Sub

    Private Sub BConvertStop_Click(sender As Object, e As EventArgs) Handles BConvertStop.Click
        SortByVPositionInsertion()
        ConvertAreaToStop()

        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
        POStatusRefresh()

        Beep()
        TVCBPM.Focus()
    End Sub

End Class
