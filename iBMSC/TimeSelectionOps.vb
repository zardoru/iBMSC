Imports System.Linq
Imports iBMSC.Editor

Partial Public Class MainWindow
    Private Sub BVCCalculate_Click(sender As Object, e As EventArgs) Handles BVCCalculate.Click
        If Not TBTimeSelect.Checked Then Exit Sub

        ValidateNotesArray()
        BPMChangeByValue(Val(TVCBPM.Text)*10000)
        ValidateNotesArray() ' XXX: what???

        RefreshPanelAll()
        POStatusRefresh()

        Beep()
        TVCBPM.Focus()
    End Sub

    Private Sub BVCApply_Click(sender As Object, e As EventArgs) Handles BVCApply.Click
        If Not TBTimeSelect.Checked Then Exit Sub

        ValidateNotesArray()
        BPMChangeTop(Val(TVCM.Text)/Val(TVCD.Text))

        RefreshPanelAll()
        POStatusRefresh()


        Beep()
        TVCM.Focus()
    End Sub

    Public Sub BPMChangeTop(xRatio As Double, Optional ByVal bAddUndo As Boolean = True,
                            Optional ByVal bOverWriteUndo As Boolean = False)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If State.TimeSelect.EndPointLength = 0 Then GoTo EndofSub
        If xRatio = 1 Or xRatio <= 0 Then GoTo EndofSub

        Dim xVLower As Double = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.StartPoint,
                                    State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        Dim xVUpper As Double = IIf(State.TimeSelect.EndPointLength < 0, State.TimeSelect.StartPoint,
                                    State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        If xVLower < 0 Then xVLower = 0
        If xVUpper >= GetMaxVPosition() Then xVUpper = GetMaxVPosition() - 1

        Dim xBPM As Integer = Notes(0).Value
        Dim i As Integer
        Dim j As Integer
        Dim xI3 As Integer

        Dim xValueL As Integer = xBPM
        Dim xValueU As Integer = xBPM

        RedoRemoveNoteAll(Notes, xUndo, xRedo)

        'Start
        If Not NTInput Then
            'Below Selection
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition > xVLower Then Exit For
                If Notes(i).ColumnIndex = ColumnType.BPM Then xBPM = Notes(i).Value
            Next
            xValueL = xBPM
            j = i

            'Within Selection
            For i = j To UBound(Notes)
                If Notes(i).VPosition > xVUpper Then Exit For
                If Notes(i).ColumnIndex = ColumnType.BPM Then
                    xBPM = Notes(i).Value
                    Notes(i).Value = Notes(i).Value*xRatio
                End If
                Notes(i).VPosition = (Notes(i).VPosition - xVLower)*xRatio + xVLower
            Next
            xValueU = xBPM
            j = i

            'Above Selection
            For i = j To UBound(Notes)
                Notes(i).VPosition += (xRatio - 1)*(xVUpper - xVLower)
            Next

            'Add BPMs
            AddNote(New Note(ColumnType.BPM, xVLower, xValueL*xRatio), False, True, False)
            AddNote(New Note(ColumnType.BPM, xVUpper + (xRatio - 1)*(xVUpper - xVLower), xValueU), False, True, False)

        Else
            Dim xAddBPML = True
            Dim xAddBPMU = True

            For i = 1 To UBound(Notes)
                'Modify notes
                If Notes(i).VPosition <= xVLower Then
                    'check BPM
                    If Notes(i).ColumnIndex = ColumnType.BPM Then
                        xValueL = Notes(i).Value
                        xValueU = Notes(i).Value
                        If Notes(i).VPosition = xVLower Then _
                            xAddBPML = False : _
                                Notes(i).Value = IIf(Notes(i).Value*xRatio <= 655359999, Notes(i).Value*xRatio,
                                                     655359999)
                    End If

                    'If longnote then adjust length
                    If Notes(i).VPosition + Notes(i).Length > xVLower Then
                        Notes(i).Length +=
                            (IIf(xVUpper < Notes(i).VPosition + Notes(i).Length, xVUpper,
                                 Notes(i).VPosition + Notes(i).Length) - xVLower)*(xRatio - 1)
                    End If

                ElseIf Notes(i).VPosition <= xVUpper Then
                    'check BPM
                    If Notes(i).ColumnIndex = ColumnType.BPM Then
                        xValueU = Notes(i).Value
                        If Notes(i).VPosition = xVUpper Then xAddBPMU = False Else _
                            Notes(i).Value = IIf(Notes(i).Value*xRatio <= 655359999, Notes(i).Value*xRatio, 655359999)
                    End If

                    'Adjust Length
                    Notes(i).Length +=
                        (IIf(xVUpper < Notes(i).Length + Notes(i).VPosition, xVUpper,
                             Notes(i).Length + Notes(i).VPosition) - Notes(i).VPosition)*(xRatio - 1)

                    'Adjust VPosition
                    Notes(i).VPosition = (Notes(i).VPosition - xVLower)*xRatio + xVLower

                Else
                    Notes(i).VPosition += (xVUpper - xVLower)*(xRatio - 1)
                End If
            Next

            'Add BPMs
            If xAddBPML Then AddNote(New Note(ColumnType.BPM, xVLower, xValueL*xRatio), False, True, False)
            If xAddBPMU Then _
                AddNote(New Note(ColumnType.BPM, (xVUpper - xVLower)*xRatio + xVLower, xValueU), False, True, False)
        End If

        'Check BPM Overflow
        For xI3 = 1 To UBound(Notes)
            If Notes(xI3).ColumnIndex = ColumnType.BPM AndAlso Notes(xI3).Value < 1 Then Notes(xI3).Value = 1
        Next

        RedoAddNoteAll(Notes, xUndo, xRedo)

        'Restore selection
        Dim pSelStart As Double = State.TimeSelect.StartPoint
        Dim pSelLength As Double = State.TimeSelect.EndPointLength
        Dim pSelHalf As Double = State.TimeSelect.HalfPointLength
        If State.TimeSelect.EndPointLength < 0 Then State.TimeSelect.StartPoint += (xRatio - 1)*(xVUpper - xVLower)
        State.TimeSelect.EndPointLength = State.TimeSelect.EndPointLength*xRatio
        State.TimeSelect.HalfPointLength = State.TimeSelect.HalfPointLength*xRatio
        State.TimeSelect.ValidateSelection(GetMaxVPosition())

        RedoChangeTimeSelection(pSelStart, pSelLength, pSelHalf, State.TimeSelect.StartPoint,
                                State.TimeSelect.EndPointLength, State.TimeSelect.HalfPointLength, True, xUndo, xRedo)


        'Restore note selection
        xVLower = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.StartPoint,
                      State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        xVUpper = IIf(State.TimeSelect.EndPointLength < 0, State.TimeSelect.StartPoint,
                      State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        If Not NTInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And
                                      Notes(xI3).VPosition < xVUpper And
                                      Columns.IsEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And
                                      Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And
                                      Columns.IsEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

        EndofSub:
        If bAddUndo Then AddUndoChain(xUndo, xBaseRedo.Next, bOverWriteUndo)
        ValidateNotesArray()
        UpdatePairing()
    End Sub

    Public Sub BPMChangeHalf(dVPosition As Double, Optional ByVal bAddUndo As Boolean = True,
                             Optional ByVal bOverWriteUndo As Boolean = False)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If State.TimeSelect.EndPointLength = 0 Then GoTo EndofSub
        If dVPosition = 0 Then GoTo EndofSub

        Dim xVLower As Double = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.StartPoint,
                                    State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        Dim xVHalf As Double = State.TimeSelect.StartPoint + State.TimeSelect.HalfPointLength
        Dim xVUpper As Double = IIf(State.TimeSelect.EndPointLength < 0, State.TimeSelect.StartPoint,
                                    State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        If dVPosition + xVHalf <= xVLower Or dVPosition + xVHalf >= xVUpper Then GoTo EndofSub

        If xVLower < 0 Then xVLower = 0
        If xVUpper >= GetMaxVPosition() Then xVUpper = GetMaxVPosition() - 1
        If xVHalf > xVUpper Then xVHalf = xVUpper
        If xVHalf < xVLower Then xVHalf = xVLower

        Dim xBPM As Integer = Notes(0).Value
        Dim i As Integer
        Dim j As Integer
        Dim xI3 As Integer

        Dim xValueL As Integer = xBPM
        Dim xValueM As Integer = xBPM
        Dim xValueU As Integer = xBPM

        Dim xRatio1 As Double = (xVHalf - xVLower + dVPosition)/(xVHalf - xVLower)
        Dim xRatio2 As Double = (xVUpper - xVHalf - dVPosition)/(xVUpper - xVHalf)

        'Save undo
        'For xI3 = 1 To UBound(K)
        'K(xI3).Selected = True
        'Next
        'xUndo = "KZ" & vbCrLf & _
        '        sCmdKs(False) & vbCrLf & _
        '        "SA_" & State.TimeSelect.Start & "_" & State.TimeSelect.Length & "_" & State.TimeSelect.Half & "_1"

        RedoRemoveNoteAll(Notes, xUndo, xRedo)

        If Not NTInput Then
            'Below Selection
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition > xVLower Then Exit For
                If Notes(i).ColumnIndex = ColumnType.BPM Then xBPM = Notes(i).Value
            Next
            xValueL = xBPM
            j = i

            'Below Half
            For i = j To UBound(Notes)
                If Notes(i).VPosition > xVHalf Then Exit For
                If Notes(i).ColumnIndex = ColumnType.BPM Then
                    xBPM = Notes(i).Value
                    Notes(i).Value = Notes(i).Value*xRatio1
                End If
                Notes(i).VPosition = (Notes(i).VPosition - xVLower)*xRatio1 + xVLower
            Next
            xValueM = xBPM
            j = i

            'Above Half
            For i = j To UBound(Notes)
                If Notes(i).VPosition > xVUpper Then Exit For
                If Notes(i).ColumnIndex = ColumnType.BPM Then
                    xBPM = Notes(i).Value
                    Notes(i).Value = IIf(Notes(i).Value*xRatio2 <= 655359999, Notes(i).Value*xRatio2, 655359999)
                End If
                Notes(i).VPosition = (Notes(i).VPosition - xVHalf)*xRatio2 + xVHalf + dVPosition
            Next
            xValueU = xBPM
            j = i

            'Above Selection
            'For i = j To UBound(K)
            '    K(i).VPosition += (xRatio - 1) * (xVUpper - xVLower)
            'Next

            'Add BPMs
            ' az: cond. removed; 
            ' IIf(xVHalf <> xVLower AndAlso xValueL * xRatio1 <= 655359999, xValueL * xRatio1, 655359999)
            AddNote(New Note(ColumnType.BPM, xVLower, xValueL*xRatio1), False, True, False)
            ' az: cond removed;
            ' IIf(xVHalf <> xVUpper AndAlso xValueM * xRatio2 <= 655359999, xValueM * xRatio2, 655359999)
            AddNote(New Note(ColumnType.BPM, xVHalf + dVPosition, xValueM*xRatio2), False, True, False)
            AddNote(New Note(ColumnType.BPM, xVUpper, xValueU), False, True, False)

        Else
            Dim xAddBPML = True
            Dim xAddBPMM = True
            Dim xAddBPMU = True

            'Modify notes
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition <= xVLower Then
                    'check BPM
                    If Notes(i).ColumnIndex = ColumnType.BPM Then
                        xValueL = Notes(i).Value
                        xValueM = Notes(i).Value
                        xValueU = Notes(i).Value
                        If Notes(i).VPosition = xVLower Then
                            xAddBPML = False

                            ' az: condition removed;
                            ' IIf(xVHalf <> xVLower AndAlso Notes(i).Value * xRatio1 <= 655359999, Notes(i).Value * xRatio1, 655359999)
                            Notes(i).Value = Notes(i).Value*xRatio1
                        End If
                    End If

                    'If longnote then adjust length
                    Dim xEnd As Double = Notes(i).VPosition + Notes(i).Length
                    If xEnd > xVUpper Then
                    ElseIf xEnd > xVHalf Then
                        Notes(i).Length = (xEnd - xVHalf)*xRatio2 + xVHalf + dVPosition - Notes(i).VPosition
                    ElseIf xEnd > xVLower Then
                        Notes(i).Length = (xEnd - xVLower)*xRatio1 + xVLower - Notes(i).VPosition
                    End If

                ElseIf Notes(i).VPosition <= xVHalf Then
                    'check BPM
                    If Notes(i).ColumnIndex = ColumnType.BPM Then
                        xValueM = Notes(i).Value
                        xValueU = Notes(i).Value
                        If Notes(i).VPosition = xVHalf Then
                            xAddBPMM = False
                            ' az: cond. remove
                            ' IIf(xVHalf <> xVUpper AndAlso Notes(i).Value * xRatio2 <= 655359999, Notes(i).Value * xRatio2, 655359999)
                            Notes(i).Value = Notes(i).Value*xRatio2
                        Else
                            ' az: cond. remove
                            ' IIf(Notes(i).Value * xRatio1 <= 655359999, Notes(i).Value * xRatio1, 655359999)
                            Notes(i).Value = Notes(i).Value*xRatio1
                        End If
                    End If

                    'Adjust Length
                    Dim xEnd As Double = Notes(i).VPosition + Notes(i).Length
                    If xEnd > xVUpper Then
                        Notes(i).Length = xEnd - xVLower - (Notes(i).VPosition - xVLower)*xRatio1
                    ElseIf xEnd > xVHalf Then
                        Notes(i).Length = (xVHalf - Notes(i).VPosition)*xRatio1 + (xEnd - xVHalf)*xRatio2
                    Else
                        Notes(i).Length *= xRatio1
                    End If

                    'Adjust VPosition
                    Notes(i).VPosition = (Notes(i).VPosition - xVLower)*xRatio1 + xVLower

                ElseIf Notes(i).VPosition <= xVUpper Then
                    'check BPM
                    If Notes(i).ColumnIndex = ColumnType.BPM Then
                        xValueU = Notes(i).Value
                        If Notes(i).VPosition = xVUpper Then xAddBPMU = False Else _
                            Notes(i).Value = IIf(Notes(i).Value*xRatio2 <= 655359999, Notes(i).Value*xRatio2, 655359999)
                    End If

                    'Adjust Length
                    Dim xEnd As Double = Notes(i).VPosition + Notes(i).Length
                    If xEnd > xVUpper Then
                        Notes(i).Length = (xVUpper - Notes(i).VPosition)*xRatio2 + xEnd - xVUpper
                    Else
                        Notes(i).Length *= xRatio2
                    End If

                    'Adjust VPosition
                    Notes(i).VPosition = (Notes(i).VPosition - xVHalf)*xRatio2 + xVHalf + dVPosition

                    'Else
                    '    K(i).VPosition += (xVUpper - xVLower) * (xRatio - 1)
                End If
            Next

            'Add BPMs
            ' IIf(xVHalf <> xVLower AndAlso xValueL * xRatio1 <= 655359999, xValueL * xRatio1, 655359999)
            If xAddBPML Then AddNote(New Note(ColumnType.BPM, xVLower, xValueL*xRatio1), False, True, False)
            ' IIf(xVHalf <> xVUpper AndAlso xValueM * xRatio2 <= 655359999, xValueM * xRatio2, 655359999)
            If xAddBPMM Then AddNote(New Note(ColumnType.BPM, xVHalf + dVPosition, xValueM*xRatio2), False, True, False)
            If xAddBPMU Then AddNote(New Note(ColumnType.BPM, xVUpper, xValueU), False, True, False)
        End If

        'Check BPM Overflow
        'For xI3 = 1 To UBound(Notes)
        '    If Notes(xI3).ColumnIndex = Columns.BGMPM Then
        '        If Notes(xI3).Value > 655359999 Then Notes(xI3).Value = 655359999
        '        If Notes(xI3).Value < 1 Then Notes(xI3).Value = 1
        '    End If
        'Next

        'Restore selection
        'If State.TimeSelect.Length < 0 Then State.TimeSelect.Start += (xRatio - 1) * (xVUpper - xVLower)
        'State.TimeSelect.Length = State.TimeSelect.Length * xRatio
        Dim pSelHalf As Double = State.TimeSelect.HalfPointLength
        State.TimeSelect.HalfPointLength += dVPosition
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RedoChangeTimeSelection(State.TimeSelect.StartPoint,
                                State.TimeSelect.EndPointLength,
                                pSelHalf,
                                State.TimeSelect.StartPoint,
                                State.TimeSelect.StartPoint,
                                State.TimeSelect.HalfPointLength,
                                True, xUndo, xRedo)

        RedoAddNoteAll(Notes, xUndo, xRedo)


        'Restore note selection
        xVLower = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.StartPoint,
                      State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        xVUpper = IIf(State.TimeSelect.EndPointLength < 0, State.TimeSelect.StartPoint,
                      State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        If Not NTInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And
                                      Notes(xI3).VPosition < xVUpper And
                                      Columns.IsEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And
                                      Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And
                                      Columns.IsEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

        EndofSub:
        If bAddUndo Then AddUndoChain(xUndo, xBaseRedo.Next, bOverWriteUndo)

        ValidateNotesArray()
        UpdatePairing()
    End Sub

    Private Sub BPMChangeByValue(xValue As Integer)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If State.TimeSelect.EndPointLength = 0 Then Return

        Dim xVLower As Double = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.StartPoint,
                                    State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        Dim xVHalf As Double = State.TimeSelect.StartPoint + State.TimeSelect.HalfPointLength
        Dim xVUpper As Double = IIf(State.TimeSelect.EndPointLength < 0, State.TimeSelect.StartPoint,
                                    State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        If xVHalf = xVUpper Then xVHalf = xVLower
        'If dVPosition + xVHalf <= xVLower Or dVPosition + xVHalf >= xVUpper Then GoTo EndofSub

        If xVLower < 0 Then xVLower = 0
        If xVUpper >= GetMaxVPosition() Then xVUpper = GetMaxVPosition() - 1
        If xVHalf > xVUpper Then xVHalf = xVUpper
        If xVHalf < xVLower Then xVHalf = xVLower

        Dim xBPM = Notes(0).Value
        Dim i As Integer
        Dim j As Integer
        Dim xI3 As Integer

        Dim xConstBPM As Double = 0

        'Below Selection
        For i = 1 To UBound(Notes)
            If Notes(i).VPosition > xVLower Then Exit For
            If Notes(i).ColumnIndex = ColumnType.BPM Then xBPM = Notes(i).Value
        Next
        j = i
        Dim xVPos() As Double = {xVLower}
        Dim xVal() = {xBPM}

        'Within Selection
        Dim xU = 0
        For i = j To UBound(Notes)
            If Notes(i).VPosition > xVUpper Then Exit For

            If Notes(i).ColumnIndex = ColumnType.BPM Then
                xU = UBound(xVPos) + 1
                ReDim Preserve xVPos(xU)
                ReDim Preserve xVal(xU)
                xVPos(xU) = Notes(i).VPosition
                xVal(xU) = Notes(i).Value
            End If
        Next
        ReDim Preserve xVPos(xU + 1)
        xVPos(xU + 1) = xVUpper

        'Calculate Constant BPM
        For i = 0 To xU
            xConstBPM += (xVPos(i + 1) - xVPos(i)) / xVal(i)
        Next
        xConstBPM = (xVUpper - xVLower) / xConstBPM

        'Compare BPM        '(xVHalf - xVLower) / xValue + (xVUpper - xVHalf) / xResult = (xVUpper - xVLower) / xConstBPM
        If (xVUpper - xVLower) / xConstBPM <= (xVHalf - xVLower) / xValue Then
            Dim Limit = ((xVHalf - xVLower) * xConstBPM / (xVUpper - xVLower) / 10000)
            MsgBox("Please enter a value that is greater than " & Limit & ".", MsgBoxStyle.Critical,
                   Strings.Messages.Err)
            Return
        End If
        Dim xTempDivider As Double = xConstBPM * (xVHalf - xVLower) - xValue * (xVUpper - xVLower)

        ' az: I want to allow negative values, maybe...
        If xTempDivider = 0 Then
            Return ' nullop this
        End If

        ' apply div. by 10k to nullify mult. caused by divider being divided by 10k
        Dim xResult = (xVHalf - xVUpper) * xValue / xTempDivider * xConstBPM ' order here is important to avoid an overflow

        RedoRemoveNoteAll(Notes, xUndo, xRedo)

        'Adjust note
        If Not NtInput Then
            'Below Selection
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition > xVLower Then Exit For
            Next
            j = i
            If j > UBound(Notes) Then GoTo EndOfAdjustment

            'Within Selection
            Dim xTempTime As Double
            Dim xTempVPos As Double
            For i = j To UBound(Notes)
                If Notes(i).VPosition >= xVUpper Then Exit For
                xTempTime = 0

                xTempVPos = Notes(i).VPosition
                For xI3 = 0 To xU
                    If xTempVPos < xVPos(xI3 + 1) Then Exit For
                    xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                Next
                xTempTime += (xTempVPos - xVPos(xI3)) / xVal(xI3)

                If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                    Notes(i).VPosition = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf
                Else
                    Notes(i).VPosition = xTempTime * xValue + xVLower
                End If
            Next

        Else
            Dim xTempTime As Double
            Dim xTempVPos As Double
            Dim xTempEnd As Double

            For i = 1 To UBound(Notes)
                If Notes(i).Length Then xTempEnd = Notes(i).VPosition + Notes(i).Length

                If Notes(i).VPosition > xVLower And Notes(i).VPosition < xVUpper Then
                    xTempTime = 0

                    xTempVPos = Notes(i).VPosition
                    For xI3 = 0 To xU
                        If xTempVPos < xVPos(xI3 + 1) Then Exit For
                        xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                    Next
                    xTempTime += (xTempVPos - xVPos(xI3)) / xVal(xI3)

                    If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                        Notes(i).VPosition = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf
                    Else
                        Notes(i).VPosition = xTempTime * xValue + xVLower
                    End If
                End If

                If Notes(i).Length Then
                    If xTempEnd > xVLower And xTempEnd < xVUpper Then
                        xTempTime = 0

                        For xI3 = 0 To xU
                            If xTempEnd < xVPos(xI3 + 1) Then Exit For
                            xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                        Next
                        xTempTime += (xTempEnd - xVPos(xI3)) / xVal(xI3)

                        If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                            Notes(i).Length = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf -
                                              Notes(i).VPosition
                        Else
                            Notes(i).Length = xTempTime * xValue + xVLower - Notes(i).VPosition
                        End If

                    Else
                        Notes(i).Length = xTempEnd - Notes(i).VPosition
                    End If
                End If

            Next
        End If

EndOfAdjustment:

        'Delete BPMs
        i = 1
        Do While i <= UBound(Notes)
            If Notes(i).VPosition > xVUpper Then Exit Do
            If Notes(i).VPosition >= xVLower And Notes(i).ColumnIndex = ColumnType.BPM Then
                For xI3 = i + 1 To UBound(Notes)
                    Notes(xI3 - 1) = Notes(xI3)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
            Else
                i += 1
            End If
        Loop

        'Add BPMs
        ReDim Preserve Notes(UBound(Notes) + 2)
        With Notes(UBound(Notes) - 1)
            .ColumnIndex = ColumnType.BPM
            .VPosition = xVHalf
            .Value = xResult
        End With
        With Notes(UBound(Notes))
            .ColumnIndex = ColumnType.BPM
            .VPosition = xVUpper
            .Value = xVal(xU)
        End With
        If xVLower <> xVHalf Then
            ReDim Preserve Notes(UBound(Notes) + 1)
            With Notes(UBound(Notes))
                .ColumnIndex = ColumnType.BPM
                .VPosition = xVLower
                .Value = xValue
            End With
        End If

        RedoAddNoteAll(Notes, xUndo, xRedo)

        'Restore note selection
        xVLower = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.StartPoint,
                      State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        xVUpper = IIf(State.TimeSelect.EndPointLength < 0, State.TimeSelect.StartPoint,
                      State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        If Not NtInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And
                                      Notes(xI3).VPosition < xVUpper And
                                      Columns.IsEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And
                                      Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And
                                      Columns.IsEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

        'EndofSub:
        AddUndoChain(xUndo, xBaseRedo.Next)
    End Sub

    Private Sub ConvertAreaToStop()
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If State.TimeSelect.EndPointLength = 0 Then Return

        Dim xVLower As Double = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.StartPoint,
                                    State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)
        Dim xVUpper As Double = IIf(State.TimeSelect.EndPointLength < 0, State.TimeSelect.StartPoint,
                                    State.TimeSelect.StartPoint + State.TimeSelect.EndPointLength)

        Dim notesInRange = From note In Notes
                           Where note.VPosition > xVLower And note.VPosition <= xVUpper
                           Select note

        If notesInRange.Count() > 0 Then
            MessageBox.Show("The selected area can't have notes anywhere but at the start.")
            Return
        End If

        RedoRemoveNoteAll(Notes, xUndo, xRedo)

        ' Translate notes
        For I = 1 To UBound(Notes)
            If Notes(I).VPosition > xVUpper Then
                Notes(I).VPosition -= State.TimeSelect.EndPointLength
            End If
        Next

        ' Add Stop
        Dim stopNote = New Note With {
                .ColumnIndex = ColumnType.STOPS,
                .VPosition = xVLower,
                .Value = State.TimeSelect.EndPointLength * 10000
                }
        Notes = Notes.Concat({stopNote}).ToArray()

        RedoAddNoteAll(Notes, xUndo, xRedo)
        AddUndoChain(xUndo, xBaseRedo.Next)
    End Sub

    Private Sub BConvertStop_Click(sender As Object, e As EventArgs) Handles BConvertStop.Click
        ValidateNotesArray()
        ConvertAreaToStop()
        ValidateNotesArray() ' XXX: Hmmm

        RefreshPanelAll()
        POStatusRefresh()

        Beep()
        TVCBPM.Focus()
    End Sub
End Class
