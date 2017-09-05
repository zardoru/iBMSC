Partial Public Class MainWindow

    Public Sub MyO2ConstBPM(ByVal vBPM As Integer)
        SortByVPositionInsertion()
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        'Save undo
        'For xI3 = 1 To UBound(K)
        ' K(xI3).Selected = True
        ' Next
        ' xUndo = "KZ" & vbCrLf & _
        '         sCmdKs(False) & vbCrLf & _
        '         "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_0"
        Me.RedoRemoveNoteAll(False, xUndo, xRedo)

        'Adjust note
        Dim xcTime As Double = 0
        Dim xcVPos As Double = 0
        Dim xcBPM As Integer = Notes(0).Value

        If Not NTInput Then
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).ColumnIndex = niBPM Then
                    xcTime += (Notes(xI1).VPosition - xcVPos) / xcBPM
                    xcVPos = Notes(xI1).VPosition
                    xcBPM = Notes(xI1).Value
                Else
                    Notes(xI1).VPosition = vBPM * (xcTime + (Notes(xI1).VPosition - xcVPos) / xcBPM)
                End If
            Next

        Else
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).ColumnIndex = niBPM Then
                    xcTime += (Notes(xI1).VPosition - xcVPos) / xcBPM
                    xcVPos = Notes(xI1).VPosition
                    xcBPM = Notes(xI1).Value
                ElseIf Notes(xI1).Length = 0 Then
                    Notes(xI1).VPosition = vBPM * (xcTime + (Notes(xI1).VPosition - xcVPos) / xcBPM)
                Else
                    Dim xNewTimeL As Double = (xcTime + (Notes(xI1).VPosition - xcVPos) / xcBPM)

                    'find bpms
                    Dim xcTime2 As Double = xcTime
                    Dim xcVPos2 As Double = xcVPos
                    Dim xcBPM2 As Integer = xcBPM
                    For xI2 = xI1 + 1 To UBound(Notes)
                        If Notes(xI2).VPosition >= Notes(xI1).VPosition + Notes(xI1).Length Then Exit For
                        If Notes(xI2).ColumnIndex = niBPM Then
                            xcTime2 += (Notes(xI2).VPosition - xcVPos2) / xcBPM2
                            xcVPos2 = Notes(xI2).VPosition
                            xcBPM2 = Notes(xI2).Value
                        End If
                    Next
                    Dim xNewTimeU As Double = (xcTime2 + (Notes(xI1).VPosition + Notes(xI1).Length - xcVPos2) / xcBPM2)

                    Notes(xI1).VPosition = vBPM * xNewTimeL
                    Notes(xI1).Length = vBPM * xNewTimeU - Notes(xI1).VPosition
                End If
            Next
        End If

        'Delete BPMs
        xI1 = 1
        Do While xI1 <= UBound(Notes)
            If Notes(xI1).ColumnIndex = niBPM Then
                For xI3 = xI1 + 1 To UBound(Notes)
                    Notes(xI3 - 1) = Notes(xI3)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
            Else
                xI1 += 1
            End If
        Loop

        'Save redo
        'For xI3 = 1 To UBound(K)
        '    K(xI3).Selected = True
        'Next
        'xRedo = "KZ" & vbCrLf & _
        '        sCmdKs(False) & vbCrLf & _
        '        "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"

        Me.RedoAddNoteAll(False, xUndo, xRedo)
        Me.RedoRelabelNote(Notes(0), vBPM, xUndo, xRedo)

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        Notes(0).Value = vBPM
        THBPM.Value = vBPM / 10000
        CalculateTotalPlayableNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Public Function MyO2GridCheck() As String()
        CalculateGreatestVPosition()
        SortByVPositionInsertion()
        Dim xResult(-1) As String
        Dim xResult2(-1) As String
        Dim Identifiers() As String = {"01", "03", "04", "06", "07", "08", "09",
                                       "16", "11", "12", "13", "14", "15", "18", "19",
                                       "26", "21", "22", "23", "24", "25", "28", "29",
                                       "36", "31", "32", "33", "34", "35", "38", "39",
                                       "46", "41", "42", "43", "44", "45", "48", "49",
                                       "56", "51", "52", "53", "54", "55", "58", "59",
                                       "66", "61", "62", "63", "64", "65", "68", "69",
                                       "76", "71", "72", "73", "74", "75", "78", "79",
                                       "86", "81", "82", "83", "84", "85", "88", "89"}

        Dim xLowerIndex As Integer = 1
        Dim xUpperIndex As Integer = 1

        If Not NTInput Then
            For xMeasure As Integer = 0 To MeasureAtDisplacement(GreatestVPosition)
                'Find Ks in the same measure
                Dim xI1 As Integer
                For xI1 = xLowerIndex To UBound(Notes)
                    If MeasureAtDisplacement(Notes(xI1).VPosition) > xMeasure Then Exit For
                Next
                xUpperIndex = xI1

                For Each xId As String In Identifiers
                    'collect vposition data
                    Dim xVPos(-1) As Double
                    For xI2 As Integer = xLowerIndex To xUpperIndex - 1
                        If GetBMSChannelBy(Notes(xI2)) = xId And Math.Abs(Notes(xI2).VPosition - MeasureAtDisplacement(Notes(xI2).VPosition) * 192) > 0 Then
                            ReDim Preserve xVPos(UBound(xVPos) + 1)
                            xVPos(UBound(xVPos)) = Notes(xI2).VPosition - xMeasure * 192 : If xVPos(UBound(xVPos)) < 0 Then xVPos(UBound(xVPos)) = 0
                        End If
                    Next

                    'find gcd
                    Dim xGCD As Double = 192
                    For xI2 As Integer = 0 To UBound(xVPos)
                        xGCD = GCD(xGCD, xVPos(xI2))
                    Next

                    'check if smaller than minGCD
                    If xGCD < 3 Then
                        'suggestion
                        Dim xAdj64 As Boolean
                        Dim xD48 As Integer = 0
                        Dim xD64 As Integer = 0
                        For xI2 As Integer = 0 To UBound(xVPos)
                            xD48 += Math.Abs(xVPos(xI2) - CInt(xVPos(xI2) / 4) * 4)
                            xD64 += Math.Abs(xVPos(xI2) - CInt(xVPos(xI2) / 3) * 3)
                        Next
                        xAdj64 = xD48 > xD64

                        'put result
                        ReDim Preserve xResult(UBound(xResult) + 1)
                        xResult(UBound(xResult)) = xMeasure & "_" &
                                                   BMSChannelToColumn(xId) & "_" &
                                                   nTitle(BMSChannelToColumn(xId)) & "_" &
                                                   CStr(CInt(192 / xGCD)) & "_" &
                                                   CInt(IsChannelLongNote(xId)) & "_" &
                                                   CInt(IsChannelHidden(xId)) & "_" &
                                                   CInt(xAdj64) & "_" &
                                                   xD64 & "_" &
                                                   xD48
                    End If
                Next

980:            xLowerIndex = xUpperIndex
990:        Next

        Else
            For xMeasure As Integer = 0 To MeasureAtDisplacement(GreatestVPosition)

                For Each xId As String In Identifiers
                    Dim xVPos(-1) As Double

                    'collect vposition data
                    For xI2 As Integer = 1 To UBound(Notes)
                        If MeasureAtDisplacement(Notes(xI2).VPosition) > xMeasure Then Exit For

                        If GetBMSChannelBy(Notes(xI2)) <> xId Then GoTo 1330
                        If IsChannelLongNote(xId) Xor CBool(Notes(xI2).Length) Then GoTo 1330

                        If MeasureAtDisplacement(Notes(xI2).VPosition) = xMeasure AndAlso Math.Abs(Notes(xI2).VPosition - MeasureAtDisplacement(Notes(xI2).VPosition) * 192) > 0 Then
                            ReDim Preserve xVPos(UBound(xVPos) + 1)
                            xVPos(UBound(xVPos)) = Notes(xI2).VPosition - xMeasure * 192 : If xVPos(UBound(xVPos)) < 0 Then xVPos(UBound(xVPos)) = 0
                        End If

                        If Not CBool(Notes(xI2).Length) Then GoTo 1330

                        If MeasureAtDisplacement(Notes(xI2).VPosition + Notes(xI2).Length) = xMeasure AndAlso Not Notes(xI2).VPosition + Notes(xI2).Length - xMeasure * 192 = 0 Then
                            ReDim Preserve xVPos(UBound(xVPos) + 1)
                            xVPos(UBound(xVPos)) = Notes(xI2).VPosition + Notes(xI2).Length - xMeasure * 192 : If xVPos(UBound(xVPos)) < 0 Then xVPos(UBound(xVPos)) = 0
                        End If
1330:               Next

                    'find gcd
                    Dim xGCD As Double = 192
                    For xI2 As Integer = 0 To UBound(xVPos)
                        xGCD = GCD(xGCD, xVPos(xI2))
                    Next

                    'check if smaller than minGCD
                    If xGCD < 3 Then
                        'suggestion
                        Dim xAdj64 As Boolean
                        Dim xD48 As Integer = 0
                        Dim xD64 As Integer = 0
                        For xI2 As Integer = 0 To UBound(xVPos)
                            xD48 += Math.Abs(xVPos(xI2) - CInt(xVPos(xI2) / 4) * 4)
                            xD64 += Math.Abs(xVPos(xI2) - CInt(xVPos(xI2) / 3) * 3)
                        Next
                        xAdj64 = xD48 > xD64

                        'put result
                        ReDim Preserve xResult(UBound(xResult) + 1)
                        xResult(UBound(xResult)) = xMeasure & "_" &
                                                   BMSChannelToColumn(xId) & "_" &
                                                   nTitle(BMSChannelToColumn(xId)) & "_" &
                                                   CStr(CInt(192 / xGCD)) & "_" &
                                                   CInt(IsChannelLongNote(xId)) & "_" &
                                                   CInt(IsChannelHidden(xId)) & "_" &
                                                   CInt(xAdj64) & "_" &
                                                   xD64 & "_" &
                                                   xD48
                    End If
                Next

1990:       Next
        End If

        Return xResult
    End Function

    Public Sub MyO2GridAdjust(ByVal xaj() As dgMyO2.Adj)
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        'Save undo
        'For xI3 As Integer = 1 To UBound(K)
        '    K(xI3).Selected = True
        'Next
        'xUndo = "KZ" & vbCrLf & _
        '        sCmdKs(False) & vbCrLf & _
        '        "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_0"
        Me.RedoRemoveNoteAll(False, xUndo, xRedo)

        'adjust
        If Not NTInput Then
            For Each xadj As dgMyO2.Adj In xaj
                For xI1 As Integer = 1 To UBound(Notes)
                    If MeasureAtDisplacement(Notes(xI1).VPosition) = xadj.Measure And
                       Notes(xI1).ColumnIndex = xadj.ColumnIndex And
                       Notes(xI1).LongNote = xadj.LongNote And
                       Notes(xI1).Hidden = xadj.Hidden Then
                        Notes(xI1).VPosition = CLng(Notes(xI1).VPosition / IIf(xadj.AdjTo64, 3, 4)) * IIf(xadj.AdjTo64, 3, 4)
                    End If
                Next
            Next

        Else
            For Each xadj As dgMyO2.Adj In xaj
                For xI1 As Integer = 1 To UBound(Notes)
                    If CBool(Notes(xI1).Length) Xor xadj.LongNote Then GoTo 1100

                    Dim xStart As Double = Notes(xI1).VPosition
                    Dim xEnd As Double = Notes(xI1).VPosition + Notes(xI1).Length
                    If MeasureAtDisplacement(Notes(xI1).VPosition) = xadj.Measure And
                       Notes(xI1).ColumnIndex = xadj.ColumnIndex And
                       Notes(xI1).Hidden = xadj.Hidden Then _
                        xStart = CLng(Notes(xI1).VPosition / IIf(xadj.AdjTo64, 3, 4)) * IIf(xadj.AdjTo64, 3, 4)

                    If Notes(xI1).Length > 0 AndAlso
                       MeasureAtDisplacement(Notes(xI1).VPosition + Notes(xI1).Length) = xadj.Measure And
                       Notes(xI1).ColumnIndex = xadj.ColumnIndex And
                       Notes(xI1).Hidden = xadj.Hidden Then _
                        xEnd = CLng((Notes(xI1).VPosition + Notes(xI1).Length) / IIf(xadj.AdjTo64, 3, 4)) * IIf(xadj.AdjTo64, 3, 4)

                    Notes(xI1).VPosition = xStart
                    If Notes(xI1).Length > 0 Then Notes(xI1).Length = xEnd - Notes(xI1).VPosition
1100:           Next
            Next
        End If

        'Save redo
        'For xI3 As Integer = 1 To UBound(K)
        ' K(xI3).Selected = True
        ' Next
        ' xRedo = "KZ" & vbCrLf & _
        '         sCmdKs(False) & vbCrLf & _
        '         "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"
        Me.RedoAddNoteAll(False, xUndo, xRedo)

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
        Beep()
    End Sub
End Class
