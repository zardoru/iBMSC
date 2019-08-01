Partial Public Class MainWindow
    Public Sub MyO2ConstBPM(vBPM As Integer)
        SortByVPositionInsertion()
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim i As Integer
        Dim j As Integer
        Dim xI3 As Integer

        RedoRemoveNoteAll(Notes, xUndo, xRedo)

        'Adjust note
        Dim xcTime As Double = 0
        Dim xcVPos As Double = 0
        Dim xcBPM As Integer = Notes(0).Value

        If Not NTInput Then
            For i = 1 To UBound(Notes)
                If Notes(i).ColumnIndex = ColumnType.BPM Then
                    xcTime += (Notes(i).VPosition - xcVPos)/xcBPM
                    xcVPos = Notes(i).VPosition
                    xcBPM = Notes(i).Value
                Else
                    Notes(i).VPosition = vBPM*(xcTime + (Notes(i).VPosition - xcVPos)/xcBPM)
                End If
            Next

        Else
            For i = 1 To UBound(Notes)
                If Notes(i).ColumnIndex = ColumnType.BPM Then
                    xcTime += (Notes(i).VPosition - xcVPos)/xcBPM
                    xcVPos = Notes(i).VPosition
                    xcBPM = Notes(i).Value
                ElseIf Notes(i).Length = 0 Then
                    Notes(i).VPosition = vBPM*(xcTime + (Notes(i).VPosition - xcVPos)/xcBPM)
                Else
                    Dim xNewTimeL As Double = (xcTime + (Notes(i).VPosition - xcVPos)/xcBPM)

                    'find bpms
                    Dim xcTime2 As Double = xcTime
                    Dim xcVPos2 As Double = xcVPos
                    Dim xcBPM2 As Integer = xcBPM
                    For j = i + 1 To UBound(Notes)
                        If Notes(j).VPosition >= Notes(i).VPosition + Notes(i).Length Then Exit For
                        If Notes(j).ColumnIndex = ColumnType.BPM Then
                            xcTime2 += (Notes(j).VPosition - xcVPos2)/xcBPM2
                            xcVPos2 = Notes(j).VPosition
                            xcBPM2 = Notes(j).Value
                        End If
                    Next
                    Dim xNewTimeU As Double = (xcTime2 + (Notes(i).VPosition + Notes(i).Length - xcVPos2)/xcBPM2)

                    Notes(i).VPosition = vBPM*xNewTimeL
                    Notes(i).Length = vBPM*xNewTimeU - Notes(i).VPosition
                End If
            Next
        End If

        'Delete BPMs
        i = 1
        Do While i <= UBound(Notes)
            If Notes(i).ColumnIndex = ColumnType.BPM Then
                For xI3 = i + 1 To UBound(Notes)
                    Notes(xI3 - 1) = Notes(xI3)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
            Else
                i += 1
            End If
        Loop


        RedoAddNoteAll(Notes, xUndo, xRedo)
        RedoRelabelNote(Notes(0), vBPM, xUndo, xRedo)

        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        Notes(0).Value = vBPM
        THBPM.Value = vBPM/10000
        CalculateTotalPlayableNotes()
        
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Public Function MyO2GridCheck() As String()
        
        SortByVPositionInsertion()
        Dim xResult(- 1) As String
        Dim xResult2(- 1) As String
        Dim Identifiers() As String = {"01", "03", "04", "06", "07", "08", "09",
                                       "16", "11", "12", "13", "14", "15", "18", "19",
                                       "26", "21", "22", "23", "24", "25", "28", "29",
                                       "36", "31", "32", "33", "34", "35", "38", "39",
                                       "46", "41", "42", "43", "44", "45", "48", "49",
                                       "56", "51", "52", "53", "54", "55", "58", "59",
                                       "66", "61", "62", "63", "64", "65", "68", "69",
                                       "76", "71", "72", "73", "74", "75", "78", "79",
                                       "86", "81", "82", "83", "84", "85", "88", "89"}

        Dim xLowerIndex = 1
        Dim xUpperIndex = 1

        If Not NTInput Then
            For xMeasure = 0 To MeasureAtDisplacement(GreatestVPosition)
                'Find Ks in the same measure
                Dim i As Integer
                For i = xLowerIndex To UBound(Notes)
                    If MeasureAtDisplacement(Notes(i).VPosition) > xMeasure Then Exit For
                Next
                xUpperIndex = i

                For Each xId As String In Identifiers
                    'collect vposition data
                    Dim xVPos(- 1) As Double
                    For j As Integer = xLowerIndex To xUpperIndex - 1
                        If _
                            Columns.GetBMSChannelBy(Notes(j)) = xId And
                            Math.Abs(Notes(j).VPosition - MeasureAtDisplacement(Notes(j).VPosition)*192) > 0 Then
                            ReDim Preserve xVPos(UBound(xVPos) + 1)
                            xVPos(UBound(xVPos)) = Notes(j).VPosition - xMeasure*192
                            If xVPos(UBound(xVPos)) < 0 Then xVPos(UBound(xVPos)) = 0
                        End If
                    Next

                    'find gcd
                    Dim xGCD As Double = 192
                    For j = 0 To UBound(xVPos)
                        xGCD = GCD(xGCD, xVPos(j))
                    Next

                    'check if smaller than minGCD
                    If xGCD < 3 Then
                        'suggestion
                        Dim xAdj64 As Boolean
                        Dim xD48 = 0
                        Dim xD64 = 0
                        For j = 0 To UBound(xVPos)
                            xD48 += Math.Abs(xVPos(j) - CInt(xVPos(j)/4)*4)
                            xD64 += Math.Abs(xVPos(j) - CInt(xVPos(j)/3)*3)
                        Next
                        xAdj64 = xD48 > xD64

                        'put result
                        ReDim Preserve xResult(UBound(xResult) + 1)
                        xResult(UBound(xResult)) = xMeasure & "_" &
                                                   Columns.BMSChannelToColumn(xId) & "_" &
                                                   Columns.GetName(Columns.BMSChannelToColumn(xId)) & "_" &
                                                   CStr(CInt(192/xGCD)) & "_" &
                                                   CInt(IsChannelLongNote(xId)) & "_" &
                                                   CInt(IsChannelHidden(xId)) & "_" &
                                                   CInt(xAdj64) & "_" &
                                                   xD64 & "_" &
                                                   xD48
                    End If
                Next

                980: xLowerIndex = xUpperIndex
                990: Next

        Else
            For xMeasure = 0 To MeasureAtDisplacement(GreatestVPosition)

                For Each xId As String In Identifiers
                    Dim xVPos(- 1) As Double

                    'collect vposition data
                    For j = 1 To UBound(Notes)
                        If MeasureAtDisplacement(Notes(j).VPosition) > xMeasure Then Exit For

                        If Columns.GetBMSChannelBy(Notes(j)) <> xId Then GoTo 1330
                        If IsChannelLongNote(xId) Xor CBool(Notes(j).Length) Then GoTo 1330

                        If _
                            MeasureAtDisplacement(Notes(j).VPosition) = xMeasure AndAlso
                            Math.Abs(Notes(j).VPosition - MeasureAtDisplacement(Notes(j).VPosition)*192) > 0 Then
                            ReDim Preserve xVPos(UBound(xVPos) + 1)
                            xVPos(UBound(xVPos)) = Notes(j).VPosition - xMeasure*192
                            If xVPos(UBound(xVPos)) < 0 Then xVPos(UBound(xVPos)) = 0
                        End If

                        If Not CBool(Notes(j).Length) Then GoTo 1330

                        If _
                            MeasureAtDisplacement(Notes(j).VPosition + Notes(j).Length) = xMeasure AndAlso
                            Not Notes(j).VPosition + Notes(j).Length - xMeasure*192 = 0 Then
                            ReDim Preserve xVPos(UBound(xVPos) + 1)
                            xVPos(UBound(xVPos)) = Notes(j).VPosition + Notes(j).Length - xMeasure*192
                            If xVPos(UBound(xVPos)) < 0 Then xVPos(UBound(xVPos)) = 0
                        End If
                        1330: Next

                    'find gcd
                    Dim xGCD As Double = 192
                    For j = 0 To UBound(xVPos)
                        xGCD = GCD(xGCD, xVPos(j))
                    Next

                    'check if smaller than minGCD
                    If xGCD < 3 Then
                        'suggestion
                        Dim xAdj64 As Boolean
                        Dim xD48 = 0
                        Dim xD64 = 0
                        For j = 0 To UBound(xVPos)
                            xD48 += Math.Abs(xVPos(j) - CInt(xVPos(j)/4)*4)
                            xD64 += Math.Abs(xVPos(j) - CInt(xVPos(j)/3)*3)
                        Next
                        xAdj64 = xD48 > xD64

                        'put result
                        ReDim Preserve xResult(UBound(xResult) + 1)
                        xResult(UBound(xResult)) = xMeasure & "_" &
                                                   Columns.BMSChannelToColumn(xId) & "_" &
                                                   Columns.GetName(Columns.BMSChannelToColumn(xId)) & "_" &
                                                   CStr(CInt(192/xGCD)) & "_" &
                                                   CInt(IsChannelLongNote(xId)) & "_" &
                                                   CInt(IsChannelHidden(xId)) & "_" &
                                                   CInt(xAdj64) & "_" &
                                                   xD64 & "_" &
                                                   xD48
                    End If
                Next

                1990: Next
        End If

        Return xResult
    End Function

    Public Sub MyO2GridAdjust(xaj() As dgMyO2.Adj)
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        RedoRemoveNoteAll(Notes, xUndo, xRedo)

        'adjust
        If Not NTInput Then
            For Each xadj As dgMyO2.Adj In xaj
                For i = 1 To UBound(Notes)
                    If MeasureAtDisplacement(Notes(i).VPosition) = xadj.Measure And
                       Notes(i).ColumnIndex = xadj.ColumnIndex And
                       Notes(i).LongNote = xadj.LongNote And
                       Notes(i).Hidden = xadj.Hidden Then
                        Notes(i).VPosition = CLng(Notes(i).VPosition/IIf(xadj.AdjTo64, 3, 4))*IIf(xadj.AdjTo64, 3, 4)
                    End If
                Next
            Next

        Else
            For Each xadj As dgMyO2.Adj In xaj
                For i = 1 To UBound(Notes)
                    If CBool(Notes(i).Length) Xor xadj.LongNote Then GoTo 1100

                    Dim xStart As Double = Notes(i).VPosition
                    Dim xEnd As Double = Notes(i).VPosition + Notes(i).Length
                    If MeasureAtDisplacement(Notes(i).VPosition) = xadj.Measure And
                       Notes(i).ColumnIndex = xadj.ColumnIndex And
                       Notes(i).Hidden = xadj.Hidden Then _
                        xStart = CLng(Notes(i).VPosition/IIf(xadj.AdjTo64, 3, 4))*IIf(xadj.AdjTo64, 3, 4)

                    If Notes(i).Length > 0 AndAlso
                       MeasureAtDisplacement(Notes(i).VPosition + Notes(i).Length) = xadj.Measure And
                       Notes(i).ColumnIndex = xadj.ColumnIndex And
                       Notes(i).Hidden = xadj.Hidden Then _
                        xEnd = CLng((Notes(i).VPosition + Notes(i).Length)/IIf(xadj.AdjTo64, 3, 4))*
                               IIf(xadj.AdjTo64, 3, 4)

                    Notes(i).VPosition = xStart
                    If Notes(i).Length > 0 Then Notes(i).Length = xEnd - Notes(i).VPosition
                    1100: Next
            Next
        End If

        RedoAddNoteAll(Notes, xUndo, xRedo)

        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        
        RefreshPanelAll()
        POStatusRefresh()
        Beep()
    End Sub
End Class
