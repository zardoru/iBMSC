Imports iBMSC.Editor

Partial Public Class MainWindow
    Private Sub AddNotesFromClipboard(Optional ByVal xSelected As Boolean = True,
                                      Optional ByVal SortAndUpdatePairing As Boolean = True)
        Dim xStrLine() As String = Split(Clipboard.GetText, vbCrLf)

        Dim i As Integer
        For i = 0 To UBound(Notes)
            Notes(i).Selected = False
        Next

        Dim verticalScroll As Long = FocusedPanel.VerticalPosition
        Dim xTempVP As Double
        Dim xKbu() As Note = Notes

        If xStrLine(0) = "iBMSC Clipboard Data" Then
            If NTInput Then ReDim Preserve Notes(0)

            'paste
            Dim xStrSub() As String
            For i = 1 To UBound(xStrLine)
                If xStrLine(i).Trim = "" Then Continue For
                xStrSub = Split(xStrLine(i), " ")
                xTempVP = Val(xStrSub(1)) + MeasureBottom(MeasureAtDisplacement(- verticalScroll) + 1)
                If UBound(xStrSub) = 5 And xTempVP >= 0 And xTempVP < GetMaxVPosition() Then
                    ReDim Preserve Notes(UBound(Notes) + 1)
                    With Notes(UBound(Notes))
                        .ColumnIndex = Val(xStrSub(0))
                        .VPosition = xTempVP
                        .Value = Val(xStrSub(2))
                        .LongNote = CBool(Val(xStrSub(3)))
                        .Hidden = CBool(Val(xStrSub(4)))
                        .Landmine = CBool(Val(xStrSub(5)))
                        .Selected = xSelected
                    End With
                End If
            Next

            'convert
            If NTInput Then
                ConvertBMSE2NT()

                For i = 1 To UBound(Notes)
                    Notes(i - 1) = Notes(i)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)

                Dim xKn() As Note = Notes
                Notes = xKbu

                Dim xIStart As Integer = Notes.Length
                ReDim Preserve Notes(UBound(Notes) + xKn.Length)

                For i = xIStart To UBound(Notes)
                    Notes(i) = xKn(i - xIStart)
                Next
            End If

        ElseIf xStrLine(0) = "iBMSC Clipboard Data xNT" Then
            If Not NTInput Then ReDim Preserve Notes(0)

            'paste
            Dim xStrSub() As String
            For i = 1 To UBound(xStrLine)
                If xStrLine(i).Trim = "" Then Continue For
                xStrSub = Split(xStrLine(i), " ")
                xTempVP = Val(xStrSub(1)) + MeasureBottom(MeasureAtDisplacement(- verticalScroll) + 1)
                If UBound(xStrSub) = 5 And xTempVP >= 0 And xTempVP < GetMaxVPosition() Then
                    ReDim Preserve Notes(UBound(Notes) + 1)
                    With Notes(UBound(Notes))
                        .ColumnIndex = Val(xStrSub(0))
                        .VPosition = xTempVP
                        .Value = Val(xStrSub(2))
                        .Length = Val(xStrSub(3))
                        .Hidden = CBool(Val(xStrSub(4)))
                        .Landmine = CBool(Val(xStrSub(5)))
                        .Selected = xSelected
                    End With
                End If
            Next

            'convert
            If Not NTInput Then
                ConvertNT2BMSE()

                For i = 1 To UBound(Notes)
                    Notes(i - 1) = Notes(i)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)

                Dim xKn() As Note = Notes
                Notes = xKbu

                Dim xIStart As Integer = Notes.Length
                ReDim Preserve Notes(UBound(Notes) + xKn.Length)

                For i = xIStart To UBound(Notes)
                    Notes(i) = xKn(i - xIStart)
                Next
            End If

        ElseIf xStrLine(0) = "BMSE ClipBoard Object Data Format" Then
            If NTInput Then ReDim Preserve Notes(0)

            'paste
            For i = 1 To UBound(xStrLine)
                ' zdr: holy crap this is obtuse
                Dim posStr = Mid(xStrLine(i), 5, 7)
                Dim vPos = Val(posStr) + MeasureBottom(MeasureAtDisplacement(- verticalScroll) + 1)

                Dim bmsIdent = Mid(xStrLine(i), 1, 3)
                Dim lineCol = Columns.BMSEChannelToColumnIndex(bmsIdent)

                Dim Value = Val(Mid(xStrLine(i), 12))*10000

                Dim attribute = Mid(xStrLine(i), 4, 1)

                Dim validCol = Len(xStrLine(i)) > 11 And lineCol > 0
                Dim inRange = vPos >= 0 And vPos < GetMaxVPosition()
                If validCol And inRange Then
                    ReDim Preserve Notes(UBound(Notes) + 1)

                    With Notes(UBound(Notes))
                        .ColumnIndex = lineCol
                        .VPosition = vPos
                        .Value = Value
                        .LongNote = attribute = "2"
                        .Hidden = attribute = "1"
                        .Selected = xSelected And Columns.IsEnabled(.ColumnIndex)
                    End With
                End If
            Next

            'convert
            If NTInput Then
                ConvertBMSE2NT()

                For i = 1 To UBound(Notes)
                    Notes(i - 1) = Notes(i)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)

                Dim xKn() As Note = Notes
                Notes = xKbu

                Dim xIStart As Integer = Notes.Length
                ReDim Preserve Notes(UBound(Notes) + xKn.Length)

                For i = xIStart To UBound(Notes)
                    Notes(i) = xKn(i - xIStart)
                Next
            End If
        End If

        If SortAndUpdatePairing Then
            ValidateNotesArray()
        End If
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub CopyNotes(Optional ByVal Unselect As Boolean = True)
        Dim xStrAll As String = "iBMSC Clipboard Data" & IIf(NTInput, " xNT", "")
        Dim i As Integer
        Dim MinMeasure As Double = 999

        For i = 1 To UBound(Notes)
            If Notes(i).Selected And MeasureAtDisplacement(Notes(i).VPosition) < MinMeasure Then _
                MinMeasure = MeasureAtDisplacement(Notes(i).VPosition)
        Next
        MinMeasure = MeasureBottom(MinMeasure)

        If Not NTInput Then
            For i = 1 To UBound(Notes)
                If Notes(i).Selected Then
                    xStrAll &= vbCrLf & Notes(i).ColumnIndex.ToString & " " &
                               (Notes(i).VPosition - MinMeasure).ToString & " " &
                               Notes(i).Value.ToString & " " &
                               CInt(Notes(i).LongNote).ToString & " " &
                               CInt(Notes(i).Hidden).ToString & " " &
                               CInt(Notes(i).Landmine).ToString
                    Notes(i).Selected = Not Unselect
                End If
            Next

        Else
            For i = 1 To UBound(Notes)
                If Notes(i).Selected Then
                    xStrAll &= vbCrLf & Notes(i).ColumnIndex.ToString & " " &
                               (Notes(i).VPosition - MinMeasure).ToString & " " &
                               Notes(i).Value.ToString & " " &
                               Notes(i).Length.ToString & " " &
                               CInt(Notes(i).Hidden).ToString & " " &
                               CInt(Notes(i).Landmine).ToString
                    Notes(i).Selected = Not Unselect
                End If
            Next
        End If

        Clipboard.SetText(xStrAll)
    End Sub
End Class
