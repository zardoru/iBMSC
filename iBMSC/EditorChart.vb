

Imports iBMSC.Editor

Partial Public Class MainWindow
    '----Header Options
    Public BmsWAV(1295) As String
    Public BmsBMP(1295) As String
    Dim BmsBPM(1295) As Long   'x10000
    Dim BmsSTOP(1295) As Long
    Dim BmsSCROLL(1295) As Long

    Public Notes() As Note = {New Note(ColumnType.BPM, -1, 1200000, 0, False)}
    Dim LnObj As Integer = 0    '0 for none, 1-1295 for 01-ZZ

    Private Function FindNoteIndex(note As Note) As Integer
        Dim i As Integer
        If NTInput Then
            For i = 1 To UBound(Notes)
                If Notes(i).equalsNT(note) Then Return i
            Next
        Else
            For i = 1 To UBound(Notes)
                If Notes(i).equalsBMSE(note) Then Return i
            Next
        End If
        Return i
    End Function

    ' az: why _three_ different sort functions? why. _why_
    Public Sub SortByVPositionInsertion() 'Insertion Sort
        If Notes.Length <= 1 Then Exit Sub
        Dim xNote As Note
        Dim i As Integer
        Dim j As Integer
        For i = 2 To Notes.Length - 1
            xNote = Notes(i)
            For j = i - 1 To 1 Step -1
                If Notes(j).VPosition > xNote.VPosition Then
                    Notes(j + 1) = Notes(j)
                    If j = 1 Then
                        Notes(j) = xNote
                        Exit For
                    End If
                Else
                    Notes(j + 1) = xNote
                    Exit For
                End If
            Next
        Next

    End Sub

    Private Sub SortByVPositionQuick(ByVal xMin As Integer, ByVal xMax As Integer) 'Quick Sort
        Dim xNote As Note
        Dim iHi As Integer
        Dim iLo As Integer
        Dim i As Integer

        ' If min >= max, the list contains 0 or 1 items so it is sorted.
        If xMin >= xMax Then Exit Sub

        ' Pick the dividing value.
        i = CInt((xMax - xMin) / 2) + xMin
        xNote = Notes(i)

        ' Swap it to the front.
        Notes(i) = Notes(xMin)

        iLo = xMin
        iHi = xMax
        Do
            ' Look down from hi for a value < med_value.
            Do While Notes(iHi).VPosition >= xNote.VPosition
                iHi = iHi - 1
                If iHi <= iLo Then Exit Do
            Loop
            If iHi <= iLo Then
                Notes(iLo) = xNote
                Exit Do
            End If

            ' Swap the lo and hi values.
            Notes(iLo) = Notes(iHi)

            ' Look up from lo for a value >= med_value.
            iLo = iLo + 1
            Do While Notes(iLo).VPosition < xNote.VPosition
                iLo = iLo + 1
                If iLo >= iHi Then Exit Do
            Loop
            If iLo >= iHi Then
                iLo = iHi
                Notes(iHi) = xNote
                Exit Do
            End If

            ' Swap the lo and hi values.
            Notes(iHi) = Notes(iLo)
        Loop

        ' Sort the two sublists.
        SortByVPositionQuick(xMin, iLo - 1)
        SortByVPositionQuick(iLo + 1, xMax)
    End Sub

    Private Sub SortByVPositionQuick3(ByVal xMin As Integer, ByVal xMax As Integer)
        Dim xxMin As Integer
        Dim xxMax As Integer
        Dim xxMid As Integer
        Dim xNote As Note
        Dim xNoteMid As Note
        Dim i As Integer
        Dim j As Integer
        Dim xI3 As Integer

        'If xMax = 0 Then
        '    xMin = LBound(K1)
        '    xMax = UBound(K1)
        'End If
        xxMin = xMin
        xxMax = xMax
        xxMid = xMax - xMin + 1
        i = CInt(Int(xxMid * Rnd())) + xMin
        j = CInt(Int(xxMid * Rnd())) + xMin
        xI3 = CInt(Int(xxMid * Rnd())) + xMin
        If Notes(i).VPosition <= Notes(j).VPosition And Notes(j).VPosition <= Notes(xI3).VPosition Then
            xxMid = j
        Else
            If Notes(j).VPosition <= Notes(i).VPosition And Notes(i).VPosition <= Notes(xI3).VPosition Then
                xxMid = i
            Else
                xxMid = xI3
            End If
        End If
        xNoteMid = Notes(xxMid)
        Do
            Do While Notes(xxMin).VPosition < xNoteMid.VPosition And xxMin < xMax
                xxMin = xxMin + 1
            Loop
            Do While xNoteMid.VPosition < Notes(xxMax).VPosition And xxMax > xMin
                xxMax = xxMax - 1
            Loop
            If xxMin <= xxMax Then
                xNote = Notes(xxMin)
                Notes(xxMin) = Notes(xxMax)
                Notes(xxMax) = xNote
                xxMin = xxMin + 1
                xxMax = xxMax - 1
            End If
        Loop Until xxMin > xxMax
        If xxMax - xMin < xMax - xxMin Then
            If xMin < xxMax Then SortByVPositionQuick3(xMin, xxMax)
            If xxMin < xMax Then SortByVPositionQuick3(xxMin, xMax)
        Else
            If xxMin < xMax Then SortByVPositionQuick3(xxMin, xMax)
            If xMin < xxMax Then SortByVPositionQuick3(xMin, xxMax)
        End If
    End Sub
End Class
