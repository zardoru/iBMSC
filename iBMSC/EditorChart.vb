
Imports iBMSC.Editor

Partial Public Class MainWindow
    '----Header Options
    Public BmsWAV(1295) As String
    Public BmsBMP(1295) As String
    Dim BmsBPM(1295) As Long   'x10000
    Dim BmsSTOP(1295) As Long
    Dim BmsSCROLL(1295) As Long

    Public Notes() As Note = {New Note(ColumnType.BPM, - 1, 1200000, 0, False)}
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
            For j = i - 1 To 1 Step - 1
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

    Private Sub SortByVPositionQuick(xMin As Integer, xMax As Integer) 'Quick Sort
        Dim xNote As Note
        Dim iHi As Integer
        Dim iLo As Integer
        Dim i As Integer

        ' If min >= max, the list contains 0 or 1 items so it is sorted.
        If xMin >= xMax Then Exit Sub

        ' Pick the dividing value.
        i = CInt((xMax - xMin)/2) + xMin
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
    
    Public Function MeasureAtDisplacement(xVPos As Double) As Integer
        Dim i As Integer
        For i = 1 To 999
            If xVPos < MeasureBottom(i) Then Exit For
        Next
        Return i - 1
    End Function

    Public Function GetMaxVPosition() As Double
        Return MeasureUpper(999)
    End Function
    
    Public Function IsLabelMatch(note As Note, note2 As Note) As Boolean
        If TBShowFileName.Checked Then
            Dim wavidx = note2.Value/10000
            Dim wav = BmsWAV(wavidx)
            If BmsWAV(note.Value/10000) = wav Then
                Return True
            End If
        Else
            If note.Value = note2.Value Then
                Return True
            End If
        End If

        Return False
    End Function
    
    
    Public Function IsLabelMatch(note As Note, index As Integer) As Boolean
        If TBShowFileName.Checked Then
            Dim wavidx = Notes(index).Value/10000
            Dim wav = BmsWAV(wavidx)
            If BmsWAV(note.Value/10000) = wav Then
                Return True
            End If
        Else
            If note.Value = Notes(index).Value Then
                Return True
            End If
        End If

        Return False
    End Function
End Class
