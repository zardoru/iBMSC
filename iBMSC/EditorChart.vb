
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
