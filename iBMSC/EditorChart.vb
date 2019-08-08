
Imports System.Linq
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
        If NtInput Then
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
            Dim wavidx = note2.Value / 10000
            Dim wav = BmsWAV(wavidx)
            If BmsWAV(note.Value / 10000) = wav Then
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
            Dim wavidx = Notes(index).Value / 10000
            Dim wav = BmsWAV(wavidx)
            If BmsWAV(note.Value / 10000) = wav Then
                Return True
            End If
        Else
            If note.Value = Notes(index).Value Then
                Return True
            End If
        End If

        Return False
    End Function

    Public Sub AddNote(note As Note, Optional xSelected As Boolean = False, Optional overWrite As Boolean = True,
                       Optional sortAndUpdatePairing As Boolean = True)

        If note.VPosition < 0 Or note.VPosition >= GetMaxVPosition() Then Exit Sub

        Dim i = 1

        If overWrite Then
            Do While i <= UBound(Notes)
                If Notes(i).VPosition = note.VPosition And
                   Notes(i).ColumnIndex = note.ColumnIndex Then
                    RemoveNote(i)
                Else
                    i += 1
                End If
            Loop
        End If

        note.Selected = note.Selected And Columns.IsEnabled(note.ColumnIndex)
        Notes = Notes.Concat({note}).ToArray()

        If sortAndUpdatePairing Then
            ValidateNotesArray()
        Else
            CalculateTotalPlayableNotes()
        End If
    End Sub

    Public Sub RemoveNote(I As Integer, Optional ByVal sortAndUpdatePairing As Boolean = True)
        State.Mouse.CurrentHoveredNoteIndex = -1 ' az: Why here???
        Dim j As Integer

        If TBWavIncrease.Checked Then
            If Notes(I).Value = LWAV.SelectedIndex * 10000 Then
                DecreaseCurrentWav()
            End If
        End If


        For j = I + 1 To UBound(Notes)
            Notes(j - 1) = Notes(j)
        Next
        ReDim Preserve Notes(UBound(Notes) - 1)
        If sortAndUpdatePairing Then
            ValidateNotesArray()
        End If
    End Sub


    Private Sub RemoveNotes(Optional ByVal sortAndUpdatePairing As Boolean = True)
        If UBound(Notes) = 0 Then Exit Sub

        State.Mouse.CurrentHoveredNoteIndex = -1
        Dim i = 1
        Dim j As Integer
        Do
            If Notes(i).Selected Then
                For j = i + 1 To UBound(Notes)
                    Notes(j - 1) = Notes(j)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
                i = 0
            End If
            i += 1
        Loop While i < UBound(Notes) + 1
        If sortAndUpdatePairing Then
            ValidateNotesArray()
        End If
        CalculateTotalPlayableNotes()
    End Sub

    Private Function GetTimeFromVPosition(vpos As Double) As Double
        Dim timingNotes = (From note In Notes
                           Where note.ColumnIndex = ColumnType.BPM Or note.ColumnIndex = ColumnType.STOPS
                           Group By column = note.ColumnIndex
                Into noteGroups = Group).ToDictionary(Function(x) x.column, Function(x) x.noteGroups)

        Dim bpmNotes = timingNotes.Item(ColumnType.BPM)

        Dim stopNotes As IEnumerable(Of Note) = Nothing

        If timingNotes.ContainsKey(ColumnType.STOPS) Then
            stopNotes = timingNotes.Item(ColumnType.STOPS)
        End If


        Dim stopContrib As Double
        Dim bpmContrib As Double

        For i = 0 To bpmNotes.Count() - 1
            ' az: sum bpm contribution first
            Dim duration = 0.0
            Dim currentNote = bpmNotes.ElementAt(i)
            Dim notevpos = Math.Max(0, currentNote.VPosition)

            If i + 1 <> bpmNotes.Count() Then
                Dim nextNote = bpmNotes.ElementAt(i + 1)
                duration = nextNote.VPosition - notevpos
            Else
                duration = vpos - notevpos
            End If

            Dim currentBps = 60 / (currentNote.Value / 10000)
            bpmContrib += currentBps * duration / 48

            If stopNotes Is Nothing Then Continue For

            Dim stops = From stp In stopNotes
                        Where stp.VPosition >= notevpos And
                          stp.VPosition < notevpos + duration

            Dim stopBeats = stops.Sum(Function(x) x.Value / 10000.0) / 48
            stopContrib += currentBps * stopBeats

        Next

        Return stopContrib + bpmContrib
    End Function
End Class
