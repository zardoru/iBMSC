Imports iBMSC.Editor

Partial Public Class MainWindow

    Public Const niMeasure As Integer = 0
    Public Const niSCROLL As Integer = 1
    Public Const niBPM As Integer = 2
    Public Const niSTOP As Integer = 3
    Public Const niS1 As Integer = 4

    Public Const niA1 As Integer = 5
    Public Const niA2 As Integer = 6
    Public Const niA3 As Integer = 7
    Public Const niA4 As Integer = 8
    Public Const niA5 As Integer = 9
    Public Const niA6 As Integer = 10
    Public Const niA7 As Integer = 11
    Public Const niA8 As Integer = 12
    Public Const niS2 As Integer = 13

    Public Const niD1 As Integer = 14
    Public Const niD2 As Integer = 15
    Public Const niD3 As Integer = 16
    Public Const niD4 As Integer = 17
    Public Const niD5 As Integer = 18
    Public Const niD6 As Integer = 19
    Public Const niD7 As Integer = 20
    Public Const niD8 As Integer = 21
    Public Const niS3 As Integer = 22

    Public Const niBGA As Integer = 23
    Public Const niLAYER As Integer = 24
    Public Const niPOOR As Integer = 25
    Public Const niS4 As Integer = 26
    Public Const niB As Integer = 27

    Public column() As Column = {New Column(0, 50, "Measure", False, True, False, True, 0, 0, &HFF00FFFF, 0, &HFF00FFFF, 0),
                              New Column(50, 60, "SCROLL", True, True, False, True, 99, 0, &HFFFF0000, 0, &HFFFF0000, 0),
                              New Column(110, 60, "BPM", True, True, False, True, 3, 0, &HFFFF0000, 0, &HFFFF0000, 0),
                              New Column(170, 50, "STOP", True, True, False, True, 9, 0, &HFFFF0000, 0, &HFFFF0000, 0),
                              New Column(220, 5, "", False, False, False, True, 0, 0, 0, 0, 0, 0),
                              New Column(225, 42, "A1", True, False, True, True, 16, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(267, 30, "A2", True, False, True, True, 11, &HFF62B0FF, &HFF000000, &HFF6AB0F7, &HFF000000, &H140033FF),
                              New Column(297, 42, "A3", True, False, True, True, 12, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(339, 45, "A4", True, False, True, True, 13, &HFFFFC862, &HFF000000, &HFFF7C66A, &HFF000000, &H16F38B0C),
                              New Column(384, 42, "A5", True, False, True, True, 14, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(426, 30, "A6", True, False, True, True, 15, &HFF62B0FF, &HFF000000, &HFF6AB0F7, &HFF000000, &H140033FF),
                              New Column(456, 42, "A7", True, False, True, True, 18, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(498, 40, "A8", True, False, True, True, 19, &HFF808080, &HFF000000, &HFF909090, &HFF000000, 0),
                              New Column(498, 5, "", False, False, False, True, 0, 0, 0, 0, 0, 0),
                              New Column(503, 42, "D1", True, False, True, False, 21, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(503, 30, "D2", True, False, True, False, 22, &HFF62B0FF, &HFF000000, &HFF6AB0F7, &HFF000000, &H140033FF),
                              New Column(503, 42, "D3", True, False, True, False, 23, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(503, 45, "D4", True, False, True, False, 24, &HFFFFC862, &HFF000000, &HFFF7C66A, &HFF000000, &H16F38B0C),
                              New Column(503, 42, "D5", True, False, True, False, 25, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(503, 30, "D6", True, False, True, False, 28, &HFF62B0FF, &HFF000000, &HFF6AB0F7, &HFF000000, &H140033FF),
                              New Column(503, 42, "D7", True, False, True, False, 29, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(503, 40, "D8", True, False, True, False, 26, &HFF808080, &HFF000000, &HFF909090, &HFF000000, 0),
                              New Column(503, 5, "", False, False, False, False, 0, 0, 0, 0, 0, 0),
                              New Column(503, 40, "BGA", True, False, False, False, 4, &HFF8CD78A, &HFF000000, &HFF90D38E, &HFF000000, 0),
                              New Column(503, 40, "LAYER", True, False, False, False, 7, &HFF8CD78A, &HFF000000, &HFF90D38E, &HFF000000, 0),
                              New Column(503, 40, "POOR", True, False, False, False, 6, &HFF8CD78A, &HFF000000, &HFF90D38E, &HFF000000, 0),
                              New Column(503, 5, "", False, False, False, False, 0, 0, 0, 0, 0, 0),
                              New Column(503, 40, "B", True, False, True, True, 1, &HFFE18080, &HFF000000, &HFFDC8585, &HFF000000, 0)}


    Public Const idflBPM As Integer = 5

    Private Function GetBMSChannelBy(note As Note) As String
        Dim iCol = note.ColumnIndex
        Dim xVal = note.Value
        Dim xLong = note.LongNote
        Dim xHidden = note.Hidden
        Dim bmsBaseChannel As Integer = GetColumn(iCol).Identifier
        Dim xLandmine = note.Landmine

        If iCol = niBPM AndAlso (xVal / 10000 <> xVal \ 10000 Or xVal >= 2560000 Or xVal < 0) Then bmsBaseChannel += idflBPM

        If iCol = niSCROLL Then Return "SC"

        ' p1 side
        If iCol >= niA1 And iCol <= niA8 Then
            If xLong Then
                Return Hex(bmsBaseChannel + Convert.ToInt32("50", 16) - 10)
            End If
            If xHidden Then
                Return Hex(bmsBaseChannel + Convert.ToInt32("30", 16) - 10)
            End If
            If xLandmine Then
                Return Hex(bmsBaseChannel + Convert.ToInt32("D0", 16) - 10)
            End If
        End If

        ' p2 side
        If iCol >= niD1 And iCol <= niD8 Then
            If xLong Then
                Return Hex(bmsBaseChannel + Convert.ToInt32("60", 16) - 20)
            End If
            If xHidden Then
                Return Hex(bmsBaseChannel + Convert.ToInt32("40", 16) - 20)
            End If
            If xLandmine Then
                Return Hex(bmsBaseChannel + Convert.ToInt32("E0", 16) - 20)
            End If
        End If

        Return Add2Zeros(bmsBaseChannel)
    End Function

    Private Function nLeft(ByVal iCol As Integer) As Integer
        If iCol < niB Then Return column(iCol).Left Else Return column(niB).Left + (iCol - niB) * column(niB).Width
    End Function
    Private Function GetColumnWidth(ByVal iCol As Integer) As Integer
        If Not GetColumn(iCol).isVisible Then Return 0
        If iCol < niB Then Return column(iCol).Width Else Return column(niB).Width
    End Function
    Private Function nTitle(ByVal iCol As Integer) As String
        If iCol < niB Then Return column(iCol).Title Else Return column(niB).Title & (iCol - niB + 1).ToString
    End Function
    Private Function nEnabled(ByVal iCol As Integer) As Boolean
        'If iCol < niB Then Return col(iCol).Enabled And col(iCol).Visible Else Return col(niB).Enabled And col(niB).Visible
        If iCol < niB Then Return column(iCol).isEnabledAfterAll Else Return column(niB).isEnabledAfterAll
    End Function
    Private Function IsColumnNumeric(ByVal iCol As Integer) As Boolean
        If iCol < niB Then Return column(iCol).isNumeric Else Return column(niB).isNumeric
    End Function
    Private Function IsColumnSound(ByVal iCol As Integer) As Boolean
        If iCol < niB Then Return column(iCol).isSound Else Return column(niB).isSound
    End Function



    Private Function GetColumn(ByVal iCol As Integer) As Column
        If iCol < niB Then Return column(iCol) Else Return column(niB)
    End Function

    Private Function BMSEChannelToColumnIndex(ByVal I As String)
        Dim Ivalue = Val(I)
        If Ivalue > 100 Then
            Return niB + Ivalue - 101
        ElseIf Ivalue < 100 And Ivalue > 0 Then
            Return BMSChannelToColumn(Mid(I, 2, 2))
        End If
        Return niB ' ??? how did a negative number get here?
    End Function

    Private Function BMSChannelToColumn(ByVal I As String) As Integer
        Select Case I
            Case "01" : Return niB
            Case "03", "08" : Return niBPM
            Case "09" : Return niSTOP
            Case "SC" : Return niSCROLL
            Case "04" : Return niBGA
            Case "07" : Return niLAYER
            Case "06" : Return niPOOR

            Case "16", "36", "56", "76", "D6" : Return niA1
            Case "11", "31", "51", "71", "D1" : Return niA2
            Case "12", "32", "52", "72", "D2" : Return niA3
            Case "13", "33", "53", "73", "D3" : Return niA4
            Case "14", "34", "54", "74", "D4" : Return niA5
            Case "15", "35", "55", "75", "D5" : Return niA6
            Case "18", "38", "58", "78", "D8" : Return niA7
            Case "19", "39", "59", "79", "D9" : Return niA8

            Case "21", "41", "61", "81", "E1" : Return niD1
            Case "22", "42", "62", "82", "E2" : Return niD2
            Case "23", "43", "63", "83", "E3" : Return niD3
            Case "24", "44", "64", "84", "E4" : Return niD4
            Case "25", "45", "65", "85", "E5" : Return niD5
            Case "28", "48", "68", "88", "E8" : Return niD6
            Case "29", "49", "69", "89", "E9" : Return niD7
            Case "26", "46", "66", "86", "E6" : Return niD8

            Case Else : Return 0
        End Select
    End Function

End Class
