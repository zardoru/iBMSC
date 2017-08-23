Module BMS
    Public Function IdentifiertoLongNote(ByVal I As String) As Boolean
        Dim xI As Integer = CInt(Val(I))
        Return xI >= 50 And xI < 90
    End Function

    Public Function IsChannelHidden(ByVal I As String) As Boolean
        Dim xI As Integer = CInt(Val(I))
        Return (xI >= 30 And xI < 50) Or (xI >= 70 And xI < 90)
    End Function
End Module
