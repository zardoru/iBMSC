Namespace Editor
    Public Structure COverride
        Public RangeL() As Integer
        Public RangeR() As Integer
        Public NoteColor() As Integer
        Public LongNoteColor() As Integer
        Public LongTextColor() As Integer
        Public BG() As Integer
        Public Len As Integer

        ' New(ByVal xLeft As Integer, ByVal xWidth As Integer, ByVal xTitle As String,
        ' ByVal xNoteCol As Boolean, ByVal xisNumeric As Boolean, ByVal xVisible As Boolean, ByVal xIdentifier As Integer,
        ' ByVal xcNote As Integer, ByVal xcText As Integer, ByVal xcLNote As Integer, ByVal xcLText As Integer, ByVal xcBG As Integer)
        ' New Column(339, 45, "A4", True, False, True, 13, &HFFFFC862, &HFF000000, &HFFF7C66A, &HFF000000, &H16F38B0C),

        Public Sub New(ByVal iRangeL As Integer,
                       ByVal iRangeR As Integer,
                       ByVal iNoteColor As Integer,
                       ByVal iLongNoteColor As Integer,
                       ByVal iLongTextColor As Integer,
                       ByVal iBG As Integer)
            ' ByVal Len As Integer)
            ReDim RangeL(0)
            ReDim RangeR(0)
            ReDim NoteColor(0)
            ReDim LongNoteColor(0)
            ReDim LongTextColor(0)
            ReDim BG(0)

            RangeL(0) = iRangeL
            RangeR(0) = iRangeR
            NoteColor(0) = iNoteColor
            LongNoteColor(0) = iLongNoteColor
            LongTextColor(0) = iLongTextColor
            BG(0) = iBG
            Len = 0
        End Sub

        Public Sub Resize(ByVal n As Integer)
            Array.Resize(RangeL, n)
            Array.Resize(RangeR, n)
            Array.Resize(NoteColor, n)
            Array.Resize(LongNoteColor, n)
            Array.Resize(LongTextColor, n)
            Array.Resize(BG, n)
            Len = n - 1
        End Sub
    End Structure
End Namespace
