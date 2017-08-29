Namespace Editor
    Structure Note
        Public VPosition As Double
        Public ColumnIndex As Integer
        Public Value As Long 'x10000
        Public LongNote As Boolean
        Public Hidden As Boolean
        Public Length As Double
        Public Landmine As Boolean

        Public LNPair As Integer
        Public Selected As Boolean
        Public HasError As Boolean

        'Public TempBoolean As Boolean
        Public TempSelected As Boolean
        Public TempMouseDown As Boolean
        Public TempIndex As Integer

        Public Function equalsBMSE(ByVal nColumnIndex As Integer, ByVal nVposition As Double,
    ByVal nValue As Long, ByVal nLongNote As Boolean, ByVal nHidden As Boolean) As Boolean
            Return VPosition = nVposition And
               ColumnIndex = nColumnIndex And
               Value = nValue And
               LongNote = nLongNote And
               Hidden = nHidden
        End Function
        Public Function equalsNT(ByVal nColumnIndex As Integer, ByVal nVposition As Double,
    ByVal nValue As Long, ByVal nLength As Double, ByVal nHidden As Boolean) As Boolean
            Return VPosition = nVposition And
               ColumnIndex = nColumnIndex And
               Value = nValue And
               Hidden = nHidden And
               Length = nLength
        End Function
        Public Sub New(ByVal nColumnIndex As Integer, ByVal nVposition As Double,
    ByVal nValue As Integer, ByVal nLongNote As Double, ByVal nHidden As Boolean)
            VPosition = nVposition
            ColumnIndex = nColumnIndex
            Value = nValue
            LongNote = nLongNote
            Length = nLongNote
            Hidden = nHidden
            Landmine = False
        End Sub
        Public Sub New(ByVal nColumnIndex As Integer, ByVal nVposition As Double,
    ByVal nValue As Integer, ByVal nLongNote As Double, ByVal nHidden As Boolean, ByVal nSelected As Boolean)
            Me.New(nColumnIndex, nVposition, nValue, nLongNote, nHidden)
            Selected = nSelected
        End Sub
    End Structure
End Namespace
