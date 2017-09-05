Namespace Editor
    Public Structure Note
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

        Public Function equalsBMSE(note As Note) As Boolean
            Return VPosition = note.VPosition And
               ColumnIndex = note.ColumnIndex And
               Value = note.Value And
               LongNote = note.LongNote And
               Hidden = note.Hidden And
               Landmine = note.Landmine
        End Function
        Public Function equalsNT(note As Note) As Boolean
            Return VPosition = note.VPosition And
               ColumnIndex = note.ColumnIndex And
               Value = note.Value And
               Hidden = note.Hidden And
               Length = note.Length And
               Landmine = note.Landmine
        End Function

        Public Sub New(nColumnIndex As Integer,
                       nVposition As Double,
                       nValue As Long,
                       Optional nLongNote As Double = 0,
                       Optional nHidden As Boolean = False,
                       Optional nSelected As Boolean = False,
                       Optional nLandmine As Boolean = False)
            VPosition = nVposition
            ColumnIndex = nColumnIndex
            Value = nValue
            LongNote = nLongNote
            Length = nLongNote
            Hidden = nHidden
            Landmine = nLandmine
        End Sub

        Friend Function ToBytes() As Byte()
            Dim MS As New MemoryStream()
            Dim bw As New BinaryWriter(MS)
            WriteBinWriter(bw)

            Return MS.GetBuffer()
        End Function

        Friend Sub WriteBinWriter(ByRef bw As BinaryWriter)
            bw.Write(VPosition)
            bw.Write(ColumnIndex)
            bw.Write(Value)
            bw.Write(LongNote)
            bw.Write(Length)
            bw.Write(Hidden)
            bw.Write(Landmine)
        End Sub

        Friend Sub FromBinReader(ByRef br As BinaryReader)
            VPosition = br.ReadDouble()
            ColumnIndex = br.ReadInt32()
            Value = br.ReadInt64()
            LongNote = br.ReadBoolean()
            Length = br.ReadDouble()
            Hidden = br.ReadBoolean()
            Landmine = br.ReadBoolean()
        End Sub

        Friend Sub FromBytes(ByRef bytes() As Byte)
            Dim br As New BinaryReader(New MemoryStream(bytes))
            FromBinReader(br)
        End Sub
    End Structure
End Namespace
