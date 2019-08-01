Imports iBMSC.Editor

Public Class UndoRedo
    Public Const opVoid As Byte = 0
    Public Const opAddNote As Byte = 1
    Public Const opRemoveNote As Byte = 2
    Public Const opChangeNote As Byte = 3
    Public Const opMoveNote As Byte = 4
    Public Const opLongNoteModify As Byte = 5
    Public Const opHiddenNoteModify As Byte = 6
    Public Const opRelabelNote As Byte = 7
    Public Const opRemoveAllNotes As Byte = 15
    Public Const opChangeMeasureLength As Byte = 16
    Public Const opChangeTimeSelection As Byte = 17
    Public Const opNT As Byte = 18
    'Public Const opChangeVisibleColumns As Byte = 19
    Public Const opWavAutoincFlag As Byte = 20

    Public Const opNoOperation As Byte = 255

    Private Const trueByte As Byte = 1
    Private Const falseByte As Byte = 0


    Public MustInherit Class LinkedURCmd
        Public [Next] As LinkedURCmd = Nothing
        Public MustOverride Function ofType() As Byte
        Public MustOverride Function toBytes() As Byte()
        'Public MustOverride Sub fromBytes(ByVal b As Byte())
    End Class


    Public Shared Function fromBytes(b() As Byte) As LinkedURCmd
        If b Is Nothing Then Return Nothing
        If b.Length = 0 Then Return Nothing

        Select Case b(0)
            Case opVoid : Return New Void(b)
            Case opAddNote : Return New AddNote(b)
            Case opRemoveNote : Return New RemoveNote(b)
            Case opChangeNote : Return New ChangeNote(b)
            Case opMoveNote : Return New MoveNote(b)
            Case opLongNoteModify : Return New LongNoteModify(b)
            Case opHiddenNoteModify : Return New HiddenNoteModify(b)
            Case opRelabelNote : Return New RelabelNote(b)
            Case opRemoveAllNotes : Return New RemoveAllNotes(b)
            Case opChangeMeasureLength : Return New ChangeMeasureLength(b)
            Case opChangeTimeSelection : Return New ChangeTimeSelection(b)
            Case opNT : Return New NT(b)
                'Case opChangeVisibleColumns : Return New ChangeVisibleColumns(b)
            Case opWavAutoincFlag : Return New WavAutoincFlag(b)
            Case opNoOperation : Return New NoOperation(b)
            Case Else : Return Nothing
        End Select
    End Function


    Public Class Void
        Inherits LinkedURCmd
        '1 = 1
        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opVoid}
        End Function

        Public Sub New()
        End Sub

        Public Sub New(b() As Byte)
        End Sub

        Public Overrides Function ofType() As Byte
            Return opVoid
        End Function
    End Class

    Public MustInherit Class LinkedURNoteCmd
        Inherits LinkedURCmd
        Public note As Note

        Public Sub New()
        End Sub

        Public Sub New(b As Note)
            note = b
        End Sub

        Public Sub New(b() As Byte)
            FromBinaryReader(New BinaryReader(New MemoryStream(b)))
        End Sub

        Public Sub FromBinaryReader(ByRef br As BinaryReader)
            br.ReadByte()
            note.FromBinReader(br)
        End Sub

        Public Sub WriteBinWriter(ByRef bw As BinaryWriter)
            bw.Write(ofType())
            bw.Write(note.ToBytes())
        End Sub

        Public MustOverride Overrides Function ofType() As Byte

        Public Overrides Function toBytes() As Byte()
            Dim ms = New MemoryStream()
            Dim bw As New BinaryWriter(ms)
            WriteBinWriter(bw)

            Return ms.GetBuffer()
        End Function
    End Class

    Public Class AddNote
        Inherits LinkedURNoteCmd

        Public Sub New(_note As Note)
            note = _note
        End Sub

        Public Sub New(b() As Byte)
            MyBase.New(b)
        End Sub

        Public Overrides Function ofType() As Byte
            Return opAddNote
        End Function
    End Class


    Public Class RemoveNote
        Inherits LinkedURNoteCmd

        Public Sub New(_note As Note)
            note = _note
        End Sub

        Public Sub New(b() As Byte)
            MyBase.New(b)
        End Sub

        Public Overrides Function ofType() As Byte
            Return opRemoveNote
        End Function
    End Class


    Public Class ChangeNote
        Inherits LinkedURNoteCmd
        Public NNote As Note

        Public Overrides Function toBytes() As Byte()
            Dim ms = New MemoryStream(MyBase.toBytes)
            Dim bw = New BinaryWriter(ms)
            WriteBinWriter(bw)
            NNote.WriteBinWriter(bw)
            Return ms.GetBuffer()
        End Function

        Public Sub New(b() As Byte)
            Dim br = New BinaryReader(New MemoryStream(b))
            FromBinaryReader(br)
            NNote.FromBinReader(br)
        End Sub

        Public Sub New(note1 As Note, note2 As Note)
            note = note1
            NNote = note2
        End Sub

        Public Overrides Function ofType() As Byte
            Return opChangeNote
        End Function
    End Class


    Public Class MoveNote
        Inherits LinkedURNoteCmd
        Public NColumnIndex As Integer = 0
        Public NVPosition As Double = 0

        Public Overrides Function toBytes() As Byte()
            Dim ms = New MemoryStream()
            Dim bw As New BinaryWriter(ms)
            WriteBinWriter(bw)
            bw.Write(NColumnIndex)
            bw.Write(NVPosition)

            Return ms.GetBuffer()
        End Function

        Public Sub New(b() As Byte)
            Dim br As New BinaryReader(New MemoryStream(b))
            FromBinaryReader(br)
            NColumnIndex = br.ReadInt32()
            NVPosition = br.ReadDouble()
        End Sub

        Public Sub New(_note As Note, _ColIndex As Integer, _VPos As Double)
            note = _note
            NColumnIndex = _ColIndex
            NVPosition = _VPos
        End Sub

        Public Overrides Function ofType() As Byte
            Return opMoveNote
        End Function
    End Class


    Public Class LongNoteModify
        Inherits LinkedURNoteCmd
        Public NVPosition As Double = 0
        Public NLongNote As Double = 0

        Public Overrides Function toBytes() As Byte()
            Dim ms = New MemoryStream()
            Dim bw = New BinaryWriter(ms)
            WriteBinWriter(bw)
            bw.Write(NVPosition)
            bw.Write(NLongNote)

            Return ms.GetBuffer()
        End Function

        Public Sub New(b() As Byte)
            Dim br = New BinaryReader(New MemoryStream(b))
            FromBinaryReader(br)
            NLongNote = br.ReadDouble()
            NVPosition = br.ReadDouble()
        End Sub

        Public Sub New(_note As Note, xNVPosition As Double, xNLongNote As Double)
            note = _note
            NVPosition = xNVPosition
            NLongNote = xNLongNote
        End Sub

        Public Overrides Function ofType() As Byte
            Return opLongNoteModify
        End Function
    End Class


    Public Class HiddenNoteModify
        Inherits LinkedURNoteCmd
        Public NHidden As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim MS = New MemoryStream()
            Dim bw = New BinaryWriter(MS)
            WriteBinWriter(bw)
            bw.Write(NHidden)
            Return MS.GetBuffer()
        End Function

        Public Sub New(b() As Byte)
            Dim br = New BinaryReader(New MemoryStream(b))
            FromBinaryReader(br)
            NHidden = br.ReadBoolean()
        End Sub

        Public Sub New(_note As Note, xNHidden As Boolean)
            note = _note
            NHidden = xNHidden
        End Sub

        Public Overrides Function ofType() As Byte
            Return opHiddenNoteModify
        End Function
    End Class


    Public Class RelabelNote
        Inherits LinkedURNoteCmd
        '1 + 25 + 4 + 1 = 31

        Public NValue As Long = 10000

        Public Overrides Function toBytes() As Byte()
            Dim ms = New MemoryStream()
            Dim bw = New BinaryWriter(ms)
            WriteBinWriter(bw)
            bw.Write(NValue)

            Return ms.GetBuffer()
        End Function

        Public Sub New(b() As Byte)
            Dim br = New BinaryReader(New MemoryStream(b))
            FromBinaryReader(br)
            NValue = br.ReadInt64
        End Sub

        Public Sub New(_note As Note, xNValue As Long)
            note = _note
            NValue = xNValue
        End Sub

        Public Overrides Function ofType() As Byte
            Return opRelabelNote
        End Function
    End Class


    Public Class RemoveAllNotes
        Inherits LinkedURCmd
        '1 = 1
        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opRemoveAllNotes}
        End Function

        Public Sub New(b() As Byte)
        End Sub

        Public Sub New()
        End Sub

        Public Overrides Function ofType() As Byte
            Return opRemoveAllNotes
        End Function
    End Class


    Public Class ChangeMeasureLength
        Inherits LinkedURCmd
        '1 + 8 + 4 + 4 * Indices.Length = 13 + 4 * Indices.Length
        Public Value As Double = 192
        Public Indices() As Integer = {}

        Public Overrides Function toBytes() As Byte()
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xUbound() As Byte = BitConverter.GetBytes(UBound(Indices))
            Dim xToBytes() As Byte = {opChangeMeasureLength,
                                      xVal(0), xVal(1), xVal(2), xVal(3), xVal(4), xVal(5), xVal(6), xVal(7),
                                      xUbound(0), xUbound(1), xUbound(2), xUbound(3)}
            ReDim Preserve xToBytes(12 + 4*Indices.Length)
            For i = 13 To UBound(xToBytes) Step 4
                Dim xId() As Byte = BitConverter.GetBytes(Indices((i - 13)\4))
                xToBytes(i + 0) = xId(0)
                xToBytes(i + 1) = xId(1)
                xToBytes(i + 2) = xId(2)
                xToBytes(i + 3) = xId(3)
            Next
            Return xToBytes
        End Function

        Public Sub New(b() As Byte)
            Value = BitConverter.ToDouble(b, 1)
            Dim xUbound As Integer = BitConverter.ToInt32(b, 9)
            ReDim Preserve Indices(xUbound)
            For i = 13 To xUbound Step 4
                Indices((i - 13)\4) = BitConverter.ToInt32(b, i)
            Next
        End Sub

        Public Sub New(xValue As Double, xIndices() As Integer)
            Value = xValue
            Indices = xIndices
        End Sub

        Public Overrides Function ofType() As Byte
            Return opChangeMeasureLength
        End Function
    End Class


    Public Class ChangeTimeSelection
        Inherits LinkedURCmd
        '1 + 8 + 8 + 8 + 1 = 26
        Public SelStart As Double = 0
        Public SelLength As Double = 0
        Public SelHalf As Double = 0
        Public Selected As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xSta() As Byte = BitConverter.GetBytes(SelStart)
            Dim xLen() As Byte = BitConverter.GetBytes(SelLength)
            Dim xHalf() As Byte = BitConverter.GetBytes(SelLength)
            toBytes = New Byte() {opChangeTimeSelection,
                                  xSta(0), xSta(1), xSta(2), xSta(3), xSta(4), xSta(5), xSta(6), xSta(7),
                                  xLen(0), xLen(1), xLen(2), xLen(3), xLen(4), xLen(5), xLen(6), xLen(7),
                                  xHalf(0), xHalf(1), xHalf(2), xHalf(3), xHalf(4), xHalf(5), xHalf(6), xHalf(7),
                                  IIf(Selected, trueByte, falseByte)}
        End Function

        Public Sub New(b() As Byte)
            SelStart = BitConverter.ToDouble(b, 1)
            SelLength = BitConverter.ToDouble(b, 9)
            SelHalf = BitConverter.ToDouble(b, 17)
            Selected = CBool(b(25))
        End Sub

        Public Sub New(xSelStart As Double, xSelLength As Double, xSelHalf As Double, xSelected As Boolean)
            SelStart = xSelStart
            SelLength = xSelLength
            SelHalf = xSelHalf
            Selected = xSelected
        End Sub

        Public Overrides Function ofType() As Byte
            Return opChangeTimeSelection
        End Function
    End Class


    Public Class NT
        Inherits LinkedURCmd
        '1 + 1 + 1 = 3
        Public BecomeNT As Boolean = False
        Public AutoConvert As Boolean = False

        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opNT,
                                  IIf(BecomeNT, trueByte, falseByte),
                                  IIf(AutoConvert, trueByte, falseByte)}
        End Function

        Public Sub New(b() As Byte)
            BecomeNT = CBool(b(1))
            AutoConvert = CBool(b(2))
        End Sub

        Public Sub New(xBecomeNT As Boolean, xAutoConvert As Boolean)
            BecomeNT = xBecomeNT
            AutoConvert = xAutoConvert
        End Sub

        Public Overrides Function ofType() As Byte
            Return opNT
        End Function
    End Class

    Public Class WavAutoincFlag
        Inherits LinkedURCmd
        Public Checked As Boolean = False

        Public Sub New(_checked As Boolean)
            Checked = _checked
        End Sub

        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opWavAutoincFlag,
                                  IIf(Checked, trueByte, falseByte)}
        End Function

        Public Sub New(b() As Byte)
            Checked = CBool(b(1))
        End Sub

        Public Overrides Function ofType() As Byte
            Return opWavAutoincFlag
        End Function
    End Class


    Public Class NoOperation
        Inherits LinkedURCmd
        '1 = 1
        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opNoOperation}
        End Function

        Public Sub New()
        End Sub

        Public Sub New(b() As Byte)
        End Sub

        Public Overrides Function ofType() As Byte
            Return opNoOperation
        End Function
    End Class
End Class
