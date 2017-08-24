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
    Public Const opNoOperation As Byte = 255

    Private Const trueByte As Byte = 1
    Private Const falseByte As Byte = 0



    Public MustInherit Class LinkedURCmd
        Public [Next] As LinkedURCmd = Nothing
        Public MustOverride Function ofType() As Byte
        Public MustOverride Function toBytes() As Byte()
        'Public MustOverride Sub fromBytes(ByVal b As Byte())
    End Class



    Public Shared Function fromBytes(ByVal b() As Byte) As LinkedURCmd
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
            Case opNoOperation : Return New NoOperation(b)
            Case Else : Return Nothing
        End Select
    End Function


    Public Class Void : Inherits LinkedURCmd
        '1 = 1
        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opVoid}
        End Function

        Public Sub New()
        End Sub

        Public Sub New(ByVal b() As Byte)
        End Sub

        Public Overrides Function ofType() As Byte
            Return opVoid
        End Function
    End Class



    Public Class AddNote : Inherits LinkedURCmd
        '1 + 25 + 1 = 27
        Public ColumnIndex As Integer = 0
        Public VPosition As Double = 0
        Public Value As Long = 10000
        Public LongNote As Double = 0
        Public Hidden As Boolean = False
        Public Selected As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xCol() As Byte = BitConverter.GetBytes(ColumnIndex)
            Dim xVPos() As Byte = BitConverter.GetBytes(VPosition)
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xLong() As Byte = BitConverter.GetBytes(LongNote)
            toBytes = New Byte() {opAddNote, _
                                  xCol(0), xCol(1), xCol(2), xCol(3), _
                                  xVPos(0), xVPos(1), xVPos(2), xVPos(3), xVPos(4), xVPos(5), xVPos(6), xVPos(7), _
                                  xVal(0), xVal(1), xVal(2), xVal(3), _
                                  xLong(0), xLong(1), xLong(2), xLong(3), xLong(4), xLong(5), xLong(6), xLong(7), _
                                  IIf(Hidden, trueByte, falseByte), _
                                  IIf(Selected, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            ColumnIndex = BitConverter.ToInt32(b, 1)
            VPosition = BitConverter.ToDouble(b, 5)
            Value = BitConverter.ToInt64(b, 13)
            LongNote = BitConverter.ToDouble(b, 21)
            Hidden = CBool(b(29))
            Selected = CBool(b(30))
        End Sub

        Public Sub New(ByVal xColumnIndex As Integer, ByVal xVPosition As Double, ByVal xValue As Long, ByVal xLongNote As Double, ByVal xHidden As Boolean,
        ByVal xSelected As Boolean)
            ColumnIndex = xColumnIndex
            VPosition = xVPosition
            Value = xValue
            LongNote = xLongNote
            Hidden = xHidden
            Selected = xSelected
        End Sub

        Public Overrides Function ofType() As Byte
            Return opAddNote
        End Function
    End Class



    Public Class RemoveNote : Inherits LinkedURCmd
        '1 + 25 = 26
        Public ColumnIndex As Integer = 0
        Public VPosition As Double = 0
        Public Value As Long = 10000
        Public LongNote As Double = 0
        Public Hidden As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xCol() As Byte = BitConverter.GetBytes(ColumnIndex)
            Dim xVPos() As Byte = BitConverter.GetBytes(VPosition)
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xLong() As Byte = BitConverter.GetBytes(LongNote)
            toBytes = New Byte() {opRemoveNote, _
                                  xCol(0), xCol(1), xCol(2), xCol(3), _
                                  xVPos(0), xVPos(1), xVPos(2), xVPos(3), xVPos(4), xVPos(5), xVPos(6), xVPos(7), _
                                  xVal(0), xVal(1), xVal(2), xVal(3), _
                                  xLong(0), xLong(1), xLong(2), xLong(3), xLong(4), xLong(5), xLong(6), xLong(7), _
                                  IIf(Hidden, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            ColumnIndex = BitConverter.ToInt32(b, 1)
            VPosition = BitConverter.ToDouble(b, 5)
            Value = BitConverter.ToInt64(b, 13)
            LongNote = BitConverter.ToDouble(b, 21)
            Hidden = CBool(b(29))
        End Sub

        Public Sub New(ByVal xColumnIndex As Integer, ByVal xVPosition As Double, ByVal xValue As Long, ByVal xLongNote As Double, ByVal xHidden As Boolean)
            ColumnIndex = xColumnIndex
            VPosition = xVPosition
            Value = xValue
            LongNote = xLongNote
            Hidden = xHidden
        End Sub

        Public Overrides Function ofType() As Byte
            Return opRemoveNote
        End Function
    End Class



    Public Class ChangeNote : Inherits LinkedURCmd
        '1 + 25 + 25 + 1 = 52
        Public ColumnIndex As Integer = 0
        Public VPosition As Double = 0
        Public Value As Integer = 10000
        Public LongNote As Double = 0
        Public Hidden As Boolean = False

        Public NColumnIndex As Integer = 0
        Public NVPosition As Double = 0
        Public NValue As Integer = 10000
        Public NLongNote As Double = 0
        Public NHidden As Boolean = False

        Public Selected As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xCol() As Byte = BitConverter.GetBytes(ColumnIndex)
            Dim xVPos() As Byte = BitConverter.GetBytes(VPosition)
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xLong() As Byte = BitConverter.GetBytes(LongNote)
            Dim xNCol() As Byte = BitConverter.GetBytes(NColumnIndex)
            Dim xNVPos() As Byte = BitConverter.GetBytes(NVPosition)
            Dim xNVal() As Byte = BitConverter.GetBytes(NValue)
            Dim xNLong() As Byte = BitConverter.GetBytes(NLongNote)
            toBytes = New Byte() {opChangeNote, _
                                  xCol(0), xCol(1), xCol(2), xCol(3), _
                                  xVPos(0), xVPos(1), xVPos(2), xVPos(3), xVPos(4), xVPos(5), xVPos(6), xVPos(7), _
                                  xVal(0), xVal(1), xVal(2), xVal(3), _
                                  xLong(0), xLong(1), xLong(2), xLong(3), xLong(4), xLong(5), xLong(6), xLong(7), _
                                  IIf(Hidden, trueByte, falseByte), _
                                  xNCol(0), xNCol(1), xNCol(2), xNCol(3), _
                                  xNVPos(0), xNVPos(1), xNVPos(2), xNVPos(3), xNVPos(4), xNVPos(5), xNVPos(6), xNVPos(7), _
                                  xNVal(0), xNVal(1), xNVal(2), xNVal(3), _
                                  xNLong(0), xNLong(1), xNLong(2), xNLong(3), xNLong(4), xNLong(5), xNLong(6), xNLong(7), _
                                  IIf(NHidden, trueByte, falseByte), _
                                  IIf(Selected, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            ColumnIndex = BitConverter.ToInt32(b, 1)
            VPosition = BitConverter.ToDouble(b, 5)
            Value = BitConverter.ToInt32(b, 13)
            LongNote = BitConverter.ToDouble(b, 17)
            Hidden = CBool(b(25))
            NColumnIndex = BitConverter.ToInt32(b, 26)
            NVPosition = BitConverter.ToDouble(b, 30)
            NValue = BitConverter.ToInt32(b, 38)
            NLongNote = BitConverter.ToDouble(b, 42)
            NHidden = CBool(b(50))
            Selected = CBool(b(51))
        End Sub

        Public Sub New(ByVal xColumnIndex As Integer, ByVal xVPosition As Double, ByVal xValue As Integer, ByVal xLongNote As Double, ByVal xHidden As Boolean, _
        ByVal xNColumnIndex As Integer, ByVal xNVPosition As Double, ByVal xNValue As Integer, ByVal xNLongNote As Double, ByVal xNHidden As Boolean, ByVal xSelected As Boolean)
            ColumnIndex = xColumnIndex
            VPosition = xVPosition
            Value = xValue
            LongNote = xLongNote
            Hidden = xHidden
            NColumnIndex = xNColumnIndex
            NVPosition = xNVPosition
            NValue = xNValue
            NLongNote = xNLongNote
            NHidden = xNHidden
            Selected = xSelected
        End Sub

        Public Overrides Function ofType() As Byte
            Return opChangeNote
        End Function
    End Class



    Public Class MoveNote : Inherits LinkedURCmd
        '1 + 25 + 4 + 8 + 1 = 39
        Public ColumnIndex As Integer = 0
        Public VPosition As Double = 0
        Public Value As Integer = 10000
        Public LongNote As Double = 0
        Public Hidden As Boolean = False

        Public NColumnIndex As Integer = 0
        Public NVPosition As Double = 0

        Public Selected As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xCol() As Byte = BitConverter.GetBytes(ColumnIndex)
            Dim xVPos() As Byte = BitConverter.GetBytes(VPosition)
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xLong() As Byte = BitConverter.GetBytes(LongNote)
            Dim xNCol() As Byte = BitConverter.GetBytes(NColumnIndex)
            Dim xNVPos() As Byte = BitConverter.GetBytes(NVPosition)
            toBytes = New Byte() {opMoveNote, _
                                  xCol(0), xCol(1), xCol(2), xCol(3), _
                                  xVPos(0), xVPos(1), xVPos(2), xVPos(3), xVPos(4), xVPos(5), xVPos(6), xVPos(7), _
                                  xVal(0), xVal(1), xVal(2), xVal(3), _
                                  xLong(0), xLong(1), xLong(2), xLong(3), xLong(4), xLong(5), xLong(6), xLong(7), _
                                  IIf(Hidden, trueByte, falseByte), _
                                  xNCol(0), xNCol(1), xNCol(2), xNCol(3), _
                                  xNVPos(0), xNVPos(1), xNVPos(2), xNVPos(3), xNVPos(4), xNVPos(5), xNVPos(6), xNVPos(7), _
                                  IIf(Selected, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            ColumnIndex = BitConverter.ToInt32(b, 1)
            VPosition = BitConverter.ToDouble(b, 5)
            Value = BitConverter.ToInt32(b, 13)
            LongNote = BitConverter.ToDouble(b, 17)
            Hidden = CBool(b(25))
            NColumnIndex = BitConverter.ToInt32(b, 26)
            NVPosition = BitConverter.ToDouble(b, 30)
            Selected = CBool(b(38))
        End Sub

        Public Sub New(ByVal xColumnIndex As Integer, ByVal xVPosition As Double, ByVal xValue As Integer, ByVal xLongNote As Double, ByVal xHidden As Boolean, _
        ByVal xNColumnIndex As Integer, ByVal xNVPosition As Double, ByVal xSelected As Boolean)
            ColumnIndex = xColumnIndex
            VPosition = xVPosition
            Value = xValue
            LongNote = xLongNote
            Hidden = xHidden
            NColumnIndex = xNColumnIndex
            NVPosition = xNVPosition
            Selected = xSelected
        End Sub

        Public Overrides Function ofType() As Byte
            Return opMoveNote
        End Function
    End Class



    Public Class LongNoteModify : Inherits LinkedURCmd
        '1 + 25 + 8 + 8 + 1 = 43
        Public ColumnIndex As Integer = 0
        Public VPosition As Double = 0
        Public Value As Integer = 10000
        Public LongNote As Double = 0
        Public Hidden As Boolean = False

        Public NVPosition As Double = 0
        Public NLongNote As Double = 0

        Public Selected As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xCol() As Byte = BitConverter.GetBytes(ColumnIndex)
            Dim xVPos() As Byte = BitConverter.GetBytes(VPosition)
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xLong() As Byte = BitConverter.GetBytes(LongNote)
            Dim xNVPos() As Byte = BitConverter.GetBytes(NVPosition)
            Dim xNLong() As Byte = BitConverter.GetBytes(NLongNote)
            toBytes = New Byte() {opLongNoteModify, _
                                  xCol(0), xCol(1), xCol(2), xCol(3), _
                                  xVPos(0), xVPos(1), xVPos(2), xVPos(3), xVPos(4), xVPos(5), xVPos(6), xVPos(7), _
                                  xVal(0), xVal(1), xVal(2), xVal(3), _
                                  xLong(0), xLong(1), xLong(2), xLong(3), xLong(4), xLong(5), xLong(6), xLong(7), _
                                  IIf(Hidden, trueByte, falseByte), _
                                  xNLong(0), xNLong(1), xNLong(2), xNLong(3), xNLong(4), xNLong(5), xNLong(6), xNLong(7), _
                                  xNVPos(0), xNVPos(1), xNVPos(2), xNVPos(3), xNVPos(4), xNVPos(5), xNVPos(6), xNVPos(7), _
                                  IIf(Selected, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            ColumnIndex = BitConverter.ToInt32(b, 1)
            VPosition = BitConverter.ToDouble(b, 5)
            Value = BitConverter.ToInt32(b, 13)
            LongNote = BitConverter.ToDouble(b, 17)
            Hidden = CBool(b(25))
            NLongNote = BitConverter.ToDouble(b, 26)
            NVPosition = BitConverter.ToDouble(b, 34)
            Selected = CBool(b(35))
        End Sub

        Public Sub New(ByVal xColumnIndex As Integer, ByVal xVPosition As Double, ByVal xValue As Integer, ByVal xLongNote As Double, ByVal xHidden As Boolean, _
        ByVal xNVPosition As Double, ByVal xNLongNote As Double, ByVal xSelected As Boolean)
            ColumnIndex = xColumnIndex
            VPosition = xVPosition
            Value = xValue
            LongNote = xLongNote
            Hidden = xHidden
            NVPosition = xNVPosition
            NLongNote = xNLongNote
            Selected = xSelected
        End Sub

        Public Overrides Function ofType() As Byte
            Return opLongNoteModify
        End Function
    End Class



    Public Class HiddenNoteModify : Inherits LinkedURCmd
        '1 + 25 + 1 + 1 = 28
        Public ColumnIndex As Integer = 0
        Public VPosition As Double = 0
        Public Value As Integer = 10000
        Public LongNote As Double = 0
        Public Hidden As Boolean = False

        Public NHidden As Boolean = False

        Public Selected As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xCol() As Byte = BitConverter.GetBytes(ColumnIndex)
            Dim xVPos() As Byte = BitConverter.GetBytes(VPosition)
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xLong() As Byte = BitConverter.GetBytes(LongNote)
            toBytes = New Byte() {opHiddenNoteModify, _
                                  xCol(0), xCol(1), xCol(2), xCol(3), _
                                  xVPos(0), xVPos(1), xVPos(2), xVPos(3), xVPos(4), xVPos(5), xVPos(6), xVPos(7), _
                                  xVal(0), xVal(1), xVal(2), xVal(3), _
                                  xLong(0), xLong(1), xLong(2), xLong(3), xLong(4), xLong(5), xLong(6), xLong(7), _
                                  IIf(Hidden, trueByte, falseByte), _
                                  IIf(NHidden, trueByte, falseByte), _
                                  IIf(Selected, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            ColumnIndex = BitConverter.ToInt32(b, 1)
            VPosition = BitConverter.ToDouble(b, 5)
            Value = BitConverter.ToInt32(b, 13)
            LongNote = BitConverter.ToDouble(b, 17)
            Hidden = CBool(b(25))
            NHidden = CBool(b(26))
            Selected = CBool(b(27))
        End Sub

        Public Sub New(ByVal xColumnIndex As Integer, ByVal xVPosition As Double, ByVal xValue As Integer, ByVal xLongNote As Double, ByVal xHidden As Boolean, _
        ByVal xNHidden As Boolean, ByVal xSelected As Boolean)
            ColumnIndex = xColumnIndex
            VPosition = xVPosition
            Value = xValue
            LongNote = xLongNote
            Hidden = xHidden
            NHidden = xNHidden
            Selected = xSelected
        End Sub

        Public Overrides Function ofType() As Byte
            Return opHiddenNoteModify
        End Function
    End Class



    Public Class RelabelNote : Inherits LinkedURCmd
        '1 + 25 + 4 + 1 = 31
        Public ColumnIndex As Integer = 0
        Public VPosition As Double = 0
        Public Value As Integer = 10000
        Public LongNote As Double = 0
        Public Hidden As Boolean = False

        Public NValue As Integer = 10000

        Public Selected As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xCol() As Byte = BitConverter.GetBytes(ColumnIndex)
            Dim xVPos() As Byte = BitConverter.GetBytes(VPosition)
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xLong() As Byte = BitConverter.GetBytes(LongNote)
            Dim xNVal() As Byte = BitConverter.GetBytes(NValue)
            toBytes = New Byte() {opRelabelNote, _
                                  xCol(0), xCol(1), xCol(2), xCol(3), _
                                  xVPos(0), xVPos(1), xVPos(2), xVPos(3), xVPos(4), xVPos(5), xVPos(6), xVPos(7), _
                                  xVal(0), xVal(1), xVal(2), xVal(3), _
                                  xLong(0), xLong(1), xLong(2), xLong(3), xLong(4), xLong(5), xLong(6), xLong(7), _
                                  IIf(Hidden, trueByte, falseByte), _
                                  xNVal(0), xNVal(1), xNVal(2), xNVal(3), _
                                  IIf(Selected, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            ColumnIndex = BitConverter.ToInt32(b, 1)
            VPosition = BitConverter.ToDouble(b, 5)
            Value = BitConverter.ToInt32(b, 13)
            LongNote = BitConverter.ToDouble(b, 17)
            Hidden = CBool(b(25))
            NValue = BitConverter.ToInt32(b, 26)
            Selected = CBool(b(30))
        End Sub

        Public Sub New(ByVal xColumnIndex As Integer, ByVal xVPosition As Double, ByVal xValue As Integer, ByVal xLongNote As Double, ByVal xHidden As Boolean, _
        ByVal xNValue As Integer, ByVal xSelected As Boolean)
            ColumnIndex = xColumnIndex
            VPosition = xVPosition
            Value = xValue
            LongNote = xLongNote
            Hidden = xHidden
            NValue = xNValue
            Selected = xSelected
        End Sub

        Public Overrides Function ofType() As Byte
            Return opRelabelNote
        End Function
    End Class



    Public Class RemoveAllNotes : Inherits LinkedURCmd
        '1 = 1
        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opRemoveAllNotes}
        End Function

        Public Sub New(ByVal b() As Byte)
        End Sub

        Public Sub New()
        End Sub

        Public Overrides Function ofType() As Byte
            Return opRemoveAllNotes
        End Function
    End Class



    Public Class ChangeMeasureLength : Inherits LinkedURCmd
        '1 + 8 + 4 + 4 * Indices.Length = 13 + 4 * Indices.Length
        Public Value As Double = 192
        Public Indices() As Integer = {}

        Public Overrides Function toBytes() As Byte()
            Dim xVal() As Byte = BitConverter.GetBytes(Value)
            Dim xUbound() As Byte = BitConverter.GetBytes(UBound(Indices))
            Dim xToBytes() As Byte = {opChangeMeasureLength, _
                                      xVal(0), xVal(1), xVal(2), xVal(3), xVal(4), xVal(5), xVal(6), xVal(7), _
                                      xUbound(0), xUbound(1), xUbound(2), xUbound(3)}
            ReDim Preserve xToBytes(12 + 4 * Indices.Length)
            For xI1 As Integer = 13 To UBound(xToBytes) Step 4
                Dim xId() As Byte = BitConverter.GetBytes(Indices((xI1 - 13) \ 4))
                xToBytes(xI1 + 0) = xId(0)
                xToBytes(xI1 + 1) = xId(1)
                xToBytes(xI1 + 2) = xId(2)
                xToBytes(xI1 + 3) = xId(3)
            Next
            Return xToBytes
        End Function

        Public Sub New(ByVal b() As Byte)
            Value = BitConverter.ToDouble(b, 1)
            Dim xUbound As Integer = BitConverter.ToInt32(b, 9)
            ReDim Preserve Indices(xUbound)
            For xI1 As Integer = 13 To xUbound Step 4
                Indices((xI1 - 13) \ 4) = BitConverter.ToInt32(b, xI1)
            Next
        End Sub

        Public Sub New(ByVal xValue As Double, ByVal xIndices() As Integer)
            Value = xValue
            Indices = xIndices
        End Sub

        Public Overrides Function ofType() As Byte
            Return opChangeMeasureLength
        End Function
    End Class



    Public Class ChangeTimeSelection : Inherits LinkedURCmd
        '1 + 8 + 8 + 8 + 1 = 26
        Public SelStart As Double = 0
        Public SelLength As Double = 0
        Public SelHalf As Double = 0
        Public Selected As Boolean = False

        Public Overrides Function toBytes() As Byte()
            Dim xSta() As Byte = BitConverter.GetBytes(SelStart)
            Dim xLen() As Byte = BitConverter.GetBytes(SelLength)
            Dim xHalf() As Byte = BitConverter.GetBytes(SelLength)
            toBytes = New Byte() {opChangeTimeSelection, _
                                  xSta(0), xSta(1), xSta(2), xSta(3), xSta(4), xSta(5), xSta(6), xSta(7), _
                                  xLen(0), xLen(1), xLen(2), xLen(3), xLen(4), xLen(5), xLen(6), xLen(7), _
                                  xHalf(0), xHalf(1), xHalf(2), xHalf(3), xHalf(4), xHalf(5), xHalf(6), xHalf(7), _
                                  IIf(Selected, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            SelStart = BitConverter.ToDouble(b, 1)
            SelLength = BitConverter.ToDouble(b, 9)
            SelHalf = BitConverter.ToDouble(b, 17)
            Selected = CBool(b(25))
        End Sub

        Public Sub New(ByVal xSelStart As Double, ByVal xSelLength As Double, ByVal xSelHalf As Double, ByVal xSelected As Boolean)
            SelStart = xSelStart
            SelLength = xSelLength
            SelHalf = xSelHalf
            Selected = xSelected
        End Sub

        Public Overrides Function ofType() As Byte
            Return opChangeTimeSelection
        End Function
    End Class



    Public Class NT : Inherits LinkedURCmd
        '1 + 1 + 1 = 3
        Public BecomeNT As Boolean = False
        Public AutoConvert As Boolean = False

        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opNT, _
                                  IIf(BecomeNT, trueByte, falseByte), _
                                  IIf(AutoConvert, trueByte, falseByte)}
        End Function

        Public Sub New(ByVal b() As Byte)
            BecomeNT = CBool(b(1))
            AutoConvert = CBool(b(2))
        End Sub

        Public Sub New(ByVal xBecomeNT As Boolean, ByVal xAutoConvert As Boolean)
            BecomeNT = xBecomeNT
            AutoConvert = xAutoConvert
        End Sub

        Public Overrides Function ofType() As Byte
            Return opNT
        End Function
    End Class



    'Public Class ChangeVisibleColumns : Inherits LinkedURCmd
    '    '1 + 1 + 1 + 4 = 7
    '    Public CGBLP As Boolean = False
    '    Public CGSTOP As Boolean = False
    '    Public CHPlayer As Integer = 0
    '
    '    Public Overrides Function toBytes() As Byte()
    '        Dim xCHP() As Byte = BitConverter.GetBytes(CHPlayer)
    '        toBytes = New Byte() {opChangeVisibleColumns, _
    '                              IIf(CGBLP, trueByte, falseByte), _
    '                              IIf(CGSTOP, trueByte, falseByte), _
    '                              xCHP(0), xCHP(1), xCHP(2), xCHP(3)}
    '    End Function
    '
    '    Public Sub New(ByVal b() As Byte)
    '        CGBLP = CBool(b(1))
    '        CGSTOP = CBool(b(2))
    '        CHPlayer = BitConverter.ToInt32(b, 3)
    '    End Sub
    '
    '    Public Sub New(ByVal xCGBLP As Boolean, ByVal xCGSTOP As Boolean, ByVal xCHPlayer As Integer)
    '        CGBLP = xCGBLP
    '        CGSTOP = xCGSTOP
    '        CHPlayer = xCHPlayer
    '    End Sub
    '
    '    Public Overrides Function ofType() As Byte
    '        Return opChangeVisibleColumns
    '    End Function
    'End Class



    Public Class NoOperation : Inherits LinkedURCmd
        '1 = 1
        Public Overrides Function toBytes() As Byte()
            toBytes = New Byte() {opNoOperation}
        End Function

        Public Sub New()
        End Sub

        Public Sub New(ByVal b() As Byte)
        End Sub

        Public Overrides Function ofType() As Byte
            Return opNoOperation
        End Function
    End Class
End Class
