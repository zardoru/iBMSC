Module XMLUtil
    Public Sub XMLWriteValue(ByVal w As XmlTextWriter, ByVal local As String, ByVal val As String)
        w.WriteStartElement(local)
        w.WriteAttributeString("Value", val)
        w.WriteEndElement()
    End Sub

    Public Sub XMLLoadAttribute(ByVal s As String, ByRef v As String)
        If s.Length = 0 Then Exit Sub
        v = s
    End Sub
    Public Sub XMLLoadAttribute(ByVal s As String, ByRef v As Boolean)
        If s.Length = 0 Then Exit Sub
        v = CBool(s)
    End Sub
    Public Sub XMLLoadAttribute(ByVal s As String, ByRef v As Integer)
        If s.Length = 0 Then Exit Sub
        v = CInt(s)
    End Sub
    Public Sub XMLLoadAttribute(ByVal s As String, ByRef v As Single)
        If s.Length = 0 Then Exit Sub
        v = CSng(s)
    End Sub
    Public Sub XMLLoadAttribute(ByVal s As String, ByRef v As Double)
        If s.Length = 0 Then Exit Sub
        v = CDbl(s)
    End Sub
    Public Sub XMLLoadAttribute(ByVal s As String, ByRef v As Long)
        If s.Length = 0 Then Exit Sub
        v = CLng(s)
    End Sub
    Public Sub XMLLoadAttribute(ByVal s As String, ByRef v As Decimal)
        If s.Length = 0 Then Exit Sub
        v = CDec(s)
    End Sub
    Public Sub XMLLoadAttribute(ByVal s As String, ByRef v As Color)
        If s.Length = 0 Then Exit Sub
        v = Color.FromArgb(CInt(s))
    End Sub

    Public Sub XMLLoadElementValue(ByVal n As XmlElement, ByRef v As Integer)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub
    Public Sub XMLLoadElementValue(ByVal n As XmlElement, ByRef v As Single)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub
    Public Sub XMLLoadElementValue(ByVal n As XmlElement, ByRef v As Color)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub
End Module
