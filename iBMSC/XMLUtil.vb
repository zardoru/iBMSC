Imports System.Globalization

Module XMLUtil
    Public Sub XMLWriteValue(w As XmlTextWriter, local As String, val As String)
        w.WriteStartElement(local)
        w.WriteAttributeString("Value", Editor.WriteDecimalWithDot(val))
        w.WriteEndElement()
    End Sub

    Public Sub XMLLoadAttribute(s As String, ByRef v As String)
        If s.Length = 0 Then Exit Sub
        v = s
    End Sub

    Public Sub XMLLoadAttribute(s As String, ByRef v As Boolean)
        If s.Length = 0 Then Exit Sub
        v = CBool(s)
    End Sub

    Public Sub XMLLoadAttribute(s As String, ByRef v As Integer)
        If s.Length = 0 Then Exit Sub
        v = CInt(s)
    End Sub

    Public Sub XMLLoadAttribute(s As String, ByRef v As Single)
        If s.Length = 0 Then Exit Sub

        Static nfi As New NumberFormatInfo()
        nfi.NumberDecimalSeparator = "."
        v = Double.Parse(s, nfi)
    End Sub

    Public Sub XMLLoadAttribute(s As String, ByRef v As Double)
        If s.Length = 0 Then Exit Sub

        Static nfi As New NumberFormatInfo()
        nfi.NumberDecimalSeparator = "."
        v = Double.Parse(s, nfi)
    End Sub

    Public Sub XMLLoadAttribute(s As String, ByRef v As Long)
        If s.Length = 0 Then Exit Sub
        v = CLng(s)
    End Sub

    Public Sub XMLLoadAttribute(s As String, ByRef v As Decimal)
        If s.Length = 0 Then Exit Sub
        v = CDec(s)
    End Sub

    Public Sub XMLLoadAttribute(s As String, ByRef v As Color)
        If s.Length = 0 Then Exit Sub
        v = Color.FromArgb(CInt(s))
    End Sub

    Public Sub XMLLoadElementValue(n As XmlElement, ByRef v As Integer)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub

    Public Sub XMLLoadElementValue(n As XmlElement, ByRef v As Single)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub

    Public Sub XMLLoadElementValue(n As XmlElement, ByRef v As Color)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub
End Module
