Imports System.Text.RegularExpressions

Namespace Editor
    Public Module Functions
        Public Function WriteDecimalWithDot(v As Double) As String
            Static nfi As New System.Globalization.NumberFormatInfo()
            nfi.NumberDecimalSeparator = "."
            Return v.ToString(nfi)
        End Function

        Public Function Add3Zeros(ByVal xNum As Integer) As String
            Dim xStr1 As String = "000" & xNum
            Return Mid(xStr1, Len(xStr1) - 2)
        End Function

        Public Function Add2Zeros(ByVal xNum As Integer) As String
            Dim xStr1 As String = "00" & xNum
            Return Mid(xStr1, Len(xStr1) - 1)
        End Function

        Public Function C10to36S(ByVal xStart As Integer) As Char
            If xStart < 10 Then Return CChar(CStr(xStart)) Else Return Chr(xStart + 55)
        End Function
        Public Function C36to10S(ByVal xChar As Char) As Integer
            Dim xAsc As Integer = Asc(UCase(xChar))
            If xAsc >= 48 And xAsc <= 57 Then
                Return xAsc - 48
            ElseIf xAsc >= 65 And xAsc <= 90 Then
                Return xAsc - 55
            End If
            Return 0
        End Function
        Public Function C10to36(ByVal xStart As Long) As String
            If xStart < 1 Then xStart = 1
            If xStart > 1295 Then xStart = 1295
            Return C10to36S(xStart \ 36) & C10to36S(xStart Mod 36)
        End Function
        Public Function C36to10(ByVal xStart As String) As Integer
            xStart = Mid("00" & xStart, Len(xStart) + 1)
            Return C36to10S(xStart.Chars(0)) * 36 + C36to10S(xStart.Chars(1))
        End Function

        Public Function EncodingToString(TextEncoding As System.Text.Encoding) As String
            If TextEncoding Is System.Text.Encoding.Default Then Return "System Ansi"
            If TextEncoding Is System.Text.Encoding.Unicode Then Return "Little Endian UTF16"
            If TextEncoding Is System.Text.Encoding.ASCII Then Return "ASCII"
            If TextEncoding Is System.Text.Encoding.BigEndianUnicode Then Return "Big Endian UTF16"
            If TextEncoding Is System.Text.Encoding.UTF32 Then Return "Little Endian UTF32"
            If TextEncoding Is System.Text.Encoding.UTF7 Then Return "UTF7"
            If TextEncoding Is System.Text.Encoding.UTF8 Then Return "UTF8"
            If TextEncoding Is System.Text.Encoding.GetEncoding(932) Then Return "SJIS"
            If TextEncoding Is System.Text.Encoding.GetEncoding(51949) Then Return "EUC-KR"
            Return "ANSI (" & TextEncoding.EncodingName & ")" & (TextEncoding Is System.Text.Encoding.Default)
        End Function

        ''' <summary>
        ''' Adjust the brightness of a color.
        ''' </summary>
        ''' <param name="cStart">Original Color.</param>
        ''' <param name="iPercent">(-100 to 100) Brightness.</param>
        ''' <param name="iTransparency">(0 - 1) Transparency.</param>
        Public Function AdjustBrightness(ByVal cStart As Color, ByVal iPercent As Single, ByVal iTransparency As Single) As System.Drawing.Color
            If cStart.A = 0 Then
                Return Color.FromArgb(0)
            Else
                Return Color.FromArgb(
                    cStart.A * iTransparency,
                    cStart.R * (100 - Math.Abs(iPercent)) * 0.01 + Math.Abs(CInt(iPercent >= 0) * iPercent) * 2.55,
                    cStart.G * (100 - Math.Abs(iPercent)) * 0.01 + Math.Abs(CInt(iPercent >= 0) * iPercent) * 2.55,
                    cStart.B * (100 - Math.Abs(iPercent)) * 0.01 + Math.Abs(CInt(iPercent >= 0) * iPercent) * 2.55)
            End If
        End Function

        Public Function IdentifiertoLongNote(ByVal I As String) As Boolean
            Dim xI As Integer = CInt(Val(I))
            Return xI >= 50 And xI < 90
        End Function

        Public Function IdentifiertoHidden(ByVal I As String) As Boolean
            Dim xI As Integer = CInt(Val(I))
            Return (xI >= 30 And xI < 50) Or (xI >= 70 And xI < 90)
        End Function

        Public Function RandomFileName(ByVal extWithDot As String) As String
            Do
                Randomize()
                RandomFileName = Now.Ticks & Mid(Rnd(), 3) & extWithDot
            Loop While File.Exists(RandomFileName) Or Directory.Exists(RandomFileName)
        End Function

        ''' <param name="xH">Hue (0-359)</param>
        ''' <param name="xS">Saturation (0-1000)</param>
        ''' <param name="xL">Lightness (0-1000)</param>
        ''' <param name="xA">Alpha (0-255)</param>
        Public Function HSL2RGB(ByVal xH As Integer, ByVal xS As Integer, ByVal xL As Integer, Optional ByVal xA As Integer = 255) As Color
            If xH > 360 Or xS > 1000 Or xL > 1000 Or xA > 255 Then Return Color.Black

            'Dim xxH As Double = xH
            Dim xxS As Double = xS / 1000
            Dim xxB As Double = (xL - 500) / 500
            Dim xR As Double
            Dim xG As Double
            Dim xB As Double

            If xH < 60 Then
                xB = -1 : xR = 1 : xG = (xH - 30) / 30
            ElseIf xH < 120 Then
                xB = -1 : xG = 1 : xR = (90 - xH) / 30
            ElseIf xH < 180 Then
                xR = -1 : xG = 1 : xB = (xH - 150) / 30
            ElseIf xH < 240 Then
                xR = -1 : xB = 1 : xG = (210 - xH) / 30
            ElseIf xH < 300 Then
                xG = -1 : xB = 1 : xR = (xH - 270) / 30
            Else
                xG = -1 : xR = 1 : xB = (330 - xH) / 30
            End If

            xR = (xR * xxS * (1 - Math.Abs(xxB)) + xxB + 1) * 255 / 2
            xG = (xG * xxS * (1 - Math.Abs(xxB)) + xxB + 1) * 255 / 2
            xB = (xB * xxS * (1 - Math.Abs(xxB)) + xxB + 1) * 255 / 2

            Return Color.FromArgb(xA, xR, xG, xB)
        End Function

        Public Function FontToString(ByVal xFont As Font) As String
            Return xFont.FontFamily.Name & "," & xFont.Size & "," & CInt(xFont.Style)
        End Function

        Public Function isFontInstalled(ByVal f As String) As Boolean
            Dim xFontCollection As New Drawing.Text.InstalledFontCollection
            For Each ff As FontFamily In xFontCollection.Families
                If f.Equals(ff.Name, StringComparison.CurrentCultureIgnoreCase) Then Return True
            Next
            Return False
        End Function


        Public Function StringToFont(ByVal xStr As String, ByVal xDefault As Font) As Font
            Dim xLine() As String = Split(xStr, ",")
            If UBound(xLine) = 2 Then
                Dim xFontStyle As System.Drawing.FontStyle = Val(xLine(2))
                Return New Font(xLine(0), CSng(Val(xLine(1))), xFontStyle, GraphicsUnit.Pixel)
            Else
                Return xDefault
            End If
        End Function

        Public Function ArrayToString(ByVal xInt() As Integer) As String
            Dim xStr As String = ""
            For xI1 As Integer = 0 To UBound(xInt)
                xStr &= xInt(xI1).ToString & IIf(xI1 = UBound(xInt), "", ",")
            Next
            Return xStr
        End Function

        Public Function ArrayToString(ByVal xBool() As Boolean) As String
            Dim xStr As String = ""
            For xI1 As Integer = 0 To UBound(xBool)
                xStr &= CInt(xBool(xI1)).ToString & IIf(xI1 = UBound(xBool), "", ",")
            Next
            Return xStr
        End Function

        Public Function ArrayToString(ByVal xColor() As Color) As String
            Dim xStr As String = ""
            For xI1 As Integer = 0 To UBound(xColor)
                xStr &= xColor(xI1).ToArgb.ToString & IIf(xI1 = UBound(xColor), "", ",")
            Next
            Return xStr
        End Function

        Public Function StringToArrayInt(ByVal xStr As String) As Integer()
            Dim xL() As String = Split(xStr, ",")
            Dim xInt(UBound(xL)) As Integer
            For xI1 As Integer = 0 To UBound(xInt)
                xInt(xI1) = Val(xL(xI1))
            Next
            Return xInt
        End Function

        Public Function StringToArrayBool(ByVal xStr As String) As Boolean()
            Dim xL() As String = Split(xStr, ",")
            Dim xBool(UBound(xL)) As Boolean
            For xI1 As Integer = 0 To UBound(xBool)
                xBool(xI1) = CBool(Val(xL(xI1)))
            Next
            Return xBool
        End Function

        Public Function GetDenominator(ByVal a As Double, Optional ByVal maxDenom As Long = &H7FFFFFFF) As Long
            Dim m00 As Long = 1
            Dim m01 As Long = 0
            Dim m10 As Long = 0
            Dim m11 As Long = 1
            Dim x As Double = a
            Dim ai As Long = Int(x)

            Do While m10 * ai + m11 <= maxDenom
                Dim t As Long
                t = m00 * ai + m01
                m01 = m00
                m00 = t
                t = m10 * ai + m11
                m11 = m10
                m10 = t

                If x = CDbl(ai) Then Exit Do
                x = 1 / (x - ai)

                If x > CDbl(&H7FFFFFFFFFFFFFFF) Then Exit Do
                ai = Int(x)
            Loop

            Return m10
        End Function


        Public Function IsBase36(str As String) As Boolean
            Static re As New Regex("^[A-Za-z0-9]+$")
            Return re.IsMatch(str)
        End Function

    End Module
End Namespace
