Public Class Grid
    Public Property IsSnapEnabled As Boolean = True
    Public Property ShowMainGrid As Boolean = True
    Public Property ShowSubGrid As Boolean = True
    Public Property ShowBackground As Boolean = True
    Public Property ShowMeasureNumber As Boolean = True
    Public Property ShowVerticalLines As Boolean = True
    Public Property ShowMeasureBars As Boolean = True
    Public Property ShowColumnCaptions As Boolean = True
    Public Property Divider As Integer = 16
    Public Property Subdivider As Integer = 4
    Public Property Slash As Integer = 192
    Public Property HeightScale As Single = 1.0!
    Public Property WidthScale As Single = 1.0!
    Public Property WheelScroll As Integer = 96
    Public Property PageUpDnScroll As Integer = 384
    Public Property ShowBgaColumn As Boolean = True
    Public Property ShowScrollColumn As Boolean = True
    Public Property ShowStopColumn As Boolean = True
    Public Property ShowBpmColumn As Boolean = True
    Public Property IPlayer As Integer = 0
End Class
