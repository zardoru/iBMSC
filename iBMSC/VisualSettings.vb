Namespace Editor

    Public Class visualSettings
        Public ColumnTitle As SolidBrush
        Public ColumnTitleFont As Font
        Public Bg As SolidBrush
        Public pGrid As Pen
        Public pSub As Pen
        Public pVLine As Pen
        Public pMLine As Pen
        Public pBGMWav As Pen

        Public SelBox As Pen
        Public PECursor As Pen
        Public PEHalf As Pen
        Public PEDeltaMouseOver As Integer
        Public PEMouseOver As Pen
        Public PESel As SolidBrush
        Public PEBPM As SolidBrush
        Public PEBPMFont As Font
        Public MiddleDeltaRelease As Integer

        Public kHeight As Integer
        Public kFont As Font
        Public kMFont As Font
        Public kLabelVShift As Integer
        Public kLabelHShift As Integer
        Public kLabelHShiftL As Integer
        Public kMouseOver As Pen
        Public kMouseOverE As Pen
        Public kSelected As Pen
        Public kOpacity As Single

        Public Sub New()
            Me.New(New SolidBrush(Color.Lime),
                    New Font("Tahoma", 11, FontStyle.Regular, GraphicsUnit.Pixel),
                    New SolidBrush(Color.Black),
                    New Pen(Color.FromArgb(893008442)),
                    New Pen(Color.FromArgb(1530542650)),
                    New Pen(Color.FromArgb(-13158601)),
                    New Pen(Color.FromArgb(1599230546)),
                    New Pen(Color.FromArgb(851493056)),
                    New Pen(Color.FromArgb(-1056964609)),
                    New Pen(Color.FromArgb(&HC0FF8080)),
                    New Pen(Color.FromArgb(&H808080FF)),
                    5,
                    New Pen(Color.FromArgb(&H80FF8080)),
                    New SolidBrush(Color.FromArgb(855605376)),
                    New SolidBrush(Color.FromArgb(855605376)),
                    New Font("Verdana", 12, FontStyle.Bold, GraphicsUnit.Pixel),
                    10,
                    10,
                    New Font("Verdana", 12, FontStyle.Bold, GraphicsUnit.Pixel),
                    New Font("Verdana", 12, FontStyle.Regular, GraphicsUnit.Pixel),
                    -2,
                    0,
                    2,
                    New Pen(Color.Lime),
                    New Pen(Color.FromArgb(-16711681)),
                    New Pen(Color.Red),
                    0.5)
        End Sub

        Public Sub New(
    ByVal voTitle As SolidBrush,
    ByVal voTitleFont As Font,
    ByVal voBg As SolidBrush,
    ByVal voGrid As Pen,
    ByVal voSub As Pen,
    ByVal voVLine As Pen,
    ByVal voMLine As Pen,
    ByVal voBGMWav As Pen, _
 _
    ByVal voSelBox As Pen,
    ByVal voPECursor As Pen,
    ByVal voPEHalf As Pen,
    ByVal voPEDeltaMouseOver As Integer,
    ByVal voPEMouseOver As Pen,
    ByVal voPESel As SolidBrush,
    ByVal voPEBPM As SolidBrush,
    ByVal voPEBPMFont As Font,
    ByVal xMiddleDeltaRelease As Integer, _
 _
    ByVal vKHeight As Integer,
    ByVal vKFont As Font,
    ByVal vKMFont As Font,
    ByVal vKLabelVShift As Integer,
    ByVal vKLabelHShift As Integer,
    ByVal vKLabelHShiftL As Integer,
    ByVal vKMouseOver As Pen,
    ByVal vKMouseOverE As Pen,
    ByVal vKSelected As Pen,
    ByVal vKOpacity As Single)

            ColumnTitle = voTitle
            ColumnTitleFont = voTitleFont
            Bg = voBg
            pGrid = voGrid
            pSub = voSub
            pVLine = voVLine
            pMLine = voMLine
            pBGMWav = voBGMWav

            SelBox = voSelBox
            PECursor = voPECursor
            PEHalf = voPEHalf
            PEDeltaMouseOver = voPEDeltaMouseOver
            PEMouseOver = voPEMouseOver
            PESel = voPESel
            PEBPM = voPEBPM
            PEBPMFont = voPEBPMFont
            MiddleDeltaRelease = xMiddleDeltaRelease

            kHeight = vKHeight
            kFont = vKFont
            kMFont = vKMFont
            kLabelVShift = vKLabelVShift
            kLabelHShift = vKLabelHShift
            kLabelHShiftL = vKLabelHShiftL
            kMouseOver = vKMouseOver
            kMouseOverE = vKMouseOverE
            kSelected = vKSelected
            kOpacity = vKOpacity
        End Sub
    End Class

End Namespace
