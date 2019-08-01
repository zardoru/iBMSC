Public NotInheritable Class AboutBox1
    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA"(hwnd As Integer, wMsg As Integer,
                                                                              wParam As Integer, lParam As Integer) _
        As Integer
    Public Declare Function ReleaseCapture Lib "user32.dll" Alias "ReleaseCapture"() As Integer
    Private Const WM_SYSCOMMAND As Integer = &H112
    Private Const SC_MOVE As Integer = &HF010
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2
    Public bBitmap As Bitmap = My.Resources.ImageError

    Private Sub AboutBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        Me.Close()
    End Sub

    Private Sub AboutBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        ReleaseCapture()
        SendMessage(Me.Handle.ToInt32, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
    End Sub

    Private Sub AboutBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        If e.Button = MouseButtons.Right Then Me.Close()
    End Sub

    Public Sub SelectBitmap()
        ' Does this bitmap contain an alpha channel?
        If bBitmap.PixelFormat <> PixelFormat.Format32bppArgb Then
            Throw New ApplicationException("The bitmap must be 32bpp with alpha-channel.")
        End If

        ' Get device contexts
        Dim screenDc As IntPtr = APIHelp.GetDC(IntPtr.Zero)
        Dim memDc As IntPtr = APIHelp.CreateCompatibleDC(screenDc)
        Dim hBitmap As IntPtr = IntPtr.Zero
        Dim hOldBitmap As IntPtr = IntPtr.Zero

        Try
            ' Get handle to the new bitmap and select it into the current device context
            hBitmap = bBitmap.GetHbitmap(Color.FromArgb(0))
            hOldBitmap = APIHelp.SelectObject(memDc, hBitmap)

            ' Set parameters for layered window update
            Dim newSize As New APIHelp.Size(bBitmap.Width, bBitmap.Height)
            ' Size window to match bitmap
            Dim sourceLocation As New APIHelp.Point(0, 0)
            Dim newLocation As New APIHelp.Point(Me.Left, Me.Top)
            ' Same as this window
            Dim blend As New APIHelp.BLENDFUNCTION()
            blend.BlendOp = APIHelp.AC_SRC_OVER
            ' Only works with a 32bpp bitmap
            blend.BlendFlags = 0
            ' Always 0
            blend.SourceConstantAlpha = 255
            ' Set to 255 for per-pixel alpha values
            blend.AlphaFormat = APIHelp.AC_SRC_ALPHA
            ' Only works when the bitmap contains an alpha channel
            ' Update the window
            APIHelp.UpdateLayeredWindow(Handle, screenDc, newLocation, newSize, memDc, sourceLocation,
                                        0, blend, APIHelp.ULW_ALPHA)
        Finally
            ' Release device context
            APIHelp.ReleaseDC(IntPtr.Zero, screenDc)
            If hBitmap <> IntPtr.Zero Then
                APIHelp.SelectObject(memDc, hOldBitmap)
                ' Remove bitmap resources
                APIHelp.DeleteObject(hBitmap)
            End If
            APIHelp.DeleteDC(memDc)
        End Try
    End Sub

    ' Let Windows drag this window for us (thinks its hitting the title bar of the window)
    'Protected Overrides Sub WndProc(ByRef message As Message)
    '    If message.Msg = APIHelp.WM_NCHITTEST Then
    '        ' Tell Windows that the user is on the title bar (caption)
    '        message.Result = CType(APIHelp.HTCAPTION, IntPtr)
    '    Else
    '        MyBase.WndProc(message)
    '    End If
    'End Sub

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            ' Add the layered extended style (WS_EX_LAYERED) to this window
            Dim createParams__1 As CreateParams = MyBase.CreateParams
            createParams__1.ExStyle = createParams__1.ExStyle Or APIHelp.WS_EX_LAYERED
            Return createParams__1
        End Get
    End Property

    Private Sub ClickToCopy_Click(sender As Object, e As EventArgs) Handles ClickToCopy.Click
        Clipboard.Clear()
        Clipboard.SetText("higan314doaz9@qq.com")
        Beep()
    End Sub

    Private Sub AboutBox1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.SelectBitmap()
    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub
End Class

Class APIHelp
    ' Required constants
    Public Const WS_EX_LAYERED As Int32 = &H80000
    Public Const HTCAPTION As Int32 = &H2
    Public Const WM_NCHITTEST As Int32 = &H84
    Public Const ULW_ALPHA As Int32 = &H2
    Public Const AC_SRC_OVER As Byte = &H0
    Public Const AC_SRC_ALPHA As Byte = &H1

    Public Enum Bool
        [False] = 0
        [True] = 1
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure Point
        Public x As Int32
        Public y As Int32

        Public Sub New(x As Int32, y As Int32)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure Size
        Public cx As Int32
        Public cy As Int32

        Public Sub New(cx As Int32, cy As Int32)
            Me.cx = cx
            Me.cy = cy
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack := 1)>
    Private Structure ARGB
        Public ReadOnly Blue As Byte
        Public ReadOnly Green As Byte
        Public ReadOnly Red As Byte
        Public ReadOnly Alpha As Byte
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack := 1)>
    Public Structure BLENDFUNCTION
        Public BlendOp As Byte
        Public BlendFlags As Byte
        Public SourceConstantAlpha As Byte
        Public AlphaFormat As Byte
    End Structure

    Public Declare Auto Function UpdateLayeredWindow Lib "user32.dll"(hwnd As IntPtr, hdcDst As IntPtr,
                                                                      ByRef pptDst As Point, ByRef psize As Size,
                                                                      hdcSrc As IntPtr, ByRef pprSrc As Point,
                                                                      crKey As Int32, ByRef pblend As BLENDFUNCTION,
                                                                      dwFlags As Int32) As Bool

    Public Declare Auto Function CreateCompatibleDC Lib "gdi32.dll"(hDC As IntPtr) As IntPtr

    Public Declare Auto Function GetDC Lib "user32.dll"(hWnd As IntPtr) As IntPtr

    <DllImport("user32.dll", ExactSpelling := True)>
    Public Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
    End Function

    Public Declare Auto Function DeleteDC Lib "gdi32.dll"(hdc As IntPtr) As Bool

    <DllImport("gdi32.dll", ExactSpelling := True)>
    Public Shared Function SelectObject(hDC As IntPtr, hObject As IntPtr) As IntPtr
    End Function

    Public Declare Auto Function DeleteObject Lib "gdi32.dll"(hObject As IntPtr) As Bool
End Class
