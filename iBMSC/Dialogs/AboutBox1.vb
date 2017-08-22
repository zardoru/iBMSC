Public NotInheritable Class AboutBox1
    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function ReleaseCapture Lib "user32.dll" Alias "ReleaseCapture" () As Integer
    Private Const WM_SYSCOMMAND As Integer = &H112
    Private Const SC_MOVE As Integer = &HF010
    Private Const WM_NCLBUTTONDOWN As Integer = &HA1
    Private Const HTCAPTION As Integer = 2
    Public bBitmap As Bitmap = My.Resources.SplashScreenx

    Private Sub AboutBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Me.Close()
    End Sub

    Private Sub AboutBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        ReleaseCapture()
        SendMessage(Me.Handle.ToInt32, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
    End Sub

    Private Sub AboutBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then Me.Close()
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
            APIHelp.UpdateLayeredWindow(Handle, screenDc, newLocation, newSize, memDc, sourceLocation, _
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

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            ' Add the layered extended style (WS_EX_LAYERED) to this window
            Dim createParams__1 As CreateParams = MyBase.CreateParams
            createParams__1.ExStyle = createParams__1.ExStyle Or APIHelp.WS_EX_LAYERED
            Return createParams__1
        End Get
    End Property

    Private Sub ClickToCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClickToCopy.Click
        Clipboard.Clear()
        Clipboard.SetText("higan314doaz9@qq.com")
        Beep()
    End Sub

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure Point
        Public x As Int32
        Public y As Int32

        Public Sub New(ByVal x As Int32, ByVal y As Int32)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure Size
        Public cx As Int32
        Public cy As Int32

        Public Sub New(ByVal cx As Int32, ByVal cy As Int32)
            Me.cx = cx
            Me.cy = cy
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Private Structure ARGB
        Public Blue As Byte
        Public Green As Byte
        Public Red As Byte
        Public Alpha As Byte
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Public Structure BLENDFUNCTION
        Public BlendOp As Byte
        Public BlendFlags As Byte
        Public SourceConstantAlpha As Byte
        Public AlphaFormat As Byte
    End Structure

    Public Declare Auto Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal hdcDst As IntPtr, ByRef pptDst As Point, ByRef psize As Size, ByVal hdcSrc As IntPtr, ByRef pprSrc As Point, _
    ByVal crKey As Int32, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Int32) As Bool

    Public Declare Auto Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As IntPtr) As IntPtr

    Public Declare Auto Function GetDC Lib "user32.dll" (ByVal hWnd As IntPtr) As IntPtr

    <DllImport("user32.dll", ExactSpelling:=True)> _
    Public Shared Function ReleaseDC(ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Integer
    End Function

    Public Declare Auto Function DeleteDC Lib "gdi32.dll" (ByVal hdc As IntPtr) As Bool

    <DllImport("gdi32.dll", ExactSpelling:=True)> _
    Public Shared Function SelectObject(ByVal hDC As IntPtr, ByVal hObject As IntPtr) As IntPtr
    End Function

    Public Declare Auto Function DeleteObject Lib "gdi32.dll" (ByVal hObject As IntPtr) As Bool
End Class
