Public NotInheritable Class SplashScreen1

    'TODO: 可轻松将此窗体设置为应用程序的初始屏幕，方法是转到
    '  “项目设计器”的“应用程序”选项卡(“项目”菜单下的“属性”)。

    'Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '根据应用程序的程序集信息在运行时设置对话框文本。  

    'TODO: 在项目属性对话框(“项目”菜单下)中的“应用程序”面板
    '  中自定义应用程序的程序集信息。 

    '应用程序标题
    'If My.Application.Info.Title <> "" Then
    '    ApplicationTitle.Text = My.Application.Info.Title
    'Else
    '    '若应用程序标题丢失，则使用不带扩展名的应用程序名
    '    ApplicationTitle.Text = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
    'End If

    '使用在设计时作为格式字符串设置到 Version 控件中的文本格式化版本信息。
    '  以便根据需要进行有效的本地化。
    '  使用以下代码，将Version 控件的设计时文本 
    '  更改为“Version {0}.{1:00}.{2}.{3}”或类似格式，将内部版本和修订信息包括在内。
    '  有关更多信息，请参阅帮助中的 String.Format()。
    '
    '    Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)

    'Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor)

    '版权信息
    'Copyright.Text = My.Application.Info.Copyright

    'End Sub

    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function ReleaseCapture Lib "user32.dll" Alias "ReleaseCapture" () As Integer
    Private Const WM_SYSCOMMAND = &H112
    Private Const SC_MOVE = &HF010&
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const HTCAPTION = 2
    Public bBitmap As Bitmap = My.Resources.About0

    Private Sub Me_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        Me.Close()
    End Sub

    Private Sub Me_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        If e.Button = MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Me.Handle.ToInt32, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0)
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            Me.Close()
        End If
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
End Class
