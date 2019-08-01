Public NotInheritable Class SplashScreen1
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        ' MyBase.OnPaint(e)
    End Sub

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        ' MyBase.OnPaintBackground(e)
        Dim rect As New Rectangle(0, 0, Width, Height)
        'e.Graphics.DrawImage(My.Resources.About0, rect)
    End Sub

    Private Sub SplashScreen1_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
    End Sub
End Class
