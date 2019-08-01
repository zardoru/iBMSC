

Public Class fLoadFileProgress
    Dim ReadOnly xPath(- 1) As String
    Dim CancelPressed As Boolean = False
    Dim ReadOnly IsSaved As Boolean = False

    Public Sub New(xxPath() As String, xIsSaved As Boolean, Optional ByVal TopMost As Boolean = True)
        InitializeComponent()
        prog.Maximum = UBound(xxPath) + 1
        xPath = xxPath
        IsSaved = xIsSaved
        Me.TopMost = TopMost
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = DialogResult.Cancel
        CancelPressed = True
        Me.Close()
    End Sub

    Private Sub fLoadFileProgress_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        On Error GoTo 0
        For i = 0 To UBound(xPath)
            Label1.Text = "Currently loading ( " & (i + 1) & " / " & (UBound(xPath) + 1) & " ): " & xPath(i)
            Dim aa As Integer = prog.Maximum
            Dim bb As Integer = prog.Value
            prog.Value = i
            Application.DoEvents()
            If CancelPressed Then Exit For

            If i = 0 AndAlso IsSaved Then MainWindow.ReadFile(xPath(i)) _
                Else Process.Start(Application.ExecutablePath, """" & xPath(i) & """") _
            'Shell("""" & Application.ExecutablePath & """ """ & xPaths(i) & """") ' 
        Next
        Me.Close()
    End Sub

    Private Sub fLoadFileProgress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Font = MainWindow.Font
        Me.Cancel_Button.Text = Strings.Cancel
    End Sub
End Class
