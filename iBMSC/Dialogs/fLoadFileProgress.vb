Imports System.Windows.Forms

Public Class fLoadFileProgress
    Dim xPath(-1) As String
    Dim CancelPressed As Boolean = False
    Dim IsSaved As Boolean = False

    Public Sub New(ByVal xxPath() As String, ByVal xIsSaved As Boolean, Optional ByVal TopMost As Boolean = True)
        InitializeComponent()
        prog.Maximum = UBound(xxPath) + 1
        xPath = xxPath
        IsSaved = xIsSaved
        Me.TopMost = TopMost
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        CancelPressed = True
        Me.Close()
    End Sub

    Private Sub fLoadFileProgress_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        On Error GoTo 0
        For i As Integer = 0 To UBound(xPath)
            Label1.Text = "Currently loading ( " & (i + 1) & " / " & (UBound(xPath) + 1) & " ): " & xPath(i)
            Dim aa As Integer = prog.Maximum
            Dim bb As Integer = prog.Value
            prog.Value = i
            Application.DoEvents()
            If CancelPressed Then Exit For

            If i = 0 AndAlso IsSaved Then MainWindow.ReadFile(xPath(i)) _
                Else System.Diagnostics.Process.Start(Application.ExecutablePath, """" & xPath(i) & """") 'Shell("""" & Application.ExecutablePath & """ """ & xPaths(i) & """") ' 
        Next
        Me.Close()
    End Sub

    Private Sub fLoadFileProgress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Font = MainWindow.Font
        Me.Cancel_Button.Text = Strings.Cancel
    End Sub
End Class
