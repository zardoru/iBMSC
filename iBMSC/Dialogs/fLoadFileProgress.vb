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
        For xI1 As Integer = 0 To UBound(xPath)
            Label1.Text = "Currently loading ( " & (xI1 + 1) & " / " & (UBound(xPath) + 1) & " ): " & xPath(xI1)
            Dim aa As Integer = prog.Maximum
            Dim bb As Integer = prog.Value
            prog.Value = xI1
            Application.DoEvents()
            If CancelPressed Then Exit For

            If xI1 = 0 AndAlso IsSaved Then MainWindow.ReadFile(xPath(xI1)) _
                Else System.Diagnostics.Process.Start(Application.ExecutablePath, """" & xPath(xI1) & """") 'Shell("""" & Application.ExecutablePath & """ """ & xPaths(xI1) & """") ' 
        Next
        Me.Close()
    End Sub

    Private Sub fLoadFileProgress_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Font = MainWindow.Font
        Me.Cancel_Button.Text = Strings.Cancel
    End Sub
End Class
