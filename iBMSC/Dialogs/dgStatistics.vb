Imports System.Windows.Forms

Public Class dgStatistics

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub dgStatistics_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Font = MainWindow.Font

        Me.Text = Strings.fStatistics.Title
        Label6.Text = Strings.fStatistics.lBPM
        Label7.Text = Strings.fStatistics.lSTOP
        Label8.Text = Strings.fStatistics.lSCROLL
        Label9.Text = Strings.fStatistics.lA
        Label10.Text = Strings.fStatistics.lD
        Label11.Text = Strings.fStatistics.lBGM
        Label1.Text = Strings.fStatistics.lTotal
        Label12.Text = Strings.fStatistics.lShort
        Label13.Text = Strings.fStatistics.lLong
        Label14.Text = Strings.fStatistics.lLnObj
        Label15.Text = Strings.fStatistics.lHidden
        Label16.Text = Strings.fStatistics.lErrors
        Label2.Text = Strings.fStatistics.lTotal
        OK_Button.Text = Strings.OK
    End Sub

    Public Sub New(ByVal data(,) As Integer)
        InitializeComponent()

        For row As Integer = 0 To 6
            For col As Integer = 0 To 5
                Dim xLabel As New Label
                xLabel.Dock = DockStyle.Fill
                xLabel.TextAlign = ContentAlignment.MiddleCenter
                xLabel.Margin = New Padding(0)
                xLabel.Font = New Font(Me.Font, FontStyle.Bold)
                If data(row, col) <> 0 Then xLabel.Text = data(row, col)
                If row Mod 2 = 0 Then xLabel.BackColor = Color.FromArgb(&H10000000)
                TableLayoutPanel1.Controls.Add(xLabel, col + 1, row + 1)
            Next
        Next
    End Sub
End Class
