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

        ' Me.Text = Strings.fStatistics.Title
        ' LRow1.Text = Strings.fStatistics.lSCROLL
        ' LRow2.Text = Strings.fStatistics.lBPM
        ' LRow3.Text = Strings.fStatistics.lSTOP
        ' LRow4.Text = Strings.fStatistics.lA1
        ' LRow5.Text = Strings.fStatistics.lA2
        ' LRow6.Text = Strings.fStatistics.lA3
        ' LRow7.Text = Strings.fStatistics.lA4
        ' LRow8.Text = Strings.fStatistics.lA5
        ' LRow9.Text = Strings.fStatistics.lA6
        ' LRow10.Text = Strings.fStatistics.lA7
        ' LRow11.Text = Strings.fStatistics.lA8
        ' LRow12.Text = Strings.fStatistics.lD1
        ' LRow13.Text = Strings.fStatistics.lD2
        ' LRow14.Text = Strings.fStatistics.lD3
        ' LRow15.Text = Strings.fStatistics.lD4
        ' LRow16.Text = Strings.fStatistics.lD5
        ' LRow17.Text = Strings.fStatistics.lD6
        ' LRow18.Text = Strings.fStatistics.lD7
        ' LRow19.Text = Strings.fStatistics.lD8
        ' LRow20.Text = Strings.fStatistics.lA
        ' LRow21.Text = Strings.fStatistics.lD
        ' LRow22.Text = Strings.fStatistics.lBGM
        ' LRow23.Text = Strings.fStatistics.lNotes
        ' LRowTotal.Text = Strings.fStatistics.lTotal
        ' LColumn1.Text = Strings.fStatistics.lShort
        ' LColumn2.Text = Strings.fStatistics.lLong
        ' LColumn3.Text = Strings.fStatistics.lLnObj
        ' LColumn4.Text = Strings.fStatistics.lHidden
        ' LColumn5.Text = Strings.fStatistics.lLandmines
        ' LColumn6.Text = Strings.fStatistics.lErrors
        ' LColumnTotal.Text = Strings.fStatistics.lTotal
        OK_Button.Text = Strings.OK
        ' StTotalText.Text = Strings.fStatistics.recTotal
        ' StTotalValue.Text = Strings.fStatistics.recTotalValue
    End Sub

    Public Sub New(ByVal data(,) As Integer, ByVal LRows() As String, ByVal LCols() As String)
        InitializeComponent()

        For row As Integer = 1 To TableLayoutPanel1.RowCount - 1
            Dim xLabel As New Label
            xLabel.Dock = DockStyle.Fill
            xLabel.TextAlign = ContentAlignment.MiddleRight
            xLabel.Margin = New Padding(0)
            xLabel.Text = LRows(row - 1)
            If row Mod 2 = 1 Then xLabel.BackColor = Color.FromArgb(&H10000000)
            TableLayoutPanel1.Controls.Add(xLabel, 0, row)
        Next
        For col As Integer = 1 To TableLayoutPanel1.ColumnCount - 1
            Dim xLabel As New Label
            xLabel.Dock = DockStyle.Fill
            xLabel.TextAlign = ContentAlignment.MiddleCenter
            xLabel.Margin = New Padding(0)
            xLabel.Text = LCols(col - 1)
            TableLayoutPanel1.Controls.Add(xLabel, col, 0)
        Next

        For row As Integer = 1 To TableLayoutPanel1.RowCount - 1
            For col As Integer = 1 To TableLayoutPanel1.ColumnCount - 1
                Dim xLabel As New Label
                xLabel.Dock = DockStyle.Fill
                xLabel.TextAlign = ContentAlignment.MiddleCenter
                xLabel.Margin = New Padding(0)
                xLabel.Font = New Font(Me.Font, FontStyle.Bold)
                If data(row - 1, col - 1) <> 0 Then xLabel.Text = data(row - 1, col - 1)
                If row Mod 2 = 1 Then xLabel.BackColor = Color.FromArgb(&H10000000)
                TableLayoutPanel1.Controls.Add(xLabel, col, row)
            Next
        Next
    End Sub

End Class
