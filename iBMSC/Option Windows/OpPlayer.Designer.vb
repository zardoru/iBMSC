<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OpPlayer
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.TPlayB = New System.Windows.Forms.TextBox
        Me.TPlay = New System.Windows.Forms.TextBox
        Me.TStop = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TPath = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.BBrowse = New System.Windows.Forms.Button
        Me.BDefault = New System.Windows.Forms.Button
        Me.BRemove = New System.Windows.Forms.Button
        Me.BAdd = New System.Windows.Forms.Button
        Me.LPlayer = New System.Windows.Forms.ListBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OK_Button.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.OK_Button.Location = New System.Drawing.Point(72, 316)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(78, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Cancel_Button.Location = New System.Drawing.Point(156, 316)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(78, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(14, 153)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(157, 15)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Play from beginning"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(14, 182)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(157, 15)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Play from current measure"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(14, 211)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(157, 15)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Stop"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TPlayB
        '
        Me.TPlayB.Location = New System.Drawing.Point(177, 150)
        Me.TPlayB.Name = "TPlayB"
        Me.TPlayB.Size = New System.Drawing.Size(192, 23)
        Me.TPlayB.TabIndex = 4
        '
        'TPlay
        '
        Me.TPlay.Location = New System.Drawing.Point(177, 179)
        Me.TPlay.Name = "TPlay"
        Me.TPlay.Size = New System.Drawing.Size(192, 23)
        Me.TPlay.TabIndex = 5
        '
        'TStop
        '
        Me.TStop.Location = New System.Drawing.Point(177, 208)
        Me.TStop.Name = "TStop"
        Me.TStop.Size = New System.Drawing.Size(192, 23)
        Me.TStop.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(14, 244)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(225, 60)
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "References (case-sensitive):" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "<apppath> = Directory of the application" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "<measure>" & _
            " = Current measure" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "<filename> = File Name"
        '
        'TPath
        '
        Me.TPath.Location = New System.Drawing.Point(51, 98)
        Me.TPath.Name = "TPath"
        Me.TPath.Size = New System.Drawing.Size(288, 23)
        Me.TPath.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 101)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 15)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Path"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BBrowse
        '
        Me.BBrowse.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BBrowse.Location = New System.Drawing.Point(345, 98)
        Me.BBrowse.Name = "BBrowse"
        Me.BBrowse.Size = New System.Drawing.Size(24, 23)
        Me.BBrowse.TabIndex = 3
        Me.BBrowse.Text = "..."
        Me.BBrowse.UseVisualStyleBackColor = True
        '
        'BDefault
        '
        Me.BDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BDefault.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BDefault.Location = New System.Drawing.Point(240, 316)
        Me.BDefault.Name = "BDefault"
        Me.BDefault.Size = New System.Drawing.Size(130, 23)
        Me.BDefault.TabIndex = 62
        Me.BDefault.Text = "Restore default"
        Me.BDefault.UseVisualStyleBackColor = True
        '
        'BRemove
        '
        Me.BRemove.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BRemove.Location = New System.Drawing.Point(243, 41)
        Me.BRemove.Name = "BRemove"
        Me.BRemove.Size = New System.Drawing.Size(117, 23)
        Me.BRemove.TabIndex = 59
        Me.BRemove.Text = "Remove"
        Me.BRemove.UseVisualStyleBackColor = True
        '
        'BAdd
        '
        Me.BAdd.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BAdd.Location = New System.Drawing.Point(243, 12)
        Me.BAdd.Name = "BAdd"
        Me.BAdd.Size = New System.Drawing.Size(117, 23)
        Me.BAdd.TabIndex = 58
        Me.BAdd.Text = "Add"
        Me.BAdd.UseVisualStyleBackColor = True
        '
        'LPlayer
        '
        Me.LPlayer.FormattingEnabled = True
        Me.LPlayer.IntegralHeight = False
        Me.LPlayer.ItemHeight = 15
        Me.LPlayer.Items.AddRange(New Object() {"uBMplay.exe", "o2play.exe"})
        Me.LPlayer.Location = New System.Drawing.Point(12, 27)
        Me.LPlayer.Name = "LPlayer"
        Me.LPlayer.Size = New System.Drawing.Size(225, 55)
        Me.LPlayer.TabIndex = 63
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Location = New System.Drawing.Point(12, 134)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(357, 1)
        Me.PictureBox1.TabIndex = 64
        Me.PictureBox1.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(85, 15)
        Me.Label5.TabIndex = 65
        Me.Label5.Text = "Current Player:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(12, 126)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(66, 15)
        Me.Label7.TabIndex = 66
        Me.Label7.Text = "Arguments"
        '
        'OpPlayer
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(382, 351)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.LPlayer)
        Me.Controls.Add(Me.BDefault)
        Me.Controls.Add(Me.BRemove)
        Me.Controls.Add(Me.BAdd)
        Me.Controls.Add(Me.BBrowse)
        Me.Controls.Add(Me.TPath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TStop)
        Me.Controls.Add(Me.TPlay)
        Me.Controls.Add(Me.TPlayB)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "OpPlayer"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Player Options"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TPlayB As System.Windows.Forms.TextBox
    Friend WithEvents TPlay As System.Windows.Forms.TextBox
    Friend WithEvents TStop As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TPath As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BBrowse As System.Windows.Forms.Button
    Friend WithEvents BDefault As System.Windows.Forms.Button
    Friend WithEvents BRemove As System.Windows.Forms.Button
    Friend WithEvents BAdd As System.Windows.Forms.Button
    Friend WithEvents LPlayer As System.Windows.Forms.ListBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label

End Class
