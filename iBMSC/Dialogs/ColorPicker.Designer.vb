<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ColorPicker
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
        Me.PCMain = New System.Windows.Forms.Panel
        Me.tbPrecision = New System.Windows.Forms.TrackBar
        Me.PCA = New System.Windows.Forms.Panel
        Me.PC1 = New System.Windows.Forms.Panel
        Me.rbH = New System.Windows.Forms.RadioButton
        Me.rbS = New System.Windows.Forms.RadioButton
        Me.rbL = New System.Windows.Forms.RadioButton
        Me.rbR = New System.Windows.Forms.RadioButton
        Me.rbG = New System.Windows.Forms.RadioButton
        Me.rbB = New System.Windows.Forms.RadioButton
        Me.inH = New System.Windows.Forms.NumericUpDown
        Me.inS = New System.Windows.Forms.NumericUpDown
        Me.inL = New System.Windows.Forms.NumericUpDown
        Me.inR = New System.Windows.Forms.NumericUpDown
        Me.inG = New System.Windows.Forms.NumericUpDown
        Me.inB = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.tStr = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.inA = New System.Windows.Forms.NumericUpDown
        Me.rbA = New System.Windows.Forms.RadioButton
        Me.pPrev = New System.Windows.Forms.Panel
        CType(Me.tbPrecision, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.inH, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.inS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.inL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.inR, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.inG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.inB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.inA, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(315, 275)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(126, 27)
        Me.OK_Button.TabIndex = 21
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(315, 306)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(126, 27)
        Me.Cancel_Button.TabIndex = 22
        Me.Cancel_Button.Text = "Cancel"
        '
        'PCMain
        '
        Me.PCMain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PCMain.Location = New System.Drawing.Point(14, 14)
        Me.PCMain.Name = "PCMain"
        Me.PCMain.Size = New System.Drawing.Size(258, 258)
        Me.PCMain.TabIndex = 1
        '
        'tbPrecision
        '
        Me.tbPrecision.LargeChange = 2
        Me.tbPrecision.Location = New System.Drawing.Point(73, 303)
        Me.tbPrecision.Minimum = 1
        Me.tbPrecision.Name = "tbPrecision"
        Me.tbPrecision.Size = New System.Drawing.Size(199, 45)
        Me.tbPrecision.TabIndex = 2
        Me.tbPrecision.Value = 2
        '
        'PCA
        '
        Me.PCA.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PCA.Location = New System.Drawing.Point(14, 278)
        Me.PCA.Name = "PCA"
        Me.PCA.Size = New System.Drawing.Size(258, 19)
        Me.PCA.TabIndex = 3
        '
        'PC1
        '
        Me.PC1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PC1.Location = New System.Drawing.Point(278, 14)
        Me.PC1.Name = "PC1"
        Me.PC1.Size = New System.Drawing.Size(19, 258)
        Me.PC1.TabIndex = 4
        '
        'rbH
        '
        Me.rbH.Appearance = System.Windows.Forms.Appearance.Button
        Me.rbH.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbH.Location = New System.Drawing.Point(315, 55)
        Me.rbH.Name = "rbH"
        Me.rbH.Size = New System.Drawing.Size(77, 23)
        Me.rbH.TabIndex = 7
        Me.rbH.Text = "H (0-359)"
        Me.rbH.UseVisualStyleBackColor = True
        '
        'rbS
        '
        Me.rbS.Appearance = System.Windows.Forms.Appearance.Button
        Me.rbS.Checked = True
        Me.rbS.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbS.Location = New System.Drawing.Point(315, 80)
        Me.rbS.Name = "rbS"
        Me.rbS.Size = New System.Drawing.Size(77, 23)
        Me.rbS.TabIndex = 8
        Me.rbS.TabStop = True
        Me.rbS.Text = "S (0-1000)"
        Me.rbS.UseVisualStyleBackColor = True
        '
        'rbL
        '
        Me.rbL.Appearance = System.Windows.Forms.Appearance.Button
        Me.rbL.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbL.Location = New System.Drawing.Point(315, 105)
        Me.rbL.Name = "rbL"
        Me.rbL.Size = New System.Drawing.Size(77, 23)
        Me.rbL.TabIndex = 9
        Me.rbL.Text = "L (0-1000)"
        Me.rbL.UseVisualStyleBackColor = True
        '
        'rbR
        '
        Me.rbR.Appearance = System.Windows.Forms.Appearance.Button
        Me.rbR.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbR.Location = New System.Drawing.Point(315, 134)
        Me.rbR.Name = "rbR"
        Me.rbR.Size = New System.Drawing.Size(77, 23)
        Me.rbR.TabIndex = 10
        Me.rbR.Text = "R (0-255)"
        Me.rbR.UseVisualStyleBackColor = True
        '
        'rbG
        '
        Me.rbG.Appearance = System.Windows.Forms.Appearance.Button
        Me.rbG.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbG.Location = New System.Drawing.Point(315, 159)
        Me.rbG.Name = "rbG"
        Me.rbG.Size = New System.Drawing.Size(77, 23)
        Me.rbG.TabIndex = 11
        Me.rbG.Text = "G (0-255)"
        Me.rbG.UseVisualStyleBackColor = True
        '
        'rbB
        '
        Me.rbB.Appearance = System.Windows.Forms.Appearance.Button
        Me.rbB.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbB.Location = New System.Drawing.Point(315, 184)
        Me.rbB.Name = "rbB"
        Me.rbB.Size = New System.Drawing.Size(77, 23)
        Me.rbB.TabIndex = 12
        Me.rbB.Text = "B (0-255)"
        Me.rbB.UseVisualStyleBackColor = True
        '
        'inH
        '
        Me.inH.Location = New System.Drawing.Point(396, 55)
        Me.inH.Maximum = New Decimal(New Integer() {360, 0, 0, 0})
        Me.inH.Name = "inH"
        Me.inH.Size = New System.Drawing.Size(45, 23)
        Me.inH.TabIndex = 14
        '
        'inS
        '
        Me.inS.Location = New System.Drawing.Point(396, 80)
        Me.inS.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.inS.Name = "inS"
        Me.inS.Size = New System.Drawing.Size(45, 23)
        Me.inS.TabIndex = 15
        '
        'inL
        '
        Me.inL.Location = New System.Drawing.Point(396, 105)
        Me.inL.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.inL.Name = "inL"
        Me.inL.Size = New System.Drawing.Size(45, 23)
        Me.inL.TabIndex = 16
        '
        'inR
        '
        Me.inR.Location = New System.Drawing.Point(396, 134)
        Me.inR.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.inR.Name = "inR"
        Me.inR.Size = New System.Drawing.Size(45, 23)
        Me.inR.TabIndex = 17
        '
        'inG
        '
        Me.inG.Location = New System.Drawing.Point(396, 159)
        Me.inG.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.inG.Name = "inG"
        Me.inG.Size = New System.Drawing.Size(45, 23)
        Me.inG.TabIndex = 18
        '
        'inB
        '
        Me.inB.Location = New System.Drawing.Point(396, 184)
        Me.inB.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.inB.Name = "inB"
        Me.inB.Size = New System.Drawing.Size(45, 23)
        Me.inB.TabIndex = 19
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 305)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(58, 15)
        Me.Label1.TabIndex = 99
        Me.Label1.Text = "Precision:"
        '
        'tStr
        '
        Me.tStr.Location = New System.Drawing.Point(340, 245)
        Me.tStr.Name = "tStr"
        Me.tStr.Size = New System.Drawing.Size(89, 23)
        Me.tStr.TabIndex = 0
        Me.tStr.Text = "FF000000"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(320, 248)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(14, 15)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "#"
        '
        'inA
        '
        Me.inA.Location = New System.Drawing.Point(396, 213)
        Me.inA.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.inA.Name = "inA"
        Me.inA.Size = New System.Drawing.Size(45, 23)
        Me.inA.TabIndex = 20
        Me.inA.Value = New Decimal(New Integer() {255, 0, 0, 0})
        '
        'rbA
        '
        Me.rbA.Appearance = System.Windows.Forms.Appearance.Button
        Me.rbA.Enabled = False
        Me.rbA.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.rbA.Location = New System.Drawing.Point(315, 213)
        Me.rbA.Name = "rbA"
        Me.rbA.Size = New System.Drawing.Size(77, 23)
        Me.rbA.TabIndex = 13
        Me.rbA.Text = "A (0-255)"
        Me.rbA.UseVisualStyleBackColor = True
        '
        'pPrev
        '
        Me.pPrev.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pPrev.Location = New System.Drawing.Point(315, 14)
        Me.pPrev.Name = "pPrev"
        Me.pPrev.Size = New System.Drawing.Size(125, 30)
        Me.pPrev.TabIndex = 25
        '
        'ColorPicker
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(453, 345)
        Me.Controls.Add(Me.inA)
        Me.Controls.Add(Me.rbA)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tStr)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.inB)
        Me.Controls.Add(Me.inG)
        Me.Controls.Add(Me.inR)
        Me.Controls.Add(Me.inL)
        Me.Controls.Add(Me.inS)
        Me.Controls.Add(Me.inH)
        Me.Controls.Add(Me.rbB)
        Me.Controls.Add(Me.rbG)
        Me.Controls.Add(Me.rbR)
        Me.Controls.Add(Me.rbL)
        Me.Controls.Add(Me.rbS)
        Me.Controls.Add(Me.rbH)
        Me.Controls.Add(Me.PC1)
        Me.Controls.Add(Me.PCA)
        Me.Controls.Add(Me.tbPrecision)
        Me.Controls.Add(Me.PCMain)
        Me.Controls.Add(Me.pPrev)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ColorPicker"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ColorPicker"
        CType(Me.tbPrecision, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.inH, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.inS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.inL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.inR, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.inG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.inB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.inA, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents PCMain As System.Windows.Forms.Panel
    Friend WithEvents tbPrecision As System.Windows.Forms.TrackBar
    Friend WithEvents PCA As System.Windows.Forms.Panel
    Friend WithEvents PC1 As System.Windows.Forms.Panel
    Friend WithEvents rbH As System.Windows.Forms.RadioButton
    Friend WithEvents rbS As System.Windows.Forms.RadioButton
    Friend WithEvents rbL As System.Windows.Forms.RadioButton
    Friend WithEvents rbR As System.Windows.Forms.RadioButton
    Friend WithEvents rbG As System.Windows.Forms.RadioButton
    Friend WithEvents rbB As System.Windows.Forms.RadioButton
    Friend WithEvents inH As System.Windows.Forms.NumericUpDown
    Friend WithEvents inS As System.Windows.Forms.NumericUpDown
    Friend WithEvents inL As System.Windows.Forms.NumericUpDown
    Friend WithEvents inR As System.Windows.Forms.NumericUpDown
    Friend WithEvents inG As System.Windows.Forms.NumericUpDown
    Friend WithEvents inB As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tStr As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents inA As System.Windows.Forms.NumericUpDown
    Friend WithEvents rbA As System.Windows.Forms.RadioButton
    Friend WithEvents pPrev As System.Windows.Forms.Panel

End Class
